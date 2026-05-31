/**
 * Generic Sheet API — wraps any Google Sheets tab without sheet-specific
 * knowledge. Records are keyed by header text (read from the header row at
 * runtime), formulas use {row} and {col:Header} placeholders that the runner
 * resolves at write time.
 *
 * Optimization model:
 *   - _init does ONE spreadsheets.get covering sheet metadata, headers row,
 *     and probe row (for data validation rules).
 *   - Idempotency map is bulk-fetched once per sheet handle (one
 *     developerMetadata.search returning all tokens), looked up locally.
 *   - _nextRow is cached on the instance and incremented locally after each
 *     write. No round trip per insert.
 *   - insertMany batches N records into a single batchUpdate
 *     (insertDimension + updateCells with N rows + N idempotency tokens).
 *   - update/delete/format accept an explicit `rows: [...]` to skip find().
 */

import { compileStyle } from "./schema.mjs";
import { colLetter, colToIdx, a1ToGridRange } from "./a1.mjs";
import {
  META_KEY_IDEMPOTENCY,
  buildDeveloperMetadata,
  buildRepeatCellFormat,
  buildMerge,
  buildUnmerge,
  buildSetBorders,
  buildAddConditionalFormat,
  buildFreeze,
  buildResizeColumns,
  buildInsertColumns,
  buildDeleteColumns,
  buildSort,
  buildSetFilter,
  buildFindReplace,
  buildSetNote,
  buildSetValidation,
} from "./requests.mjs";

/**
 * Construct + initialize a Sheet handle bound to an injected `client` (which
 * carries the dry-run capture). No process-global cache — the session
 * (makeSheetsApi) owns per-exec handle caching. Returns a Proxy that throws on
 * any public method call once the tab has been deleted (issue #7 limitation #2).
 */
export async function createSheet(spreadsheetId, name, opts, client) {
  const s = new Sheet(spreadsheetId, name, opts, client);
  await s._init();
  return guardDeleted(s);
}

/** Wrap a Sheet so public method calls throw after the tab is deleted. */
function guardDeleted(sheet) {
  return new Proxy(sheet, {
    get(target, prop, receiver) {
      const value = Reflect.get(target, prop, receiver);
      if (
        target._deleted &&
        typeof value === "function" &&
        typeof prop === "string" &&
        !prop.startsWith("_")
      ) {
        return () => { throw new Error(`Sheet "${target.name}" was deleted`); };
      }
      return value;
    },
  });
}

class Sheet {
  constructor(spreadsheetId, name, { headerRow = 1 } = {}, client) {
    this.spreadsheetId = spreadsheetId;
    this.name = name;
    this.headerRow = headerRow;
    this.client = client;
    this.sheetId = null;
    this.rowCount = 0;
    this.colCount = 0;
    this.headers = [];
    this.headerToIdx = new Map();
    this._ambiguous = new Map();
    this._validations = new Map();
    this._nextRowCache = null;
    this._idempotencyMap = null;
    this._deleted = false;
    this._onStructureChange = null; // set by the session; (handle) => Promise
  }

  _markDeleted() {
    this._deleted = true;
  }

  /**
   * Single-call init: pulls sheet metadata + header row values + probe row
   * data validation in one spreadsheets.get round trip.
   */
  async _init() {
    const probeRow = this.headerRow + 1;
    const ranges = [
      `${quoteSheetName(this.name)}!${this.headerRow}:${this.headerRow}`,
      `${quoteSheetName(this.name)}!${probeRow}:${probeRow}`,
    ];
    const fields =
      "spreadsheetId," +
      "sheets(" +
        "properties(sheetId,title,gridProperties)," +
        "data(startRow,startColumn,rowData(values(formattedValue,dataValidation)))" +
      ")";

    const meta = await this.client.spreadsheetsGet(this.spreadsheetId, { ranges, includeGridData: true, fields });

    const sheetMeta = (meta.sheets || []).find(
      s => normHeader(s.properties?.title) === normHeader(this.name)
    );
    if (!sheetMeta) throw new Error(`Sheet "${this.name}" not found in spreadsheet`);

    this.sheetId = sheetMeta.properties.sheetId;
    this.rowCount = sheetMeta.properties.gridProperties?.rowCount ?? 0;
    this.colCount = sheetMeta.properties.gridProperties?.columnCount ?? 0;

    const data = sheetMeta.data || [];
    const headerCells = data[0]?.rowData?.[0]?.values || [];
    const headerStartCol = data[0]?.startColumn ?? 0;
    const probeCells = data[1]?.rowData?.[0]?.values || [];
    const probeStartCol = data[1]?.startColumn ?? 0;

    const maxIdx = Math.max(
      headerStartCol + headerCells.length,
      probeStartCol + probeCells.length,
    );
    this.headers = new Array(maxIdx).fill("");
    for (let i = 0; i < headerCells.length; i++) {
      this.headers[headerStartCol + i] = headerCells[i]?.formattedValue ?? "";
    }
    // indexHeaders is the single source of truth for header→column mapping and
    // collision detection; recomputed each _init() so a re-init (after a column
    // op) reflects the shifted positions.
    const { headerToIdx, ambiguous } = indexHeaders(this.headers);
    this.headerToIdx = headerToIdx;
    this._ambiguous = ambiguous;
    this._validations = new Map();

    for (let i = 0; i < probeCells.length; i++) {
      const colIdx = probeStartCol + i;
      const dv = probeCells[i]?.dataValidation;
      if (!dv?.condition) continue;
      const cond = dv.condition;
      if (cond.type !== "ONE_OF_LIST") continue;
      const headerText = this.headers[colIdx];
      if (!headerText) continue;
      this._validations.set(normHeader(headerText), {
        type: "ONE_OF_LIST",
        values: (cond.values || []).map(v => v.userEnteredValue),
        strict: dv.strict ?? false,
      });
    }
  }

  describe() {
    return {
      sheet: this.name,
      sheetId: this.sheetId,
      headerRow: this.headerRow,
      rowCount: this.rowCount,
      headers: this.headers
        .map((text, index) => ({ index, letter: colLetter(index), text }))
        .filter(h => h.text !== ""),
    };
  }

  _col(header) {
    const k = normHeader(header);
    if (this._ambiguous.has(k)) {
      const cols = this._ambiguous.get(k).map(colLetter).join(", ");
      throw new Error(`Ambiguous header "${header}" — appears in columns ${cols}. Make headers unique.`);
    }
    const idx = this.headerToIdx.get(k);
    if (idx == null) {
      const known = this.headers.filter(Boolean).join(", ");
      throw new Error(`Unknown header "${header}". Known: ${known}`);
    }
    return idx;
  }

  async _getNextRow() {
    if (this._nextRowCache != null) return this._nextRowCache;

    const indices = [...this.headerToIdx.values()];
    if (indices.length === 0) throw new Error(`Sheet "${this.name}" has no headers`);
    const minIdx = Math.min(...indices);
    const maxIdx = Math.max(...indices);
    const range = `${quoteSheetName(this.name)}!${colLetter(minIdx)}${this.headerRow + 1}:${colLetter(maxIdx)}`;
    const rows = await this.client.valuesGet(this.spreadsheetId, range);

    let lastFilledOffset = 0;
    for (let i = rows.length - 1; i >= 0; i--) {
      const row = rows[i] || [];
      const anyFilled = indices.some(idx => {
        const v = row[idx - minIdx];
        return v != null && String(v).trim() !== "";
      });
      if (anyFilled) {
        lastFilledOffset = i + 1;
        break;
      }
    }
    this._nextRowCache = this.headerRow + lastFilledOffset + 1;
    return this._nextRowCache;
  }

  async _loadIdempotency() {
    if (this._idempotencyMap) return this._idempotencyMap;
    const matches = await this.client.developerMetadataSearch(this.spreadsheetId, [
      {
        developerMetadataLookup: {
          metadataKey: META_KEY_IDEMPOTENCY,
        },
      },
    ]);
    // Tokens are scoped to this sheet by location.sheetId — NOT by a title prefix
    // — so they survive a tab rename (issue #7 limitation #3). The whole value is
    // the key.
    const map = new Map();
    for (const m of matches) {
      const md = m.developerMetadata;
      if (md?.metadataValue == null) continue;
      const loc = md.location?.dimensionRange;
      if (loc?.sheetId !== this.sheetId) continue;
      map.set(md.metadataValue, loc.startIndex + 1);
    }
    this._idempotencyMap = map;
    return map;
  }

  _validateRecord(record, errors) {
    for (const [header, value] of Object.entries(record)) {
      this._col(header);
      if (typeof value === "string" && value.startsWith("=")) continue;

      const rule = this._validations.get(normHeader(header));
      if (!rule || rule.type !== "ONE_OF_LIST") continue;
      if (!rule.values.includes(String(value))) {
        if (rule.strict) {
          errors.push(`"${header}" value "${value}" not in [${rule.values.join(", ")}]`);
        }
      }
    }
  }

  async insertMany(records, opts = {}) {
    if (!Array.isArray(records)) throw new Error("insertMany expects an array");
    if (records.length === 0) {
      return { inserted: [], skipped: [], rows: [] };
    }

    const errors = [];
    for (let i = 0; i < records.length; i++) {
      const rec = records[i];
      const before = errors.length;
      this._validateRecord(rec, errors);
      if (errors.length > before) errors[errors.length - 1] = `record[${i}]: ${errors[errors.length - 1]}`;
    }
    if (errors.length > 0) {
      const e = new Error(`Validation failed: ${errors.join("; ")}`);
      e.errors = errors;
      throw e;
    }

    const keyFn = opts.idempotencyKey;
    const keys = new Array(records.length).fill(null);
    let needIdemMap = false;
    if (typeof keyFn === "function") {
      for (let i = 0; i < records.length; i++) {
        const k = keyFn(records[i], i);
        keys[i] = k != null ? String(k) : null;
        if (keys[i] != null) needIdemMap = true;
      }
    }

    let idemMap = null;
    if (needIdemMap) idemMap = await this._loadIdempotency();

    const skipped = [];
    const toInsert = [];
    for (let i = 0; i < records.length; i++) {
      const key = keys[i];
      if (key && idemMap.has(key)) {
        skipped.push({ index: i, key, existingRow: idemMap.get(key) });
      } else {
        toInsert.push({ index: i, record: records[i], key });
      }
    }

    if (toInsert.length === 0) {
      return { inserted: [], skipped, rows: [] };
    }

    const startRow = await this._getNextRow();
    const rowFormat = opts.format;

    const requests = this._buildBatchInsertRequests({ startRow, toInsert, rowFormat });

    // Routes through the client: under dry-run the requests are captured and a
    // synthetic reply returned, so the bookkeeping below runs either way and two
    // inserts in one dry-run script plan distinct rows.
    await this.client.batchUpdate(this.spreadsheetId, requests);

    this._nextRowCache = startRow + toInsert.length;
    if (idemMap) {
      for (let i = 0; i < toInsert.length; i++) {
        if (toInsert[i].key) idemMap.set(toInsert[i].key, startRow + i);
      }
    }

    return {
      inserted: toInsert.map((t, i) => ({ index: t.index, row: startRow + i, key: t.key })),
      skipped,
      rows: toInsert.map((_, i) => startRow + i),
    };
  }

  async insert(record, opts = {}) {
    const manyOpts = {
      ...(opts.format && { format: opts.format }),
    };
    if (opts.idempotencyKey != null) {
      manyOpts.idempotencyKey = () => String(opts.idempotencyKey);
    }
    const result = await this.insertMany([record], manyOpts);
    if (result.skipped.length > 0) {
      return { row: result.skipped[0].existingRow, idempotencyHit: result.skipped[0].existingRow, inserted: false };
    }
    return { row: result.inserted[0].row, idempotencyHit: null, inserted: true };
  }

  _buildBatchInsertRequests({ startRow, toInsert, rowFormat }) {
    const sheetId = this.sheetId;
    const startIndex = startRow - 1;
    const N = toInsert.length;
    const requests = [];

    requests.push({
      insertDimension: {
        range: {
          sheetId,
          dimension: "ROWS",
          startIndex,
          endIndex: startIndex + N,
        },
        inheritFromBefore: startIndex > this.headerRow,
      },
    });

    let maxColIdx = 0;
    const perRecord = toInsert.map(({ record }) => {
      const cellsByIdx = new Map();
      for (const [header, value] of Object.entries(record)) {
        const idx = this._col(header);
        if (idx > maxColIdx) maxColIdx = idx;
        cellsByIdx.set(idx, value);
      }
      return cellsByIdx;
    });

    const rows = perRecord.map((cellsByIdx, i) => {
      const targetRow = startRow + i;
      const values = new Array(maxColIdx + 1).fill(null).map((_, colIdx) =>
        cellsByIdx.has(colIdx)
          ? this._encodeValue(cellsByIdx.get(colIdx), targetRow)
          : { userEnteredValue: { stringValue: "" } }
      );
      return { values };
    });

    requests.push({
      updateCells: {
        start: { sheetId, rowIndex: startIndex, columnIndex: 0 },
        rows,
        fields: "userEnteredValue",
      },
    });

    if (rowFormat) {
      const { cellFormat, fields } = compileStyle(rowFormat);
      requests.push({
        repeatCell: {
          range: {
            sheetId,
            startRowIndex: startIndex,
            endRowIndex: startIndex + N,
            startColumnIndex: 0,
            endColumnIndex: maxColIdx + 1,
          },
          cell: cellFormat,
          fields,
        },
      });
    }

    for (let i = 0; i < toInsert.length; i++) {
      const key = toInsert[i].key;
      if (!key) continue;
      requests.push(buildDeveloperMetadata(sheetId, key, startIndex + i));
    }

    return requests;
  }

  _encodeValue(value, targetRow) {
    if (value === null || value === undefined || value === "") {
      return { userEnteredValue: { stringValue: "" } };
    }
    if (typeof value === "number") return { userEnteredValue: { numberValue: value } };
    if (typeof value === "boolean") return { userEnteredValue: { boolValue: value } };
    const s = String(value);
    if (s.startsWith("=")) {
      return { userEnteredValue: { formulaValue: this._compilePlaceholders(s, targetRow) } };
    }
    return { userEnteredValue: { stringValue: s } };
  }

  _compilePlaceholders(formula, targetRow) {
    return formula.replace(/\{(row|col:[^}]+)\}/g, (m, key) => {
      if (key === "row") return String(targetRow);
      if (key.startsWith("col:")) {
        const header = key.slice(4);
        return colLetter(this._col(header));
      }
      return m;
    });
  }

  async find(where = {}) {
    const range = `${quoteSheetName(this.name)}!A${this.headerRow + 1}:ZZ`;
    const rows = await this.client.valuesGet(this.spreadsheetId, range);
    const filterEntries = Object.entries(where);

    const results = [];
    rows.forEach((row, i) => {
      let match = true;
      for (const [header, expected] of filterEntries) {
        const idx = this._col(header);
        const actual = row[idx];
        if (actual !== expected && String(actual ?? "") !== String(expected ?? "")) {
          match = false;
          break;
        }
      }
      if (match) results.push({ row: this.headerRow + 1 + i, ...this._rowToRecord(row) });
    });
    return results;
  }

  _rowToRecord(row) {
    const r = {};
    for (const [, idx] of this.headerToIdx) {
      const original = this.headers[idx];
      r[original] = row[idx] ?? null;
    }
    return r;
  }

  async _resolveRows({ rows, where }) {
    if (rows != null) {
      const arr = Array.isArray(rows) ? rows : [rows];
      return arr.map(r => Number(r));
    }
    if (where != null) {
      const matches = await this.find(where);
      return matches.map(m => m.row);
    }
    throw new Error("Provide either `rows` or `where`");
  }

  async update({ where, rows, set }) {
    const errors = [];
    this._validateRecord(set ?? {}, errors);
    if (errors.length > 0) throw new Error(`Validation failed: ${errors.join("; ")}`);

    const targetRows = await this._resolveRows({ rows, where });
    if (targetRows.length === 0) return { updated: 0, rows: [] };

    const requests = [];
    for (const row of targetRows) {
      const rowIndex = row - 1;
      for (const [header, value] of Object.entries(set)) {
        const idx = this._col(header);
        const cell = this._encodeValue(value, row);
        requests.push({
          updateCells: {
            start: { sheetId: this.sheetId, rowIndex, columnIndex: idx },
            rows: [{ values: [cell] }],
            fields: "userEnteredValue",
          },
        });
      }
    }

    if (requests.length === 0) return { updated: 0, rows: [] };
    await this.client.batchUpdate(this.spreadsheetId, requests);
    return { updated: targetRows.length, rows: targetRows };
  }

  async delete({ where, rows }) {
    const targetRows = await this._resolveRows({ rows, where });
    if (targetRows.length === 0) return { deleted: 0, rows: [] };

    const sortedDesc = [...targetRows].sort((a, b) => b - a);
    const requests = sortedDesc.map(r => ({
      deleteDimension: {
        range: { sheetId: this.sheetId, dimension: "ROWS", startIndex: r - 1, endIndex: r },
      },
    }));
    await this.client.batchUpdate(this.spreadsheetId, requests);

    this._nextRowCache = null;
    if (this._idempotencyMap) this._idempotencyMap = null;

    return { deleted: targetRows.length, rows: targetRows };
  }

  async format({ where, rows, range, set }) {
    // Arbitrary range / single column / off-header cells: format exactly that range.
    if (range != null) return this.formatRange(range, set);

    // Whole-row path (backward compatible): style every column of the target rows.
    const targetRows = await this._resolveRows({ rows, where });
    if (targetRows.length === 0) return { formatted: 0, rows: [] };

    const { cellFormat, fields } = compileStyle(set);
    const maxColIdx = Math.max(...this.headerToIdx.values());
    const requests = targetRows.map(r => ({
      repeatCell: {
        range: {
          sheetId: this.sheetId,
          startRowIndex: r - 1,
          endRowIndex: r,
          startColumnIndex: 0,
          endColumnIndex: maxColIdx + 1,
        },
        cell: cellFormat,
        fields,
      },
    }));
    await this.client.batchUpdate(this.spreadsheetId, requests);
    return { formatted: targetRows.length, rows: targetRows };
  }

  /** Format an arbitrary A1 range ("B2:D10", "C:C", "A1") — not limited to whole rows. */
  async formatRange(a1, style) {
    const { cellFormat, fields } = compileStyle(style);
    const grid = a1ToGridRange(this.sheetId, a1);
    await this.client.batchUpdate(this.spreadsheetId, buildRepeatCellFormat(grid, cellFormat, fields));
    return { ok: true, range: a1 };
  }

  // ── Presentation / structure sugar (all compile to batchUpdate ⇒ dry-run aware) ──

  async merge(range, type = "MERGE_ALL") {
    await this.client.batchUpdate(this.spreadsheetId, buildMerge(a1ToGridRange(this.sheetId, range), type));
    return { ok: true, range, mergeType: type };
  }

  async unmerge(range) {
    await this.client.batchUpdate(this.spreadsheetId, buildUnmerge(a1ToGridRange(this.sheetId, range)));
    return { ok: true, range };
  }

  async setBorders(range, spec = "all") {
    await this.client.batchUpdate(this.spreadsheetId, buildSetBorders(a1ToGridRange(this.sheetId, range), spec));
    return { ok: true, range };
  }

  async addConditionalFormat(range, rule) {
    await this.client.batchUpdate(this.spreadsheetId, buildAddConditionalFormat(a1ToGridRange(this.sheetId, range), rule));
    return { ok: true, range };
  }

  async freeze({ rows, cols } = {}) {
    await this.client.batchUpdate(this.spreadsheetId, buildFreeze(this.sheetId, { rows, cols }));
    return { ok: true, rows, cols };
  }

  /** range: a column range ("B:D" / "C:C"). opts: { width } | { auto: true }. */
  async resizeColumns(range, opts = {}) {
    const grid = a1ToGridRange(this.sheetId, range);
    if (grid.startColumnIndex == null) throw new Error(`resizeColumns: "${range}" must specify columns`);
    const endCol = grid.endColumnIndex ?? grid.startColumnIndex + 1;
    await this.client.batchUpdate(this.spreadsheetId, buildResizeColumns(this.sheetId, grid.startColumnIndex, endCol, opts));
    return { ok: true, range };
  }

  async insertColumns(at, count = 1) {
    const idx = this._colPos(at);
    await this.client.batchUpdate(this.spreadsheetId, buildInsertColumns(this.sheetId, idx, count));
    await this._invalidateStructure();
    return { ok: true, at: idx, count };
  }

  async deleteColumns(at, count = 1) {
    const idx = this._colPos(at);
    await this.client.batchUpdate(this.spreadsheetId, buildDeleteColumns(this.sheetId, idx, count));
    await this._invalidateStructure();
    return { ok: true, at: idx, count };
  }

  /** specs: [{ column: header|letter|index, order?: "ASC"|"DESC" }]. */
  async sort(range, specs) {
    const grid = a1ToGridRange(this.sheetId, range);
    const norm = (specs || []).map(s => ({
      dimensionIndex: s.dimensionIndex != null ? s.dimensionIndex : this._colPos(s.column),
      order: s.order,
    }));
    await this.client.batchUpdate(this.spreadsheetId, buildSort(grid, norm));
    return { ok: true, range };
  }

  async setFilter(range) {
    await this.client.batchUpdate(this.spreadsheetId, buildSetFilter(a1ToGridRange(this.sheetId, range)));
    return { ok: true, range };
  }

  /** Scoped to THIS sheet by default. opts: { range?, allSheets?, matchCase?, matchEntireCell?, searchByRegex?, includeFormulas? }. */
  async findReplace(find, replacement, opts = {}) {
    const frOpts = {
      find,
      replacement,
      matchCase: opts.matchCase,
      matchEntireCell: opts.matchEntireCell,
      searchByRegex: opts.searchByRegex,
      includeFormulas: opts.includeFormulas,
    };
    if (opts.allSheets) frOpts.allSheets = true;
    else if (opts.range) frOpts.range = a1ToGridRange(this.sheetId, opts.range);
    else frOpts.sheetId = this.sheetId;

    const res = await this.client.batchUpdate(this.spreadsheetId, buildFindReplace(frOpts));
    return res?.replies?.[0]?.findReplace ?? { ok: true };
  }

  async setNote(cell, text) {
    await this.client.batchUpdate(this.spreadsheetId, buildSetNote(a1ToGridRange(this.sheetId, cell), text));
    return { ok: true, cell };
  }

  /** Column position from a header name, a column letter, or a 0-based index. */
  _colPos(at) {
    if (typeof at === "number") return at;
    const k = normHeader(at);
    if (this._ambiguous.has(k)) {
      const cols = this._ambiguous.get(k).map(colLetter).join(", ");
      throw new Error(`Ambiguous header "${at}" — appears in columns ${cols}. Make headers unique.`);
    }
    if (this.headerToIdx.has(k)) return this.headerToIdx.get(k);
    return colToIdx(String(at));
  }

  /**
   * After a column insert/delete the cached header positions/width are stale.
   * Re-read this handle's metadata, then ask the session to refresh any other
   * cached handles pointing at the same tab (different headerRow) so they don't
   * keep stale column positions.
   */
  async _invalidateStructure() {
    this._nextRowCache = null;
    this._idempotencyMap = null;
    await this._init();
    await this._onStructureChange?.(this);
  }

  /**
   * Set or clear a column's data validation. Applies to the whole column below
   * the header row. spec.type is a Sheets ConditionType:
   *   ONE_OF_LIST { values, strict?, showCustomUi? }      (dropdown)
   *   ONE_OF_RANGE { range, strict? }                     (dropdown from a range)
   *   NUMBER_BETWEEN/NOT_BETWEEN { min, max }
   *   NUMBER_GREATER/LESS/EQ/… { value }
   *   DATE_BETWEEN { min, max } | DATE_AFTER/BEFORE/… { value } | DATE_IS_VALID {}
   *   TEXT_CONTAINS/STARTS_WITH/EQ/… { value } | TEXT_IS_EMAIL/URL {}
   *   BOOLEAN {}                                          (checkbox)
   *   CUSTOM_FORMULA { formula }
   *   "clear" | { clear: true }                           (remove validation)
   */
  async setValidation(header, spec) {
    const idx = this._col(header);
    const range = {
      sheetId: this.sheetId,
      startRowIndex: this.headerRow,
      startColumnIndex: idx,
      endColumnIndex: idx + 1,
    };
    await this.client.batchUpdate(this.spreadsheetId, buildSetValidation(range, spec));

    // Keep the in-memory validation cache (consulted by insert/update) in sync.
    // Only ONE_OF_LIST is enforced locally; other types and clears just evict.
    const k = normHeader(header);
    if (spec?.type === "clear" || spec?.clear === true) {
      this._validations.delete(k);
    } else if (spec.type === "ONE_OF_LIST") {
      this._validations.set(k, {
        type: "ONE_OF_LIST",
        values: spec.values.map(v => String(v)),
        strict: spec.strict ?? false,
      });
    } else {
      this._validations.delete(k);
    }
    return { ok: true };
  }

  /**
   * Richer read: pull formatting, notes, merges, validation and hyperlinks for
   * an arbitrary A1 range. Returns { data, merges } from the Sheets grid data.
   */
  async readFormatting(a1, { fields } = {}) {
    const range = `${quoteSheetName(this.name)}!${a1}`;
    const mask = fields ||
      "sheets(merges,data(startRow,startColumn,rowData(values(" +
      "formattedValue,userEnteredValue,effectiveFormat,note,dataValidation,hyperlink))))";
    const meta = await this.client.spreadsheetsGet(this.spreadsheetId, { ranges: [range], includeGridData: true, fields: mask });
    const sheet = (meta.sheets || [])[0] || {};
    return { data: sheet.data || [], merges: sheet.merges || [] };
  }

  async readRange(a1, { valueRender = "UNFORMATTED_VALUE" } = {}) {
    const range = `${quoteSheetName(this.name)}!${a1}`;
    return this.client.valuesGet(this.spreadsheetId, range, { valueRenderOption: valueRender });
  }

  async writeRange(a1, values, { raw = false } = {}) {
    if (!Array.isArray(values) || !Array.isArray(values[0])) {
      throw new Error("writeRange: values must be a 2D array");
    }
    const range = `${quoteSheetName(this.name)}!${a1}`;
    const res = await this.client.valuesUpdate(this.spreadsheetId, range, values, { raw });
    return {
      updatedRange: res.updatedRange,
      updatedRows: res.updatedRows ?? 0,
      updatedColumns: res.updatedColumns ?? 0,
      updatedCells: res.updatedCells ?? 0,
    };
  }
}

export function normHeader(s) {
  return String(s ?? "").trim().toLowerCase().replace(/ё/g, "е");
}

/**
 * Index a header row by normalized text. Headers that collide after
 * normalization (case-insensitive, ё→е) are NOT placed in headerToIdx — they go
 * to `ambiguous` (normKey → all colliding 0-based column indices) so a write
 * keyed by an ambiguous header fails loudly instead of hitting the wrong column
 * (issue #7 limitation #4). Empty/whitespace headers are ignored.
 */
export function indexHeaders(headers) {
  const seen = new Map(); // normKey → [idx, ...]
  headers.forEach((h, idx) => {
    const k = normHeader(h);
    if (!k) return;
    if (seen.has(k)) seen.get(k).push(idx);
    else seen.set(k, [idx]);
  });
  const headerToIdx = new Map();
  const ambiguous = new Map();
  for (const [k, idxs] of seen) {
    if (idxs.length > 1) ambiguous.set(k, idxs);
    else headerToIdx.set(k, idxs[0]);
  }
  return { headerToIdx, ambiguous };
}

function quoteSheetName(name) {
  if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) return name;
  return `'${name.replace(/'/g, "''")}'`;
}
