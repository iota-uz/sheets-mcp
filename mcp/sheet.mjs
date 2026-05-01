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

import { compileStyle, colLetter } from "./schema.mjs";
import {
  batchUpdate,
  spreadsheetsGet,
  developerMetadataSearch,
  valuesGet,
} from "./sheets-client.mjs";

const META_KEY_IDEMPOTENCY = "iota:idempotency";

const sheetCache = new Map();

/**
 * Get a Sheet handle. Cached per process.
 */
export async function sheet(name, opts = {}) {
  const cacheKey = `${name}|${opts.headerRow ?? 1}`;
  if (sheetCache.has(cacheKey)) return sheetCache.get(cacheKey);
  const s = new Sheet(name, opts);
  await s._init();
  sheetCache.set(cacheKey, s);
  return s;
}

export function clearCache() {
  sheetCache.clear();
}

class Sheet {
  constructor(name, { headerRow = 1 } = {}) {
    this.name = name;
    this.headerRow = headerRow;
    this.sheetId = null;
    this.rowCount = 0;
    this.colCount = 0;
    this.headers = [];
    this.headerToIdx = new Map();
    this._validations = new Map();   // headerText (normalized) → { type, values, strict }
    this._nextRowCache = null;        // 1-based; lazy
    this._idempotencyMap = null;      // Map<key, row>; lazy
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

    const meta = await spreadsheetsGet({ ranges, includeGridData: true, fields });

    const sheetMeta = (meta.sheets || []).find(
      s => normHeader(s.properties?.title) === normHeader(this.name)
    );
    if (!sheetMeta) throw new Error(`Sheet "${this.name}" not found in spreadsheet`);

    this.sheetId = sheetMeta.properties.sheetId;
    this.rowCount = sheetMeta.properties.gridProperties?.rowCount ?? 0;
    this.colCount = sheetMeta.properties.gridProperties?.columnCount ?? 0;

    const data = sheetMeta.data || [];
    // data[0] = header row, data[1] = probe row (matches request order)
    const headerCells = data[0]?.rowData?.[0]?.values || [];
    const headerStartCol = data[0]?.startColumn ?? 0;
    const probeCells = data[1]?.rowData?.[0]?.values || [];
    const probeStartCol = data[1]?.startColumn ?? 0;

    // Build headers array indexed by absolute column index
    const maxIdx = Math.max(
      headerStartCol + headerCells.length,
      probeStartCol + probeCells.length,
    );
    this.headers = new Array(maxIdx).fill("");
    for (let i = 0; i < headerCells.length; i++) {
      this.headers[headerStartCol + i] = headerCells[i]?.formattedValue ?? "";
    }
    this.headers.forEach((h, idx) => {
      const k = normHeader(h);
      if (k) this.headerToIdx.set(k, idx);
    });

    // Build validations indexed by header text
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
    const idx = this.headerToIdx.get(normHeader(header));
    if (idx == null) {
      const known = this.headers.filter(Boolean).join(", ");
      throw new Error(`Unknown header "${header}". Known: ${known}`);
    }
    return idx;
  }

  /**
   * Get the next available row number (1-based). Lazily initialized from
   * a single valuesGet, then incremented locally on each successful write.
   */
  async _getNextRow() {
    if (this._nextRowCache != null) return this._nextRowCache;

    const indices = [...this.headerToIdx.values()];
    if (indices.length === 0) throw new Error(`Sheet "${this.name}" has no headers`);
    const minIdx = Math.min(...indices);
    const maxIdx = Math.max(...indices);
    const range = `${quoteSheetName(this.name)}!${colLetter(minIdx)}${this.headerRow + 1}:${colLetter(maxIdx)}`;
    const rows = await valuesGet(range);

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

  /**
   * Lazy-load the full idempotency token map for this sheet. One round trip,
   * results stay cached for the lifetime of this Sheet handle.
   */
  async _loadIdempotency() {
    if (this._idempotencyMap) return this._idempotencyMap;
    const matches = await developerMetadataSearch([
      {
        developerMetadataLookup: {
          metadataKey: META_KEY_IDEMPOTENCY,
        },
      },
    ]);
    const map = new Map();
    const prefix = `${this.name}:`;
    for (const m of matches) {
      const md = m.developerMetadata;
      if (!md?.metadataValue?.startsWith(prefix)) continue;
      const loc = md.location?.dimensionRange;
      if (loc?.sheetId !== this.sheetId) continue;
      const key = md.metadataValue.slice(prefix.length);
      map.set(key, loc.startIndex + 1);
    }
    this._idempotencyMap = map;
    return map;
  }

  _validateRecord(record, errors) {
    for (const [header, value] of Object.entries(record)) {
      this._col(header);  // throws on unknown header
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

  /**
   * Insert N records atomically. ONE batchUpdate covers all inserts and
   * idempotency tokens. ~3 round trips total regardless of N (init done
   * upstream; idempotency + nextRow + batchUpdate).
   *
   * records: Array<{ "Header": value | "=formula" }>
   * opts:
   *   idempotencyKey?: (record, index) => string   per-record key generator
   *   format?: StyleObject                          applied to each new row
   *   dryRun?: boolean
   */
  async insertMany(records, opts = {}) {
    if (!Array.isArray(records)) throw new Error("insertMany expects an array");
    if (records.length === 0) {
      return { inserted: [], skipped: [], rows: [] };
    }

    // 1. Validate every record locally
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

    // 2. Bulk idempotency check
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

    // 3. Compute target rows
    const startRow = await this._getNextRow();
    const rowFormat = opts.format;

    const requests = this._buildBatchInsertRequests({ startRow, toInsert, rowFormat });

    if (opts.dryRun) {
      return {
        inserted: toInsert.map((t, i) => ({ index: t.index, row: startRow + i, key: t.key })),
        skipped,
        rows: toInsert.map((_, i) => startRow + i),
        dryRun: true,
        requests,
      };
    }

    await batchUpdate(requests);

    // 4. Update local caches
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

  /**
   * Single-record insert. Sugar over insertMany.
   */
  async insert(record, opts = {}) {
    const manyOpts = {
      ...(opts.format && { format: opts.format }),
      ...(opts.dryRun && { dryRun: opts.dryRun }),
    };
    if (opts.idempotencyKey != null) {
      manyOpts.idempotencyKey = () => String(opts.idempotencyKey);
    }
    const result = await this.insertMany([record], manyOpts);
    if (result.skipped.length > 0) {
      return { row: result.skipped[0].existingRow, idempotencyHit: result.skipped[0].existingRow, inserted: false };
    }
    return { row: result.inserted[0].row, idempotencyHit: null, inserted: true, ...(opts.dryRun && { dryRun: true, requests: result.requests }) };
  }

  _buildBatchInsertRequests({ startRow, toInsert, rowFormat }) {
    const sheetId = this.sheetId;
    const startIndex = startRow - 1;
    const N = toInsert.length;
    const requests = [];

    // 1. Insert N empty rows starting at startIndex
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

    // 2. Compute max column across all records
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

    // Build rows with cell values, encoding formulas with the actual target row
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

    // 3. Optional row format (applied across all new rows in one repeatCell)
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

    // 4. Idempotency tokens for non-duplicate records that have a key
    for (let i = 0; i < toInsert.length; i++) {
      const key = toInsert[i].key;
      if (!key) continue;
      requests.push({
        createDeveloperMetadata: {
          developerMetadata: {
            metadataKey: META_KEY_IDEMPOTENCY,
            metadataValue: `${this.name}:${key}`,
            location: {
              dimensionRange: {
                sheetId,
                dimension: "ROWS",
                startIndex: startIndex + i,
                endIndex: startIndex + i + 1,
              },
            },
            visibility: "PROJECT",
          },
        },
      });
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

  /**
   * Find rows matching a where-clause. Reads the data area once.
   *
   * where: { "Header": value, "Header2": value }
   */
  async find(where = {}) {
    const range = `${quoteSheetName(this.name)}!A${this.headerRow + 1}:ZZ`;
    const rows = await valuesGet(range);
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

  /**
   * Resolve { rows?, where? } to an array of 1-based row numbers.
   * `rows` short-circuits find() — useful when the agent already knows the row.
   */
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

  /**
   * Update cells in selected rows.
   *
   * { where, set } | { rows, set }
   */
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
    await batchUpdate(requests);
    return { updated: targetRows.length, rows: targetRows };
  }

  /**
   * Delete selected rows (deleteDimension).
   *
   * { where } | { rows }
   */
  async delete({ where, rows }) {
    const targetRows = await this._resolveRows({ rows, where });
    if (targetRows.length === 0) return { deleted: 0, rows: [] };

    const sortedDesc = [...targetRows].sort((a, b) => b - a);
    const requests = sortedDesc.map(r => ({
      deleteDimension: {
        range: { sheetId: this.sheetId, dimension: "ROWS", startIndex: r - 1, endIndex: r },
      },
    }));
    await batchUpdate(requests);

    // Invalidate row-position-dependent caches
    this._nextRowCache = null;
    if (this._idempotencyMap) this._idempotencyMap = null;

    return { deleted: targetRows.length, rows: targetRows };
  }

  /**
   * Format selected rows.
   */
  async format({ where, rows, set }) {
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
    await batchUpdate(requests);
    return { formatted: targetRows.length, rows: targetRows };
  }

  /**
   * Set data validation on a column.
   * spec: { type: "ONE_OF_LIST", values: string[], strict?: boolean }
   */
  async setValidation(header, spec) {
    const idx = this._col(header);
    if (spec.type !== "ONE_OF_LIST") {
      throw new Error(`setValidation: unsupported type "${spec.type}" — only ONE_OF_LIST for now`);
    }
    const requests = [{
      setDataValidation: {
        range: {
          sheetId: this.sheetId,
          startRowIndex: this.headerRow,
          startColumnIndex: idx,
          endColumnIndex: idx + 1,
        },
        rule: {
          condition: {
            type: "ONE_OF_LIST",
            values: spec.values.map(v => ({ userEnteredValue: String(v) })),
          },
          showCustomUi: true,
          strict: spec.strict ?? false,
        },
      },
    }];
    await batchUpdate(requests);

    // Update local cache
    this._validations.set(normHeader(header), {
      type: "ONE_OF_LIST",
      values: spec.values.map(v => String(v)),
      strict: spec.strict ?? false,
    });
    return { ok: true };
  }
}

function normHeader(s) {
  return String(s ?? "").trim().toLowerCase().replace(/ё/g, "е");
}

function quoteSheetName(name) {
  if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) return name;
  return `'${name.replace(/'/g, "''")}'`;
}
