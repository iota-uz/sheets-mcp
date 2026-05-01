/**
 * Generic Sheet API — wraps any Google Sheets tab without sheet-specific
 * knowledge. Records are keyed by header text (read from the header row at
 * runtime), formulas use {row} and {col:Header} placeholders that the runner
 * resolves at write time.
 *
 * No schema files. No "role" concept. No hardcoded column names anywhere.
 * Sheet-specific knowledge (categorization rules, formula templates,
 * column lists) lives in CLAUDE.md / SKILL.md alongside the business logic.
 *
 * Validation: pulled from in-sheet data validation rules (setDataValidation).
 * The sheet itself is the source of truth for enums. Agents call
 * sheet.setValidation(...) to configure these.
 */

import { compileStyle, colorToRgbFloat, colLetter } from "./schema.mjs";
import {
  getSpreadsheet,
  batchUpdate,
  valuesGet,
  developerMetadataSearch,
} from "./sheets-client.mjs";

const META_KEY_IDEMPOTENCY = "iota:idempotency";

const sheetCache = new Map();

/**
 * Get a Sheet handle. Cached per process; resolves the sheet metadata once
 * (sheetId, rowCount, headers) and rebinds on each script run via runner cache reset.
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
    this.headers = [];           // row[headerRow] from the sheet (string[])
    this.headerToIdx = new Map(); // exact match, lowercased, ё→е normalized
    this._validations = null;     // lazy: column header → validation rule
  }

  async _init() {
    const meta = await getSpreadsheet();
    const sheetMeta = meta.sheets.find(
      s => normHeader(s.properties.title) === normHeader(this.name)
    );
    if (!sheetMeta) {
      throw new Error(`Sheet "${this.name}" not found in spreadsheet`);
    }
    this.sheetId = sheetMeta.properties.sheetId;
    this.rowCount = sheetMeta.properties.gridProperties.rowCount;
    this.colCount = sheetMeta.properties.gridProperties.columnCount;

    const range = `${quoteSheetName(this.name)}!${this.headerRow}:${this.headerRow}`;
    const rows = await valuesGet(range, { valueRenderOption: "FORMATTED_VALUE" });
    this.headers = rows[0] || [];
    this.headers.forEach((h, idx) => {
      const k = normHeader(h);
      if (k) this.headerToIdx.set(k, idx);
    });
  }

  /**
   * Public metadata for sheets_describe.
   */
  describe() {
    return {
      sheet: this.name,
      sheetId: this.sheetId,
      headerRow: this.headerRow,
      rowCount: this.rowCount,
      headers: this.headers.map((h, idx) => ({
        index: idx,
        letter: colLetter(idx),
        text: h,
      })),
    };
  }

  /**
   * Resolve a header text to a 0-based column index. Throws if not found.
   * Uses normalized matching: trim + lowercase + ё→е, so "Счёт" and "Счет"
   * resolve to the same column.
   */
  _col(header) {
    const idx = this.headerToIdx.get(normHeader(header));
    if (idx == null) {
      const known = [...this.headerToIdx.keys()].join(", ");
      throw new Error(`Unknown header "${header}". Known: ${known}`);
    }
    return idx;
  }

  /**
   * Find the next empty row (1-based). Scans all known-header columns and
   * picks the max last-filled row across them. A single column might be
   * sparse (e.g., placeholder column A "Column 11" with no recent data),
   * so we look at all of them to find the true table end.
   */
  async _nextRow() {
    if (this.headerToIdx.size === 0) throw new Error(`Sheet "${this.name}" has no headers`);

    const indices = [...this.headerToIdx.values()];
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
    return this.headerRow + lastFilledOffset + 1;
  }

  /**
   * Lazy-load and cache data-validation rules on each column.
   * Returns Map<headerText, { type, values, strict }>.
   */
  async _loadValidations() {
    if (this._validations) return this._validations;
    const result = new Map();

    // Sample one row in the data area (first data row) for each column's validation.
    // dataValidation is typically uniform across a column when set via setDataValidation,
    // so checking one cell is enough. Header row is excluded.
    const probeRow = this.headerRow + 1;
    const meta = await fetchValidations(this.sheetId, probeRow);

    for (const [headerText, idx] of this.headerToIdx) {
      const cell = meta[idx];
      if (!cell?.dataValidation?.condition) continue;
      const cond = cell.dataValidation.condition;
      if (cond.type === "ONE_OF_LIST") {
        result.set(headerText, {
          type: "ONE_OF_LIST",
          values: (cond.values || []).map(v => v.userEnteredValue),
          strict: cell.dataValidation.strict ?? false,
        });
      }
    }

    this._validations = result;
    return result;
  }

  /**
   * Validate a record against in-sheet data validation rules (read from
   * spreadsheets.get). Throws if any field violates a strict rule.
   */
  async _validate(record) {
    const rules = await this._loadValidations();
    const errors = [];

    for (const [header, value] of Object.entries(record)) {
      // Resolve header (catches typos before we try to compile placeholders)
      this._col(header);

      // Skip formula values
      if (typeof value === "string" && value.startsWith("=")) continue;

      const rule = rules.get(normHeader(header)) ?? rules.get(header);
      if (!rule) continue;
      if (rule.type !== "ONE_OF_LIST") continue;
      if (!rule.values.includes(String(value))) {
        const msg = `"${header}" value "${value}" not in allowed list [${rule.values.join(", ")}]`;
        if (rule.strict) errors.push(msg);
      }
    }

    if (errors.length > 0) {
      const e = new Error(`Validation failed: ${errors.join("; ")}`);
      e.errors = errors;
      throw e;
    }
  }

  /**
   * Insert one record. Atomic: insertDimension + updateCells (with formulas
   * referencing the new row's actual position) + optional idempotency
   * DeveloperMetadata, all in one batchUpdate.
   *
   * record: { "Header text": value | "=formula" }
   *   formulas may use {row} and {col:Header} placeholders, resolved here.
   *
   * opts:
   *   idempotencyKey?: string — opaque agent-supplied dedup token
   *   format?: StyleObject   — applied to the new row
   *   dryRun?: boolean       — return planned requests without sending
   */
  async insert(record, opts = {}) {
    await this._validate(record);

    const idempotencyKey = opts.idempotencyKey ?? null;
    if (idempotencyKey) {
      const hit = await this._findByIdempotency(idempotencyKey);
      if (hit != null) return { row: hit, idempotencyHit: hit, inserted: false };
    }

    const lastRow = await this._nextRow();
    const targetRow = lastRow;

    const requests = this._buildInsertRequests({ targetRow, record, idempotencyKey, rowFormat: opts.format });

    if (opts.dryRun) {
      return { row: targetRow, dryRun: true, requests, inserted: false };
    }

    await batchUpdate(requests);
    return { row: targetRow, idempotencyHit: null, inserted: true };
  }

  _buildInsertRequests({ targetRow, record, idempotencyKey, rowFormat }) {
    const sheetId = this.sheetId;
    const rowIndex = targetRow - 1;
    const requests = [];

    // 1. Insert the new row
    requests.push({
      insertDimension: {
        range: { sheetId, dimension: "ROWS", startIndex: rowIndex, endIndex: rowIndex + 1 },
        inheritFromBefore: rowIndex > this.headerRow,
      },
    });

    // 2. Per-column cell values
    let maxColIdx = 0;
    const cellsByIdx = new Map();
    for (const [header, value] of Object.entries(record)) {
      const idx = this._col(header);
      if (idx > maxColIdx) maxColIdx = idx;
      cellsByIdx.set(idx, this._encodeValue(value, targetRow));
    }

    const values = new Array(maxColIdx + 1).fill(null).map((_, i) =>
      cellsByIdx.has(i) ? cellsByIdx.get(i) : { userEnteredValue: { stringValue: "" } }
    );

    requests.push({
      updateCells: {
        start: { sheetId, rowIndex, columnIndex: 0 },
        rows: [{ values }],
        fields: "userEnteredValue",
      },
    });

    // 3. Optional row format
    if (rowFormat) {
      const { cellFormat, fields } = compileStyle(rowFormat);
      requests.push({
        repeatCell: {
          range: {
            sheetId,
            startRowIndex: rowIndex,
            endRowIndex: rowIndex + 1,
            startColumnIndex: 0,
            endColumnIndex: maxColIdx + 1,
          },
          cell: cellFormat,
          fields,
        },
      });
    }

    // 4. Idempotency token via DeveloperMetadata on the row
    if (idempotencyKey) {
      requests.push({
        createDeveloperMetadata: {
          developerMetadata: {
            metadataKey: META_KEY_IDEMPOTENCY,
            metadataValue: `${this.name}:${idempotencyKey}`,
            location: {
              dimensionRange: {
                sheetId,
                dimension: "ROWS",
                startIndex: rowIndex,
                endIndex: rowIndex + 1,
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
    if (typeof value === "number") {
      return { userEnteredValue: { numberValue: value } };
    }
    if (typeof value === "boolean") {
      return { userEnteredValue: { boolValue: value } };
    }
    const s = String(value);
    if (s.startsWith("=")) {
      return { userEnteredValue: { formulaValue: this._compilePlaceholders(s, targetRow) } };
    }
    return { userEnteredValue: { stringValue: s } };
  }

  /**
   * Resolve {row} and {col:Header} placeholders in a formula string.
   */
  _compilePlaceholders(formula, targetRow) {
    return formula.replace(/\{(row|col:[^}]+)\}/g, (m, key) => {
      if (key === "row") return String(targetRow);
      if (key.startsWith("col:")) {
        const header = key.slice(4);
        const idx = this._col(header);
        return colLetter(idx);
      }
      return m;
    });
  }

  async _findByIdempotency(key) {
    const matches = await developerMetadataSearch([
      {
        developerMetadataLookup: {
          metadataKey: META_KEY_IDEMPOTENCY,
          metadataValue: `${this.name}:${key}`,
        },
      },
    ]);
    for (const m of matches) {
      const loc = m.developerMetadata?.location;
      if (loc?.dimensionRange?.sheetId === this.sheetId) {
        return loc.dimensionRange.startIndex + 1;
      }
    }
    return null;
  }

  /**
   * Find rows matching a where-clause (subset of headers, exact match).
   * Returns array of { row, "Header text": value, ... }.
   */
  async find(where = {}) {
    const range = `${quoteSheetName(this.name)}!A${this.headerRow + 1}:ZZ`;
    const rows = await valuesGet(range);

    const results = [];
    rows.forEach((row, i) => {
      const record = this._rowToRecord(row);
      let match = true;
      for (const [header, expected] of Object.entries(where)) {
        const idx = this._col(header);
        const actual = row[idx];
        if (actual !== expected && String(actual ?? "") !== String(expected ?? "")) {
          match = false;
          break;
        }
      }
      if (match) results.push({ row: this.headerRow + 1 + i, ...record });
    });
    return results;
  }

  _rowToRecord(row) {
    const r = {};
    for (const [headerKey, idx] of this.headerToIdx) {
      const original = this.headers[idx];
      r[original] = row[idx] ?? null;
    }
    return r;
  }

  /**
   * Update cells in rows matching the where clause.
   * `set` keys are header texts.
   */
  async update({ where, set }) {
    await this._validate(set);
    const matches = await this.find(where);
    if (matches.length === 0) return { updated: 0 };

    const sheetId = this.sheetId;
    const requests = [];

    for (const m of matches) {
      const rowIndex = m.row - 1;
      for (const [header, value] of Object.entries(set)) {
        const idx = this._col(header);
        const cell = this._encodeValue(value, m.row);
        requests.push({
          updateCells: {
            start: { sheetId, rowIndex, columnIndex: idx },
            rows: [{ values: [cell] }],
            fields: "userEnteredValue",
          },
        });
      }
    }

    if (requests.length === 0) return { updated: 0 };
    await batchUpdate(requests);
    return { updated: matches.length, rows: matches.map(m => m.row) };
  }

  /**
   * Delete rows matching the where clause (deleteDimension).
   */
  async delete({ where }) {
    const matches = await this.find(where);
    if (matches.length === 0) return { deleted: 0 };
    const sortedDesc = [...matches].sort((a, b) => b.row - a.row);
    const requests = sortedDesc.map(m => ({
      deleteDimension: {
        range: { sheetId: this.sheetId, dimension: "ROWS", startIndex: m.row - 1, endIndex: m.row },
      },
    }));
    await batchUpdate(requests);
    return { deleted: matches.length, rows: matches.map(m => m.row) };
  }

  /**
   * Apply formatting to rows matching the where clause.
   */
  async format({ where, set }) {
    const matches = await this.find(where);
    if (matches.length === 0) return { formatted: 0 };

    const { cellFormat, fields } = compileStyle(set);
    const maxColIdx = Math.max(...this.headerToIdx.values());
    const requests = matches.map(m => ({
      repeatCell: {
        range: {
          sheetId: this.sheetId,
          startRowIndex: m.row - 1,
          endRowIndex: m.row,
          startColumnIndex: 0,
          endColumnIndex: maxColIdx + 1,
        },
        cell: cellFormat,
        fields,
      },
    }));
    await batchUpdate(requests);
    return { formatted: matches.length, rows: matches.map(m => m.row) };
  }

  /**
   * Set data validation on a column. After this, future inserts/updates that
   * pass an invalid value will be rejected pre-flight.
   *
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
          startRowIndex: this.headerRow,  // exclude header row
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
    this._validations = null;  // bust cache
    return { ok: true };
  }
}

/**
 * Helper: read data validation rules for a single probe row.
 * Returns Array<CellData> aligned with column index.
 */
async function fetchValidations(sheetId, probeRow) {
  // Use spreadsheets.get with grid-data filtered to that row + dataValidation field.
  // Done via a low-level call so we don't pull all values.
  const { google } = await import("googleapis");
  const { loadAuth } = await import("./auth.mjs");
  const { loadEnv } = await import("./env.mjs");
  const env = loadEnv();
  const auth = loadAuth();
  const sheetsApi = google.sheets({ version: "v4", auth });
  const res = await sheetsApi.spreadsheets.get({
    spreadsheetId: env.SPREADSHEET_ID,
    fields: `sheets(properties(sheetId),data(startRow,startColumn,rowData(values(dataValidation))))`,
    ranges: [`${probeRow}:${probeRow}`],
  });
  const sheet = (res.data.sheets || []).find(s => s.properties?.sheetId === sheetId);
  if (!sheet) return [];
  const rowData = sheet.data?.[0]?.rowData?.[0]?.values || [];
  return rowData;
}

function normHeader(s) {
  return String(s ?? "").trim().toLowerCase().replace(/ё/g, "е");
}

function quoteSheetName(name) {
  if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) return name;
  return `'${name.replace(/'/g, "''")}'`;
}
