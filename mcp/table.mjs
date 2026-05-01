/**
 * Table API — typed schema-aware wrapper over Google Sheets.
 *
 * Each `Table` represents one sheet declared in schemas/<name>.json.
 * Operations compile to a single spreadsheets.batchUpdate when atomicity
 * matters (insert+formula), or to lighter values.* calls otherwise.
 *
 * Column→role mapping is resolved from DeveloperMetadata (written by
 * scripts/bootstrap-schema.mjs). Header text is a human hint; the binding
 * survives column reordering by the user.
 */

import {
  loadSchema,
  applyComputed,
  primaryKey as primaryKeyOf,
  validateRecord,
  compileFormula,
  compileStyle,
  colLetter,
  normHeader,
} from "./schema.mjs";
import {
  getSpreadsheet,
  batchUpdate,
  valuesGet,
  valuesBatchGet,
  developerMetadataSearch,
} from "./sheets-client.mjs";

const META_KEY_ROLE = "iota:col:role";
const META_KEY_IDEMPOTENCY = "iota:idempotency";

const tableCache = new Map();

/**
 * Get or build a Table for the named sheet. Cached per process — resolution
 * (metadata search + sheetId lookup) runs once.
 */
export async function table(sheetName) {
  if (tableCache.has(sheetName)) return tableCache.get(sheetName);
  const t = new Table(sheetName);
  await t._init();
  tableCache.set(sheetName, t);
  return t;
}

/**
 * Clear the cache. Useful after bootstrap or schema edits.
 */
export function clearCache() {
  tableCache.clear();
}

class Table {
  constructor(sheetName) {
    this.schema = loadSchema(sheetName);
    this.sheetName = this.schema.sheet;
    this.sheetId = null;
    this.gridRowCount = 0;
    this.gridColCount = 0;
    // role → 0-based column index, resolved from metadata or fuzzy fallback
    this.roleToCol = {};
    this.colToRole = {};
  }

  async _init() {
    const meta = await getSpreadsheet();
    const sheet = meta.sheets.find(
      s => normHeader(s.properties.title) === normHeader(this.sheetName)
    );
    if (!sheet) {
      throw new Error(`Sheet "${this.sheetName}" not found in spreadsheet`);
    }
    this.sheetId = sheet.properties.sheetId;
    this.gridRowCount = sheet.properties.gridProperties.rowCount;
    this.gridColCount = sheet.properties.gridProperties.columnCount;

    // Resolve column roles. Try DeveloperMetadata first (the canonical source);
    // fall back to header-row fuzzy match if no metadata exists yet (e.g. before
    // bootstrap was run).
    await this._resolveColumns();
  }

  async _resolveColumns() {
    // Try metadata search scoped to this sheet
    const matches = await developerMetadataSearch([
      {
        developerMetadataLookup: {
          metadataKey: META_KEY_ROLE,
          locationType: "COLUMN",
          metadataLocation: { sheetId: this.sheetId },
        },
      },
    ]);

    const fromMetadata = {};
    for (const m of matches) {
      if (!m.developerMetadata) continue;
      const role = m.developerMetadata.metadataValue;
      const loc = m.developerMetadata.location;
      if (!loc || !loc.dimensionRange) continue;
      if (loc.dimensionRange.sheetId !== this.sheetId) continue;
      const colIdx = loc.dimensionRange.startIndex;
      if (typeof colIdx === "number") fromMetadata[role] = colIdx;
    }

    if (Object.keys(fromMetadata).length > 0) {
      this.roleToCol = fromMetadata;
    } else {
      // Fallback: fuzzy match header row
      this.roleToCol = await this._fuzzyMatchHeaders();
    }

    for (const [role, idx] of Object.entries(this.roleToCol)) {
      this.colToRole[idx] = role;
    }
  }

  async _fuzzyMatchHeaders() {
    const headerRow = this.schema.headerRow;
    const range = `${quoteSheetName(this.sheetName)}!A${headerRow}:ZZ${headerRow}`;
    const rows = await valuesGet(range, { valueRenderOption: "FORMATTED_VALUE" });
    const headers = rows[0] || [];

    const headerToIdx = new Map();
    headers.forEach((h, i) => {
      const k = normHeader(h);
      if (k) headerToIdx.set(k, i);
    });

    const result = {};
    for (const [role, col] of Object.entries(this.schema.columns)) {
      const k = normHeader(col.header);
      const idx = headerToIdx.get(k);
      if (idx != null) result[role] = idx;
    }
    return result;
  }

  /**
   * Public schema description. Returned to MCP clients via sheets.describe.
   */
  describe() {
    return {
      sheet: this.sheetName,
      sheetId: this.sheetId,
      headerRow: this.schema.headerRow,
      primaryKey: this.schema.primaryKey,
      columns: Object.fromEntries(
        Object.entries(this.schema.columns).map(([role, col]) => [
          role,
          {
            header: col.header,
            type: col.type,
            optional: col.optional,
            values: col.values,
            columnIndex: this.roleToCol[role] ?? null,
            columnLetter: this.roleToCol[role] != null ? colLetter(this.roleToCol[role]) : null,
            ...(col.numberFormat && { numberFormat: col.numberFormat }),
            ...(col.formula && { formula: col.formula }),
          },
        ])
      ),
    };
  }

  /**
   * Find the next empty row in the data area (after header row).
   * Reads the date/primary column to find the last populated row, then +1.
   */
  async _nextRow() {
    // Use the first non-formula, non-computed column as the "presence indicator"
    const probeRole = Object.entries(this.schema.columns).find(
      ([, c]) => c.type !== "formula" && c.type !== "computed"
    )?.[0];
    if (!probeRole) throw new Error("Schema has no non-derived columns to probe");

    const colIdx = this.roleToCol[probeRole];
    if (colIdx == null) throw new Error(`Probe column "${probeRole}" not bound to a sheet column`);

    const letter = colLetter(colIdx);
    const range = `${quoteSheetName(this.sheetName)}!${letter}${this.schema.headerRow + 1}:${letter}`;
    const rows = await valuesGet(range);

    let lastFilledOffset = 0;
    for (let i = rows.length - 1; i >= 0; i--) {
      const v = rows[i] && rows[i][0];
      if (v != null && String(v).trim() !== "") {
        lastFilledOffset = i + 1;
        break;
      }
    }
    return this.schema.headerRow + lastFilledOffset; // 1-based row, the new row goes here as next
  }

  /**
   * Insert one record. Returns { row, idempotencyHit?: row }.
   *
   * Uses one batchUpdate with:
   *   1. insertDimension(row+1, inheritFromBefore=true) — adds a fresh row at end
   *   2. updateCells(row+1) — writes literals + formulas referencing that row
   *   3. createDeveloperMetadata(row+1) — idempotency token (if provided)
   *
   * If `idempotencyKey` is provided and a prior row carries the same token,
   * the insert is skipped and { row: existingRow, idempotencyHit: existingRow }
   * is returned. The check uses developerMetadata.search — deterministic, not
   * a heuristic match.
   */
  async insert(record, opts = {}) {
    const { idempotencyKey: providedKey, format: rowFormat } = opts;
    const idempotencyKey = providedKey ?? primaryKeyOf(this.schema, record);

    // Validate
    const { valid, errors } = validateRecord(this.schema, record);
    if (!valid) {
      const e = new Error(`insert validation failed: ${errors.join("; ")}`);
      e.errors = errors;
      throw e;
    }

    // Apply computed columns (period, etc.)
    applyComputed(this.schema, record);

    // Idempotency check
    if (idempotencyKey) {
      const hit = await this._findByIdempotency(idempotencyKey);
      if (hit != null) return { row: hit, idempotencyHit: hit, inserted: false };
    }

    const lastRow = await this._nextRow();
    const targetRow = lastRow + 1;

    const requests = this._buildInsertRequests({ targetRow, record, idempotencyKey, rowFormat });

    if (opts.dryRun) {
      return { row: targetRow, dryRun: true, requests, inserted: false };
    }

    await batchUpdate(requests);
    return { row: targetRow, idempotencyHit: null, inserted: true };
  }

  _buildInsertRequests({ targetRow, record, idempotencyKey, rowFormat }) {
    const sheetId = this.sheetId;
    const rowIndex = targetRow - 1; // 0-based for API

    const requests = [];

    // 1. Insert dimension (a new empty row at rowIndex)
    requests.push({
      insertDimension: {
        range: {
          sheetId,
          dimension: "ROWS",
          startIndex: rowIndex,
          endIndex: rowIndex + 1,
        },
        inheritFromBefore: rowIndex > this.schema.headerRow,
      },
    });

    // 2. Build per-column cell values
    const maxCol = Math.max(...Object.values(this.roleToCol));
    const cells = new Array(maxCol + 1).fill({ userEnteredValue: { stringValue: "" } });

    // role → column letter map for formula compilation
    const roleToLetter = Object.fromEntries(
      Object.entries(this.roleToCol).map(([r, i]) => [r, colLetter(i)])
    );

    for (const [role, col] of Object.entries(this.schema.columns)) {
      const colIdx = this.roleToCol[role];
      if (colIdx == null) continue; // role not bound — skip silently

      let cellValue;

      if (col.type === "formula") {
        const formula = compileFormula(col.formula, roleToLetter, targetRow);
        cellValue = { userEnteredValue: { formulaValue: formula } };
      } else if (record[role] === undefined || record[role] === null || record[role] === "") {
        cellValue = { userEnteredValue: { stringValue: "" } };
      } else if (col.type === "number") {
        const n = typeof record[role] === "number" ? record[role] : parseFloat(record[role]);
        cellValue = { userEnteredValue: { numberValue: n } };
      } else if (col.type === "date" || col.type === "computed") {
        // Dates & computed values written as strings; sheet's number-format
        // (set during bootstrap) handles display.
        cellValue = { userEnteredValue: { stringValue: String(record[role]) } };
      } else {
        cellValue = { userEnteredValue: { stringValue: String(record[role]) } };
      }

      cells[colIdx] = cellValue;
    }

    requests.push({
      updateCells: {
        start: { sheetId, rowIndex, columnIndex: 0 },
        rows: [{ values: cells }],
        fields: "userEnteredValue",
      },
    });

    // 3. Optional per-row format
    if (rowFormat) {
      const { cellFormat, fields } = compileStyle(rowFormat);
      requests.push({
        repeatCell: {
          range: {
            sheetId,
            startRowIndex: rowIndex,
            endRowIndex: rowIndex + 1,
            startColumnIndex: 0,
            endColumnIndex: maxCol + 1,
          },
          cell: cellFormat,
          fields,
        },
      });
    }

    // 4. Idempotency token (DeveloperMetadata on the row)
    if (idempotencyKey) {
      requests.push({
        createDeveloperMetadata: {
          developerMetadata: {
            metadataKey: META_KEY_IDEMPOTENCY,
            metadataValue: `${this.sheetName}:${idempotencyKey}`,
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

  async _findByIdempotency(key) {
    const matches = await developerMetadataSearch([
      {
        developerMetadataLookup: {
          metadataKey: META_KEY_IDEMPOTENCY,
          metadataValue: `${this.sheetName}:${key}`,
        },
      },
    ]);
    for (const m of matches) {
      const loc = m.developerMetadata?.location;
      if (loc?.dimensionRange?.sheetId === this.sheetId) {
        return loc.dimensionRange.startIndex + 1; // 1-based row
      }
    }
    return null;
  }

  /**
   * Find rows matching a where-clause (subset of fields, exact match).
   * Returns array of { row, ...record }.
   *
   * Currently does a full-sheet read + JS filter. Acceptable for sheets with
   * <10k rows; could be optimized later with developerMetadata-based indexes.
   */
  async find(where = {}) {
    const range = `${quoteSheetName(this.sheetName)}!A${this.schema.headerRow + 1}:ZZ`;
    const rows = await valuesGet(range);

    const results = [];
    rows.forEach((row, i) => {
      const record = this._rowToRecord(row);
      let match = true;
      for (const [k, v] of Object.entries(where)) {
        if (record[k] !== v && String(record[k]) !== String(v)) {
          match = false;
          break;
        }
      }
      if (match) results.push({ row: this.schema.headerRow + 1 + i, ...record });
    });

    return results;
  }

  _rowToRecord(row) {
    const record = {};
    for (const [role, idx] of Object.entries(this.roleToCol)) {
      record[role] = row[idx] ?? null;
    }
    return record;
  }

  /**
   * Update specific fields on rows matching the where clause.
   */
  async update({ where, set }) {
    const matches = await this.find(where);
    if (matches.length === 0) return { updated: 0 };

    const sheetId = this.sheetId;
    const requests = [];

    for (const m of matches) {
      const rowIndex = m.row - 1;
      for (const [role, value] of Object.entries(set)) {
        const colIdx = this.roleToCol[role];
        if (colIdx == null) continue;
        const col = this.schema.columns[role];

        let cellValue;
        if (col?.type === "number") {
          cellValue = { userEnteredValue: { numberValue: typeof value === "number" ? value : parseFloat(value) } };
        } else {
          cellValue = { userEnteredValue: { stringValue: String(value) } };
        }

        requests.push({
          updateCells: {
            start: { sheetId, rowIndex, columnIndex: colIdx },
            rows: [{ values: [cellValue] }],
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
   * Format cells in rows matching the where clause.
   */
  async format({ where, set }) {
    const matches = await this.find(where);
    if (matches.length === 0) return { formatted: 0 };

    const { cellFormat, fields } = compileStyle(set);
    const sheetId = this.sheetId;
    const maxCol = Math.max(...Object.values(this.roleToCol));

    const requests = matches.map(m => ({
      repeatCell: {
        range: {
          sheetId,
          startRowIndex: m.row - 1,
          endRowIndex: m.row,
          startColumnIndex: 0,
          endColumnIndex: maxCol + 1,
        },
        cell: cellFormat,
        fields,
      },
    }));

    await batchUpdate(requests);
    return { formatted: matches.length, rows: matches.map(m => m.row) };
  }
}

/**
 * Quote a sheet name for use in A1 notation. Adds single quotes if the name
 * contains anything besides ASCII letters/digits, including Cyrillic, spaces,
 * or punctuation. Embedded quotes are doubled.
 */
function quoteSheetName(name) {
  if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) return name;
  return `'${name.replace(/'/g, "''")}'`;
}
