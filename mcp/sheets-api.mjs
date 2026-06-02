/**
 * The `sheets` global bound into the sheets_exec sandbox — one per exec.
 *
 * makeSheetsApi(spreadsheetId, client) is the per-exec SESSION. It owns:
 *   - a handle registry (byKey) so repeated sheets.sheet(name) calls reuse a
 *     handle, scoped to this exec (no process-global cache → concurrent execs
 *     don't clobber each other; issue #7 limitation #1).
 *   - the structural ops, which keep the registry coherent: rename updates and
 *     re-keys live handles, delete marks them dead (issue #7 limitation #2).
 *
 * Everything mutating goes through the injected `client`, whose batchUpdate /
 * valuesUpdate honor the dry-run capture. Reads pass through.
 */

import { createSheet, normHeader } from "./sheet.mjs";
import { a1ToGridRange, gridRangeToA1 } from "./a1.mjs";
import {
  buildAddSheet,
  buildDeleteSheet,
  buildRenameSheet,
  buildDuplicateSheet,
  buildAddTable,
  buildUpdateTable,
  buildDeleteTable,
  buildRestoreCells,
  compileTableColumns,
} from "./requests.mjs";

/** Leading sheet name of an A1 range ("Sheet!A1:I" → "Sheet", "'A b'!A1" → "A b"), else null. */
function sheetPrefixOf(a1) {
  if (typeof a1 !== "string") return null;
  if (a1[0] === "'") {
    let i = 1, name = "";
    while (i < a1.length) {
      if (a1[i] === "'") {
        if (a1[i + 1] === "'") { name += "'"; i += 2; continue; }
        break;
      }
      name += a1[i]; i++;
    }
    return (a1[i] === "'" && a1[i + 1] === "!") ? name : null;
  }
  const bang = a1.indexOf("!");
  return bang >= 0 ? a1.slice(0, bang) : null;
}

export function makeSheetsApi(spreadsheetId, client) {
  // Per-exec handle registry. Key: `${normHeader(name)}|${headerRow}`.
  const byKey = new Map();

  const keyFor = (name, headerRow) => `${normHeader(name)}|${headerRow}`;

  // After a column op on one handle, refresh sibling handles to the same tab
  // (opened with a different headerRow) so they don't keep stale positions.
  async function reinitSiblings(sheetId, except) {
    const jobs = [];
    for (const h of byKey.values()) {
      if (h !== except && h.sheetId === sheetId && !h._deleted) jobs.push(h._init());
    }
    await Promise.all(jobs);
  }

  async function getSheet(name, opts = {}) {
    const headerRow = opts.headerRow ?? 1;
    const key = keyFor(name, headerRow);
    if (byKey.has(key)) return byKey.get(key);
    const handle = await createSheet(spreadsheetId, name, opts, client);
    handle._onStructureChange = (h) => reinitSiblings(h.sheetId, h);
    byKey.set(key, handle);
    return handle;
  }

  /**
   * Resolve a sheet title OR numeric id to a sheetId.
   *   number → already an id (returned as-is, no round trip)
   *   string → matched by normHeader(title); throws on not-found / ambiguous.
   */
  async function resolveSheetId(nameOrId) {
    if (typeof nameOrId === "number") return nameOrId;
    const name = String(nameOrId);
    const meta = await client.getSpreadsheet(spreadsheetId);
    const target = normHeader(name);
    const matches = (meta.sheets || []).filter(s => normHeader(s.properties?.title) === target);
    if (matches.length === 0) {
      const known = (meta.sheets || []).map(s => s.properties?.title).filter(Boolean).join(", ");
      throw new Error(`Sheet "${name}" not found. Known: ${known}`);
    }
    if (matches.length > 1) {
      throw new Error(`Sheet name "${name}" is ambiguous — matches ${matches.length} tabs. Pass the numeric sheetId.`);
    }
    return matches[0].properties.sheetId;
  }

  // Resolve an A1 range carrying a sheet prefix ("Sheet!A1:I") → GridRange.
  async function resolveRangeToGrid(rangeA1) {
    const name = sheetPrefixOf(rangeA1);
    if (name == null) {
      throw new Error(`range "${rangeA1}" must include a sheet name, e.g. "Sheet!A1:I"`);
    }
    const sheetId = await resolveSheetId(name);
    return a1ToGridRange(sheetId, rangeA1); // a1ToGridRange strips the prefix itself
  }

  // Resolve a table name OR tableId string → { tableId, range, sheetTitle }.
  // Scans every sheet's tables; sheetTitle comes from the owning sheet's props
  // (matched by range.sheetId) so callers can build an A1 range for reads.
  async function resolveTable(nameOrId) {
    const meta = await client.getSpreadsheet(spreadsheetId);
    const sheets = meta.sheets || [];
    const tables = sheets.flatMap(s => s.tables || []);
    const t = tables.find(t => t.name === nameOrId) ?? tables.find(t => t.tableId === nameOrId);
    if (!t) {
      const known = tables.map(t => t.name).filter(Boolean).join(", ");
      throw new Error(`Table "${nameOrId}" not found. Known: ${known}`);
    }
    const owner = sheets.find(s => s.properties?.sheetId === t.range?.sheetId);
    return { tableId: t.tableId, range: t.range, sheetTitle: owner?.properties?.title };
  }

  // Thin wrapper for callers that only need the id.
  async function resolveTableId(nameOrId) {
    return (await resolveTable(nameOrId)).tableId;
  }

  // Negative sheetIds/"dryrun:" tableIds minted by the dry-run capture signal a
  // not-yet-real object — surface that as dryRun on the structural-op result.
  async function addSheet(title, opts = {}) {
    const res = await client.batchUpdate(spreadsheetId, buildAddSheet({ title, ...opts }));
    const props = res?.replies?.[0]?.addSheet?.properties;
    if (!props || props.sheetId == null) return { sheetId: null, title, dryRun: true };
    const out = { sheetId: props.sheetId, title: props.title, index: props.index };
    if (props.sheetId < 0) out.dryRun = true;
    return out;
  }

  async function addTable(tableName, rangeA1, opts = {}) {
    const range = await resolveRangeToGrid(rangeA1);
    const table = { name: tableName, range, columnProperties: compileTableColumns(opts.columns ?? []) };
    if (opts.tableId != null) table.tableId = opts.tableId;
    const res = await client.batchUpdate(spreadsheetId, buildAddTable(table));
    const t = res?.replies?.[0]?.addTable?.table;
    if (!t || t.tableId == null) return { tableId: null, name: tableName, dryRun: true };
    const out = { tableId: t.tableId, name: t.name };
    if (typeof t.tableId === "string" && t.tableId.startsWith("dryrun:")) out.dryRun = true;
    return out;
  }

  /**
   * Remove a native Table. Google's DeleteTableRequest ALSO clears the table's
   * cell data (header + all rows) — silent data loss (issue #11). So by default
   * we read the cells (value + format + note), delete the table, and restore
   * them in one atomic batchUpdate, mirroring the Sheets UI "Delete table"
   * (which keeps data). Pass { deleteData: true } for Google's native behavior
   * (clears the range too). Structured-reference formulas ([@[Col]]) can't
   * survive the table's removal — an inherent Sheets limitation.
   */
  async function deleteTable(nameOrId, opts = {}) {
    const { tableId, range, sheetTitle } = await resolveTable(nameOrId);

    if (opts.deleteData) {
      await client.batchUpdate(spreadsheetId, buildDeleteTable(tableId));
      return { ok: true, tableId, preserved: false };
    }

    // Read the table's current cells so we can restore them after the delete.
    let rows = [];
    if (range && sheetTitle) {
      const a1 = gridRangeToA1(sheetTitle, range);
      const data = await client.spreadsheetsGet(spreadsheetId, {
        ranges: [a1],
        includeGridData: true,
        fields: "sheets(data(rowData(values(userEnteredValue,userEnteredFormat,note))))",
      });
      rows = data?.sheets?.[0]?.data?.[0]?.rowData ?? [];
    }

    const requests = buildDeleteTable(tableId);
    if (rows.length > 0) {
      requests.push(...buildRestoreCells(range, rows));
    }
    await client.batchUpdate(spreadsheetId, requests);
    return { ok: true, tableId, preserved: true, rows: rows.length };
  }

  /**
   * Convert a native Table back to a styled PLAIN range WITHOUT data loss
   * (issue #12 B4). Reuses deleteTable's preserve path. Use this to escape the
   * "setDataValidation is not allowed on cells in typed columns" limit: after
   * untable the range is plain, so sheet.setValidation applies raw rules again.
   * Cell formatting survives; the Table's BANDING does not (re-style with
   * sheet.format / sheet.setBorders).
   */
  async function untable(nameOrId) {
    const { range } = await resolveTable(nameOrId);
    const r = await deleteTable(nameOrId); // default = preserve cells
    return { ok: true, tableId: r.tableId, range, rows: r.rows ?? 0, untabled: true };
  }

  return {
    // ── per-sheet handles ──
    sheet: getSheet,
    spreadsheetId: () => spreadsheetId,

    // ── raw escape hatch (all via the client → dry-run aware) ──
    batchUpdate: (requests, opts) => client.batchUpdate(spreadsheetId, requests, opts),
    valuesBatchGet: (ranges, opts) => client.valuesBatchGet(spreadsheetId, ranges, opts),
    // Default fetch returns sheet/dev-metadata; pass { fields } / { ranges } for
    // full control (named ranges, grid data) via the lower-level get.
    getSpreadsheet: (opts = {}) =>
      (opts.fields || opts.ranges)
        ? client.spreadsheetsGet(spreadsheetId, opts)
        : client.getSpreadsheet(spreadsheetId, opts),
    developerMetadataSearch: (filters) => client.developerMetadataSearch(spreadsheetId, filters),

    // ── structural ops ──
    addSheet,

    /** Create the tab only if a tab with this title doesn't already exist. */
    async ensureSheet(title, opts = {}) {
      const meta = await client.getSpreadsheet(spreadsheetId);
      const target = normHeader(title);
      const found = (meta.sheets || []).find(s => normHeader(s.properties?.title) === target);
      if (found) {
        return { sheetId: found.properties.sheetId, title: found.properties.title, existed: true };
      }
      return { ...(await addSheet(title, opts)), existed: false };
    },

    async deleteSheet(nameOrId) {
      const sheetId = await resolveSheetId(nameOrId);
      await client.batchUpdate(spreadsheetId, buildDeleteSheet(sheetId));
      for (const [k, h] of [...byKey]) {
        if (h.sheetId === sheetId) { h._markDeleted(); byKey.delete(k); }
      }
      return { ok: true, sheetId };
    },

    async renameSheet(nameOrId, newTitle) {
      const sheetId = await resolveSheetId(nameOrId);
      await client.batchUpdate(spreadsheetId, buildRenameSheet(sheetId, newTitle));
      // Update + re-key every cached handle for this tab so handles held in the
      // script keep working after the rename.
      for (const [k, h] of [...byKey]) {
        if (h.sheetId === sheetId) {
          byKey.delete(k);
          h.name = newTitle;
          byKey.set(keyFor(newTitle, h.headerRow), h);
        }
      }
      return { ok: true, sheetId, title: newTitle };
    },

    async duplicateSheet(nameOrId, newTitle) {
      const sourceSheetId = await resolveSheetId(nameOrId);
      const res = await client.batchUpdate(
        spreadsheetId,
        buildDuplicateSheet(sourceSheetId, newTitle != null ? { newTitle } : {}),
      );
      const props = res?.replies?.[0]?.duplicateSheet?.properties;
      if (!props || props.sheetId == null) return { sheetId: null, title: newTitle ?? null, dryRun: true };
      const out = { sheetId: props.sheetId, title: props.title };
      if (props.sheetId < 0) out.dryRun = true;
      return out;
    },

    // ── tables (native Sheets v4 Table) ──
    addTable,

    /** Create the table only if one with this name doesn't already exist. */
    async ensureTable(name, rangeA1, opts = {}) {
      const meta = await client.getSpreadsheet(spreadsheetId);
      const tables = (meta.sheets || []).flatMap(s => s.tables || []);
      const found = tables.find(t => t.name === name);
      if (found) return { tableId: found.tableId, name: found.name, existed: true };
      return { ...(await addTable(name, rangeA1, opts)), existed: false };
    },

    async updateTable(nameOrId, opts = {}) {
      const tableId = await resolveTableId(nameOrId);
      const table = { tableId };
      const fields = [];
      if (opts.name != null) { table.name = opts.name; fields.push("name"); }
      if (opts.columns != null) { table.columnProperties = compileTableColumns(opts.columns); fields.push("columnProperties"); }
      if (opts.range != null) { table.range = await resolveRangeToGrid(opts.range); fields.push("range"); }
      if (fields.length === 0) throw new Error("updateTable: provide name, columns, and/or range");
      await client.batchUpdate(spreadsheetId, buildUpdateTable(table, fields.join(",")));
      return { ok: true, tableId };
    },

    deleteTable,
    untable,
  };
}
