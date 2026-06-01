/**
 * Low-level Google Sheets calls. `makeClient` is the only export — it bundles
 * these into a per-exec, dry-run-aware client that the Sheet API consumes.
 *
 * Auth-scoped (one Google client per process); every call takes a
 * spreadsheetId so one server can serve many spreadsheets.
 */

import { google } from "googleapis";
import { loadAuth } from "./auth.mjs";

let cachedSheets = null;

function getSheets() {
  if (cachedSheets) return cachedSheets;
  const auth = loadAuth();
  cachedSheets = google.sheets({ version: "v4", auth });
  return cachedSheets;
}

/**
 * Get full spreadsheet metadata: sheet IDs, properties, dev metadata.
 */
async function getSpreadsheet(spreadsheetId, { includeGridData = false } = {}) {
  const sheets = getSheets();
  const res = await sheets.spreadsheets.get({
    spreadsheetId,
    includeGridData,
    fields:
      "spreadsheetId,sheets(properties(sheetId,title,gridProperties)," +
      "tables(tableId,name,range,columnProperties(columnIndex,columnName,columnType))," +
      "merges,developerMetadata),developerMetadata",
  });
  return res.data;
}

/**
 * Lower-level spreadsheets.get with full control over ranges + fields.
 * Used by Sheet._init to fetch sheet metadata + header/probe rows in one call.
 */
async function spreadsheetsGet(spreadsheetId, { ranges, includeGridData = false, fields } = {}) {
  const sheets = getSheets();
  const params = { spreadsheetId };
  if (ranges) params.ranges = ranges;
  if (includeGridData) params.includeGridData = includeGridData;
  if (fields) params.fields = fields;
  const res = await sheets.spreadsheets.get(params);
  return res.data;
}

/**
 * Send one or more batchUpdate requests atomically.
 * Requests array follows Google Sheets API v4 Request schema.
 *
 * Pure API call. Dry-run capture is handled one layer up by makeClient, so this
 * is never reached when previewing.
 */
async function batchUpdate(spreadsheetId, requests, { responseIncludeGridData = false } = {}) {
  if (!Array.isArray(requests) || requests.length === 0) {
    throw new Error("batchUpdate requires a non-empty requests array");
  }
  const sheets = getSheets();
  const res = await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests,
      responseIncludeGridData,
    },
  });
  return res.data;
}

/**
 * Read values for a range. Returns the raw values array (or empty array).
 * `valueRenderOption`: "FORMATTED_VALUE" (default), "UNFORMATTED_VALUE", "FORMULA".
 */
async function valuesGet(spreadsheetId, range, { valueRenderOption = "UNFORMATTED_VALUE" } = {}) {
  const sheets = getSheets();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range,
    valueRenderOption,
  });
  return res.data.values || [];
}

/**
 * Read multiple ranges in one round trip.
 */
async function valuesBatchGet(spreadsheetId, ranges, { valueRenderOption = "UNFORMATTED_VALUE" } = {}) {
  const sheets = getSheets();
  const res = await sheets.spreadsheets.values.batchGet({
    spreadsheetId,
    ranges,
    valueRenderOption,
  });
  return res.data.valueRanges || [];
}

/**
 * Search DeveloperMetadata by data filters.
 * `filters` is an array of { developerMetadataLookup: {...} } objects.
 */
async function developerMetadataSearch(spreadsheetId, filters) {
  const sheets = getSheets();
  const res = await sheets.spreadsheets.developerMetadata.search({
    spreadsheetId,
    requestBody: { dataFilters: filters },
  });
  return res.data.matchedDeveloperMetadata || [];
}

/**
 * Update a single value range. Used for ad-hoc cell writes that don't need
 * the full atomicity of batchUpdate. Pure API call (dry-run handled by makeClient).
 */
async function valuesUpdate(spreadsheetId, range, values, { raw = false } = {}) {
  const sheets = getSheets();
  const res = await sheets.spreadsheets.values.update({
    spreadsheetId,
    range,
    valueInputOption: raw ? "RAW" : "USER_ENTERED",
    requestBody: { values },
  });
  return res.data;
}

/**
 * Per-exec client: bundles every Google call, scoped to one dry-run `capture`.
 *
 * Mutations (batchUpdate, valuesUpdate) consult `capture`: when set, the
 * intended op is recorded as a tagged entry and a synthetic reply is returned
 * WITHOUT touching the API. Reads always pass through. Because the capture is
 * closed over per call to makeClient(), concurrent sheets_exec runs never see
 * each other's writes (issue #7 limitation #1).
 *
 * Tagged capture entries:
 *   { kind: "batchUpdate", spreadsheetId, requests }
 *   { kind: "valuesUpdate", spreadsheetId, range, values, valueInputOption }
 *
 * Under capture, addSheet/duplicateSheet/addTable replies carry a synthetic
 * PLACEHOLDER id (negative sheetId, "dryrun:N" tableId) so a dependent op later
 * in the same dry-run script can reference the not-yet-real object and still be
 * planned (issue #10 #8). Reads of that virtual object still can't be previewed.
 */
export function makeClient({ capture = null } = {}) {
  let placeholderSeq = 0;
  const syntheticReply = (r) => {
    if (r.addSheet) {
      return { addSheet: { properties: { ...(r.addSheet.properties || {}), sheetId: -(++placeholderSeq) } } };
    }
    if (r.duplicateSheet) {
      return { duplicateSheet: { properties: { sheetId: -(++placeholderSeq), title: r.duplicateSheet.newSheetName ?? null } } };
    }
    if (r.addTable) {
      return { addTable: { table: { ...(r.addTable.table || {}), tableId: `dryrun:${++placeholderSeq}` } } };
    }
    return {};
  };

  return {
    async batchUpdate(spreadsheetId, requests, opts = {}) {
      if (!Array.isArray(requests) || requests.length === 0) {
        throw new Error("batchUpdate requires a non-empty requests array");
      }
      if (capture) {
        capture.push({ kind: "batchUpdate", spreadsheetId, requests });
        return { spreadsheetId, replies: requests.map(syntheticReply) };
      }
      return batchUpdate(spreadsheetId, requests, opts);
    },

    async valuesUpdate(spreadsheetId, range, values, opts = {}) {
      const valueInputOption = opts.raw ? "RAW" : "USER_ENTERED";
      if (capture) {
        capture.push({ kind: "valuesUpdate", spreadsheetId, range, values, valueInputOption });
        const updatedColumns = values.reduce((m, r) => Math.max(m, Array.isArray(r) ? r.length : 0), 0);
        const updatedCells = values.reduce((s, r) => s + (Array.isArray(r) ? r.length : 0), 0);
        return { updatedRange: range, updatedRows: values.length, updatedColumns, updatedCells };
      }
      return valuesUpdate(spreadsheetId, range, values, opts);
    },

    // Reads — always pass through (dry-run previews writes, not reads).
    getSpreadsheet,
    spreadsheetsGet,
    valuesGet,
    valuesBatchGet,
    developerMetadataSearch,
  };
}
