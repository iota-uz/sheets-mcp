/**
 * Low-level Google Sheets client. Internal use only — Table API wraps this.
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
export async function getSpreadsheet(spreadsheetId, { includeGridData = false } = {}) {
  const sheets = getSheets();
  const res = await sheets.spreadsheets.get({
    spreadsheetId,
    includeGridData,
    fields: "spreadsheetId,sheets(properties(sheetId,title,gridProperties),developerMetadata),developerMetadata",
  });
  return res.data;
}

/**
 * Lower-level spreadsheets.get with full control over ranges + fields.
 * Used by Sheet._init to fetch sheet metadata + header/probe rows in one call.
 */
export async function spreadsheetsGet(spreadsheetId, { ranges, includeGridData = false, fields } = {}) {
  const sheets = getSheets();
  const params = { spreadsheetId };
  if (ranges) params.ranges = ranges;
  if (includeGridData) params.includeGridData = includeGridData;
  if (fields) params.fields = fields;
  const res = await sheets.spreadsheets.get(params);
  return res.data;
}

// Dry-run capture: when set, batchUpdate records the request body and returns a
// synthetic response instead of calling the API. Toggled by the runner.
let dryRunCapture = null;

export function setDryRunMode(capture) {
  dryRunCapture = capture; // null to disable; an array to capture into
}

/**
 * Send one or more batchUpdate requests atomically.
 * Requests array follows Google Sheets API v4 Request schema.
 */
export async function batchUpdate(spreadsheetId, requests, { responseIncludeGridData = false } = {}) {
  if (!Array.isArray(requests) || requests.length === 0) {
    throw new Error("batchUpdate requires a non-empty requests array");
  }

  if (dryRunCapture) {
    dryRunCapture.push({ spreadsheetId, requests });
    return { spreadsheetId, replies: requests.map(() => ({})) };
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
export async function valuesGet(spreadsheetId, range, { valueRenderOption = "UNFORMATTED_VALUE" } = {}) {
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
export async function valuesBatchGet(spreadsheetId, ranges, { valueRenderOption = "UNFORMATTED_VALUE" } = {}) {
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
export async function developerMetadataSearch(spreadsheetId, filters) {
  const sheets = getSheets();
  const res = await sheets.spreadsheets.developerMetadata.search({
    spreadsheetId,
    requestBody: { dataFilters: filters },
  });
  return res.data.matchedDeveloperMetadata || [];
}

/**
 * Update a single value range. Used for ad-hoc cell writes that don't need
 * the full atomicity of batchUpdate.
 */
export async function valuesUpdate(spreadsheetId, range, values, { raw = false } = {}) {
  const sheets = getSheets();
  const res = await sheets.spreadsheets.values.update({
    spreadsheetId,
    range,
    valueInputOption: raw ? "RAW" : "USER_ENTERED",
    requestBody: { values },
  });
  return res.data;
}
