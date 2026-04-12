/**
 * Google Sheets API wrapper.
 *
 * Provides two layers:
 *   1. Generic operations (used by MCP tools): read, append, update, info
 *   2. IOTA-specific helpers (used by import-statement.mjs script)
 */

import { google } from "googleapis";
import { loadAuth } from "./auth.mjs";
import { loadEnv } from "./env.mjs";

// ───────────────────────────────────────────────────────────────────────────
// Core
// ───────────────────────────────────────────────────────────────────────────

function getSpreadsheetId() {
  const env = loadEnv();
  if (!env.SPREADSHEET_ID) throw new Error("SPREADSHEET_ID not set in .env");
  return env.SPREADSHEET_ID;
}

function getSheetsClient() {
  const auth = loadAuth();
  const id = getSpreadsheetId();
  const sheets = google.sheets({ version: "v4", auth });
  return { sheets, id };
}

// ───────────────────────────────────────────────────────────────────────────
// Generic operations (exposed via MCP)
// ───────────────────────────────────────────────────────────────────────────

/**
 * Read rows from a sheet.
 * @param {string} sheet    Sheet name, e.g. "расходы"
 * @param {string} [range]  Optional A1 range, e.g. "B1:K100". Defaults to "A1:Z".
 * @param {number} [lastN]  If set, return only the last N rows.
 */
export async function read(sheet, range, lastN) {
  const { sheets, id } = getSheetsClient();
  const fullRange = range ? `${sheet}!${range}` : `${sheet}!A1:Z`;
  const res = await sheets.spreadsheets.values.get({ spreadsheetId: id, range: fullRange });
  const rows = res.data.values || [];

  if (lastN && lastN > 0) {
    const start = Math.max(0, rows.length - lastN);
    return {
      sheet,
      total: rows.length,
      rows: rows.slice(start).map((r, i) => ({ row: start + i + 1, values: r })),
    };
  }

  return { sheet, total: rows.length, rows };
}

/**
 * Append a row to a sheet.
 * @param {string} sheet    Sheet name
 * @param {string} range    Target range, e.g. "B:K" or "A:H"
 * @param {Array}  values   Array of cell values for one row
 */
export async function append(sheet, range, values) {
  const { sheets, id } = getSheetsClient();
  const res = await sheets.spreadsheets.values.append({
    spreadsheetId: id,
    range: `${sheet}!${range}`,
    valueInputOption: "USER_ENTERED",
    insertDataOption: "INSERT_ROWS",
    requestBody: { values: [values] },
  });
  return { sheet, range: res.data.updates.updatedRange, values };
}

/**
 * Update a single cell.
 * @param {string} sheet  Sheet name
 * @param {string} cell   Cell address, e.g. "H2229"
 * @param {*} value       New value
 */
export async function update(sheet, cell, value) {
  const { sheets, id } = getSheetsClient();
  const res = await sheets.spreadsheets.values.update({
    spreadsheetId: id,
    range: `${sheet}!${cell}`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [[value]] },
  });
  return { sheet, cell, value, updatedRange: res.data.updatedRange };
}

/**
 * Get spreadsheet metadata: list of sheet names + first 3 rows of a given sheet.
 * @param {string} [sheet]  If provided, include first 3 rows of this sheet.
 */
export async function info(sheet) {
  const { sheets, id } = getSheetsClient();
  const meta = await sheets.spreadsheets.get({ spreadsheetId: id });
  const sheetNames = meta.data.sheets.map((s) => s.properties.title);

  const result = { spreadsheetId: id, sheets: sheetNames };

  if (sheet) {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: id,
      range: `${sheet}!A1:Z3`,
    });
    result.preview = { sheet, rows: res.data.values || [] };
  }

  return result;
}

// ───────────────────────────────────────────────────────────────────────────
// IOTA-specific helpers (used by import-statement.mjs)
// ───────────────────────────────────────────────────────────────────────────

export function toNumber(value) {
  if (value == null || value === "") return "";
  if (typeof value === "number") return value;
  const cleaned = String(value).replace(/\s/g, "").replace(",", ".");
  const parsed = parseFloat(cleaned);
  return Number.isNaN(parsed) ? value : parsed;
}

export function periodFromDate(date) {
  if (!date) return "";
  const parts = String(date).split(".");
  if (parts.length !== 3) return date;
  return `01.${parts[1]}.${parts[2]}`;
}

export function calcUsd(amount, rate) {
  if (amount == null || amount === "" || !rate) return "";
  return Math.round((Number(amount) / Number(rate)) * 100) / 100;
}

/** Append a row to расходы (columns B..K). */
export async function appendRashod(data) {
  const amount = toNumber(data.amount);
  const rate = data.currency === "USD" ? 1 : toNumber(data.rate);
  const usd = data.amount_usd != null ? toNumber(data.amount_usd) : calcUsd(amount, rate);
  const period = data.period || periodFromDate(data.date);

  return append("расходы", "B:K", [
    data.date || "", data.category || "", amount, rate || "",
    data.currency || "", usd, period, data.account || "",
    data.employee || "", data.description || "",
  ]);
}

/** Append a row to платежи (columns A..K). */
export async function appendPlatezh(data) {
  const amount = toNumber(data.amount);
  const rate = data.currency === "USD" ? 1 : toNumber(data.rate);
  const usd = data.amount_usd != null ? toNumber(data.amount_usd) : calcUsd(amount, rate);
  const period = data.period || periodFromDate(data.date);

  return append("платежи", "A:K", [
    data.iteration || "", data.project || "", amount, data.category || "",
    rate || "", data.currency || "", usd, data.date || "",
    period, data.incoming_account || "", data.comment || "",
  ]);
}

/** Append a row to ддс (columns A..H). */
export async function appendDds(data) {
  const period = data.period || periodFromDate(data.date);
  return append("ддс", "A:H", [
    data.date || "", toNumber(data.amount),
    data.source_account || "UzCard (Kapital)",
    data.target_account || "Uzum Visa UZS",
    data.type || "P2P", period,
    data.rate != null ? toNumber(data.rate) : "", data.comment || "",
  ]);
}

/** Append a row to итерации (columns A..G). */
export async function appendIteration(data) {
  return append("итерации", "A:G", [
    data.project || "", data.iteration_id || "",
    toNumber(data.iteration_num), "",
    toNumber(data.amount), toNumber(data.amount),
    data.currency || "SOM",
  ]);
}

/** Get next iteration number for a project. */
export async function getNextIteration(project) {
  const result = await read("итерации", "A:C");
  let maxNum = 0;
  for (const row of result.rows) {
    const cells = Array.isArray(row) ? row : row.values || row;
    if ((cells[0] || "").toLowerCase() === project.toLowerCase()) {
      const num = parseInt(cells[2], 10);
      if (!isNaN(num) && num > maxNum) maxNum = num;
    }
  }
  return { lastNum: maxNum, nextNum: maxNum + 1, nextId: `${project}-${maxNum + 1}` };
}
