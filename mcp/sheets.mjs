/**
 * Google Sheets API wrapper — чистый generic слой.
 *
 * Только базовые операции: read, append, update, info.
 * Никакого знания о конкретной структуре IOTA-таблицы:
 * ни имён листов, ни позиций колонок, ни правил валидации.
 *
 * Любая логика "умной записи" (auto-fill, dup-check) живёт в helper-скриптах
 * (scripts/smart-append.mjs и др.), которые строят структуру динамически
 * по заголовкам первой строки.
 */

import { google } from "googleapis";
import { loadAuth } from "./auth.mjs";
import { loadEnv } from "./env.mjs";

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

/**
 * Read rows from a sheet.
 * @param {string} sheet    Sheet name
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
 * Append a single row to a sheet. Pure generic — no auto-fill, no dup-check.
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
 * @param {boolean} [raw]  If true, uses RAW input mode (no parsing — preserves
 *                         text "03.03.2026" as literal, useful when Sheets
 *                         auto-reformats dates to unwanted format like "3.3.2026").
 *                         Default false (USER_ENTERED — parses formulas/dates).
 */
export async function update(sheet, cell, value, raw = false) {
  const { sheets, id } = getSheetsClient();
  const res = await sheets.spreadsheets.values.update({
    spreadsheetId: id,
    range: `${sheet}!${cell}`,
    valueInputOption: raw ? "RAW" : "USER_ENTERED",
    requestBody: { values: [[value]] },
  });
  return { sheet, cell, value, updatedRange: res.data.updatedRange };
}

/**
 * Clear all values in a range (cells stay but become empty).
 */
export async function clear(sheet, range) {
  const { sheets, id } = getSheetsClient();
  const res = await sheets.spreadsheets.values.clear({
    spreadsheetId: id,
    range: `${sheet}!${range}`,
  });
  return { sheet, range: res.data.clearedRange };
}

/**
 * Get spreadsheet metadata: list of sheet names. Optional preview of first 3 rows.
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

/**
 * Utility: parse a number from Sheets display format ("$1 234,56", "748 879 434,82").
 */
export function toNumber(value) {
  if (value == null || value === "") return "";
  if (typeof value === "number") return value;
  const cleaned = String(value).replace(/[\s$]/g, "").replace(",", ".");
  const parsed = parseFloat(cleaned);
  return Number.isNaN(parsed) ? value : parsed;
}
