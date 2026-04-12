/**
 * Bank statement (.xlsx) parser for IOTA expense imports.
 *
 * Reads a bank statement export, locates the data header row by looking
 * for the recognizable Russian column names, and returns the rows as
 * structured objects:
 *
 *   {
 *     row_num,        // sequential number (column "№ пп")
 *     doc_date,       // "ДД.ММ.ГГГГ"
 *     proc_date,      // "ДД.ММ.ГГГГ"
 *     doc_number,     // "№ док"
 *     account_name,   // counterparty name
 *     inn,
 *     account_number,
 *     mfo,
 *     debit,          // outgoing (number)
 *     credit,         // incoming (number)
 *     description,    // free-text "Назначение платежа"
 *   }
 *
 * Header detection is loose: any row whose joined text contains
 * "№ пп" + "Дата документа" + "Обороты по" is treated as the header.
 */

import xlsxModule from "xlsx";
import path from "path";
import { fileURLToPath } from "url";

const xlsx = xlsxModule.default || xlsxModule;

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = path.resolve(__dirname, "..", "..");

// ───────────────────────────────────────────────────────────────────────────
// Helpers
// ───────────────────────────────────────────────────────────────────────────

/**
 * Resolve a user-provided file path to an absolute path.
 * Relative paths are resolved against the project root (parent of build/).
 */
function resolvePath(filePath) {
  if (!filePath) throw new Error("file_path is required");
  return path.isAbsolute(filePath) ? filePath : path.resolve(PROJECT_ROOT, filePath);
}

/**
 * Normalize a header label: collapse whitespace, lowercase.
 */
function normalize(text) {
  return String(text ?? "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

/**
 * Locate the header row index by matching key column names.
 * Returns -1 if no header row is found.
 */
function findHeaderRowIndex(rows) {
  for (let i = 0; i < rows.length; i++) {
    const joined = normalize(rows[i].join(" | "));
    if (
      joined.includes("№ пп") &&
      joined.includes("дата документа") &&
      joined.includes("обороты по")
    ) {
      return i;
    }
  }
  return -1;
}

/**
 * Build a map of normalized column name → column index from a header row.
 */
function buildColumnMap(headerRow) {
  const map = {};
  headerRow.forEach((cell, index) => {
    const key = normalize(cell);
    if (key) map[key] = index;
  });
  return map;
}

/**
 * Convert any cell to a clean string (trim, replace newlines with spaces).
 */
function asString(value) {
  if (value == null) return "";
  return String(value).replace(/\s+/g, " ").trim();
}

/**
 * Convert any cell to a number. Handles strings like "1 500,50" and "1500.5".
 * Returns 0 for empty/unparseable values.
 */
function asNumber(value) {
  if (value == null || value === "") return 0;
  if (typeof value === "number") return value;
  const cleaned = String(value).replace(/\s/g, "").replace(",", ".");
  const parsed = parseFloat(cleaned);
  return Number.isNaN(parsed) ? 0 : parsed;
}

// ───────────────────────────────────────────────────────────────────────────
// Public API
// ───────────────────────────────────────────────────────────────────────────

/**
 * Read a bank statement .xlsx file and return its parsed rows.
 *
 * @param {string} filePath  Absolute path or path relative to the project root.
 * @returns {{
 *   file: string,
 *   sheet: string,
 *   total_rows: number,
 *   data_rows: number,
 *   rows: Array<Object>
 * }}
 */
export function readBankStatement(filePath) {
  const absPath = resolvePath(filePath);
  const workbook = xlsx.readFile(absPath, { cellDates: false });

  // Take the first sheet — bank exports normally only have one.
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: "" });

  const headerIndex = findHeaderRowIndex(rows);
  if (headerIndex === -1) {
    throw new Error(
      "Header row not found in Excel file. Expected columns: '№ пп', 'Дата документа', 'Обороты по дебету/кредиту'."
    );
  }

  const colMap = buildColumnMap(rows[headerIndex]);
  const colIndex = (...candidates) => {
    for (const c of candidates) {
      const key = normalize(c);
      if (key in colMap) return colMap[key];
    }
    return -1;
  };

  const idxNum = colIndex("№ пп");
  const idxDocDate = colIndex("дата документа");
  const idxProcDate = colIndex("дата обработки");
  const idxDocNum = colIndex("№ док");
  const idxAccountName = colIndex("наименование счёта", "наименование счета");
  const idxInn = colIndex("инн");
  const idxAccountNum = colIndex("№ счёта", "№ счета");
  const idxMfo = colIndex("мфо");
  const idxDebit = colIndex("обороты по дебету");
  const idxCredit = colIndex("обороты по кредиту");
  const idxDescription = colIndex("назначение платежа");

  const dataRows = [];
  for (let i = headerIndex + 1; i < rows.length; i++) {
    const row = rows[i];
    // Skip rows without both a sequence number AND a document date —
    // those are footer/total rows ("Итого за весь период", balances, etc.)
    const numCell = row[idxNum];
    const dateCell = row[idxDocDate];
    const hasNum = numCell !== "" && numCell != null && !isNaN(Number(numCell));
    const hasDate = dateCell !== "" && dateCell != null;
    if (!hasNum || !hasDate) continue;

    dataRows.push({
      row_num: asNumber(numCell),
      doc_date: asString(row[idxDocDate]),
      proc_date: asString(row[idxProcDate]),
      doc_number: asString(row[idxDocNum]),
      account_name: asString(row[idxAccountName]),
      inn: asString(row[idxInn]),
      account_number: asString(row[idxAccountNum]),
      mfo: asString(row[idxMfo]),
      debit: asNumber(row[idxDebit]),
      credit: asNumber(row[idxCredit]),
      description: asString(row[idxDescription]),
    });
  }

  return {
    file: absPath,
    sheet: sheetName,
    total_rows: rows.length,
    data_rows: dataRows.length,
    rows: dataRows,
  };
}
