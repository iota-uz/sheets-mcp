#!/usr/bin/env node
/**
 * Smart append — добавляет строку в любой лист Google Sheets с auto-fill и dup-check.
 *
 * Никакого хардкода имён листов или позиций колонок. Читает заголовки первой
 * строки и определяет всё динамически по паттернам названий (по-русски/по-английски).
 *
 * Usage:
 *   node smart-append.mjs --sheet=<имя_листа> --values='<JSON>'
 *
 * Values — объект, ключи = названия заголовков:
 *   --values='{"Дата":"15.04.2026","Сумма":100000,"Валюта":"SOM","Счёт":"Kapital UZS (р/с)"}'
 *
 * Auto-fill (заполняется если ячейка пустая):
 *   - Курс: читает последнее значение колонки. Для валюты USD = 1.
 *   - Сумма ($): вычисляется как Сумма / Курс.
 *   - Расчётный период: 01-е число месяца колонки Дата.
 *
 * Dup-check: читает последние 50 строк, сравнивает по "ключевым" колонкам
 *   (Итерация как unique key если есть, иначе Дата + Сумма + Категория/Проект).
 *
 * Output JSON:
 *   { appended: true, sheet, range, values }      — записано
 *   { skipped: true, duplicate_row: N, ... }      — дубль, ничего не записано
 */

import { read, append } from "../mcp/sheets.mjs";

// ─── Паттерны заголовков (по ним определяем "роль" колонки) ──────────────
// Важно: JS regex \b не работает с кириллицей — используем explicit границы.
const PATTERNS = {
  date:        /^\s*дата\s*$/i,            // "Дата" или "Дата " (trailing space)
  amount:      /^\s*сумма\s*$/i,           // ровно "Сумма" (не "Сумма ($)")
  usd:         /(сумма.*\$|\$.*сумма|usd\s*$|\$\s*$)/i,
  rate:        /курс/i,
  period:      /период\s*$/i,              // любое "* период" (Расчётный/Рассчетный/Расчетный)
  currency:    /валют/i,
  category:    /категор/i,
  project:     /проект/i,
  iteration:   /итерац|iteration/i,
  description: /описан|назначен|коммент/i, // "Описание", "Назначение платежа", "Комментарий"
};

// ─── CLI parsing ─────────────────────────────────────────────────────────
const args = process.argv.slice(2);
let sheet = null;
let valuesJson = null;
for (const a of args) {
  if (a.startsWith("--sheet=")) sheet = a.slice(8);
  else if (a.startsWith("--values=")) valuesJson = a.slice(9);
}
if (!sheet || !valuesJson) {
  console.error("Usage: --sheet=<name> --values=<JSON>");
  process.exit(1);
}

let userValues;
try {
  userValues = JSON.parse(valuesJson);
} catch (e) {
  console.error(`Invalid JSON in --values: ${e.message}`);
  process.exit(1);
}

// ─── Helpers ────────────────────────────────────────────────────────────
function num(v) {
  if (v == null || v === "") return NaN;
  if (typeof v === "number") return v;
  return parseFloat(String(v).replace(/[\s$]/g, "").replace(",", "."));
}

function colLetter(i) {
  // Supports 0..701 (A..ZZ). Good enough for any reasonable sheet.
  if (i < 26) return String.fromCharCode(65 + i);
  return String.fromCharCode(65 + Math.floor(i / 26) - 1) + String.fromCharCode(65 + (i % 26));
}

// ─── Read headers (scan first 10 rows, find the real header row) ────────
const topRes = await read(sheet, "A1:Z10");
let headerRow = [];
for (const r of topRes.rows) {
  if (!Array.isArray(r)) continue;
  // Heuristic: at least 3 short text cells (not numbers, not empty)
  const labels = r.filter(c => {
    const s = String(c ?? "").trim();
    return s.length > 0 && s.length < 60 && isNaN(parseFloat(s));
  });
  if (labels.length >= 3) { headerRow = r; break; }
}
if (headerRow.length === 0) {
  console.error(`Sheet "${sheet}" has no detectable header row in first 10 rows`);
  process.exit(1);
}

// Игнорируем автогенерированные Google-Sheets мусорные заголовки.
// "Column 1", "Столбец 5" и т.п. — это placeholders когда реальный заголовок пуст.
// Если их включить в колонки, writeRange смещается влево (обычно на колонку А),
// а Google API при append игнорирует пустую ведущую колонку → все значения сдвигаются.
const JUNK_HEADER = /^(column|столбец)\s+\d+$/i;

const colByHeader = {};   // "Дата" → 0
const colByRole = {};     // "date" → 0
headerRow.forEach((h, idx) => {
  const name = String(h ?? "").trim();
  if (!name || JUNK_HEADER.test(name)) return;
  colByHeader[name] = idx;
  for (const [role, re] of Object.entries(PATTERNS)) {
    if (!(role in colByRole) && re.test(name)) colByRole[role] = idx;
  }
});

// Determine write range (A..last header column)
const headerIndices = Object.values(colByHeader);
if (headerIndices.length === 0) {
  console.error(`No headers detected in sheet "${sheet}"`);
  process.exit(1);
}
const minIdx = Math.min(...headerIndices);
const maxIdx = Math.max(...headerIndices);
const writeRange = `${colLetter(minIdx)}:${colLetter(maxIdx)}`;

// ─── Build values array in sheet's column order ──────────────────────────
const values = new Array(maxIdx - minIdx + 1).fill("");
const unknown = [];
for (const [key, val] of Object.entries(userValues)) {
  if (key in colByHeader) {
    values[colByHeader[key] - minIdx] = val;
  } else {
    unknown.push(key);
  }
}
if (unknown.length > 0) {
  console.error(`Warning: unknown headers ignored: ${unknown.join(", ")}`);
}

// ─── Auto-fill ───────────────────────────────────────────────────────────

// 1. Rate (if column exists and is empty)
// Важно: при подборе последнего курса берём только строки с ТОЙ ЖЕ валютой,
// что мы записываем. Иначе для SOM-записи можно случайно схватить курс=1
// из последней USD-строки.
if (colByRole.rate != null && (values[colByRole.rate - minIdx] === "" || values[colByRole.rate - minIdx] == null)) {
  const currency = colByRole.currency != null
    ? String(values[colByRole.currency - minIdx] ?? "").toUpperCase()
    : "";
  if (currency === "USD") {
    values[colByRole.rate - minIdx] = 1;
  } else if (colByRole.currency != null) {
    const rateColIdx = colByRole.rate;
    const currColIdx = colByRole.currency;
    const minC = Math.min(rateColIdx, currColIdx);
    const maxC = Math.max(rateColIdx, currColIdx);
    const res = await read(sheet, `${colLetter(minC)}:${colLetter(maxC)}`, 50);
    for (let i = res.rows.length - 1; i >= 0; i--) {
      const row = res.rows[i].values || res.rows[i];
      const rowCurr = String(row[currColIdx - minC] ?? "").toUpperCase();
      if (rowCurr !== currency) continue;
      const n = num(row[rateColIdx - minC]);
      if (!isNaN(n) && n > 0) {
        values[colByRole.rate - minIdx] = n;
        break;
      }
    }
  }
}

// 2. Period (first of month from date column)
if (colByRole.period != null && colByRole.date != null) {
  const slot = colByRole.period - minIdx;
  if (values[slot] === "" || values[slot] == null) {
    const date = String(values[colByRole.date - minIdx] ?? "").trim();
    const parts = date.split(".");
    if (parts.length === 3) values[slot] = `01.${parts[1]}.${parts[2]}`;
  }
}

// 3. USD sum (amount / rate)
if (colByRole.usd != null && colByRole.amount != null && colByRole.rate != null) {
  const slot = colByRole.usd - minIdx;
  if (values[slot] === "" || values[slot] == null) {
    const amount = num(values[colByRole.amount - minIdx]);
    const rate = num(values[colByRole.rate - minIdx]);
    if (!isNaN(amount) && !isNaN(rate) && rate > 0) {
      values[slot] = Math.round((amount / rate) * 100) / 100;
    }
  }
}

// ─── Duplicate check ─────────────────────────────────────────────────────
const dupRes = await read(sheet, writeRange, 50);
let isDup = false;
let dupRow = null;

// A) Unique key match (e.g., iteration id)
if (colByRole.iteration != null) {
  const slot = colByRole.iteration - minIdx;
  const newVal = String(values[slot] ?? "").trim().toLowerCase();
  if (newVal) {
    for (const row of dupRes.rows) {
      const cells = row.values || row;
      if (String(cells[slot] ?? "").trim().toLowerCase() === newVal) {
        isDup = true;
        dupRow = row.row;
        break;
      }
    }
  }
}

// B) Date + amount + category/project (+ description if present) match.
// Описание включается чтобы не флагать 3 отдельные комиссии 1000 SOM в один день
// как дубликаты — они одинаковы по дате+сумме+категории, но описание разное (разные номера доков).
if (!isDup && colByRole.date != null && colByRole.amount != null) {
  const dateSlot = colByRole.date - minIdx;
  const amtSlot = colByRole.amount - minIdx;
  const groupRole = colByRole.category ?? colByRole.project ?? null;
  const groupSlot = groupRole != null ? groupRole - minIdx : null;
  const descSlot = colByRole.description != null ? colByRole.description - minIdx : null;
  const newDate = String(values[dateSlot] ?? "").trim();
  const newAmt = num(values[amtSlot]);
  const newGroup = groupSlot != null ? String(values[groupSlot] ?? "").trim().toLowerCase() : null;
  const newDesc = descSlot != null ? String(values[descSlot] ?? "").trim().toLowerCase() : null;

  if (newDate && !isNaN(newAmt)) {
    for (const row of dupRes.rows) {
      const cells = row.values || row;
      if (String(cells[dateSlot] ?? "").trim() !== newDate) continue;
      if (Math.abs(num(cells[amtSlot]) - newAmt) > 0.01) continue;
      if (groupSlot != null && String(cells[groupSlot] ?? "").trim().toLowerCase() !== newGroup) continue;
      // Если у обеих строк есть непустое описание и они отличаются — это НЕ дубль.
      if (descSlot != null && newDesc) {
        const existDesc = String(cells[descSlot] ?? "").trim().toLowerCase();
        if (existDesc && existDesc !== newDesc) continue;
      }
      isDup = true;
      dupRow = row.row;
      break;
    }
  }
}

if (isDup) {
  console.log(JSON.stringify({
    skipped: true,
    reason: "duplicate",
    duplicate_row: dupRow,
    sheet, writeRange,
  }, null, 2));
  process.exit(0);
}

// ─── Append ──────────────────────────────────────────────────────────────
// Если в листе есть "разрыв" в данных (например блок 3-64 + блок 99-108),
// обычный append("A:H") находит пустоту на row 65 и пишет туда. Чтобы запись
// попадала в КОНЕЦ листа, сами находим последний ряд с данными и передаём
// range начинающийся с него — тогда Google API корректно определяет конец
// таблицы и пишет следом.
const fullRead = await read(sheet, writeRange, 5000);
let lastRow = 0;
for (let i = fullRead.rows.length - 1; i >= 0; i--) {
  const row = fullRead.rows[i];
  const cells = row.values || row;
  if (Array.isArray(cells) && cells.some(c => c != null && String(c).trim())) {
    lastRow = row.row;
    break;
  }
}
const minCol = colLetter(minIdx);
const maxCol = colLetter(maxIdx);
const appendRange = lastRow > 0
  ? `${minCol}${lastRow}:${maxCol}`
  : writeRange;

const result = await append(sheet, appendRange, values);
console.log(JSON.stringify({ appended: true, ...result }, null, 2));
