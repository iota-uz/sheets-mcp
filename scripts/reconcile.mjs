#!/usr/bin/env node
/**
 * Generic balance reconciliation.
 *
 * Открывает структуру таблицы ДИНАМИЧЕСКИ — не хардкодит ни имён листов,
 * ни позиций колонок. Работает с любой таблицей, где есть:
 *   - Один лист-сводка со структурой "название+сумма+валюта" в верхней части
 *     и строками-итогами внизу (метки "баланс"/"разница"/"фактический").
 *   - Транзакционные листы с колонками "Дата"/"Сумма"/"Счёт" (возможно с
 *     префиксом "входящий"/"исходящий" для листов-переводов).
 *
 * Usage:
 *   node scripts/reconcile.mjs                      # текстовый отчёт, текущий месяц
 *   node scripts/reconcile.mjs --json               # JSON-вывод для интеграций
 *   node scripts/reconcile.mjs --period=01.04.2026  # сверка за другой месяц
 *   node scripts/reconcile.mjs --tolerance=5        # допуск $ для паттернов
 */

import { read, info, toNumber } from "../mcp/sheets.mjs";

// ─── Паттерны колонок (определяют роль колонки по заголовку) ─────────────
const COL = {
  date:         /^\s*дата\b/i,
  amount:       /^\s*сумма\b(?!.*\$)/i,
  usd:          /(сумма.*\$|\$.*сумма|\$$)/i,
  rate:         /курс/i,
  period:       /расч.*период|^\s*период\b/i,
  currency:     /валют/i,
  category:     /категор/i,
  project:      /проект/i,
  description:  /описани|назначени|коммент/i,
  account:      /^\s*счёт|^\s*счет|^\s*account/i,
  accountFrom:  /исход.*счёт|исход.*счет|источник/i,
  accountTo:    /вход.*счёт|вход.*счет|получател/i,
};

// ─── Паттерны строк-итогов в сводном листе ──────────────────────────────
const SUMMARY_LABELS = {
  tableBalance:  /баланс.*таблиц|баланс.*согласно/i,
  actualBalance: /фактическ.*баланс/i,
  difference:    /^\s*разница|^\s*расхожд/i,
};

// ─── CLI parsing ────────────────────────────────────────────────────────
const args = process.argv.slice(2);
let period = null, tolerance = 1.0, ratePct = 2.0, asJson = false;
for (const a of args) {
  if (a === "--json") asJson = true;
  else if (a.startsWith("--period="))    period    = a.slice(9);
  else if (a.startsWith("--tolerance=")) tolerance = parseFloat(a.slice(12));
  else if (a.startsWith("--ratePct="))   ratePct   = parseFloat(a.slice(10));
}
if (!period) {
  const d = new Date();
  period = `01.${String(d.getMonth() + 1).padStart(2, "0")}.${d.getFullYear()}`;
}

// ─── Helpers ────────────────────────────────────────────────────────────
function num(v) {
  const n = toNumber(v);
  return typeof n === "number" ? n : NaN;
}
function str(v) { return String(v ?? "").trim(); }
function normAccount(n) {
  return str(n).toLowerCase().replace(/\([^)]*\)/g, "").replace(/\s+/g, " ").trim();
}
function classifyHeaders(headerRow) {
  // headerRow: массив строк-заголовков
  // Возвращает { role: columnIndex } + { headerName: columnIndex }
  const byRole = {};
  const byName = {};
  headerRow.forEach((h, idx) => {
    const name = str(h);
    if (!name) return;
    byName[name] = idx;
    // Более специфичные паттерны имеют приоритет:
    // сначала accountFrom/accountTo, потом generic account
    if (COL.accountFrom.test(name)) { if (!("accountFrom" in byRole)) byRole.accountFrom = idx; return; }
    if (COL.accountTo.test(name))   { if (!("accountTo"   in byRole)) byRole.accountTo   = idx; return; }
    for (const [role, re] of Object.entries(COL)) {
      if (role === "accountFrom" || role === "accountTo") continue;
      if (!(role in byRole) && re.test(name)) byRole[role] = idx;
    }
  });
  return { byRole, byName };
}

// ─── Находим сводный лист и транзакционные листы ────────────────────────
const meta = await info();
const sheetNames = meta.sheets;

const sheetScans = {};  // sheetName → { kind, layout, headerRow, allRows }
let summarySheet = null;

for (const name of sheetNames) {
  try {
    const full = await read(name);
    const rows = full.rows;
    if (rows.length < 2) continue;

    const headerRow = rows[0];
    const { byRole } = classifyHeaders(headerRow);

    // Транзакционный лист: есть date + amount + хотя бы один account
    const hasTxShape =
      "date" in byRole &&
      "amount" in byRole &&
      ("account" in byRole || "accountFrom" in byRole || "accountTo" in byRole);

    if (hasTxShape) {
      sheetScans[name] = { kind: "transactions", layout: byRole, rows };
      continue;
    }

    // Сводный лист: первая строка-название ("Счета"?) или строки "название+сумма+валюта"
    // + строки-итоги с метками "разница"/"баланс" где-то ниже
    let hasSummaryLabels = false;
    for (const r of rows) {
      if (!Array.isArray(r)) continue;
      const label = str(r[0]);
      if (SUMMARY_LABELS.difference.test(label) || SUMMARY_LABELS.tableBalance.test(label)) {
        hasSummaryLabels = true;
        break;
      }
    }
    if (hasSummaryLabels) {
      sheetScans[name] = { kind: "summary", rows };
      if (!summarySheet) summarySheet = name;
    }
  } catch (e) {
    // игнорируем недоступные листы
  }
}

if (!summarySheet) {
  console.error("Не найден сводный лист (с метками 'Баланс'/'Разница').");
  process.exit(1);
}

// ─── Парсим сводный лист ────────────────────────────────────────────────
function parseSummarySheet(rows) {
  const accounts = [];
  const summary = {};
  for (const r of rows) {
    if (!Array.isArray(r)) continue;
    const label = str(r[0]);
    if (!label) continue;

    // Итоговая строка: метка в A, значение где-то в последних колонках
    let matchedSummary = false;
    for (const [key, re] of Object.entries(SUMMARY_LABELS)) {
      if (re.test(label)) {
        // Ищем число в строке, начиная с правого края
        for (let i = r.length - 1; i >= 1; i--) {
          const n = num(r[i]);
          if (!isNaN(n)) { summary[key] = n; break; }
        }
        matchedSummary = true;
        break;
      }
    }
    if (matchedSummary) continue;

    // Строка-счёт: имя в A, число в B, валюта в C
    const amount = num(r[1]);
    const currency = str(r[2]);
    if (!isNaN(amount) && amount > 0 && currency && currency.length <= 6) {
      // usd баланс обычно в следующей колонке после валюты или через одну
      let usdBalance = 0;
      for (let i = 3; i < Math.min(r.length, 6); i++) {
        const n = num(r[i]);
        if (!isNaN(n)) { usdBalance = n; break; }
      }
      accounts.push({ name: label, amount, currency, usdBalance });
    }
  }
  return { accounts, summary };
}

const { accounts, summary: summaryTotals } = parseSummarySheet(sheetScans[summarySheet].rows);
const difference = summaryTotals.difference ?? 0;
const absDiff = Math.abs(difference);

// ─── Строим поток по счетам из всех транзакционных листов ────────────────
const flow = new Map();
const ensureFlow = (key) => {
  if (!flow.has(key)) flow.set(key, { incoming: 0, outgoing: 0, count: 0 });
  return flow.get(key);
};

const periodTx = [];

for (const [sheetName, scan] of Object.entries(sheetScans)) {
  if (scan.kind !== "transactions") continue;
  const { layout, rows } = scan;

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!Array.isArray(r)) continue;

    const usd = layout.usd != null ? num(r[layout.usd]) : NaN;
    const raw = layout.amount != null ? num(r[layout.amount]) : NaN;
    const rate = layout.rate != null ? num(r[layout.rate]) : NaN;
    const usdAmount = !isNaN(usd) ? usd : (!isNaN(raw) && !isNaN(rate) && rate > 0 ? raw / rate : 0);
    const currency = layout.currency != null ? str(r[layout.currency]) : "";
    const txPeriod = layout.period != null ? str(r[layout.period]) : "";
    const description = layout.description != null ? str(r[layout.description]) : "";

    // Determine direction
    if (layout.accountFrom != null && layout.accountTo != null) {
      // Transfer: out of one account, into another
      const from = normAccount(r[layout.accountFrom]);
      const to = normAccount(r[layout.accountTo]);
      if (from) { const b = ensureFlow(from); b.outgoing += usdAmount; b.count++; }
      if (to)   { const b = ensureFlow(to);   b.incoming += usdAmount; b.count++; }
    } else if (layout.accountTo != null) {
      // Incoming-only (платежи): money into accountTo
      const to = normAccount(r[layout.accountTo]);
      if (to) { const b = ensureFlow(to); b.incoming += usdAmount; b.count++; }
    } else if (layout.accountFrom != null) {
      // Outgoing-only
      const from = normAccount(r[layout.accountFrom]);
      if (from) { const b = ensureFlow(from); b.outgoing += usdAmount; b.count++; }
    } else if (layout.account != null) {
      // Generic "Счёт" — наиболее частый случай для расходов; считаем outgoing
      const a = normAccount(r[layout.account]);
      if (a) { const b = ensureFlow(a); b.outgoing += usdAmount; b.count++; }
    }

    // Collect transactions in the selected period (for pattern scanning)
    if (txPeriod === period) {
      periodTx.push({
        sheet: sheetName,
        row: i + 1,
        usd: usdAmount,
        raw: !isNaN(raw) ? raw : 0,
        rate: !isNaN(rate) ? rate : 0,
        currency,
        description,
      });
    }
  }
}

// ─── Per-account view ────────────────────────────────────────────────────
const perAccount = accounts.map(a => {
  const f = flow.get(normAccount(a.name)) || { incoming: 0, outgoing: 0, count: 0 };
  return {
    name: a.name,
    currency: a.currency,
    actualAmount: a.amount,
    actualUsd: a.usdBalance,
    incomingUsd: f.incoming,
    outgoingUsd: f.outgoing,
    netFlowUsd: f.incoming - f.outgoing,
    txCount: f.count,
  };
});

// ─── Orphans (счета в транзакциях, которых нет в сводном) ───────────────
const knownKeys = new Set(accounts.map(a => normAccount(a.name)));
const orphans = [];
for (const [key, data] of flow.entries()) {
  if (!knownKeys.has(key) && key) {
    orphans.push({
      account: key,
      incomingUsd: data.incoming,
      outgoingUsd: data.outgoing,
      txCount: data.count,
    });
  }
}

// ─── Pattern scan ────────────────────────────────────────────────────────
const suspects = [];
if (absDiff >= tolerance) {
  const close = (a, b) => Math.abs(a - b) <= tolerance;
  const closePct = (a, b, pct) => Math.abs(a - b) <= b * (pct / 100);

  for (const t of periodTx) {
    if (close(t.usd, absDiff)) suspects.push({ ...t, reason: "exact_match",
      note: "Сумма транзакции в $ равна разнице — возможно запись лишняя или пропущена" });
    if (close(t.usd * 2, absDiff)) suspects.push({ ...t, reason: "doubled",
      note: "Сумма × 2 = разница — возможно запись продублирована" });
    if (close(t.usd / 2, absDiff)) suspects.push({ ...t, reason: "halved",
      note: "Сумма / 2 = разница — возможно сумма вдвое больше нужного" });
    // Для любой НЕ-USD валюты: если сырая сумма близка к разнице в $, возможно перепутана валюта
    if (t.currency && t.currency.toUpperCase() !== "USD" && t.raw > 0 && closePct(t.raw, absDiff, ratePct)) {
      suspects.push({ ...t, reason: "currency_error",
        note: `Значение в ${t.currency} близко к разнице в $ — возможно перепутана валюта` });
    }
    if (t.rate > 0) {
      const rawTarget = absDiff * t.rate;
      if (closePct(t.raw * 10, rawTarget, ratePct) || closePct(t.raw / 10, rawTarget, ratePct)) {
        suspects.push({ ...t, reason: "decimal_shift",
          note: "Сумма отличается от разницы в 10 раз — возможно пропущен или лишний ноль" });
      }
    }
  }
}

// ─── Top-5 largest in period ────────────────────────────────────────────
const topLargest = [...periodTx].sort((a, b) => b.usd - a.usd).slice(0, 5);

// ─── Result ──────────────────────────────────────────────────────────────
const result = {
  period,
  discoveredSheets: {
    summary: summarySheet,
    transactions: Object.entries(sheetScans).filter(([, s]) => s.kind === "transactions").map(([n]) => n),
  },
  summary: {
    tableBalanceUsd:  summaryTotals.tableBalance  ?? 0,
    actualBalanceUsd: summaryTotals.actualBalance ?? 0,
    differenceUsd:    difference,
    absDifferenceUsd: absDiff,
  },
  perAccount,
  orphans,
  inPeriodCount: periodTx.length,
  suspects,
  topLargestInPeriod: topLargest,
};

// ─── Output ──────────────────────────────────────────────────────────────
if (asJson) {
  console.log(JSON.stringify(result, null, 2));
} else {
  const fmtUsd = (n) => `$${(n ?? 0).toFixed(2)}`;
  const fmtAmount = (n) => (n ?? 0).toLocaleString("ru-RU", { maximumFractionDigits: 2 });
  const pad = (s, w) => String(s).padEnd(w);

  console.log(`\n=== Сверка — период ${result.period} ===\n`);
  console.log(`Сводный лист:        ${result.discoveredSheets.summary}`);
  console.log(`Транзакционные:      ${result.discoveredSheets.transactions.join(", ")}`);
  console.log();
  console.log(`Баланс по таблице:   ${fmtUsd(result.summary.tableBalanceUsd)}`);
  console.log(`Фактический баланс:  ${fmtUsd(result.summary.actualBalanceUsd)}`);
  console.log(`Разница:             ${fmtUsd(result.summary.differenceUsd)}`);
  console.log(`|Δ| для поиска:      ${fmtUsd(result.summary.absDifferenceUsd)}\n`);

  console.log("=== По счетам (за всё время) ===\n");
  console.log(pad("счёт", 28) + pad("актуально", 22) + pad("поступ.", 14) + pad("расходы", 14) + pad("нетто", 14) + "тр.");
  console.log("-".repeat(92));
  for (const a of result.perAccount) {
    console.log(
      pad(a.name.slice(0, 26), 28) +
      pad(`${fmtAmount(a.actualAmount)} ${a.currency}`, 22) +
      pad(fmtUsd(a.incomingUsd), 14) +
      pad(fmtUsd(a.outgoingUsd), 14) +
      pad(fmtUsd(a.netFlowUsd), 14) +
      String(a.txCount)
    );
  }

  if (result.orphans.length > 0) {
    console.log("\n=== Орфаны (счета вне сводного листа) ===\n");
    for (const o of result.orphans) {
      console.log(`  "${o.account}"  поступ=${fmtUsd(o.incomingUsd)}  расх=${fmtUsd(o.outgoingUsd)}  (${o.txCount} тр.)`);
    }
  } else {
    console.log("\n=== Орфанов нет ===");
  }

  console.log(`\n=== Транзакций в периоде: ${result.inPeriodCount} ===\n`);
  if (result.suspects.length === 0) {
    console.log("Подозрительных строк по паттернам не найдено.");
  } else {
    console.log(`Найдено подозрительных строк: ${result.suspects.length}\n`);
    for (const s of result.suspects) {
      console.log(`[${s.sheet} row ${s.row}] ${s.reason}`);
      console.log(`  → ${s.note}`);
      console.log(`  сумма: ${fmtAmount(s.raw)} ${s.currency}  (${fmtUsd(s.usd)})`);
      if (s.description) console.log(`  ${s.description.slice(0, 120)}`);
      console.log();
    }
  }

  console.log("\n=== Топ-5 крупнейших в периоде ===\n");
  for (const t of result.topLargestInPeriod) {
    console.log(`[${t.sheet} row ${t.row}] ${fmtUsd(t.usd)}  (${fmtAmount(t.raw)} ${t.currency})`);
    if (t.description) console.log(`  ${t.description.slice(0, 120)}`);
  }
}
