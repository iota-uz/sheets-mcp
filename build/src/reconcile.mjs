/**
 * Balance reconciliation for the "Счета" summary sheet.
 *
 * Three-layer check:
 *   1. Per-account — sums the money flow through each account from the
 *      expenses / payments / transfers sheets and shows the actual balance
 *      next to it. Also detects "orphan" transactions that reference an
 *      account which is not listed in "Счета".
 *   2. Per-period — narrows the hunt to the given period (current month
 *      by default) using the "Расчётный период" column.
 *   3. Pattern match — scans in-period transactions for amounts that
 *      match the discrepancy: exact, doubled, halved, or decimal-shift.
 */

import { read } from "./sheets.mjs";

// ───────────────────────────────────────────────────────────────────────────
// Helpers
// ───────────────────────────────────────────────────────────────────────────

/** Parse a number from the Sheets display format ("$1 234,56", "748 879 434,82"). */
function num(val) {
  if (val == null || val === "") return 0;
  if (typeof val === "number") return val;
  const cleaned = String(val).replace(/[\s$]/g, "").replace(",", ".");
  const n = parseFloat(cleaned);
  return isNaN(n) ? 0 : n;
}

function str(val) {
  return String(val ?? "").trim();
}

/** Расчётный период for the current month ("01.04.2026"). */
function currentPeriod() {
  const d = new Date();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  return `01.${mm}.${d.getFullYear()}`;
}

/** Normalize an account name for comparison (drop parens, case, extra spaces). */
function normAccount(name) {
  return str(name)
    .toLowerCase()
    .replace(/\([^)]*\)/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

// ───────────────────────────────────────────────────────────────────────────
// Transaction sheet layouts (indices within the read range)
// ───────────────────────────────────────────────────────────────────────────

const EXPENSES  = { sheet: "расходы",  range: "B:K", date: 0, category: 1, amount: 2, rate: 3, currency: 4, usd: 5, period: 6, account: 7, description: 9 };
const PAYMENTS  = { sheet: "платежи",  range: "A:K", amount: 2, rate: 4, currency: 5, usd: 6, date: 7, period: 8, account: 9, description: 10 };
const TRANSFERS = { sheet: "ддс",      range: "A:H", date: 0, amount: 1, from: 2, to: 3, period: 5, rate: 6, description: 7 };

// ───────────────────────────────────────────────────────────────────────────
// Dynamic parsing of the "Счета" sheet (not tied to specific row numbers)
// ───────────────────────────────────────────────────────────────────────────

const SUMMARY_LABELS = {
  income:        /^всего\s+поступ/i,
  expenses:      /^всего\s+расход/i,
  debtIn:        /^получено\s+в\s+долг/i,
  debtOut:       /^выдано\s+в\s+долг/i,
  tableBalance:  /^баланс\s+согласно/i,
  actualBalance: /^фактическ.*баланс/i,
  difference:    /^разница/i,
};

function parseAccountsSheet(rows) {
  const accounts = [];
  const summary = {};

  for (const r of rows) {
    if (!Array.isArray(r)) continue;
    const label = str(r[0]);
    if (!label) continue;

    let matchedSummary = false;
    for (const [key, re] of Object.entries(SUMMARY_LABELS)) {
      if (re.test(label)) {
        summary[key] = num(r[6]);
        matchedSummary = true;
        break;
      }
    }
    if (matchedSummary) continue;

    const amount = num(r[1]);
    const currency = str(r[2]);
    if (amount > 0 && currency) {
      accounts.push({
        name: label,
        amount,
        currency,
        usdBalance: num(r[3]),
      });
    }
  }

  return { accounts, summary };
}

// ───────────────────────────────────────────────────────────────────────────
// Per-account money flow (aggregated over all transactions)
// ───────────────────────────────────────────────────────────────────────────

function buildAccountFlow(expensesRows, paymentsRows, transfersRows) {
  const flow = new Map();
  const ensure = (key) => {
    if (!flow.has(key)) flow.set(key, { incoming: 0, outgoing: 0, transfersIn: 0, transfersOut: 0, count: 0 });
    return flow.get(key);
  };

  for (const r of expensesRows) {
    const key = normAccount(r[EXPENSES.account]);
    if (!key) continue;
    const b = ensure(key);
    b.outgoing += num(r[EXPENSES.usd]);
    b.count++;
  }
  for (const r of paymentsRows) {
    const key = normAccount(r[PAYMENTS.account]);
    if (!key) continue;
    const b = ensure(key);
    b.incoming += num(r[PAYMENTS.usd]);
    b.count++;
  }
  for (const r of transfersRows) {
    const rate = num(r[TRANSFERS.rate]);
    const raw = num(r[TRANSFERS.amount]);
    const usd = rate > 0 ? raw / rate : 0;
    const from = normAccount(r[TRANSFERS.from]);
    const to = normAccount(r[TRANSFERS.to]);
    if (from) { const b = ensure(from); b.transfersOut += usd; b.count++; }
    if (to)   { const b = ensure(to);   b.transfersIn  += usd; b.count++; }
  }

  return flow;
}

// ───────────────────────────────────────────────────────────────────────────
// Pattern scan — hunts for transactions that could explain the diff
// ───────────────────────────────────────────────────────────────────────────

function scanSuspects(transactions, absDiff, tolerance, ratePct) {
  if (absDiff < tolerance) return [];
  const out = [];
  const close = (a, b) => Math.abs(a - b) <= tolerance;
  const closePct = (a, b, pct) => Math.abs(a - b) <= b * (pct / 100);

  for (const t of transactions) {
    const { sheet, row, usd, raw, rate, currency, description } = t;

    if (close(usd, absDiff)) {
      out.push({ sheet, row, reason: "exact_match",
                 note: "Сумма транзакции в $ равна разнице — возможно запись лишняя или пропущена",
                 usd, raw, currency, description });
    }
    if (close(usd * 2, absDiff)) {
      out.push({ sheet, row, reason: "doubled",
                 note: "Сумма × 2 = разница — возможно запись продублирована",
                 usd, raw, currency, description });
    }
    if (close(usd / 2, absDiff)) {
      out.push({ sheet, row, reason: "halved",
                 note: "Сумма / 2 = разница — возможно сумма вдвое больше нужного",
                 usd, raw, currency, description });
    }
    if (currency.toUpperCase() === "SOM" && raw > 0 && closePct(raw, absDiff, ratePct)) {
      out.push({ sheet, row, reason: "currency_error",
                 note: "Значение в SOM близко к разнице в $ — возможно перепутана валюта",
                 usd, raw, currency, description });
    }
    if (rate > 0) {
      const rawTarget = absDiff * rate;
      if (closePct(raw * 10, rawTarget, ratePct) || closePct(raw / 10, rawTarget, ratePct)) {
        out.push({ sheet, row, reason: "decimal_shift",
                   note: "Сумма отличается от разницы в 10 раз — возможно пропущен или лишний ноль",
                   usd, raw, currency, description });
      }
    }
  }

  const seen = new Set();
  return out.filter(s => {
    const k = `${s.sheet}|${s.row}|${s.reason}`;
    if (seen.has(k)) return false;
    seen.add(k);
    return true;
  });
}

// ───────────────────────────────────────────────────────────────────────────
// Public API
// ───────────────────────────────────────────────────────────────────────────

/**
 * Run the full reconciliation.
 *
 * @param {Object} [options]
 * @param {string} [options.period]     Расчётный период (defaults to current month)
 * @param {number} [options.tolerance]  USD tolerance for pattern matching (default 1.0)
 * @param {number} [options.ratePct]    % tolerance for currency-error detection (default 2)
 */
export async function reconcile(options = {}) {
  const { tolerance = 1.0, ratePct = 2 } = options;
  const period = options.period || currentPeriod();

  // 1) Read "Счета" dynamically ----------------------------------------------
  const accountsSheet = await read("Счета", "A1:G50");
  const { accounts, summary } = parseAccountsSheet(accountsSheet.rows);

  const difference = summary.difference ?? 0;
  const absDiff = Math.abs(difference);

  // 2) Read all transactions -------------------------------------------------
  const [expensesAll, paymentsAll, transfersAll] = await Promise.all([
    read(EXPENSES.sheet,  EXPENSES.range),
    read(PAYMENTS.sheet,  PAYMENTS.range),
    read(TRANSFERS.sheet, TRANSFERS.range),
  ]);

  // 3) Per-account flow (all-time) ------------------------------------------
  const flow = buildAccountFlow(expensesAll.rows, paymentsAll.rows, transfersAll.rows);

  const perAccount = accounts.map(a => {
    const f = flow.get(normAccount(a.name)) || { incoming: 0, outgoing: 0, transfersIn: 0, transfersOut: 0, count: 0 };
    const netFlow = f.incoming - f.outgoing + f.transfersIn - f.transfersOut;
    return {
      name: a.name,
      currency: a.currency,
      actualAmount: a.amount,
      actualUsd: a.usdBalance,
      incomingUsd: f.incoming,
      outgoingUsd: f.outgoing,
      transfersInUsd: f.transfersIn,
      transfersOutUsd: f.transfersOut,
      netFlowUsd: netFlow,
      txCount: f.count,
    };
  });

  // Transactions that reference an account not present in "Счета"
  const knownKeys = new Set(accounts.map(a => normAccount(a.name)));
  const orphans = [];
  for (const [key, data] of flow.entries()) {
    if (!knownKeys.has(key) && key) {
      orphans.push({
        account: key,
        incomingUsd: data.incoming,
        outgoingUsd: data.outgoing,
        transfersInUsd: data.transfersIn,
        transfersOutUsd: data.transfersOut,
        txCount: data.count,
      });
    }
  }

  // 4) In-period transactions (for the pattern scanner) ---------------------
  const pickPeriod = (allRows, layout) =>
    allRows.map((v, i) => ({ row: i + 1, values: v }))
      .filter(t => Array.isArray(t.values) && str(t.values[layout.period]) === period);

  const periodExpenses  = pickPeriod(expensesAll.rows,  EXPENSES);
  const periodPayments  = pickPeriod(paymentsAll.rows,  PAYMENTS);
  const periodTransfers = pickPeriod(transfersAll.rows, TRANSFERS);

  const periodTx = [];
  periodExpenses.forEach(t => periodTx.push({
    sheet: EXPENSES.sheet, row: t.row,
    usd: num(t.values[EXPENSES.usd]), raw: num(t.values[EXPENSES.amount]),
    rate: num(t.values[EXPENSES.rate]), currency: str(t.values[EXPENSES.currency]),
    description: str(t.values[EXPENSES.description]),
  }));
  periodPayments.forEach(t => periodTx.push({
    sheet: PAYMENTS.sheet, row: t.row,
    usd: num(t.values[PAYMENTS.usd]), raw: num(t.values[PAYMENTS.amount]),
    rate: num(t.values[PAYMENTS.rate]), currency: str(t.values[PAYMENTS.currency]),
    description: str(t.values[PAYMENTS.description]),
  }));
  periodTransfers.forEach(t => {
    const rate = num(t.values[TRANSFERS.rate]);
    const raw = num(t.values[TRANSFERS.amount]);
    periodTx.push({
      sheet: TRANSFERS.sheet, row: t.row,
      usd: rate > 0 ? raw / rate : 0, raw,
      rate, currency: "SOM",
      description: str(t.values[TRANSFERS.description]),
    });
  });

  const suspects = scanSuspects(periodTx, absDiff, tolerance, ratePct);
  const topLargest = [...periodTx].sort((a, b) => b.usd - a.usd).slice(0, 5);

  return {
    period,
    summary: {
      tableBalanceUsd: summary.tableBalance ?? 0,
      actualBalanceUsd: summary.actualBalance ?? 0,
      differenceUsd: difference,
      absDifferenceUsd: absDiff,
    },
    perAccount,
    orphans,
    inPeriodCounts: {
      expenses: periodExpenses.length,
      payments: periodPayments.length,
      transfers: periodTransfers.length,
    },
    suspects,
    topLargestInPeriod: topLargest,
  };
}
