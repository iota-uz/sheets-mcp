/**
 * CLI wrapper for the reconcile module. Prints a readable report.
 *
 * Usage:
 *   node build/scripts/reconcile.mjs                      # current month
 *   node build/scripts/reconcile.mjs --period=01.04.2026  # specific period
 *   node build/scripts/reconcile.mjs --tolerance=5        # pattern-match tolerance
 */

import { reconcile } from "../src/reconcile.mjs";

const args = process.argv.slice(2);
const opts = {};
for (const a of args) {
  if (a.startsWith("--period=")) opts.period = a.slice(9);
  else if (a.startsWith("--tolerance=")) opts.tolerance = parseFloat(a.slice(12));
  else if (a.startsWith("--ratePct=")) opts.ratePct = parseFloat(a.slice(10));
}

const fmtUsd = (n) => `$${(n ?? 0).toFixed(2)}`;
const fmtAmount = (n) => (n ?? 0).toLocaleString("ru-RU", { maximumFractionDigits: 2 });
const pad = (s, w) => String(s).padEnd(w);

const r = await reconcile(opts);

console.log(`\n=== Сверка — период ${r.period} ===\n`);
console.log(`Баланс по таблице:   ${fmtUsd(r.summary.tableBalanceUsd)}`);
console.log(`Фактический баланс:  ${fmtUsd(r.summary.actualBalanceUsd)}`);
console.log(`Разница:             ${fmtUsd(r.summary.differenceUsd)}`);
console.log(`|Δ| для поиска:      ${fmtUsd(r.summary.absDifferenceUsd)}\n`);

console.log("=== По счетам (за всё время) ===\n");
console.log(pad("счёт", 28) + pad("актуально", 22) + pad("поступ.", 14) + pad("расходы", 14) + pad("ддс вх.", 12) + pad("ддс исх.", 12) + pad("нетто", 14) + "тр.");
console.log("-".repeat(116));
for (const a of r.perAccount) {
  console.log(
    pad(a.name.slice(0, 26), 28) +
    pad(`${fmtAmount(a.actualAmount)} ${a.currency}`, 22) +
    pad(fmtUsd(a.incomingUsd), 14) +
    pad(fmtUsd(a.outgoingUsd), 14) +
    pad(fmtUsd(a.transfersInUsd), 12) +
    pad(fmtUsd(a.transfersOutUsd), 12) +
    pad(fmtUsd(a.netFlowUsd), 14) +
    String(a.txCount)
  );
}

if (r.orphans.length > 0) {
  console.log("\n=== Орфаны (счета, которых НЕТ в листе 'Счета', но на них есть транзакции) ===\n");
  for (const o of r.orphans) {
    console.log(`  "${o.account}"  поступ=${fmtUsd(o.incomingUsd)}  расх=${fmtUsd(o.outgoingUsd)}  ддс_вх=${fmtUsd(o.transfersInUsd)}  ддс_исх=${fmtUsd(o.transfersOutUsd)}  (${o.txCount} тр.)`);
  }
} else {
  console.log("\n=== Орфанов не найдено — все счета в транзакциях есть в листе 'Счета' ===");
}

console.log(`\n=== Транзакций в периоде: расходы=${r.inPeriodCounts.expenses}, платежи=${r.inPeriodCounts.payments}, ддс=${r.inPeriodCounts.transfers} ===\n`);

if (r.suspects.length === 0) {
  console.log("Подозрительных строк по паттернам не найдено.");
} else {
  console.log(`Найдено подозрительных строк: ${r.suspects.length}\n`);
  for (const s of r.suspects) {
    console.log(`[${s.sheet} row ${s.row}] ${s.reason}`);
    console.log(`  → ${s.note}`);
    console.log(`  сумма: ${fmtAmount(s.raw)} ${s.currency}  (${fmtUsd(s.usd)})`);
    if (s.description) console.log(`  описание: ${s.description.slice(0, 120)}`);
    console.log();
  }
}

console.log("\n=== Топ-5 крупнейших в периоде (для глазного контроля) ===\n");
for (const t of r.topLargestInPeriod) {
  console.log(`[${t.sheet} row ${t.row}] ${fmtUsd(t.usd)}  (${fmtAmount(t.raw)} ${t.currency})`);
  if (t.description) console.log(`  ${t.description.slice(0, 120)}`);
}
