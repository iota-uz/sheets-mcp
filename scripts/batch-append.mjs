#!/usr/bin/env node
/**
 * Batch append — записывает много строк в лист через smart-append логику.
 *
 * Входные данные: JSON-файл с массивом объектов.
 * Каждый объект — мапа header_name → value (так же как для одиночного smart-append).
 *
 * Usage:
 *   # с файлом
 *   node scripts/batch-append.mjs --sheet=расходы --file=reports/march.json
 *
 *   # с stdin (pipe)
 *   echo '[{"Дата":"01.03.2026","Сумма":1000,...}, ...]' | node scripts/batch-append.mjs --sheet=расходы
 *
 * Вывод: построчный прогресс + итоговая сводка (записано/пропущено/ошибок).
 */

import fs from "fs";
import { spawnSync } from "child_process";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const SMART_APPEND = path.join(__dirname, "smart-append.mjs");

const args = process.argv.slice(2);
let sheet = null, file = null;
for (const a of args) {
  if (a.startsWith("--sheet=")) sheet = a.slice(8);
  else if (a.startsWith("--file=")) file = a.slice(7);
}
if (!sheet) {
  console.error("Usage: --sheet=<name> [--file=<path>] (stdin if no file)");
  process.exit(1);
}

let raw;
if (file) raw = fs.readFileSync(file, "utf-8");
else raw = fs.readFileSync(0, "utf-8");   // read stdin

let rows;
try { rows = JSON.parse(raw); }
catch (e) {
  console.error(`Invalid JSON: ${e.message}`);
  process.exit(1);
}
if (!Array.isArray(rows)) {
  console.error("Input must be a JSON array of objects");
  process.exit(1);
}

console.log(`\nBatch append → "${sheet}" (${rows.length} rows)\n`);

let ok = 0, dup = 0, err = 0;
const dupList = [];
const errList = [];

async function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

for (let i = 0; i < rows.length; i++) {
  const row = rows[i];
  // Retry: Google Sheets API иногда даёт rate limit / transient errors.
  // 3 попытки с экспоненциальным backoff.
  let proc, out, res;
  let attempts = 0;
  while (attempts < 3) {
    proc = spawnSync("node", [SMART_APPEND, `--sheet=${sheet}`, `--values=${JSON.stringify(row)}`], { encoding: "utf-8" });
    out = (proc.stdout || "").trim();
    try { res = JSON.parse(out); break; }
    catch (e) {
      attempts++;
      if (attempts < 3) { await sleep(1000 * attempts); continue; }
    }
  }
  if (!res) {
    err++;
    const stderr = (proc?.stderr || "").trim().slice(0, 200);
    errList.push({ i: i + 1, row, stdout: out.slice(0, 200), stderr });
    console.log(`✗ #${i + 1}  ERR (after 3 retries): ${stderr || "(no stderr)"}`);
    continue;
  }
  if (res.appended) {
    ok++;
    console.log(`✓ #${i + 1}  → ${res.range}`);
  } else if (res.skipped) {
    dup++;
    dupList.push({ i: i + 1, row: res.duplicate_row });
    console.log(`⊘ #${i + 1}  дубль в row ${res.duplicate_row}`);
  } else {
    err++;
    errList.push({ i: i + 1, row, unknown_response: res });
    console.log(`? #${i + 1}  неизвестный ответ`);
  }
}

console.log(`\nИтог: ${ok} записано, ${dup} пропущено (дубли), ${err} ошибок`);
if (dup > 0) console.log("Дубли:", JSON.stringify(dupList));
if (err > 0) console.log("Ошибки:", JSON.stringify(errList, null, 2));
