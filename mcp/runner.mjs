/**
 * Script runner — executes agent-supplied JS against the generic Sheet API.
 *
 * Each call is a fresh evaluation in a vm context with only the `sheets`
 * library and a captured `console`. No fs, network (except via sheets), or
 * process access.
 *
 * The user code is wrapped in `async function() { ... }` and awaited; the
 * return value plus stdout/stderr come back as the result.
 *
 * dryRun: when true, batchUpdate calls go into a captured array instead of
 * being sent. The runner returns the captured request bodies.
 */

import vm from "vm";
import { sheet as makeSheet, clearCache } from "./sheet.mjs";
import { spreadsheetId as getSpreadsheetId, setDryRunMode } from "./sheets-client.mjs";

export async function exec(code, { dryRun = false, timeoutMs = 30000 } = {}) {
  const stdout = [];
  const stderr = [];
  const captured = { batchUpdates: [] };

  const sheets = {
    sheet: makeSheet,
    spreadsheetId: getSpreadsheetId,
  };

  const sandboxConsole = {
    log:   (...a) => stdout.push(a.map(stringify).join(" ")),
    info:  (...a) => stdout.push(a.map(stringify).join(" ")),
    warn:  (...a) => stderr.push(a.map(stringify).join(" ")),
    error: (...a) => stderr.push(a.map(stringify).join(" ")),
  };

  if (dryRun) {
    setDryRunMode(captured.batchUpdates);
    clearCache();
  }

  const ctx = vm.createContext({
    sheets,
    console: sandboxConsole,
    Promise, Date, Math, JSON, Object, Array, String, Number, Boolean,
    setTimeout, setImmediate, clearTimeout, clearImmediate,
  });

  let result;
  let error = null;
  try {
    const wrapped = `(async () => {\n${code}\n})()`;
    const promise = vm.runInContext(wrapped, ctx, { timeout: timeoutMs, displayErrors: true });
    result = await Promise.race([
      promise,
      new Promise((_, rej) => setTimeout(() => rej(new Error(`Script timeout after ${timeoutMs}ms`)), timeoutMs)),
    ]);
  } catch (e) {
    error = { message: e.message, stack: e.stack, name: e.name ?? "Error" };
  } finally {
    if (dryRun) setDryRunMode(null);
    clearCache();
  }

  return {
    ok: !error,
    result,
    stdout: stdout.join("\n"),
    stderr: stderr.join("\n"),
    ...(error && { error }),
    ...(dryRun && { dryRun: true, planned: captured.batchUpdates }),
  };
}

function stringify(v) {
  if (typeof v === "string") return v;
  if (v instanceof Error) return `${v.name}: ${v.message}`;
  try { return JSON.stringify(v); } catch { return String(v); }
}
