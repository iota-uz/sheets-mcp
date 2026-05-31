/**
 * Script runner — executes agent-supplied JS against the generic Sheet API.
 *
 * Each call is bound to a specific spreadsheetId. The sandbox exposes
 * `sheets.sheet(name, opts?)` and `sheets.spreadsheetId()` already bound,
 * so agent scripts don't repeat the ID per call.
 *
 * dryRun: when true, batchUpdate calls go into a captured array instead of
 * being sent. The runner returns the captured request bodies.
 */

import vm from "vm";
import { clearCache } from "./sheet.mjs";
import { makeSheetsApi } from "./sheets-api.mjs";
import { setDryRunMode } from "./sheets-client.mjs";

export async function exec(spreadsheetId, code, { dryRun = false, timeoutMs = 30000 } = {}) {
  if (!spreadsheetId) throw new Error("exec requires a spreadsheetId");

  const stdout = [];
  const stderr = [];
  // One ordered log of tagged planned ops ({ kind: "batchUpdate" | "valuesUpdate", ... }).
  const captured = [];

  const sheets = makeSheetsApi(spreadsheetId);

  const sandboxConsole = {
    log:   (...a) => stdout.push(a.map(stringify).join(" ")),
    info:  (...a) => stdout.push(a.map(stringify).join(" ")),
    warn:  (...a) => stderr.push(a.map(stringify).join(" ")),
    error: (...a) => stderr.push(a.map(stringify).join(" ")),
  };

  if (dryRun) {
    setDryRunMode(captured);
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
  let timeoutTimer = null;
  try {
    const wrapped = `(async () => {\n${code}\n})()`;
    const promise = vm.runInContext(wrapped, ctx, { timeout: timeoutMs, displayErrors: true });
    result = await Promise.race([
      promise,
      new Promise((_, rej) => {
        timeoutTimer = setTimeout(() => rej(new Error(`Script timeout after ${timeoutMs}ms`)), timeoutMs);
      }),
    ]);
  } catch (e) {
    error = { message: e.message, stack: e.stack, name: e.name ?? "Error" };
  } finally {
    // Clear the timeout timer so a fast script doesn't leave it dangling and
    // hold the event loop open for the full timeoutMs.
    if (timeoutTimer) clearTimeout(timeoutTimer);
    if (dryRun) setDryRunMode(null);
    clearCache();
  }

  return {
    ok: !error,
    result,
    stdout: stdout.join("\n"),
    stderr: stderr.join("\n"),
    ...(error && { error }),
    ...(dryRun && {
      dryRun: true,
      // Back-compat: `planned` stays the batchUpdate request bodies, same shape as before.
      planned: captured
        .filter(e => e.kind === "batchUpdate")
        .map(({ spreadsheetId: id, requests }) => ({ spreadsheetId: id, requests })),
      // Full, honest log of every intended mutation in order (incl. writeRange value writes).
      plannedOps: captured,
    }),
  };
}

function stringify(v) {
  if (typeof v === "string") return v;
  if (v instanceof Error) return `${v.name}: ${v.message}`;
  try { return JSON.stringify(v); } catch { return String(v); }
}
