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
import { makeClient } from "./sheets-client.mjs";
import { makeSheetsApi } from "./sheets-api.mjs";

export async function exec(spreadsheetId, code, { dryRun = false, timeoutMs = 30000 } = {}) {
  if (!spreadsheetId) throw new Error("exec requires a spreadsheetId");

  const stdout = [];
  const stderr = [];
  // Per-exec dry-run capture: an ordered log of tagged planned ops
  // ({ kind: "batchUpdate" | "valuesUpdate", ... }), or null when committing.
  // Held by this exec's client only, so concurrent execs never interfere.
  const capture = dryRun ? [] : null;
  const client = makeClient({ capture });
  const sheets = makeSheetsApi(spreadsheetId, client);

  const sandboxConsole = {
    log:   (...a) => stdout.push(a.map(stringify).join(" ")),
    info:  (...a) => stdout.push(a.map(stringify).join(" ")),
    warn:  (...a) => stderr.push(a.map(stringify).join(" ")),
    error: (...a) => stderr.push(a.map(stringify).join(" ")),
  };

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
  }

  return {
    ok: !error,
    result,
    stdout: stdout.join("\n"),
    stderr: stderr.join("\n"),
    ...(error && { error }),
    ...(dryRun && {
      dryRun: true,
      // `planned` = the batchUpdate request bodies; `plannedOps` = the full
      // ordered log of every intended mutation (incl. writeRange value writes).
      planned: capture
        .filter(e => e.kind === "batchUpdate")
        .map(({ spreadsheetId: id, requests }) => ({ spreadsheetId: id, requests })),
      plannedOps: capture,
    }),
  };
}

function stringify(v) {
  if (typeof v === "string") return v;
  if (v instanceof Error) return `${v.name}: ${v.message}`;
  try { return JSON.stringify(v); } catch { return String(v); }
}
