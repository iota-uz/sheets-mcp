/**
 * Pure result-size guard for sheets_exec output (issue #12 I1).
 *
 * The MCP serializes the exec result to JSON; a very large result blows the
 * client's token budget and is dumped to a file, interrupting the flow.
 * guardResultSize caps the SERIALIZED size and, when over, replaces the value
 * with a small preview carrying a hint on how to get the data in bounded chunks.
 * Truncation replaces the value (keeping valid JSON) rather than clipping the
 * serialized string.
 */

const DEFAULT_MAX_CHARS = 50000;

/**
 * @param {*} result the exec result value
 * @param {{maxChars?:number, full?:boolean}} opts maxChars 0 disables the cap
 * @returns {{value:*, truncated:boolean, originalChars?:number, hint?:string}}
 */
export function guardResultSize(result, { maxChars = DEFAULT_MAX_CHARS, full = false } = {}) {
  if (full || maxChars === 0) return { value: result, truncated: false };

  let serialized;
  try {
    serialized = JSON.stringify(result);
  } catch {
    // Non-serializable (circular, BigInt, …): leave it for the caller's own
    // stringify to handle/throw — not our concern to truncate.
    return { value: result, truncated: false };
  }
  if (serialized == null || serialized.length <= maxChars) {
    return { value: result, truncated: false };
  }

  return {
    value: previewOf(result, maxChars),
    truncated: true,
    originalChars: serialized.length,
    hint:
      `result was ${serialized.length} chars, over the ${maxChars}-char cap. ` +
      `Re-run sheets_exec with maxBytes:0 for the full result, pass head:N to keep ` +
      `the first N rows, or page reads with sheet.readRange(range, { limit, offset }).`,
  };
}

/**
 * Thread the sheets_exec `maxBytes`/`head` options through a runner result.
 * Applies `head` (keep first N rows of an array result) before the size guard.
 * Returns the original object untouched when nothing was capped.
 */
export function capResult(out, { maxBytes, head } = {}) {
  let result = out?.result;
  let headApplied = null;

  if (typeof head === "number" && head >= 0 && Array.isArray(result) && result.length > head) {
    headApplied = { kept: head, omitted: result.length - head };
    result = result.slice(0, head);
  }

  const guard = guardResultSize(result, {
    maxChars: typeof maxBytes === "number" ? maxBytes : undefined,
    full: maxBytes === 0,
  });

  if (!guard.truncated && !headApplied) return out;
  return {
    ...out,
    result: guard.value,
    ...(guard.truncated && { truncated: true, resultChars: guard.originalChars, truncationHint: guard.hint }),
    ...(headApplied && { head: headApplied }),
  };
}

/** Build a small, still-serializable preview of an oversized result. */
function previewOf(result, maxChars) {
  if (Array.isArray(result)) {
    // Keep as many leading elements as fit under the cap (binary search on the
    // serialized prefix length).
    let lo = 0, hi = result.length, keep = 0;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      if (serializedLen(result.slice(0, mid)) <= maxChars) { keep = mid; lo = mid + 1; }
      else hi = mid - 1;
    }
    return { _truncated: true, kept: keep, omitted: result.length - keep, preview: result.slice(0, keep) };
  }
  // Non-array: clip the string form.
  let s;
  try { s = JSON.stringify(result); } catch { s = String(result); }
  return { _truncated: true, preview: s.slice(0, maxChars) };
}

function serializedLen(v) {
  try { return JSON.stringify(v)?.length ?? Infinity; }
  catch { return Infinity; }
}
