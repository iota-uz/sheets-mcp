/**
 * Best-effort, NON-EXHAUSTIVE structural linter for Sheets v4 batchUpdate
 * requests. Pure: no I/O, no schema fetch. It catches a handful of shapes Google
 * rejects at commit time (missing required fields), so a dryRun can surface them
 * early instead of "passing" and then failing on the live call (issue #12 B1).
 *
 * Absence of warnings does NOT imply the live call will succeed — Google is the
 * only authority, and many failures (e.g. unmerging a partial merge) need state
 * the linter can't see. Rules are intentionally shallow + additive: an unmodeled
 * request kind never produces a warning.
 */

export const LINT_CODES = {
  MISSING_RANGE: "missing-range",
  MISSING_FIELDS: "missing-fields",
  MISSING_TARGET: "missing-target",
  MERGE_NO_TYPE: "merge-missing-type",
};

/**
 * Lint a single Sheets v4 Request object.
 * @param {object} request one entry of a batchUpdate `requests` array
 * @returns {Array<{code,message}>} zero or more findings (no positional info)
 */
export function lintRequest(request) {
  if (!request || typeof request !== "object") return [];
  const kind = Object.keys(request)[0];
  if (!kind) return [];
  const body = request[kind];
  const out = [];

  switch (kind) {
    case "unmergeCells":
    case "mergeCells":
      if (!isObject(body?.range)) {
        out.push({ code: LINT_CODES.MISSING_RANGE, message: `${kind}: requires a \`range\`` });
      }
      if (kind === "mergeCells" && body?.mergeType == null) {
        out.push({ code: LINT_CODES.MERGE_NO_TYPE, message: "mergeCells: requires a `mergeType` (MERGE_ALL|MERGE_COLUMNS|MERGE_ROWS)" });
      }
      break;

    case "updateCells":
      if (!body?.fields) {
        out.push({ code: LINT_CODES.MISSING_FIELDS, message: "updateCells: requires a non-empty `fields` mask" });
      }
      if (body?.start == null && body?.range == null) {
        out.push({ code: LINT_CODES.MISSING_TARGET, message: "updateCells: requires `start` or `range`" });
      }
      break;

    case "repeatCell":
      if (!isObject(body?.range)) {
        out.push({ code: LINT_CODES.MISSING_RANGE, message: "repeatCell: requires a `range`" });
      }
      if (!body?.fields) {
        out.push({ code: LINT_CODES.MISSING_FIELDS, message: "repeatCell: requires a non-empty `fields` mask" });
      }
      break;

    default:
      break; // unmodeled kind → no warning (additive-only contract)
  }
  return out;
}

/**
 * Lint a captured plannedOps log (the runner's dry-run capture). Walks each
 * { kind:"batchUpdate", requests:[…] } entry; valuesUpdate entries carry no
 * request shape and are skipped. Findings are tagged with their position.
 * @param {Array<object>} plannedOps
 * @returns {Array<{opIndex,requestIndex,kind,code,message}>}
 */
export function lintPlannedOps(plannedOps) {
  if (!Array.isArray(plannedOps)) return [];
  const findings = [];
  plannedOps.forEach((op, opIndex) => {
    if (op?.kind !== "batchUpdate" || !Array.isArray(op.requests)) return;
    op.requests.forEach((req, requestIndex) => {
      const kind = isObject(req) ? Object.keys(req)[0] ?? null : null;
      for (const f of lintRequest(req)) {
        findings.push({ opIndex, requestIndex, kind, ...f });
      }
    });
  });
  return findings;
}

function isObject(v) {
  return v != null && typeof v === "object";
}
