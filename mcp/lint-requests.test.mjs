import { test } from "node:test";
import assert from "node:assert/strict";
import { lintRequest, lintPlannedOps, LINT_CODES } from "./lint-requests.mjs";
import { buildUpdateCells, buildUnmerge } from "./requests.mjs";

const RANGE = { sheetId: 0, startRowIndex: 0, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: 2 };

test("lintRequest — real builder output lints clean", () => {
  assert.deepEqual(lintRequest(buildUnmerge(RANGE)[0]), []);
  assert.deepEqual(lintRequest(buildUpdateCells(RANGE, [["a", "b"]])[0]), []);
});

test("lintRequest — flags an unmergeCells with no range", () => {
  const found = lintRequest({ unmergeCells: {} });
  assert.equal(found.length, 1);
  assert.equal(found[0].code, LINT_CODES.MISSING_RANGE);
});

test("lintRequest — flags updateCells missing fields and target", () => {
  const codes = lintRequest({ updateCells: { rows: [], fields: "" } }).map(f => f.code);
  assert.ok(codes.includes(LINT_CODES.MISSING_FIELDS));
  assert.ok(codes.includes(LINT_CODES.MISSING_TARGET));
});

test("lintRequest — mergeCells needs a mergeType", () => {
  const codes = lintRequest({ mergeCells: { range: RANGE } }).map(f => f.code);
  assert.deepEqual(codes, [LINT_CODES.MERGE_NO_TYPE]);
});

test("lintRequest — unmodeled kinds never warn", () => {
  assert.deepEqual(lintRequest({ addSheet: { properties: {} } }), []);
  assert.deepEqual(lintRequest(null), []);
});

test("lintPlannedOps — tags opIndex/requestIndex and skips valuesUpdate", () => {
  const capture = [
    { kind: "valuesUpdate", range: "A1", values: [["x"]] },
    { kind: "batchUpdate", requests: [buildUnmerge(RANGE)[0], { unmergeCells: {} }] },
  ];
  const findings = lintPlannedOps(capture);
  assert.equal(findings.length, 1);
  assert.deepEqual(
    { opIndex: findings[0].opIndex, requestIndex: findings[0].requestIndex, kind: findings[0].kind, code: findings[0].code },
    { opIndex: 1, requestIndex: 1, kind: "unmergeCells", code: LINT_CODES.MISSING_RANGE },
  );
});

test("lintPlannedOps — non-array input is safe", () => {
  assert.deepEqual(lintPlannedOps(null), []);
});
