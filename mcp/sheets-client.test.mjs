import { test } from "node:test";
import assert from "node:assert/strict";
import { normalizeValueRanges, enrichBatchError } from "./sheets-client.mjs";

// ── B5: empty valueRanges always carry values:[] ─────────────────────────────

test("normalizeValueRanges — empty range gets values:[]", () => {
  const out = normalizeValueRanges([
    { range: "A1:A1" },                              // empty — no values key
    { range: "B1:B2", values: [["x"], ["y"]] },      // populated
  ]);
  assert.deepEqual(out[0], { range: "A1:A1", values: [] });
  assert.deepEqual(out[1], { range: "B1:B2", values: [["x"], ["y"]] });
});

test("normalizeValueRanges — undefined/null input → []", () => {
  assert.deepEqual(normalizeValueRanges(undefined), []);
  assert.deepEqual(normalizeValueRanges(null), []);
});

// ── I4: batchUpdate errors localized to op + A1 ──────────────────────────────

test("enrichBatchError — maps requests[N] to op kind + A1 range", () => {
  const requests = Array.from({ length: 40 }, () => ({ noop: {} }));
  requests[39] = { setDataValidation: { range: { sheetId: 2, startRowIndex: 1, endRowIndex: 100, startColumnIndex: 1, endColumnIndex: 2 } } };
  const err = enrichBatchError(new Error("Invalid requests[39].setDataValidation: bad rule"), requests);

  assert.equal(err.requestIndex, 39);
  assert.equal(err.opKind, "setDataValidation");
  assert.equal(err.range, "B2:B100");
  assert.equal(err.sheetId, 2);
  assert.match(err.message, /requests\[39\] setDataValidation \(range B2:B100\)/);
});

test("enrichBatchError — unparseable message returned unchanged", () => {
  const orig = new Error("network blew up");
  assert.equal(enrichBatchError(orig, []), orig);
});

test("enrichBatchError — op without a range omits the range clause", () => {
  const requests = [{ findReplace: { find: "x", replacement: "y" } }];
  const err = enrichBatchError(new Error("Invalid requests[0].findReplace: nope"), requests);
  assert.equal(err.opKind, "findReplace");
  assert.equal(err.range, undefined);
  assert.match(err.message, /requests\[0\] findReplace:/);
});
