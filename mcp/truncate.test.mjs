import { test } from "node:test";
import assert from "node:assert/strict";
import { guardResultSize, capResult } from "./truncate.mjs";

test("guardResultSize — under cap is untouched", () => {
  const r = guardResultSize({ a: 1 }, { maxChars: 1000 });
  assert.deepEqual(r, { value: { a: 1 }, truncated: false });
});

test("guardResultSize — over cap truncates with a hint", () => {
  const big = Array.from({ length: 500 }, (_, i) => ({ i, pad: "x".repeat(50) }));
  const r = guardResultSize(big, { maxChars: 500 });
  assert.equal(r.truncated, true);
  assert.ok(r.originalChars > 500);
  assert.ok(/maxBytes:0/.test(r.hint));
  assert.equal(r.value._truncated, true);
  assert.ok(r.value.preview.length < big.length);
  assert.equal(r.value.kept + r.value.omitted, big.length);
});

test("guardResultSize — maxChars:0 / full never truncates", () => {
  const big = "y".repeat(100000);
  assert.equal(guardResultSize(big, { maxChars: 0 }).truncated, false);
  assert.equal(guardResultSize(big, { full: true }).truncated, false);
});

test("capResult — head slices an array result before capping", () => {
  const out = { ok: true, result: [1, 2, 3, 4, 5] };
  const capped = capResult(out, { head: 2 });
  assert.deepEqual(capped.result, [1, 2]);
  assert.deepEqual(capped.head, { kept: 2, omitted: 3 });
});

test("capResult — returns the original object when nothing capped", () => {
  const out = { ok: true, result: [1, 2, 3] };
  assert.equal(capResult(out, {}), out);
});

test("capResult — maxBytes:0 disables the cap", () => {
  const out = { ok: true, result: "z".repeat(80000) };
  assert.equal(capResult(out, { maxBytes: 0 }), out);
});
