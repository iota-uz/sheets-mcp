import { test } from "node:test";
import assert from "node:assert/strict";
import { colLetter, colToIdx, a1ToGridRange, gridRangeToA1 } from "./a1.mjs";

test("gridRangeToA1 — bounded range round-trips with a1ToGridRange", () => {
  const range = { sheetId: 4, startRowIndex: 0, endRowIndex: 6, startColumnIndex: 0, endColumnIndex: 9 };
  assert.equal(gridRangeToA1("Data", range), "Data!A1:I6");
  // quotes titles needing them, doubling internal quotes
  assert.equal(gridRangeToA1("обязательства", range), "'обязательства'!A1:I6");
  assert.equal(gridRangeToA1("A b", range), "'A b'!A1:I6");
  assert.equal(gridRangeToA1("O'Brien", range), "'O''Brien'!A1:I6");
  // round-trip: A1 → GridRange drops sheetId, so compare the index fields
  const back = a1ToGridRange(4, gridRangeToA1("Data", range));
  assert.deepEqual(back, range);
});

test("gridRangeToA1 — rejects an unbounded range", () => {
  assert.throws(() => gridRangeToA1("Data", { sheetId: 4, startColumnIndex: 0, endColumnIndex: 2 }), /fully-bounded/);
});

test("colLetter — boundaries", () => {
  assert.equal(colLetter(0), "A");
  assert.equal(colLetter(25), "Z");
  assert.equal(colLetter(26), "AA");
  assert.equal(colLetter(27), "AB");
  assert.equal(colLetter(51), "AZ");
  assert.equal(colLetter(52), "BA");
  assert.equal(colLetter(701), "ZZ");   // last 2-letter — the old impl broke past here
  assert.equal(colLetter(702), "AAA");
  assert.equal(colLetter(728), "ABA");
});

test("colLetter — rejects bad input", () => {
  assert.throws(() => colLetter(-1));
  assert.throws(() => colLetter(1.5));
  assert.throws(() => colLetter("A"));
});

test("colToIdx — boundaries + case-insensitive", () => {
  assert.equal(colToIdx("A"), 0);
  assert.equal(colToIdx("Z"), 25);
  assert.equal(colToIdx("AA"), 26);
  assert.equal(colToIdx("ZZ"), 701);
  assert.equal(colToIdx("AAA"), 702);
  assert.equal(colToIdx("aba"), 728);
  assert.throws(() => colToIdx("A1"));
  assert.throws(() => colToIdx(""));
});

test("colLetter ↔ colToIdx round-trip", () => {
  for (const i of [0, 1, 25, 26, 27, 51, 52, 100, 701, 702, 728, 1000, 16383]) {
    assert.equal(colToIdx(colLetter(i)), i, `round-trip failed at ${i}`);
  }
});

test("a1ToGridRange — single cell", () => {
  assert.deepEqual(a1ToGridRange(7, "A1"), {
    sheetId: 7, startColumnIndex: 0, endColumnIndex: 1, startRowIndex: 0, endRowIndex: 1,
  });
});

test("a1ToGridRange — bounded rect", () => {
  assert.deepEqual(a1ToGridRange(0, "B2:D10"), {
    sheetId: 0, startColumnIndex: 1, startRowIndex: 1, endColumnIndex: 4, endRowIndex: 10,
  });
});

test("a1ToGridRange — full columns (no row bounds)", () => {
  assert.deepEqual(a1ToGridRange(3, "C:C"), { sheetId: 3, startColumnIndex: 2, endColumnIndex: 3 });
  assert.deepEqual(a1ToGridRange(3, "C:E"), { sheetId: 3, startColumnIndex: 2, endColumnIndex: 5 });
});

test("a1ToGridRange — full rows (no column bounds)", () => {
  assert.deepEqual(a1ToGridRange(3, "2:2"), { sheetId: 3, startRowIndex: 1, endRowIndex: 2 });
  assert.deepEqual(a1ToGridRange(3, "3:7"), { sheetId: 3, startRowIndex: 2, endRowIndex: 7 });
});

test("a1ToGridRange — open-ended", () => {
  assert.deepEqual(a1ToGridRange(1, "A1:C"), {
    sheetId: 1, startColumnIndex: 0, startRowIndex: 0, endColumnIndex: 3,
  });
  assert.deepEqual(a1ToGridRange(1, "A:C5"), {
    sheetId: 1, startColumnIndex: 0, endColumnIndex: 3, endRowIndex: 5,
  });
});

test("a1ToGridRange — strips sheet prefix", () => {
  assert.deepEqual(a1ToGridRange(9, "Sheet1!A1:B2"), {
    sheetId: 9, startColumnIndex: 0, startRowIndex: 0, endColumnIndex: 2, endRowIndex: 2,
  });
  assert.deepEqual(a1ToGridRange(9, "'My Sheet'!A1"), {
    sheetId: 9, startColumnIndex: 0, startRowIndex: 0, endColumnIndex: 1, endRowIndex: 1,
  });
  // escaped quote inside a quoted sheet name
  assert.deepEqual(a1ToGridRange(9, "'O''Brien'!C3"), {
    sheetId: 9, startColumnIndex: 2, startRowIndex: 2, endColumnIndex: 3, endRowIndex: 3,
  });
});

test("a1ToGridRange — malformed throws", () => {
  assert.throws(() => a1ToGridRange(0, ""));
  assert.throws(() => a1ToGridRange(0, "   "));
  assert.throws(() => a1ToGridRange(0, "!!"));
  assert.throws(() => a1ToGridRange(0, "A0"));       // row must be ≥ 1
  assert.throws(() => a1ToGridRange(0, "A1:B2:C3"));  // more than one ":"
});

test("a1ToGridRange — normalizes reversed bounded ranges", () => {
  assert.deepEqual(a1ToGridRange(0, "D10:B2"), a1ToGridRange(0, "B2:D10"));
  assert.deepEqual(a1ToGridRange(0, "D2:B5"), {
    sheetId: 0, startColumnIndex: 1, endColumnIndex: 4, startRowIndex: 1, endRowIndex: 5,
  });
});

