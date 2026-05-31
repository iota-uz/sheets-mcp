/**
 * Hermetic Sheet tests — a fake client (canned metadata, records calls) is
 * injected into createSheet, so no googleapis / network is involved.
 */
import { test } from "node:test";
import assert from "node:assert/strict";
import { createSheet, indexHeaders, normHeader } from "./sheet.mjs";

function fakeClient({ sheets = [], values = [], metadata = [] } = {}) {
  const calls = [];
  const sheetData = sheets.map(s => ({
    properties: { sheetId: s.sheetId, title: s.title, gridProperties: { rowCount: 1000, columnCount: 26 } },
    data: [
      { startColumn: 0, rowData: [{ values: (s.headers ?? []).map(h => ({ formattedValue: h })) }] },
      { startColumn: 0, rowData: [{ values: [] }] },
    ],
  }));
  return {
    calls,
    async spreadsheetsGet(id) { calls.push({ method: "spreadsheetsGet" }); return { spreadsheetId: id, sheets: sheetData }; },
    async getSpreadsheet(id) { calls.push({ method: "getSpreadsheet" }); return { spreadsheetId: id, sheets: sheetData.map(s => ({ properties: s.properties })) }; },
    async valuesGet(id, range) { calls.push({ method: "valuesGet", range }); return values; },
    async valuesBatchGet() { return []; },
    async developerMetadataSearch() { calls.push({ method: "developerMetadataSearch" }); return metadata; },
    async batchUpdate(id, requests) { calls.push({ method: "batchUpdate", requests }); return { replies: requests.map(() => ({})) }; },
    async valuesUpdate(id, range) { calls.push({ method: "valuesUpdate", range }); return { updatedRange: range }; },
  };
}

// ── indexHeaders (pure) ──────────────────────────────────────────────────────

test("indexHeaders — unique headers map cleanly", () => {
  const { headerToIdx, ambiguous } = indexHeaders(["A", "B", "C"]);
  assert.deepEqual([...headerToIdx], [["a", 0], ["b", 1], ["c", 2]]);
  assert.equal(ambiguous.size, 0);
});

test("indexHeaders — collisions go to ambiguous, not headerToIdx", () => {
  const { headerToIdx, ambiguous } = indexHeaders(["Name", "name", "Age"]);
  assert.deepEqual([...headerToIdx], [["age", 2]]);           // only the unique one
  assert.deepEqual(ambiguous.get(normHeader("name")), [0, 1]);
});

test("indexHeaders — empty/whitespace headers are skipped (not ambiguous)", () => {
  const { headerToIdx, ambiguous } = indexHeaders(["", "  ", "X", ""]);
  assert.deepEqual([...headerToIdx], [["x", 2]]);
  assert.equal(ambiguous.size, 0);
});

test("indexHeaders — ё/е normalize to the same key → ambiguous", () => {
  const { ambiguous } = indexHeaders(["ёж", "еж"]);
  assert.deepEqual(ambiguous.get("еж"), [0, 1]);
});

// ── ambiguous header behavior on a live handle ───────────────────────────────

test("ambiguous header throws on use, but describe() still lists all", async () => {
  const client = fakeClient({ sheets: [{ sheetId: 1, title: "T", headers: ["Name", "name", "Age"] }] });
  const s = await createSheet("SS", "T", {}, client);

  assert.equal(s.describe().headers.length, 3);           // all listed, dups included
  assert.equal(s._col("Age"), 2);                          // non-colliding header works
  assert.throws(() => s._col("name"), /Ambiguous header/); // collision throws with columns
  await assert.rejects(s.insert({ "name": "x" }), /Ambiguous header/);
  await assert.doesNotReject(s.insert({ "Age": 5 }));      // writing a unique header is fine
});

// ── dead-handle guard (Proxy) ────────────────────────────────────────────────

const PUBLIC_METHODS = [
  "describe", "find", "insert", "insertMany", "update", "delete", "format", "formatRange",
  "merge", "unmerge", "setBorders", "addConditionalFormat", "freeze", "resizeColumns",
  "insertColumns", "deleteColumns", "sort", "setFilter", "findReplace", "setNote",
  "setValidation", "readFormatting", "readRange", "writeRange",
];

test("a deleted handle throws on every public method", async () => {
  const client = fakeClient({ sheets: [{ sheetId: 1, title: "T", headers: ["X"] }] });
  const s = await createSheet("SS", "T", {}, client);
  s._markDeleted();
  for (const m of PUBLIC_METHODS) {
    assert.throws(() => s[m](), /was deleted/, `method ${m} should throw after delete`);
  }
});

// ── idempotency keyed by sheetId (no title prefix) ───────────────────────────

test("insert writes a developer-metadata token whose value is the raw key", async () => {
  const client = fakeClient({ sheets: [{ sheetId: 7, title: "T", headers: ["Name"] }] });
  const s = await createSheet("SS", "T", {}, client);
  await s.insert({ "Name": "x" }, { idempotencyKey: "k1" });

  const bu = client.calls.filter(c => c.method === "batchUpdate").at(-1);
  const meta = bu.requests.find(r => r.createDeveloperMetadata);
  assert.equal(meta.createDeveloperMetadata.developerMetadata.metadataValue, "k1");
});

test("_loadIdempotency isolates tokens by sheetId (survives rename, no cross-tab)", async () => {
  const metadata = [
    { developerMetadata: { metadataValue: "k1", location: { dimensionRange: { sheetId: 7, startIndex: 5 } } } },
    { developerMetadata: { metadataValue: "k1", location: { dimensionRange: { sheetId: 8, startIndex: 9 } } } },
  ];
  const client = fakeClient({ sheets: [{ sheetId: 7, title: "T", headers: ["Name"] }], metadata });
  const s = await createSheet("SS", "T", {}, client);
  const map = await s._loadIdempotency();
  assert.equal(map.size, 1);
  assert.equal(map.get("k1"), 6);   // startIndex 5 → row 6, only the sheetId-7 token
});

// ── insert bookkeeping advances across calls (unified dry-run path) ──────────

test("two inserts plan distinct rows (nextRow cache advances unconditionally)", async () => {
  const client = fakeClient({ sheets: [{ sheetId: 1, title: "T", headers: ["X"] }], values: [] });
  const s = await createSheet("SS", "T", {}, client);
  const r1 = await s.insert({ "X": 1 });
  const r2 = await s.insert({ "X": 2 });
  assert.equal(r1.row, 2);
  assert.equal(r2.row, 3);

  const inserts = client.calls
    .filter(c => c.method === "batchUpdate")
    .map(c => c.requests.find(r => r.insertDimension).insertDimension.range.startIndex);
  assert.deepEqual(inserts, [1, 2]);
});
