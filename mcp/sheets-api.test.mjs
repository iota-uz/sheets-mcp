/**
 * Hermetic session tests — a fake client that also SIMULATES rename/delete
 * mutations on its in-memory sheet list, so we can assert the handle registry
 * keeps live handles coherent (issue #7 limitation #2) with no network.
 */
import { test } from "node:test";
import assert from "node:assert/strict";
import { makeSheetsApi } from "./sheets-api.mjs";

function fakeClient({ sheets = [] } = {}) {
  const calls = [];
  const sheetData = sheets.map(s => ({
    properties: { sheetId: s.sheetId, title: s.title, gridProperties: { rowCount: 1000, columnCount: 26 } },
    headers: s.headers ?? [],
  }));
  const initShape = () => sheetData.map(s => ({
    properties: s.properties,
    data: [
      { startColumn: 0, rowData: [{ values: s.headers.map(h => ({ formattedValue: h })) }] },
      { startColumn: 0, rowData: [{ values: [] }] },
    ],
  }));
  return {
    calls,
    async spreadsheetsGet(id) { return { spreadsheetId: id, sheets: initShape() }; },
    async getSpreadsheet(id) { return { spreadsheetId: id, sheets: sheetData.map(s => ({ properties: s.properties })) }; },
    async valuesGet(id, range) { calls.push({ method: "valuesGet", range }); return []; },
    async valuesBatchGet() { return []; },
    async developerMetadataSearch() { return []; },
    async batchUpdate(id, requests) {
      calls.push({ method: "batchUpdate", requests });
      for (const r of requests) {
        if (r.updateSheetProperties?.properties?.title != null) {
          const { sheetId, title } = r.updateSheetProperties.properties;
          const e = sheetData.find(s => s.properties.sheetId === sheetId);
          if (e) e.properties.title = title;
        }
        if (r.deleteSheet) {
          const i = sheetData.findIndex(s => s.properties.sheetId === r.deleteSheet.sheetId);
          if (i >= 0) sheetData.splice(i, 1);
        }
      }
      return { replies: requests.map(() => ({})) };
    },
    async valuesUpdate(id, range) { return { updatedRange: range }; },
  };
}

test("renameSheet updates a held handle, re-keys it, and old name 404s", async () => {
  const client = fakeClient({ sheets: [{ sheetId: 1, title: "A", headers: ["X"] }] });
  const api = makeSheetsApi("SS", client);
  const h = await api.sheet("A");

  await api.renameSheet("A", "B");
  assert.equal(h.name, "B");

  // the held handle now reads the renamed tab
  client.calls.length = 0;
  await h.find({});
  const vg = client.calls.find(c => c.method === "valuesGet");
  assert.ok(vg.range.startsWith("B!"), `expected a B! range, got ${vg.range}`);

  // registry: new name returns the SAME handle; old name no longer resolves
  assert.equal(await api.sheet("B"), h);
  await assert.rejects(api.sheet("A"), /not found/);
});

test("deleteSheet dead-guards the held handle and drops it from the registry", async () => {
  const client = fakeClient({ sheets: [{ sheetId: 2, title: "Tab", headers: ["X"] }] });
  const api = makeSheetsApi("SS", client);
  const h = await api.sheet("Tab");

  await api.deleteSheet("Tab");
  assert.throws(() => h.describe(), /was deleted/);
  await assert.rejects(api.sheet("Tab"), /not found/);
});

test("rename updates every cached handle for the tab (multiple headerRows)", async () => {
  const client = fakeClient({ sheets: [{ sheetId: 3, title: "M", headers: ["X"] }] });
  const api = makeSheetsApi("SS", client);
  const h1 = await api.sheet("M", { headerRow: 1 });
  const h3 = await api.sheet("M", { headerRow: 3 });

  await api.renameSheet("M", "N");
  assert.equal(h1.name, "N");
  assert.equal(h3.name, "N");
  assert.equal(await api.sheet("N", { headerRow: 3 }), h3);
});

test("sheet(name) caches per session — same handle on repeat", async () => {
  const client = fakeClient({ sheets: [{ sheetId: 1, title: "A", headers: ["X"] }] });
  const api = makeSheetsApi("SS", client);
  assert.equal(await api.sheet("A"), await api.sheet("A"));
});
