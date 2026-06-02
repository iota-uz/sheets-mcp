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
    tables: s.tables ?? [],
  }));
  const initShape = () => sheetData.map(s => ({
    properties: s.properties,
    data: [
      { startColumn: 0, rowData: [{ values: s.headers.map(h => ({ formattedValue: h })) }] },
      { startColumn: 0, rowData: [{ values: [] }] },
    ],
  }));
  const replyFor = (r) => {
    if (r.addTable) return { addTable: { table: { tableId: "tbl_x", name: r.addTable.table.name } } };
    if (r.addSheet) return { addSheet: { properties: { sheetId: 999, title: r.addSheet.properties.title, index: 0 } } };
    return {};
  };
  return {
    calls,
    async spreadsheetsGet(id) { return { spreadsheetId: id, sheets: initShape() }; },
    async getSpreadsheet(id) { return { spreadsheetId: id, sheets: sheetData.map(s => ({ properties: s.properties, tables: s.tables ?? [] })) }; },
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
      return { replies: requests.map(replyFor) };
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

// ── issue #10 #3: typed Table API ────────────────────────────────────────────

test("addTable parses a prefixed range and compiles friendly columns", async () => {
  const client = fakeClient({ sheets: [{ sheetId: 4, title: "Data", headers: ["A", "B", "C"] }] });
  const api = makeSheetsApi("SS", client);
  const res = await api.addTable("Об", "Data!A1:C", { columns: [
    { name: "Контрагент", type: "TEXT" },
    { name: "Категория", type: "DROPDOWN", values: ["Налоги"] },
    { name: "Сумма", type: "NUMBER" },
  ] });
  assert.deepEqual(res, { tableId: "tbl_x", name: "Об" });

  const add = client.calls.at(-1).requests[0].addTable.table;
  assert.equal(add.range.sheetId, 4);
  assert.equal(add.range.startColumnIndex, 0);
  assert.equal(add.range.endColumnIndex, 3);
  assert.equal(add.columnProperties[2].columnType, "DOUBLE");
  assert.ok(add.columnProperties[1].dataValidationRule);
});

test("addTable rejects a range without a sheet prefix", async () => {
  const client = fakeClient({ sheets: [{ sheetId: 4, title: "Data", headers: ["A"] }] });
  const api = makeSheetsApi("SS", client);
  await assert.rejects(api.addTable("T", "A1:C", { columns: [] }), /must include a sheet name/);
});

const boundedTable = () => ([
  { tableId: "t1", name: "Об", range: { sheetId: 4, startRowIndex: 0, endRowIndex: 6, startColumnIndex: 0, endColumnIndex: 2 } },
]);

test("updateTable / deleteTable resolve a table by name", async () => {
  const client = fakeClient({ sheets: [{ sheetId: 4, title: "Data", headers: ["A", "B"], tables: boundedTable() }] });
  const api = makeSheetsApi("SS", client);

  await api.updateTable("Об", { name: "Обязательства" });
  const upd = client.calls.at(-1).requests[0].updateTable;
  assert.equal(upd.fields, "name");
  assert.deepEqual(upd.table, { tableId: "t1", name: "Обязательства" });

  await assert.rejects(api.deleteTable("Nope"), /not found/);
});

// issue #11: deleteTable must NOT destroy cell data. Default path deletes the
// table then restores its cells in one atomic batch; deleteData opts back into
// Google's native (destructive) behavior.
test("deleteTable preserves cell data by default (issue #11 regression)", async () => {
  const client = fakeClient({ sheets: [{ sheetId: 4, title: "Data", headers: ["A", "B"], tables: boundedTable() }] });
  const api = makeSheetsApi("SS", client);

  const res = await api.deleteTable("Об");
  assert.deepEqual(res, { ok: true, tableId: "t1", preserved: true, rows: 1 });

  const { requests } = client.calls.at(-1);
  assert.deepEqual(requests[0].deleteTable, { tableId: "t1" });
  // deleteTable runs FIRST, then cells are restored.
  assert.ok(requests[1].updateCells, "expected an updateCells restore request");
  assert.equal(requests[1].updateCells.fields, "userEnteredValue,userEnteredFormat,note");
  assert.deepEqual(requests[1].updateCells.range, boundedTable()[0].range);
});

test("deleteTable({ deleteData: true }) emits only deleteTable", async () => {
  const client = fakeClient({ sheets: [{ sheetId: 4, title: "Data", headers: ["A", "B"], tables: boundedTable() }] });
  const api = makeSheetsApi("SS", client);

  const res = await api.deleteTable("Об", { deleteData: true });
  assert.deepEqual(res, { ok: true, tableId: "t1", preserved: false });

  const { requests } = client.calls.at(-1);
  assert.equal(requests.length, 1);
  assert.deepEqual(requests[0].deleteTable, { tableId: "t1" });
  assert.ok(!requests.some(r => r.updateCells), "no restore when deleteData:true");
});

test("deleteTable on an empty-body table emits only deleteTable", async () => {
  const client = fakeClient({ sheets: [{ sheetId: 4, title: "Data", headers: ["A", "B"], tables: boundedTable() }] });
  // No cells to restore.
  client.spreadsheetsGet = async (id) => ({ spreadsheetId: id, sheets: [{ data: [{ rowData: [] }] }] });
  const api = makeSheetsApi("SS", client);

  const res = await api.deleteTable("Об");
  assert.deepEqual(res, { ok: true, tableId: "t1", preserved: true, rows: 0 });
  const { requests } = client.calls.at(-1);
  assert.equal(requests.length, 1);
  assert.deepEqual(requests[0].deleteTable, { tableId: "t1" });
});

// issue #12 B4: untable() — drop the Table wrapper, keep the data (reuses the
// #11 preserve path; no setDataValidation emitted).
test("untable preserves cells and emits deleteTable + restore (issue #12 B4)", async () => {
  const client = fakeClient({ sheets: [{ sheetId: 4, title: "Data", headers: ["A", "B"], tables: boundedTable() }] });
  const api = makeSheetsApi("SS", client);

  const res = await api.untable("Об");
  assert.deepEqual(res, { ok: true, tableId: "t1", range: boundedTable()[0].range, rows: 1, untabled: true });

  const { requests } = client.calls.at(-1);
  assert.deepEqual(requests[0].deleteTable, { tableId: "t1" });
  assert.ok(requests[1].updateCells, "expected a restore updateCells");
  assert.equal(requests.some(r => r.setDataValidation), false);
});

test("untable on an empty-body table → rows:0, only deleteTable", async () => {
  const client = fakeClient({ sheets: [{ sheetId: 4, title: "Data", headers: ["A", "B"], tables: boundedTable() }] });
  client.spreadsheetsGet = async (id) => ({ spreadsheetId: id, sheets: [{ data: [{ rowData: [] }] }] });
  const api = makeSheetsApi("SS", client);

  const res = await api.untable("Об");
  assert.deepEqual(res, { ok: true, tableId: "t1", range: boundedTable()[0].range, rows: 0, untabled: true });
  assert.equal(client.calls.at(-1).requests.length, 1);
});

// ── issue #10 #7: ensureSheet / ensureTable (idempotent structural ops) ───────

test("ensureSheet returns the existing tab, or creates when missing", async () => {
  const client = fakeClient({ sheets: [{ sheetId: 1, title: "A", headers: ["X"] }] });
  const api = makeSheetsApi("SS", client);

  const existing = await api.ensureSheet("A");
  assert.deepEqual(existing, { sheetId: 1, title: "A", existed: true });

  const created = await api.ensureSheet("B");
  assert.equal(created.existed, false);
  assert.equal(created.sheetId, 999);  // from the fake addSheet reply
});

test("ensureTable returns the existing table, or creates when missing", async () => {
  const tables = [{ tableId: "t1", name: "Об", range: { sheetId: 4 } }];
  const client = fakeClient({ sheets: [{ sheetId: 4, title: "Data", headers: ["A"], tables }] });
  const api = makeSheetsApi("SS", client);

  const existing = await api.ensureTable("Об", "Data!A1:A");
  assert.deepEqual(existing, { tableId: "t1", name: "Об", existed: true });

  const created = await api.ensureTable("New", "Data!A1:A", { columns: [{ name: "A" }] });
  assert.equal(created.existed, false);
  assert.equal(created.tableId, "tbl_x");
});
