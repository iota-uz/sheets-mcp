/**
 * No-network integration tests for the dry-run wiring: runner → makeClient
 * capture → makeSheetsApi. Every path here avoids the real Google client (under
 * dry-run, batchUpdate/valuesUpdate never touch the network, and addSheet does
 * not resolve a sheet name).
 */
import { test } from "node:test";
import assert from "node:assert/strict";
import { exec } from "./runner.mjs";
import { makeClient } from "./sheets-client.mjs";

test("runner dry-run — addSheet is captured, not committed, returns a placeholder id", async () => {
  const out = await exec(
    "FAKE_ID",
    `return await sheets.addSheet("обязательства", { tabColor: "#34a853", frozenRows: 1 });`,
    { dryRun: true },
  );
  assert.equal(out.ok, true);
  assert.equal(out.result.title, "обязательства");
  assert.equal(out.result.dryRun, true);
  assert.ok(out.result.sheetId < 0, "dry-run addSheet returns a negative placeholder sheetId");

  assert.equal(out.plannedOps.length, 1);
  assert.equal(out.plannedOps[0].kind, "batchUpdate");
  const req = out.plannedOps[0].requests[0];
  assert.equal(req.addSheet.properties.title, "обязательства");
  assert.equal(req.addSheet.properties.gridProperties.frozenRowCount, 1);
});

test("runner dry-run — a create-then-populate chain plans using the placeholder id (issue #10 #8)", async () => {
  const out = await exec(
    "FAKE_ID",
    `const s = await sheets.addSheet("New");
     return await sheets.batchUpdate([{ updateCells: {
       start: { sheetId: s.sheetId, rowIndex: 0, columnIndex: 0 },
       rows: [{ values: [{ userEnteredValue: { stringValue: "x" } }] }],
       fields: "userEnteredValue",
     } }]);`,
    { dryRun: true },
  );
  assert.equal(out.ok, true);
  assert.equal(out.plannedOps.length, 2);
  const dependent = out.plannedOps[1].requests[0].updateCells.start.sheetId;
  assert.ok(dependent < 0, "the follow-up op references the minted placeholder sheetId");
});

test("runner dry-run — raw escape hatch batchUpdate is captured, not committed", async () => {
  const out = await exec(
    "FAKE_ID",
    `return await sheets.batchUpdate([{ addSheet: { properties: { title: "raw" } } }]);`,
    { dryRun: true },
  );
  assert.equal(out.ok, true);
  assert.equal(out.plannedOps.length, 1);
  assert.equal(out.plannedOps[0].requests[0].addSheet.properties.title, "raw");
});

test("concurrent dry-run execs do not cross-contaminate captures (limitation #1)", async () => {
  const [a, b] = await Promise.all([
    exec("ID_A", `await sheets.addSheet("A1"); return await sheets.addSheet("A2");`, { dryRun: true }),
    exec("ID_B", `return await sheets.addSheet("B1");`, { dryRun: true }),
  ]);
  const titles = (o) => o.plannedOps.map(e => e.requests[0].addSheet.properties.title);
  assert.deepEqual(titles(a), ["A1", "A2"]);
  assert.deepEqual(titles(b), ["B1"]);
});

test("runner — a thrown script error is returned as ok:false, not a crash", async () => {
  const out = await exec("FAKE_ID", `throw new Error("boom");`, { dryRun: true });
  assert.equal(out.ok, false);
  assert.equal(out.error.message, "boom");
});

test("makeClient capture — valuesUpdate + batchUpdate recorded + synthetic, never sent", async () => {
  const capture = [];
  const client = makeClient({ capture });
  const res = await client.valuesUpdate("ID", "Sheet1!A1:B1", [["x", "y"]]);
  await client.batchUpdate("ID", [{ addSheet: { properties: { title: "t" } } }]);

  assert.deepEqual(res, { updatedRange: "Sheet1!A1:B1", updatedRows: 1, updatedColumns: 2, updatedCells: 2 });
  assert.deepEqual(capture.map(e => e.kind), ["valuesUpdate", "batchUpdate"]);
  assert.equal(capture[0].valueInputOption, "USER_ENTERED");
});

test("makeClient with no capture builds a real-call client (reads bound)", () => {
  const client = makeClient();
  // Shape check only — these would hit Google if invoked; just assert the surface.
  for (const m of ["batchUpdate", "valuesUpdate", "getSpreadsheet", "spreadsheetsGet", "valuesGet", "valuesBatchGet", "developerMetadataSearch"]) {
    assert.equal(typeof client[m], "function", `client.${m}`);
  }
});
