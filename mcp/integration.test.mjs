/**
 * No-network integration tests for the dry-run wiring: runner → sheets-api
 * factory → tagged capture in the low-level client. These exercise the layers
 * the pure unit tests don't, and intentionally avoid any path that would call
 * the real Google client (under dry-run, batchUpdate/valuesUpdate never touch
 * the network, and addSheet does not resolve a sheet name).
 */
import { test } from "node:test";
import assert from "node:assert/strict";
import { exec } from "./runner.mjs";
import { setDryRunMode, valuesUpdate, batchUpdate } from "./sheets-client.mjs";

test("runner dry-run — addSheet is captured, not committed, and returns a placeholder", async () => {
  const out = await exec(
    "FAKE_ID",
    `return await sheets.addSheet("обязательства", { tabColor: "#34a853", frozenRows: 1 });`,
    { dryRun: true },
  );
  assert.equal(out.ok, true);
  assert.deepEqual(out.result, { sheetId: null, title: "обязательства", dryRun: true });

  // back-compat `planned` holds the batchUpdate request body…
  assert.equal(out.planned.length, 1);
  const req = out.planned[0].requests[0];
  assert.equal(req.addSheet.properties.title, "обязательства");
  assert.equal(req.addSheet.properties.gridProperties.frozenRowCount, 1);

  // …and the full tagged log agrees.
  assert.deepEqual(out.plannedOps.map(e => e.kind), ["batchUpdate"]);
});

test("runner — a thrown script error is returned as ok:false, not a crash", async () => {
  const out = await exec("FAKE_ID", `throw new Error("boom");`, { dryRun: true });
  assert.equal(out.ok, false);
  assert.equal(out.error.message, "boom");
});

test("valuesUpdate honors dry-run — captured + synthetic, never sent (footgun #1)", async () => {
  const cap = [];
  setDryRunMode(cap);
  try {
    const res = await valuesUpdate("ID", "Sheet1!A1:B1", [["x", "y"]]);
    await batchUpdate("ID", [{ addSheet: { properties: { title: "t" } } }]);
    assert.deepEqual(res, { updatedRange: "Sheet1!A1:B1", updatedRows: 1, updatedColumns: 2, updatedCells: 2 });
    assert.deepEqual(cap.map(e => e.kind), ["valuesUpdate", "batchUpdate"]);
    assert.equal(cap[0].valueInputOption, "USER_ENTERED");
  } finally {
    setDryRunMode(null);
  }
});
