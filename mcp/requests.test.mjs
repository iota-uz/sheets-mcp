import { test } from "node:test";
import assert from "node:assert/strict";
import { compileStyle } from "./schema.mjs";
import {
  buildAddSheet, buildDeleteSheet, buildRenameSheet, buildDuplicateSheet,
  buildRepeatCellFormat, buildMerge, buildUnmerge, buildSetBorders,
  buildAddConditionalFormat, buildFreeze, buildResizeColumns,
  buildInsertColumns, buildDeleteColumns, buildSort, buildSetFilter,
  buildFindReplace, buildSetNote, buildSetValidation, conditionFromSpec,
  buildDeveloperMetadata, META_KEY_IDEMPOTENCY,
  compileTableColumns, buildAddTable, buildUpdateTable, buildDeleteTable,
  buildSetRowHeight, buildCopyFormat, buildUpdateCells, buildRestoreCells,
} from "./requests.mjs";

const RANGE = { sheetId: 1, startRowIndex: 1, endRowIndex: 10, startColumnIndex: 1, endColumnIndex: 4 };

test("buildAddSheet — properties + grid + tab color", () => {
  assert.deepEqual(
    buildAddSheet({ title: "X", rows: 10, cols: 5, index: 2, frozenRows: 1, tabColor: "#ff0000" }),
    [{
      addSheet: {
        properties: {
          title: "X", index: 2,
          gridProperties: { rowCount: 10, columnCount: 5, frozenRowCount: 1 },
          tabColorStyle: { rgbColor: { red: 1, green: 0, blue: 0 } },
        },
      },
    }],
  );
});

test("buildAddSheet — minimal (title only)", () => {
  assert.deepEqual(buildAddSheet({ title: "T" }), [{ addSheet: { properties: { title: "T" } } }]);
});

test("buildDeleteSheet / buildRenameSheet / buildDuplicateSheet", () => {
  assert.deepEqual(buildDeleteSheet(5), [{ deleteSheet: { sheetId: 5 } }]);
  assert.deepEqual(buildRenameSheet(5, "New"),
    [{ updateSheetProperties: { properties: { sheetId: 5, title: "New" }, fields: "title" } }]);
  assert.deepEqual(buildDuplicateSheet(5, { newTitle: "Copy" }),
    [{ duplicateSheet: { sourceSheetId: 5, newSheetName: "Copy" } }]);
  assert.deepEqual(buildDuplicateSheet(5), [{ duplicateSheet: { sourceSheetId: 5 } }]);
});

test("buildRepeatCellFormat — carries cellFormat + fields mask", () => {
  const { cellFormat, fields } = compileStyle({ textFormat: { bold: true } });
  assert.deepEqual(buildRepeatCellFormat(RANGE, cellFormat, fields),
    [{ repeatCell: { range: RANGE, cell: cellFormat, fields } }]);
  assert.equal(fields, "userEnteredFormat.textFormat");
});

test("buildMerge / buildUnmerge", () => {
  assert.deepEqual(buildMerge(RANGE), [{ mergeCells: { range: RANGE, mergeType: "MERGE_ALL" } }]);
  assert.deepEqual(buildMerge(RANGE, "MERGE_COLUMNS"),
    [{ mergeCells: { range: RANGE, mergeType: "MERGE_COLUMNS" } }]);
  assert.deepEqual(buildUnmerge(RANGE), [{ unmergeCells: { range: RANGE } }]);
});

test("buildSetBorders — 'all' sets every edge SOLID black", () => {
  const [req] = buildSetBorders(RANGE, "all");
  const black = { style: "SOLID", colorStyle: { rgbColor: { red: 0, green: 0, blue: 0 } } };
  for (const e of ["top", "bottom", "left", "right", "innerHorizontal", "innerVertical"]) {
    assert.deepEqual(req.updateBorders[e], black, `edge ${e}`);
  }
  assert.deepEqual(req.updateBorders.range, RANGE);
});

test("buildSetBorders — 'outer' omits inner edges", () => {
  const [req] = buildSetBorders(RANGE, "outer");
  assert.ok(req.updateBorders.top && req.updateBorders.right);
  assert.equal(req.updateBorders.innerHorizontal, undefined);
  assert.equal(req.updateBorders.innerVertical, undefined);
});

test("buildSetBorders — per-edge object with defaults", () => {
  const [req] = buildSetBorders(RANGE, { style: "DASHED", color: "#ff0000", bottom: true });
  assert.deepEqual(req.updateBorders.bottom,
    { style: "DASHED", colorStyle: { rgbColor: { red: 1, green: 0, blue: 0 } } });
  assert.equal(req.updateBorders.top, undefined);
});

test("buildFreeze — rows and cols masks", () => {
  assert.deepEqual(buildFreeze(7, { rows: 1 }),
    [{ updateSheetProperties: { properties: { sheetId: 7, gridProperties: { frozenRowCount: 1 } }, fields: "gridProperties.frozenRowCount" } }]);
  const [req] = buildFreeze(7, { rows: 1, cols: 2 });
  assert.equal(req.updateSheetProperties.fields, "gridProperties.frozenRowCount,gridProperties.frozenColumnCount");
  assert.throws(() => buildFreeze(7, {}));
});

test("buildResizeColumns — width vs auto", () => {
  assert.deepEqual(buildResizeColumns(7, 1, 4, { width: 120 }),
    [{ updateDimensionProperties: { range: { sheetId: 7, dimension: "COLUMNS", startIndex: 1, endIndex: 4 }, properties: { pixelSize: 120 }, fields: "pixelSize" } }]);
  assert.deepEqual(buildResizeColumns(7, 1, 4, { auto: true }),
    [{ autoResizeDimensions: { dimensions: { sheetId: 7, dimension: "COLUMNS", startIndex: 1, endIndex: 4 } } }]);
  assert.throws(() => buildResizeColumns(7, 1, 4, {}));
});

test("buildInsertColumns / buildDeleteColumns", () => {
  assert.deepEqual(buildInsertColumns(7, 2, 3),
    [{ insertDimension: { range: { sheetId: 7, dimension: "COLUMNS", startIndex: 2, endIndex: 5 }, inheritFromBefore: true } }]);
  assert.equal(buildInsertColumns(7, 0, 1)[0].insertDimension.inheritFromBefore, false);
  assert.deepEqual(buildDeleteColumns(7, 2, 1),
    [{ deleteDimension: { range: { sheetId: 7, dimension: "COLUMNS", startIndex: 2, endIndex: 3 } } }]);
});

test("buildSort — normalizes order + column→dimensionIndex", () => {
  assert.deepEqual(buildSort(RANGE, [{ column: 1, order: "DESC" }, { column: 0 }]),
    [{ sortRange: { range: RANGE, sortSpecs: [
      { dimensionIndex: 1, sortOrder: "DESCENDING" },
      { dimensionIndex: 0, sortOrder: "ASCENDING" },
    ] } }]);
});

test("buildSetFilter", () => {
  assert.deepEqual(buildSetFilter(RANGE), [{ setBasicFilter: { filter: { range: RANGE } } }]);
});

test("buildFindReplace — scope precedence + flags", () => {
  assert.deepEqual(buildFindReplace({ find: "a", replacement: "b", sheetId: 7, matchCase: true }),
    [{ findReplace: { find: "a", replacement: "b", sheetId: 7, matchCase: true } }]);
  // range beats sheetId; allSheets beats both
  assert.equal(buildFindReplace({ find: "a", range: RANGE, sheetId: 7 })[0].findReplace.range, RANGE);
  assert.equal(buildFindReplace({ find: "a", allSheets: true, sheetId: 7 })[0].findReplace.allSheets, true);
});

test("buildSetNote", () => {
  assert.deepEqual(buildSetNote(RANGE, "hi"),
    [{ repeatCell: { range: RANGE, cell: { note: "hi" }, fields: "note" } }]);
});

test("buildDeveloperMetadata — value is the raw key (no title prefix), row-scoped", () => {
  assert.deepEqual(buildDeveloperMetadata(7, "k1", 4), {
    createDeveloperMetadata: {
      developerMetadata: {
        metadataKey: META_KEY_IDEMPOTENCY,
        metadataValue: "k1",
        location: { dimensionRange: { sheetId: 7, dimension: "ROWS", startIndex: 4, endIndex: 5 } },
        visibility: "PROJECT",
      },
    },
  });
});

test("conditionFromSpec — argument sourcing per type", () => {
  assert.deepEqual(conditionFromSpec({ type: "ONE_OF_LIST", values: ["a", "b"] }),
    { type: "ONE_OF_LIST", values: [{ userEnteredValue: "a" }, { userEnteredValue: "b" }] });
  assert.deepEqual(conditionFromSpec({ type: "NUMBER_BETWEEN", min: 1, max: 10 }),
    { type: "NUMBER_BETWEEN", values: [{ userEnteredValue: "1" }, { userEnteredValue: "10" }] });
  assert.deepEqual(conditionFromSpec({ type: "NUMBER_GREATER", value: 5 }),
    { type: "NUMBER_GREATER", values: [{ userEnteredValue: "5" }] });
  assert.deepEqual(conditionFromSpec({ type: "ONE_OF_RANGE", range: "Sheet1!A1:A9" }),
    { type: "ONE_OF_RANGE", values: [{ userEnteredValue: "=Sheet1!A1:A9" }] });
  assert.deepEqual(conditionFromSpec({ type: "CUSTOM_FORMULA", formula: "=A1>0" }),
    { type: "CUSTOM_FORMULA", values: [{ userEnteredValue: "=A1>0" }] });
  assert.deepEqual(conditionFromSpec({ type: "BOOLEAN" }), { type: "BOOLEAN" }); // no values
});

test("conditionFromSpec — rejects missing operands", () => {
  assert.throws(() => conditionFromSpec({ type: "ONE_OF_LIST", values: [] }), /non-empty/);
  assert.throws(() => conditionFromSpec({ type: "ONE_OF_RANGE" }), /requires \{ range \}/);
  assert.throws(() => conditionFromSpec({ type: "CUSTOM_FORMULA" }), /requires \{ formula \}/);
  assert.throws(() => conditionFromSpec({ type: "NUMBER_BETWEEN", min: 1 }), /requires \{ min, max \}/);
  assert.throws(() => conditionFromSpec({ type: "NUMBER_GREATER" }), /requires \{ value \}/);
  assert.throws(() => conditionFromSpec({}), /spec.type is required/);
});

test("buildFindReplace — rejects an unscoped request", () => {
  assert.throws(() => buildFindReplace({ find: "a", replacement: "b" }), /scope required/);
});

test("buildSetValidation — rule shape, defaults, clear", () => {
  assert.deepEqual(buildSetValidation(RANGE, { type: "ONE_OF_LIST", values: ["x"] }),
    [{ setDataValidation: { range: RANGE, rule: {
      condition: { type: "ONE_OF_LIST", values: [{ userEnteredValue: "x" }] },
      strict: false, showCustomUi: true,
    } } }]);

  // non-list types default showCustomUi false
  assert.equal(buildSetValidation(RANGE, { type: "NUMBER_BETWEEN", min: 1, max: 9 })[0]
    .setDataValidation.rule.showCustomUi, false);

  // BOOLEAN → checkbox ui on
  assert.equal(buildSetValidation(RANGE, { type: "BOOLEAN" })[0]
    .setDataValidation.rule.showCustomUi, true);

  // clear → no rule (all three forms, incl. the bare "clear" string — issue #10 bug #1)
  assert.deepEqual(buildSetValidation(RANGE, { clear: true }), [{ setDataValidation: { range: RANGE } }]);
  assert.deepEqual(buildSetValidation(RANGE, { type: "clear" }), [{ setDataValidation: { range: RANGE } }]);
  assert.deepEqual(buildSetValidation(RANGE, "clear"), [{ setDataValidation: { range: RANGE } }]);
});

// ── Tables ───────────────────────────────────────────────────────────────────

test("compileTableColumns — types, aliases, dropdown values → validation", () => {
  const cols = compileTableColumns([
    { name: "Контрагент", type: "TEXT" },
    { name: "Категория", type: "DROPDOWN", values: ["Налоги", "Резерв"] },
    { name: "Сумма", type: "NUMBER" },            // NUMBER → DOUBLE alias
    { name: "Дата", type: "date" },               // lowercased passthrough
    { name: "Авто", values: ["a", "b"] },         // values w/o type → DROPDOWN implied
  ]);
  assert.deepEqual(cols[0], { columnIndex: 0, columnName: "Контрагент", columnType: "TEXT" });
  assert.deepEqual(cols[1], {
    columnIndex: 1, columnName: "Категория", columnType: "DROPDOWN",
    dataValidationRule: { condition: { type: "ONE_OF_LIST", values: [{ userEnteredValue: "Налоги" }, { userEnteredValue: "Резерв" }] }, strict: false },
  });
  assert.equal(cols[2].columnType, "DOUBLE");
  assert.equal(cols[3].columnType, "DATE");
  assert.equal(cols[4].columnType, "DROPDOWN");
  assert.ok(cols[4].dataValidationRule);
});

test("buildAddTable / buildUpdateTable / buildDeleteTable", () => {
  const table = { name: "T", range: RANGE, columnProperties: [{ columnIndex: 0, columnName: "A", columnType: "TEXT" }] };
  assert.deepEqual(buildAddTable(table), [{ addTable: { table } }]);
  assert.deepEqual(buildUpdateTable({ tableId: "t1", name: "T2" }, "name"),
    [{ updateTable: { table: { tableId: "t1", name: "T2" }, fields: "name" } }]);
  assert.deepEqual(buildDeleteTable("t1"), [{ deleteTable: { tableId: "t1" } }]);
});

// ── Row height / copy format / updateCells ───────────────────────────────────

test("buildSetRowHeight", () => {
  assert.deepEqual(buildSetRowHeight(7, 21, 22, 39),
    [{ updateDimensionProperties: { range: { sheetId: 7, dimension: "ROWS", startIndex: 21, endIndex: 22 }, properties: { pixelSize: 39 }, fields: "pixelSize" } }]);
});

test("buildCopyFormat — PASTE_FORMAT only", () => {
  const src = { sheetId: 7, startRowIndex: 17, endRowIndex: 18 };
  const dst = { sheetId: 7, startRowIndex: 21, endRowIndex: 22 };
  assert.deepEqual(buildCopyFormat(src, dst),
    [{ copyPaste: { source: src, destination: dst, pasteType: "PASTE_FORMAT", pasteOrientation: "NORMAL" } }]);
});

test("buildUpdateCells — formulas → formulaValue, fields default userEnteredValue", () => {
  const [req] = buildUpdateCells(RANGE, [['=IF([@[Сумма]]="";"";1)', 42, "txt"]]);
  assert.equal(req.updateCells.fields, "userEnteredValue");
  assert.deepEqual(req.updateCells.range, RANGE);
  assert.deepEqual(req.updateCells.rows[0].values, [
    { userEnteredValue: { formulaValue: '=IF([@[Сумма]]="";"";1)' } },
    { userEnteredValue: { numberValue: 42 } },
    { userEnteredValue: { stringValue: "txt" } },
  ]);
});

test("buildRestoreCells — echoes CellData rows verbatim, value+format+note fields", () => {
  const rows = [{ values: [{ userEnteredValue: { numberValue: 1 }, note: "n" }] }];
  const [req] = buildRestoreCells(RANGE, rows);
  assert.equal(req.updateCells.fields, "userEnteredValue,userEnteredFormat,note");
  assert.deepEqual(req.updateCells.range, RANGE);
  assert.equal(req.updateCells.rows, rows); // no re-encoding — same reference
});

test("buildAddConditionalFormat — boolean rule with format", () => {
  assert.deepEqual(
    buildAddConditionalFormat(RANGE, { condition: { type: "NUMBER_GREATER", value: 100 }, format: { backgroundColor: "#ff0000" } }),
    [{ addConditionalFormatRule: { index: 0, rule: {
      ranges: [RANGE],
      booleanRule: {
        condition: { type: "NUMBER_GREATER", values: [{ userEnteredValue: "100" }] },
        format: { backgroundColorStyle: { rgbColor: { red: 1, green: 0, blue: 0 } } },
      },
    } } }],
  );
});

test("buildAddConditionalFormat — gradient rule", () => {
  const [req] = buildAddConditionalFormat(RANGE, {
    gradient: { min: { color: "#ffffff", type: "MIN" }, max: { color: "#00ff00", type: "MAX" } },
  });
  assert.deepEqual(req.addConditionalFormatRule.rule.gradientRule, {
    minpoint: { colorStyle: { rgbColor: { red: 1, green: 1, blue: 1 } }, type: "MIN" },
    maxpoint: { colorStyle: { rgbColor: { red: 0, green: 1, blue: 0 } }, type: "MAX" },
  });
});
