/**
 * Pure Sheets v4 Request builders. Each function returns an array of Request
 * objects (the shape `spreadsheets.batchUpdate` accepts) and performs NO I/O —
 * callers resolve sheet names → sheetId and A1 → GridRange before calling.
 *
 * Keeping these pure + dependency-light is what makes them unit-testable and
 * lets every higher-level method compile to a single atomic batchUpdate that
 * automatically respects dry-run.
 */

import { compileStyle, colorToRgbFloat } from "./schema.mjs";

// DeveloperMetadata key used for idempotency tokens (one per inserted row).
export const META_KEY_IDEMPOTENCY = "sheets-mcp:idempotency";

// ── Structural (spreadsheet & sheet) ────────────────────────────────────────

export function buildAddSheet(opts = {}) {
  const props = {};
  if (opts.title != null) props.title = opts.title;
  if (opts.index != null) props.index = opts.index;
  if (opts.sheetId != null) props.sheetId = opts.sheetId;

  const grid = {};
  if (opts.rows != null) grid.rowCount = opts.rows;
  if (opts.cols != null) grid.columnCount = opts.cols;
  if (opts.frozenRows != null) grid.frozenRowCount = opts.frozenRows;
  if (opts.frozenCols != null) grid.frozenColumnCount = opts.frozenCols;
  if (Object.keys(grid).length) props.gridProperties = grid;

  if (opts.tabColor != null) props.tabColorStyle = { rgbColor: colorToRgbFloat(opts.tabColor) };

  return [{ addSheet: { properties: props } }];
}

export function buildDeleteSheet(sheetId) {
  return [{ deleteSheet: { sheetId } }];
}

export function buildRenameSheet(sheetId, newTitle) {
  return [{ updateSheetProperties: { properties: { sheetId, title: newTitle }, fields: "title" } }];
}

export function buildDuplicateSheet(sourceSheetId, opts = {}) {
  const dup = { sourceSheetId };
  if (opts.newTitle != null) dup.newSheetName = opts.newTitle;
  if (opts.index != null) dup.insertSheetIndex = opts.index;
  return [{ duplicateSheet: dup }];
}

// ── Tables (native Sheets v4 Table resource) ────────────────────────────────

// Friendly aliases → Sheets v4 ColumnType. Unknown values pass through uppercased
// so any current/future enum (TEXT, DOUBLE, CURRENCY, PERCENT, DATE, TIME,
// DATE_TIME, BOOLEAN, DROPDOWN, …) works without a hardcoded allowlist.
const COLUMN_TYPE_ALIASES = { NUMBER: "DOUBLE", BOOL: "BOOLEAN", LIST: "DROPDOWN" };

/** A table column's DataValidationRule from a friendly validation spec. */
export function buildColumnValidationRule(spec) {
  const rule = { condition: conditionFromSpec(spec), strict: spec.strict ?? false };
  if (spec.showCustomUi !== undefined) rule.showCustomUi = spec.showCustomUi;
  if (spec.inputMessage) rule.inputMessage = spec.inputMessage;
  return rule;
}

/**
 * Compile friendly column specs into Sheets v4 Table columnProperties.
 *   columns: [{ name, type?, values?, validation? }]
 *     type     → columnType (NUMBER→DOUBLE etc.; else uppercased passthrough)
 *     values   → dropdown choices: implies DROPDOWN + a ONE_OF_LIST rule
 *     validation → an explicit conditionFromSpec-style spec → dataValidationRule
 */
export function compileTableColumns(columns = []) {
  return columns.map((c, i) => {
    const col = {
      columnIndex: c.columnIndex ?? i,
      columnName: c.name ?? c.columnName ?? "",
    };
    let type = c.type ?? c.columnType;
    if (type != null) {
      type = String(type).toUpperCase();
      col.columnType = COLUMN_TYPE_ALIASES[type] ?? type;
    }
    if (Array.isArray(c.values) && c.values.length > 0) {
      if (col.columnType == null) col.columnType = "DROPDOWN";
      col.dataValidationRule = buildColumnValidationRule({ type: "ONE_OF_LIST", values: c.values });
    } else if (c.validation) {
      col.dataValidationRule = buildColumnValidationRule(c.validation);
    }
    return col;
  });
}

/** table: { name, range: GridRange, columnProperties, tableId?, rowsProperties? }. */
export function buildAddTable(table) {
  return [{ addTable: { table } }];
}

/** table: { tableId, ...changed fields }. fields: update mask (e.g. "name,columnProperties"). */
export function buildUpdateTable(table, fields) {
  return [{ updateTable: { table, fields } }];
}

export function buildDeleteTable(tableId) {
  return [{ deleteTable: { tableId } }];
}

/**
 * One idempotency token attached to a row (0-based startIndex). The token value
 * is the raw key — scoped to a row of a specific sheetId via its location, so it
 * survives a tab rename (issue #7 limitation #3) and never needs a title prefix.
 */
export function buildDeveloperMetadata(sheetId, key, startIndex) {
  return {
    createDeveloperMetadata: {
      developerMetadata: {
        metadataKey: META_KEY_IDEMPOTENCY,
        metadataValue: String(key),
        location: {
          dimensionRange: { sheetId, dimension: "ROWS", startIndex, endIndex: startIndex + 1 },
        },
        visibility: "PROJECT",
      },
    },
  };
}

// ── Formatting / presentation ───────────────────────────────────────────────

/** range: GridRange, cellFormat/fields: output of compileStyle(). */
export function buildRepeatCellFormat(range, cellFormat, fields) {
  return [{ repeatCell: { range, cell: cellFormat, fields } }];
}

export function buildMerge(range, mergeType = "MERGE_ALL") {
  return [{ mergeCells: { range, mergeType } }];
}

export function buildUnmerge(range) {
  return [{ unmergeCells: { range } }];
}

const BORDER_EDGES = ["top", "bottom", "left", "right", "innerHorizontal", "innerVertical"];

function makeBorder(edgeSpec, defStyle, defColor) {
  if (!edgeSpec) return null;
  const s = edgeSpec === true ? {} : edgeSpec;
  return {
    style: s.style || defStyle || "SOLID",
    colorStyle: { rgbColor: colorToRgbFloat(s.color || defColor || "#000000") },
  };
}

/**
 * spec:
 *   "all"   → every edge + inner, SOLID black
 *   "outer" → top/bottom/left/right only
 *   "DASHED"/"SOLID_THICK"/… → that style on every edge + inner
 *   { style?, color?, top?, bottom?, left?, right?, innerHorizontal?, innerVertical? }
 *       each edge: true | { style?, color? }; top-level style/color are defaults.
 */
export function buildSetBorders(range, spec) {
  let edges;
  let defStyle;
  let defColor;
  if (spec === "all" || spec === true) {
    edges = { top: true, bottom: true, left: true, right: true, innerHorizontal: true, innerVertical: true };
  } else if (spec === "outer") {
    edges = { top: true, bottom: true, left: true, right: true };
  } else if (typeof spec === "string") {
    edges = {};
    for (const e of BORDER_EDGES) edges[e] = { style: spec };
  } else {
    edges = spec || {};
    defStyle = edges.style;
    defColor = edges.color;
  }

  const req = { updateBorders: { range } };
  for (const e of BORDER_EDGES) {
    const b = makeBorder(edges[e], defStyle, defColor);
    if (b) req.updateBorders[e] = b;
  }
  return [{ updateBorders: req.updateBorders }];
}

function interpolationPoint(p) {
  if (!p) return undefined;
  const out = { colorStyle: { rgbColor: colorToRgbFloat(p.color) }, type: p.type || "NUMBER" };
  if (p.value !== undefined) out.value = String(p.value);
  return out;
}

function buildGradientRule(g) {
  const rule = {};
  const min = g.min || g.minpoint;
  const mid = g.mid || g.midpoint;
  const max = g.max || g.maxpoint;
  if (min) rule.minpoint = interpolationPoint(min);
  if (mid) rule.midpoint = interpolationPoint(mid);
  if (max) rule.maxpoint = interpolationPoint(max);
  return rule;
}

/**
 * range: GridRange. rule:
 *   gradient → { gradient: { min, mid?, max } }  (each point: { color, type?, value? })
 *   boolean  → { condition: <validation-style spec>, format: <StyleObject> }
 *   index?   → insertion order (default 0, prepended)
 */
export function buildAddConditionalFormat(range, rule) {
  const ruleObj = { ranges: [range] };
  if (rule.gradient || rule.type === "gradient" || rule.type === "GRADIENT") {
    ruleObj.gradientRule = buildGradientRule(rule.gradient || rule);
  } else {
    const booleanRule = { condition: conditionFromSpec(rule.condition || rule) };
    if (rule.format) booleanRule.format = compileStyle(rule.format).cellFormat.userEnteredFormat;
    ruleObj.booleanRule = booleanRule;
  }
  return [{ addConditionalFormatRule: { rule: ruleObj, index: rule.index ?? 0 } }];
}

// ── Dimensions (rows & columns) ─────────────────────────────────────────────

export function buildFreeze(sheetId, { rows, cols } = {}) {
  const grid = {};
  const fields = [];
  if (rows != null) { grid.frozenRowCount = rows; fields.push("gridProperties.frozenRowCount"); }
  if (cols != null) { grid.frozenColumnCount = cols; fields.push("gridProperties.frozenColumnCount"); }
  if (fields.length === 0) throw new Error("freeze: provide rows and/or cols");
  return [{ updateSheetProperties: { properties: { sheetId, gridProperties: grid }, fields: fields.join(",") } }];
}

export function buildResizeColumns(sheetId, startCol, endCol, { width, auto } = {}) {
  const dimensions = { sheetId, dimension: "COLUMNS", startIndex: startCol, endIndex: endCol };
  if (auto) return [{ autoResizeDimensions: { dimensions } }];
  if (width == null) throw new Error("resizeColumns: provide { width } or { auto: true }");
  return [{ updateDimensionProperties: { range: dimensions, properties: { pixelSize: width }, fields: "pixelSize" } }];
}

export function buildInsertColumns(sheetId, at, count = 1) {
  return [{
    insertDimension: {
      range: { sheetId, dimension: "COLUMNS", startIndex: at, endIndex: at + count },
      inheritFromBefore: at > 0,
    },
  }];
}

export function buildDeleteColumns(sheetId, at, count = 1) {
  return [{ deleteDimension: { range: { sheetId, dimension: "COLUMNS", startIndex: at, endIndex: at + count } } }];
}

/** Set row heights (0-based, half-open row range) to pixelSize. */
export function buildSetRowHeight(sheetId, startRow, endRow, pixelSize) {
  return [{
    updateDimensionProperties: {
      range: { sheetId, dimension: "ROWS", startIndex: startRow, endIndex: endRow },
      properties: { pixelSize },
      fields: "pixelSize",
    },
  }];
}

/** Copy ONLY the formatting from sourceRange onto destRange (both GridRange). */
export function buildCopyFormat(sourceRange, destRange) {
  return [{
    copyPaste: {
      source: sourceRange,
      destination: destRange,
      pasteType: "PASTE_FORMAT",
      pasteOrientation: "NORMAL",
    },
  }];
}

// ── Data operations ─────────────────────────────────────────────────────────

function normalizeSortOrder(order) {
  if (!order) return "ASCENDING";
  return String(order).toUpperCase().startsWith("DESC") ? "DESCENDING" : "ASCENDING";
}

/** range: GridRange. specs: [{ column | dimensionIndex, order?: "ASC"|"DESC" }]. */
export function buildSort(range, specs) {
  const sortSpecs = (specs || []).map(s => ({
    dimensionIndex: s.dimensionIndex != null ? s.dimensionIndex : s.column,
    sortOrder: normalizeSortOrder(s.order),
  }));
  return [{ sortRange: { range, sortSpecs } }];
}

export function buildSetFilter(range) {
  return [{ setBasicFilter: { filter: { range } } }];
}

/** Exactly one of { sheetId, range, allSheets } must be set (caller defaults sheetId). */
export function buildFindReplace(opts) {
  if (!opts.allSheets && !opts.range && opts.sheetId == null) {
    throw new Error("findReplace: scope required — pass sheetId, range, or allSheets");
  }
  const fr = { find: opts.find, replacement: opts.replacement ?? "" };
  if (opts.allSheets) fr.allSheets = true;
  else if (opts.range) fr.range = opts.range;
  else if (opts.sheetId != null) fr.sheetId = opts.sheetId;
  if (opts.matchCase != null) fr.matchCase = opts.matchCase;
  if (opts.matchEntireCell != null) fr.matchEntireCell = opts.matchEntireCell;
  if (opts.searchByRegex != null) fr.searchByRegex = opts.searchByRegex;
  if (opts.includeFormulas != null) fr.includeFormulas = opts.includeFormulas;
  return [{ findReplace: fr }];
}

export function buildSetNote(range, text) {
  return [{ repeatCell: { range, cell: { note: text ?? "" }, fields: "note" } }];
}

/** Encode a JS scalar into a Sheets CellData userEnteredValue (formula if it starts with "="). */
function encodeCellValue(v) {
  if (v === null || v === undefined || v === "") return { userEnteredValue: { stringValue: "" } };
  if (typeof v === "number") return { userEnteredValue: { numberValue: v } };
  if (typeof v === "boolean") return { userEnteredValue: { boolValue: v } };
  const s = String(v);
  if (s.startsWith("=")) return { userEnteredValue: { formulaValue: s } };
  return { userEnteredValue: { stringValue: s } };
}

/**
 * Write a 2D array into a GridRange via updateCells. Unlike values.update, this
 * binds Table structured references ([@[Col]]) in the row's table context, so
 * such formulas don't render #ERROR!. fields defaults to "userEnteredValue" so
 * existing cell formatting is preserved.
 */
export function buildUpdateCells(range, values, fields = "userEnteredValue") {
  const rows = values.map(row => ({ values: row.map(encodeCellValue) }));
  return [{ updateCells: { range, rows, fields } }];
}

/**
 * Write already-shaped CellData rows (as returned by spreadsheets.get gridData:
 * [{ values: [CellData, …] }]) back into a GridRange. Unlike buildUpdateCells,
 * this does NOT encode scalars — it echoes full CellData so userEnteredValue,
 * userEnteredFormat, and notes survive. Used by deleteTable to restore the
 * table's cells after the (data-destroying) DeleteTableRequest (issue #11).
 */
export function buildRestoreCells(range, rows, fields = "userEnteredValue,userEnteredFormat,note") {
  return [{ updateCells: { range, rows, fields } }];
}

// ── Data validation ─────────────────────────────────────────────────────────

function toRangeFormula(range) {
  const r = String(range);
  return r.startsWith("=") ? r : `=${r}`;
}

// Condition types that take no operand (so a missing value isn't an error).
const NO_ARG_CONDITIONS = new Set([
  "BOOLEAN", "BLANK", "NOT_BLANK", "TEXT_IS_EMAIL", "TEXT_IS_URL", "DATE_IS_VALID",
]);

/**
 * Map a friendly validation spec to a Sheets BooleanCondition.
 * spec.type is a Sheets ConditionType (e.g. "NUMBER_BETWEEN", "DATE_AFTER",
 * "TEXT_CONTAINS", "BOOLEAN", "ONE_OF_LIST", "ONE_OF_RANGE", "CUSTOM_FORMULA").
 * Argument source: BETWEEN reads {min,max}; ONE_OF_LIST reads {values};
 * ONE_OF_RANGE reads {range}; CUSTOM_FORMULA reads {formula}; everything else
 * reads {value} (or {values}). Conditions with no args (BOOLEAN, BLANK,
 * NOT_BLANK, *_IS_VALID, TEXT_IS_EMAIL/URL) carry no values.
 */
export function conditionFromSpec(spec) {
  const type = spec.type;
  if (!type) throw new Error("validation: spec.type is required");

  let values;
  if (type === "ONE_OF_LIST") {
    if (!Array.isArray(spec.values) || spec.values.length === 0) {
      throw new Error("validation ONE_OF_LIST: requires a non-empty { values }");
    }
    values = spec.values.map(v => ({ userEnteredValue: String(v) }));
  } else if (type === "ONE_OF_RANGE") {
    if (!spec.range) throw new Error("validation ONE_OF_RANGE: requires { range }");
    values = [{ userEnteredValue: toRangeFormula(spec.range) }];
  } else if (type === "CUSTOM_FORMULA") {
    const formula = spec.formula ?? spec.value;
    if (formula == null) throw new Error("validation CUSTOM_FORMULA: requires { formula }");
    values = [{ userEnteredValue: String(formula) }];
  } else if (/_BETWEEN$/.test(type)) {
    if (spec.min == null || spec.max == null) {
      throw new Error(`validation ${type}: requires { min, max }`);
    }
    values = [spec.min, spec.max].map(v => ({ userEnteredValue: String(v) }));
  } else if (spec.value !== undefined) {
    values = [{ userEnteredValue: String(spec.value) }];
  } else if (Array.isArray(spec.values)) {
    values = spec.values.map(v => ({ userEnteredValue: String(v) }));
  } else if (!NO_ARG_CONDITIONS.has(type)) {
    throw new Error(`validation ${type}: requires { value }`);
  }

  const condition = { type };
  if (values) condition.values = values;
  return condition;
}

/**
 * range: GridRange. spec:
 *   "clear" | { type: "clear" } | { clear: true }  → remove validation (setDataValidation w/o rule)
 *   otherwise → conditionFromSpec(spec) with { strict?, showCustomUi?, inputMessage? }
 * showCustomUi defaults true for ONE_OF_LIST / ONE_OF_RANGE / BOOLEAN.
 */
export function buildSetValidation(range, spec) {
  if (spec === "clear" || spec?.type === "clear" || spec?.clear === true) {
    return [{ setDataValidation: { range } }];
  }
  const condition = conditionFromSpec(spec);
  const rule = { condition, strict: spec.strict ?? false };
  rule.showCustomUi = spec.showCustomUi !== undefined
    ? spec.showCustomUi
    : (spec.type === "ONE_OF_LIST" || spec.type === "ONE_OF_RANGE" || spec.type === "BOOLEAN");
  if (spec.inputMessage) rule.inputMessage = spec.inputMessage;
  return [{ setDataValidation: { range, rule } }];
}
