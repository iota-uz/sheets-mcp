/**
 * Schema loader and helpers.
 *
 * Schemas live in schemas/<name>.json. Each declares the columns of a sheet
 * (role → header + type + format + validation), the primary key, and
 * sheet-level properties (frozen rows, header style, conditional formats).
 *
 * The schema is the API contract between agent code and the spreadsheet.
 * Headers may be renamed/reordered in the UI; the role binding survives via
 * DeveloperMetadata written by scripts/bootstrap-schema.mjs.
 */

import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const SCHEMA_DIR = path.resolve(__dirname, "..", "schemas");

const cache = new Map();

/**
 * Load schema by sheet name. Cached after first read.
 */
export function loadSchema(name) {
  if (cache.has(name)) return cache.get(name);

  const file = path.join(SCHEMA_DIR, `${name}.json`);
  if (!fs.existsSync(file)) {
    throw new Error(`No schema declared for sheet "${name}" (looked at ${file})`);
  }

  const raw = JSON.parse(fs.readFileSync(file, "utf8"));
  const schema = normalizeSchema(raw);
  cache.set(name, schema);
  return schema;
}

/**
 * List all declared schemas.
 */
export function listSchemas() {
  if (!fs.existsSync(SCHEMA_DIR)) return [];
  return fs.readdirSync(SCHEMA_DIR)
    .filter(f => f.endsWith(".json"))
    .map(f => f.replace(/\.json$/, ""));
}

/**
 * Normalize a raw schema: fill defaults, validate shape.
 */
function normalizeSchema(raw) {
  if (!raw.sheet) throw new Error("Schema missing 'sheet' field");
  if (!raw.columns) throw new Error(`Schema "${raw.sheet}" missing 'columns'`);

  const headerRow = raw.headerRow ?? 1;
  const primaryKey = raw.primaryKey ?? [];
  const columns = {};

  for (const [role, col] of Object.entries(raw.columns)) {
    if (!col.header) throw new Error(`Schema "${raw.sheet}" column "${role}" missing 'header'`);
    if (!col.type) throw new Error(`Schema "${raw.sheet}" column "${role}" missing 'type'`);

    columns[role] = {
      role,
      header: col.header,
      type: col.type,
      optional: col.optional ?? false,
      values: col.values,
      formula: col.formula,
      compute: col.compute,
      numberFormat: col.numberFormat,
      horizontalAlignment: col.horizontalAlignment,
      validation: col.validation,
      refSheet: col.refSheet,
      refColumn: col.refColumn,
    };
  }

  // Verify primary key references valid roles
  for (const k of primaryKey) {
    if (!columns[k]) throw new Error(`Schema "${raw.sheet}" primaryKey references unknown role "${k}"`);
  }

  return {
    sheet: raw.sheet,
    headerRow,
    primaryKey,
    columns,
    sheetProperties: raw.sheetProperties ?? {},
    conditionalFormats: raw.conditionalFormats ?? [],
  };
}

/**
 * Validate a record against a schema. Returns { valid, errors }.
 * Errors are descriptive — agent can fix and retry.
 */
export function validateRecord(schema, record) {
  const errors = [];

  for (const [role, col] of Object.entries(schema.columns)) {
    const value = record[role];
    const present = value !== undefined && value !== null && value !== "";

    // Skip computed/formula columns — they're auto-derived
    if (col.type === "computed" || col.type === "formula") continue;

    if (!present) {
      if (!col.optional) errors.push(`Missing required field "${role}" (${col.header})`);
      continue;
    }

    if (col.type === "enum" && col.values) {
      if (!col.values.includes(value)) {
        const msg = `Field "${role}" value "${value}" not in allowed set [${col.values.join(", ")}]`;
        if (col.validation === "strict") errors.push(msg);
        // 'warn' — silent, just informational
      }
    }

    if (col.type === "number" && typeof value !== "number" && isNaN(parseFloat(value))) {
      errors.push(`Field "${role}" must be a number, got ${typeof value}: ${JSON.stringify(value)}`);
    }

    if (col.type === "date" && typeof value === "string") {
      // Accept DD.MM.YYYY format
      if (!/^\d{2}\.\d{2}\.\d{4}$/.test(value)) {
        errors.push(`Field "${role}" must be DD.MM.YYYY date, got "${value}"`);
      }
    }
  }

  // Check for unknown fields
  for (const key of Object.keys(record)) {
    if (!schema.columns[key]) {
      errors.push(`Unknown field "${key}" — not declared in schema`);
    }
  }

  return { valid: errors.length === 0, errors };
}

/**
 * Build a primary-key string for a record. Used as idempotency key.
 */
export function primaryKey(schema, record) {
  if (schema.primaryKey.length === 0) return null;
  return schema.primaryKey.map(k => String(record[k] ?? "")).join("");
}

/**
 * Compute derived columns ('computed' type) from input record.
 * Returns the record augmented with computed fields. Mutates input.
 */
export function applyComputed(schema, record) {
  for (const [role, col] of Object.entries(schema.columns)) {
    if (col.type !== "computed" || !col.compute) continue;
    if (record[role] !== undefined) continue; // user already set it

    const m = col.compute.match(/^(\w+)\((\w+)\)$/);
    if (!m) continue;
    const [, fn, arg] = m;
    const argValue = record[arg];
    if (argValue == null) continue;

    if (fn === "firstOfMonth") {
      const parts = String(argValue).split(".");
      if (parts.length === 3) record[role] = `01.${parts[1]}.${parts[2]}`;
    }
  }
  return record;
}

/**
 * Build a formula string from a column's formula template, substituting
 * {role} references with the actual A1 column letter and {row} with the row.
 *
 * Caller passes a roleToColumn map: { date: "B", currency: "F", ... }.
 */
export function compileFormula(template, roleToColumn, row) {
  return template.replace(/\{(\w+)\}/g, (match, key) => {
    if (key === "row") return String(row);
    if (roleToColumn[key]) return roleToColumn[key];
    return match;
  });
}

/**
 * Convert a hex color (#rrggbb) or shorthand (#rgb) to {red, green, blue} floats.
 * Accepts named colors for a small set.
 */
const NAMED_COLORS = {
  white: "#ffffff", black: "#000000", red: "#ff0000",
  green: "#00aa00", blue: "#0000ff", yellow: "#ffff00",
  grey: "#888888", gray: "#888888", transparent: null,
};

export function colorToRgbFloat(input) {
  if (input == null) return null;
  if (typeof input === "object" && "red" in input) return input;

  let hex = input;
  if (typeof input === "string" && NAMED_COLORS[input.toLowerCase()] !== undefined) {
    hex = NAMED_COLORS[input.toLowerCase()];
    if (hex === null) return null;
  }

  if (typeof hex !== "string" || !hex.startsWith("#")) {
    throw new Error(`Invalid color: ${JSON.stringify(input)}`);
  }

  let h = hex.slice(1);
  if (h.length === 3) h = h.split("").map(c => c + c).join("");
  if (h.length !== 6) throw new Error(`Invalid hex color: ${input}`);

  return {
    red: parseInt(h.slice(0, 2), 16) / 255,
    green: parseInt(h.slice(2, 4), 16) / 255,
    blue: parseInt(h.slice(4, 6), 16) / 255,
  };
}

/**
 * Normalize a friendly style object into a Sheets API CellFormat shape.
 * Returns { cellFormat, fields } — fields is the precise mask for the API.
 */
export function compileStyle(style) {
  const fmt = {};
  const fields = [];

  if (style.backgroundColor !== undefined) {
    fmt.backgroundColorStyle = { rgbColor: colorToRgbFloat(style.backgroundColor) };
    fields.push("userEnteredFormat.backgroundColorStyle");
  }

  if (style.horizontalAlignment) {
    fmt.horizontalAlignment = style.horizontalAlignment;
    fields.push("userEnteredFormat.horizontalAlignment");
  }

  if (style.verticalAlignment) {
    fmt.verticalAlignment = style.verticalAlignment;
    fields.push("userEnteredFormat.verticalAlignment");
  }

  if (style.numberFormat) {
    if (typeof style.numberFormat === "string") {
      // Heuristic: detect type from pattern
      const t = /^DD|^YYYY|^MM/.test(style.numberFormat) ? "DATE"
              : /\$|₽|€/.test(style.numberFormat) ? "CURRENCY"
              : "NUMBER";
      fmt.numberFormat = { type: t, pattern: style.numberFormat };
    } else {
      fmt.numberFormat = style.numberFormat;
    }
    fields.push("userEnteredFormat.numberFormat");
  }

  if (style.textFormat) {
    const tf = {};
    if (style.textFormat.bold !== undefined) tf.bold = style.textFormat.bold;
    if (style.textFormat.italic !== undefined) tf.italic = style.textFormat.italic;
    if (style.textFormat.fontSize !== undefined) tf.fontSize = style.textFormat.fontSize;
    if (style.textFormat.foregroundColor !== undefined) {
      tf.foregroundColorStyle = { rgbColor: colorToRgbFloat(style.textFormat.foregroundColor) };
    }
    fmt.textFormat = tf;
    fields.push("userEnteredFormat.textFormat");
  }

  if (style.wrapStrategy) {
    fmt.wrapStrategy = style.wrapStrategy;
    fields.push("userEnteredFormat.wrapStrategy");
  }

  return { cellFormat: { userEnteredFormat: fmt }, fields: fields.join(",") };
}

/**
 * Convert a column index (0-based) to A1 letter.
 */
export function colLetter(i) {
  if (i < 26) return String.fromCharCode(65 + i);
  return String.fromCharCode(65 + Math.floor(i / 26) - 1) + String.fromCharCode(65 + (i % 26));
}

/**
 * Normalize a header for fuzzy matching: lowercase, trim, ё→е.
 */
export function normHeader(s) {
  return String(s ?? "").trim().toLowerCase().replace(/ё/g, "е");
}
