/**
 * Generic helpers for the Sheet API: style normalizer, color parser.
 *
 * Kept separate from sheet.mjs so the helpers are stateless and reusable.
 * No sheet-specific knowledge here.
 *
 * Column-letter math lives in a1.mjs (the single source of truth); colLetter
 * is re-exported here so existing `import { colLetter } from "./schema.mjs"`
 * call sites keep working.
 */

export { colLetter } from "./a1.mjs";

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
