/**
 * A1 notation helpers — pure, no I/O, no sheet-specific knowledge.
 *
 * The Sheet API speaks A1 ("B2:D10"), the Sheets v4 batchUpdate API speaks
 * GridRange (0-based, half-open). These helpers convert between the two and
 * provide bijective column-letter math (the single source of truth; schema.mjs
 * re-exports colLetter from here).
 */

/**
 * 0-based column index → A1 letters. Bijective base-26.
 *   0→A, 25→Z, 26→AA, 701→ZZ, 702→AAA
 */
export function colLetter(i) {
  if (!Number.isInteger(i) || i < 0) {
    throw new Error(`colLetter: expected a non-negative integer, got ${JSON.stringify(i)}`);
  }
  let n = i + 1;
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/**
 * Inverse of colLetter. Case-insensitive.
 *   "A"→0, "Z"→25, "AA"→26, "ZZ"→701, "AAA"→702
 */
export function colToIdx(letters) {
  if (typeof letters !== "string" || !/^[A-Za-z]+$/.test(letters)) {
    throw new Error(`colToIdx: expected column letters, got ${JSON.stringify(letters)}`);
  }
  let n = 0;
  for (const ch of letters.toUpperCase()) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n - 1;
}

/**
 * Strip an optional leading "SheetName!" / "'Sheet Name'!" prefix.
 * The sheetId passed to a1ToGridRange is authoritative, so the prefix (if any)
 * is discarded.
 */
function stripSheetPrefix(a1) {
  if (a1[0] === "'") {
    let i = 1;
    while (i < a1.length) {
      if (a1[i] === "'") {
        if (a1[i + 1] === "'") { i += 2; continue; } // escaped quote ''
        break;
      }
      i++;
    }
    if (a1[i] === "'" && a1[i + 1] === "!") return a1.slice(i + 2);
    return a1;
  }
  const bang = a1.indexOf("!");
  return bang >= 0 ? a1.slice(bang + 1) : a1;
}

/**
 * Parse one A1 endpoint into { col?, row? } (both 0-based). At least one of
 * col/row must be present. "C"→{col:2}, "5"→{row:4}, "C5"→{col:2,row:4}.
 */
function parseRef(token) {
  const m = /^([A-Za-z]+)?(\d+)?$/.exec(token);
  if (!m || (!m[1] && !m[2])) {
    throw new Error(`a1ToGridRange: malformed reference "${token}"`);
  }
  const ref = {};
  if (m[1]) ref.col = colToIdx(m[1]);
  if (m[2]) {
    const row = parseInt(m[2], 10);
    if (row < 1) throw new Error(`a1ToGridRange: row must be ≥ 1 in "${token}"`);
    ref.row = row - 1;
  }
  return ref;
}

/**
 * Convert an A1 range to a 0-based, half-open GridRange bound to sheetId.
 * Only the bounds that are actually constrained are included (Sheets treats
 * omitted bounds as unbounded).
 *
 * Accepts:
 *   "A1"            single cell
 *   "B2:D10"        bounded rect
 *   "C:C" / "C:E"   full column(s)
 *   "2:2" / "3:7"   full row(s)
 *   "A1:C" / "A:C5" open-ended
 *   "Sheet!A1:B2" / "'Sheet Name'!A1"  prefix stripped
 */
/**
 * Quote a sheet title for A1 use: wrap in single quotes (doubling any internal
 * quote) when it isn't a bare token that A1 accepts unquoted.
 */
function quoteSheetTitle(title) {
  return /^[A-Za-z0-9_]+$/.test(title) ? title : `'${String(title).replace(/'/g, "''")}'`;
}

/**
 * Build an A1 range string ("'Title'!A1:I6") from a fully-bounded GridRange.
 * Inverse of a1ToGridRange for the bounded case — used where the Sheets API
 * accepts A1 only (spreadsheets.get `ranges`). Requires all four indices.
 */
export function gridRangeToA1(title, range) {
  const { startRowIndex, endRowIndex, startColumnIndex, endColumnIndex } = range || {};
  if (startRowIndex == null || endRowIndex == null || startColumnIndex == null || endColumnIndex == null) {
    throw new Error(`gridRangeToA1: requires a fully-bounded GridRange, got ${JSON.stringify(range)}`);
  }
  const start = `${colLetter(startColumnIndex)}${startRowIndex + 1}`;
  const end = `${colLetter(endColumnIndex - 1)}${endRowIndex}`;
  return `${quoteSheetTitle(title)}!${start}:${end}`;
}

export function a1ToGridRange(sheetId, a1) {
  if (typeof a1 !== "string" || a1.trim() === "") {
    throw new Error(`a1ToGridRange: expected a non-empty A1 string, got ${JSON.stringify(a1)}`);
  }
  const body = stripSheetPrefix(a1.trim());
  const range = { sheetId };

  if (body.includes(":")) {
    const parts = body.split(":");
    if (parts.length !== 2) {
      throw new Error(`a1ToGridRange: malformed range "${a1}" — expected a single ":"`);
    }
    const left = parseRef(parts[0]);
    const right = parseRef(parts[1]);
    if (left.col != null) range.startColumnIndex = left.col;
    if (left.row != null) range.startRowIndex = left.row;
    if (right.col != null) range.endColumnIndex = right.col + 1;
    if (right.row != null) range.endRowIndex = right.row + 1;
    // Normalize reversed but fully-bounded ranges ("D10:B2" → "B2:D10"); a
    // start>end GridRange is rejected by the API. Open-ended ranges (one side
    // missing a bound) are left untouched.
    if (range.startColumnIndex != null && range.endColumnIndex != null) {
      const a = range.startColumnIndex;
      const b = range.endColumnIndex - 1;
      range.startColumnIndex = Math.min(a, b);
      range.endColumnIndex = Math.max(a, b) + 1;
    }
    if (range.startRowIndex != null && range.endRowIndex != null) {
      const a = range.startRowIndex;
      const b = range.endRowIndex - 1;
      range.startRowIndex = Math.min(a, b);
      range.endRowIndex = Math.max(a, b) + 1;
    }
  } else {
    const ref = parseRef(body);
    if (ref.col != null) {
      range.startColumnIndex = ref.col;
      range.endColumnIndex = ref.col + 1;
    }
    if (ref.row != null) {
      range.startRowIndex = ref.row;
      range.endRowIndex = ref.row + 1;
    }
  }
  return range;
}
