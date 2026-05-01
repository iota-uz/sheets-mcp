#!/usr/bin/env node
/**
 * Bootstrap a sheet's schema into the spreadsheet.
 *
 * Reads schemas/<name>.json, fuzzy-matches declared column headers to the
 * sheet's actual header row, then writes:
 *   - DeveloperMetadata (PROJECT visibility) tagging each column with its
 *     declared role. Survives column reorders. The Table API resolves
 *     roles via this metadata at runtime.
 *   - Number formats + horizontal alignment per column (repeatCell)
 *   - Data validation dropdowns for enum columns with validation:strict
 *   - Header row style (bold, background)
 *   - Frozen header rows
 *   - Conditional format rules
 *
 * Idempotent: deletes prior iota:col:role metadata for the sheet before
 * recreating, so re-running picks up schema edits cleanly.
 *
 * Usage:
 *   node scripts/bootstrap-schema.mjs <sheet> [--dry-run]
 *   node scripts/bootstrap-schema.mjs --all [--dry-run]
 */

import {
  loadSchema,
  listSchemas,
  compileStyle,
  colLetter,
  normHeader,
} from "../mcp/schema.mjs";
import {
  getSpreadsheet,
  batchUpdate,
  valuesGet,
  developerMetadataSearch,
} from "../mcp/sheets-client.mjs";

const META_KEY_ROLE = "iota:col:role";

const args = process.argv.slice(2);
const dryRun = args.includes("--dry-run");
const safeMode = args.includes("--safe-mode");      // metadata + frozen only, skip formatting
const all = args.includes("--all");
const sheetArg = args.find(a => !a.startsWith("--"));

if (!sheetArg && !all) {
  console.error("Usage: bootstrap-schema.mjs <sheet> [--dry-run] | --all [--dry-run]");
  process.exit(1);
}

const targets = all ? listSchemas() : [sheetArg];

for (const name of targets) {
  console.log(`\n=== ${name} ${dryRun ? "(dry-run)" : ""}`);
  await bootstrap(name);
}

async function bootstrap(sheetName) {
  const schema = loadSchema(sheetName);
  const meta = await getSpreadsheet();

  const sheetMeta = meta.sheets.find(
    s => normHeader(s.properties.title) === normHeader(schema.sheet)
  );
  if (!sheetMeta) {
    console.error(`  ERROR: sheet "${schema.sheet}" not found in spreadsheet`);
    return;
  }
  const sheetId = sheetMeta.properties.sheetId;

  // Read header row to fuzzy-match
  const headerRange = `${quoteSheetName(schema.sheet)}!A${schema.headerRow}:ZZ${schema.headerRow}`;
  const rows = await valuesGet(headerRange, { valueRenderOption: "FORMATTED_VALUE" });
  const headers = rows[0] || [];

  const headerToIdx = new Map();
  headers.forEach((h, i) => {
    const k = normHeader(h);
    if (k) headerToIdx.set(k, i);
  });

  // Resolve role → column index
  const roleToCol = {};
  const unresolved = [];
  for (const [role, col] of Object.entries(schema.columns)) {
    const idx = headerToIdx.get(normHeader(col.header));
    if (idx != null) roleToCol[role] = idx;
    else unresolved.push({ role, header: col.header });
  }

  if (unresolved.length > 0) {
    console.log(`  Unresolved (header not found): ${unresolved.map(u => `${u.role}=${u.header}`).join(", ")}`);
  }

  const resolved = Object.keys(roleToCol).length;
  console.log(`  Resolved ${resolved}/${Object.keys(schema.columns).length} columns from headers`);

  // Build the batch
  const requests = [];

  // 1. Delete existing iota:col:role metadata on this sheet (clean slate)
  const existingMeta = await developerMetadataSearch([
    {
      developerMetadataLookup: {
        metadataKey: META_KEY_ROLE,
        locationType: "COLUMN",
        metadataLocation: { sheetId },
      },
    },
  ]);
  const existingForThisSheet = existingMeta.filter(
    m => m.developerMetadata?.location?.dimensionRange?.sheetId === sheetId
  );
  console.log(`  Existing role-tags: ${existingForThisSheet.length} (will be replaced)`);

  for (const m of existingForThisSheet) {
    requests.push({
      deleteDeveloperMetadata: {
        dataFilter: {
          developerMetadataLookup: {
            metadataId: m.developerMetadata.metadataId,
          },
        },
      },
    });
  }

  // 2. Create new role-tags
  for (const [role, colIdx] of Object.entries(roleToCol)) {
    requests.push({
      createDeveloperMetadata: {
        developerMetadata: {
          metadataKey: META_KEY_ROLE,
          metadataValue: role,
          location: {
            dimensionRange: {
              sheetId,
              dimension: "COLUMNS",
              startIndex: colIdx,
              endIndex: colIdx + 1,
            },
          },
          visibility: "PROJECT",
        },
      },
    });
  }

  // 3. Per-column number format + alignment (repeatCell on full data range)
  const dataStartRow = schema.headerRow; // 0-based: header is at row N (1-based) → 0-based index N-1; data starts after
  if (safeMode) {
    console.log(`  Safe-mode: skipping per-column formats, validation, header style`);
  }
  for (const [role, col] of Object.entries(schema.columns)) {
    if (safeMode) break;
    const colIdx = roleToCol[role];
    if (colIdx == null) continue;

    const stylePart = {};
    if (col.numberFormat) stylePart.numberFormat = col.numberFormat;
    if (col.horizontalAlignment) stylePart.horizontalAlignment = col.horizontalAlignment;
    if (Object.keys(stylePart).length === 0) continue;

    const { cellFormat, fields } = compileStyle(stylePart);

    requests.push({
      repeatCell: {
        range: {
          sheetId,
          startRowIndex: dataStartRow, // skip header
          // no endRowIndex → applies to remainder of sheet
          startColumnIndex: colIdx,
          endColumnIndex: colIdx + 1,
        },
        cell: cellFormat,
        fields,
      },
    });
  }

  // 4. Data validation for strict enum columns
  for (const [role, col] of Object.entries(schema.columns)) {
    if (safeMode) break;
    if (col.type !== "enum" || col.validation !== "strict") continue;
    const colIdx = roleToCol[role];
    if (colIdx == null) continue;

    requests.push({
      setDataValidation: {
        range: {
          sheetId,
          startRowIndex: dataStartRow,
          startColumnIndex: colIdx,
          endColumnIndex: colIdx + 1,
        },
        rule: {
          condition: {
            type: "ONE_OF_LIST",
            values: col.values.map(v => ({ userEnteredValue: v })),
          },
          showCustomUi: true,
          strict: true,
        },
      },
    });
  }

  // 5. Header row style
  if (!safeMode && schema.sheetProperties?.headerStyle) {
    const { cellFormat, fields } = compileStyle(schema.sheetProperties.headerStyle);
    requests.push({
      repeatCell: {
        range: {
          sheetId,
          startRowIndex: schema.headerRow - 1,
          endRowIndex: schema.headerRow,
          startColumnIndex: 0,
          // applies to all columns to ZZ; in practice only used columns
        },
        cell: cellFormat,
        fields,
      },
    });
  }

  // 6. Frozen rows
  if (schema.sheetProperties?.frozenRows != null) {
    requests.push({
      updateSheetProperties: {
        properties: {
          sheetId,
          gridProperties: { frozenRowCount: schema.sheetProperties.frozenRows },
        },
        fields: "gridProperties.frozenRowCount",
      },
    });
  }

  // 7. Conditional formats — would go here, deferred until first need

  console.log(`  Planned requests: ${requests.length}`);
  if (dryRun) {
    console.log(JSON.stringify(requests, null, 2).slice(0, 2000));
    if (JSON.stringify(requests).length > 2000) console.log("  ... (truncated)");
    return;
  }

  if (requests.length === 0) {
    console.log("  Nothing to do");
    return;
  }

  await batchUpdate(requests);
  console.log("  Applied");
}

function quoteSheetName(name) {
  if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) return name;
  return `'${name.replace(/'/g, "''")}'`;
}
