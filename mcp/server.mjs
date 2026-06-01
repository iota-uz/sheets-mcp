#!/usr/bin/env node
/**
 * Generic Google Sheets MCP server.
 *
 * The MCP exposes a typed but sheet-agnostic Sheet API — no hardcoded
 * column names, no schema files. Records are keyed by the actual header
 * text of the target sheet. Validation rules are pulled from in-sheet
 * data validation set via setValidation().
 *
 * Tools:
 *   - sheets_describe         metadata + headers for a sheet
 *   - sheets_exec             run JS scripts against the Sheet API
 *   - discord_read_messages   read a configured Discord channel
 */

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";

import { exec } from "./runner.mjs";
import { createSheet } from "./sheet.mjs";
import { makeClient } from "./sheets-client.mjs";
import { readChannelMessages } from "./discord.mjs";

const SHEETS_LIB_DOC = `
The \`sheets\` global is bound to the spreadsheetId you passed in:

  sheets.sheet(name, { headerRow?: 1 })   → Promise<Sheet>
  sheets.spreadsheetId()                  → string

  Structural ops (create/manage tabs):
  sheets.addSheet(title, opts?)           → { sheetId, title }
        opts: { rows?, cols?, index?, tabColor?, frozenRows?, frozenCols? }
  sheets.ensureSheet(title, opts?)        → { sheetId, title, existed } — idempotent addSheet (re-runnable)
  sheets.deleteSheet(nameOrId)            → { ok }
  sheets.renameSheet(nameOrId, newTitle) → { ok }
  sheets.duplicateSheet(nameOrId, newTitle?) → { sheetId }

  Tables (native Sheets v4 Table — typed columns, dropdowns, banding):
  sheets.addTable(name, "Sheet!A1:I", { columns })  → { tableId, name }
        columns: [{ name, type?, values? }]  type ∈ TEXT|DOUBLE(NUMBER)|CURRENCY|PERCENT|
        DATE|TIME|DATE_TIME|BOOLEAN|DROPDOWN;  values:[…] ⇒ DROPDOWN + ONE_OF_LIST rule
  sheets.ensureTable(name, "Sheet!A1:I", { columns })  → { tableId, name, existed } — idempotent
  sheets.updateTable(nameOrId, { name?, columns?, range? })  → { ok, tableId }
  sheets.deleteTable(nameOrId)            → { ok, tableId }
  sheet.toTable(name, { columns?, rows? }) → wrap THIS tab's header range as a Table

  Raw escape hatch (full Sheets v4 power — compile your own requests):
  sheets.batchUpdate(requests)           → runs any Sheets v4 Request[] (dry-run aware)
  sheets.valuesBatchGet(ranges, opts?)   → multi-range read in one round trip
  sheets.getSpreadsheet(opts?)           → spreadsheet metadata (all sheets, named ranges, …)
  sheets.developerMetadataSearch(filters)→ dev-metadata lookup

Sheet API:
  sheet.describe()                            → { sheet, sheetId, headerRow, rowCount, headers: [{index, letter, text}] }
        + table: { tableId, name, columns: [{name, type, letter}] }  when the tab has a native Table

  sheet.insertMany(records, opts?)            → Promise<{ inserted, skipped, rows }>
        records: Array<{ "Header": value | "=formula" }>
        opts:    { idempotencyKey?: (record, i) => string, format?: StyleObject }
        ALL records compile to ONE batchUpdate (insertDimension + updateCells +
        per-row idempotency tokens). Use this for batches — it's ~N× faster
        than a loop of insert().

  sheet.insert(record, opts?)                 → Promise<{ row, inserted, idempotencyHit }>
        Sugar over insertMany for one record.
        opts: { idempotencyKey?: string, format?: StyleObject }

        Formula values use {row} and {col:HeaderName} placeholders.
        Validation runs against in-sheet data validation rules (configure via setValidation).

  sheet.find(where)                           → Promise<Array<{ row, "Header": value, ... }>>

  sheet.update({ where?, rows?, set })        → Promise<{ updated, rows }>
        Pass \`rows: [123, 456]\` to skip find() when you already know the rows.
  sheet.delete({ where?, rows? })             → Promise<{ deleted, rows }>
  sheet.format({ where?, rows?, range?, set })→ Promise<{ formatted, rows }>
        where/rows style whole rows; pass \`range\` (A1) to style an arbitrary range.
  sheet.formatRange(a1, style)                → style any range/column ("B2:D10", "C:C").

  Presentation & analysis sugar (each → one atomic batchUpdate, dry-run aware):
  sheet.merge(range, type?)  sheet.unmerge(range)        type: MERGE_ALL|MERGE_COLUMNS|MERGE_ROWS
  sheet.setBorders(range, spec)              spec: "all" | "outer" | "DASHED" | { top?, bottom?, …, style?, color? }
  sheet.addConditionalFormat(range, rule)    rule: { condition, format } | { gradient: { min, mid?, max } }
  sheet.freeze({ rows?, cols? })
  sheet.resizeColumns(range, { width } | { auto: true })   range: "B:D"
  sheet.insertColumns(at, count?)  sheet.deleteColumns(at, count?)   at: header | letter | index
  sheet.sort(range, [{ column, order?: "ASC"|"DESC" }])
  sheet.setFilter(range)
  sheet.findReplace(find, replace, opts?)    opts: { range?, allSheets?, matchCase?, searchByRegex?, … } (this sheet by default)
  sheet.setNote(cell, text)
  sheet.setRowHeight(rowOrA1, px)             rowOrA1: 22 | "22" | "22:24"
  sheet.copyFormat(srcA1, dstA1)             copy ONLY formatting ("A18:G18" → "A22:G22")

  sheet.setValidation(header, spec)           → set/clear a column's data validation
        spec.type ∈ ONE_OF_LIST { values } | ONE_OF_RANGE { range } | BOOLEAN {} |
        NUMBER_BETWEEN { min, max } | NUMBER_GREATER/LESS/EQ { value } |
        DATE_BETWEEN { min, max } | DATE_AFTER/BEFORE { value } |
        TEXT_CONTAINS/EQ { value } | CUSTOM_FORMULA { formula } | "clear"
        (+ optional strict?, showCustomUi?)
        On a native-Table typed column this auto-routes to updateTable (raw
        setDataValidation is rejected on typed columns). Dropdown chip↔arrow
        display is a UI-only toggle (not in the API) — Table dropdowns show chips.

  sheet.readRange(a1, opts?)                  → Promise<any[][]>
        Raw A1 read (e.g. "A2:B90"). opts.valueRender:
        "UNFORMATTED_VALUE" (default) | "FORMATTED_VALUE" | "FORMULA".
  sheet.readFormatting(a1, opts?)             → Promise<{ data, merges }>
        Read formatting, notes, merges, validation & hyperlinks for a range.
  sheet.writeRange(a1, values, opts?)         → Promise<{ updatedRange, updatedCells, ... }>
        Raw A1 write of a 2D array. opts.raw=true to skip USER_ENTERED parsing.
        Cells holding a Table structured ref (=…[@[Col]]…) auto-route through
        updateCells so they bind in the row's table context instead of #ERROR!;
        opts.bind=true forces that path. (Honors dryRun — captured, not committed.)

Volatile cells (GOOGLEFINANCE, IMPORT*, historical lookups) compute only in the
browser — reading them via the API often returns #N/A, and a write can invalidate
a previously-cached value. That's expected, NOT a server error; verify such values
in the UI rather than treating #N/A as a failure.

StyleObject keys:
  backgroundColor: "#hex" | "red" | {red,green,blue}
  horizontalAlignment: "LEFT" | "CENTER" | "RIGHT"
  numberFormat: "DD.MM.YYYY" | "#,##0.00" | "[$$-409]#,##0.00" | { type, pattern }
  textFormat: { bold, italic, fontSize, foregroundColor }
  wrapStrategy: "WRAP" | "OVERFLOW_CELL" | "CLIP"

Records use the sheet's actual header text — no role/schema indirection.
The MCP itself has no knowledge of any specific spreadsheet structure;
sheet-specific knowledge (column lists, categorization rules, formula
templates) lives in the consumer's CLAUDE.md / SKILL.md.

Use sheets_describe to see headers and sheet metadata before writing scripts.
`.trim();

const tools = [
  {
    name: "sheets_describe",
    description:
      "Describe a Google Sheet — return its sheetId, header row, row count, " +
      "and the actual header texts with their column indices/letters. With no " +
      "`sheet` argument, lists all sheets in the spreadsheet.",
    inputSchema: {
      type: "object",
      properties: {
        spreadsheetId: { type: "string", description: "Target Google Sheets spreadsheet ID." },
        sheet: { type: "string", description: "Sheet name. Omit for spreadsheet-wide list of sheet titles." },
        headerRow: { type: "number", description: "1-based row containing headers (default 1)." },
      },
      required: ["spreadsheetId"],
    },
  },
  {
    name: "sheets_exec",
    description:
      "Run JavaScript against the typed Sheet API.\n\n" + SHEETS_LIB_DOC + "\n\n" +
      "Set dryRun: true to capture every intended mutation without committing — returned as " +
      "`plannedOps`, an ordered log of every intended mutation (batchUpdate request bodies and " +
      "writeRange value writes), each tagged with its `kind`. " +
      "The script is wrapped in `async () => { <your code> }` and awaited; whatever you `return` " +
      "comes back as `result`. Console.log/info/warn/error are captured.",
    inputSchema: {
      type: "object",
      properties: {
        spreadsheetId: { type: "string", description: "Target Google Sheets spreadsheet ID. Bound to the `sheets` global for this call." },
        code: { type: "string", description: "JS code to execute. Has access to `sheets`, `console`." },
        dryRun: { type: "boolean", description: "If true, every mutating call (batchUpdate + value writes) is recorded into plannedOps, not sent." },
        timeoutMs: { type: "number", description: "Execution timeout in milliseconds (default 30000)." },
      },
      required: ["spreadsheetId", "code"],
    },
  },
  {
    name: "discord_read_messages",
    description:
      "Read recent messages from a Discord channel (channel + bot configured via " +
      "DISCORD_BOT_TOKEN and DISCORD_CHANNEL_ID env vars). " +
      "Downloads image attachments to reports/discord/ for visual inspection via Read.",
    inputSchema: {
      type: "object",
      properties: {
        limit: { type: "number", description: "Number of messages to fetch (default 10, max 100)." },
        after: { type: "string", description: "Only fetch messages after this message ID." },
      },
    },
  },
];

const handlers = {
  sheets_describe: async (a) => {
    if (!a?.spreadsheetId) throw new Error("`spreadsheetId` is required");
    const client = makeClient();
    if (!a?.sheet) {
      const meta = await client.getSpreadsheet(a.spreadsheetId);
      return {
        spreadsheetId: meta.spreadsheetId,
        sheets: meta.sheets.map(s => ({ title: s.properties.title, sheetId: s.properties.sheetId })),
      };
    }
    const s = await createSheet(a.spreadsheetId, a.sheet, { headerRow: a.headerRow }, client);
    return s.describe();
  },

  sheets_exec: async (a) => {
    if (!a?.spreadsheetId) throw new Error("`spreadsheetId` is required");
    if (typeof a?.code !== "string") throw new Error("`code` must be a string");
    return exec(a.spreadsheetId, a.code, { dryRun: !!a.dryRun, timeoutMs: a.timeoutMs ?? 30000 });
  },

  discord_read_messages: (a) => readChannelMessages({ limit: a?.limit, after: a?.after }),
};

const server = new Server(
  { name: "sheets", version: "4.0.0" },
  { capabilities: { tools: {} } },
);

server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools }));

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;
  const handler = handlers[name];
  if (!handler) {
    return { content: [{ type: "text", text: `Unknown tool: ${name}` }], isError: true };
  }
  try {
    const result = await handler(args ?? {});
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  } catch (err) {
    return {
      content: [{ type: "text", text: `ERROR: ${err.message}\n${err.stack ?? ""}` }],
      isError: true,
    };
  }
});

const transport = new StdioServerTransport();
await server.connect(transport);
console.error("[sheets-mcp] Server v4 started — generic Sheet API (sheets_describe + sheets_exec)");
