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
 *   - discord_read_messages   #finances channel
 */

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";

import { exec } from "./runner.mjs";
import { sheet } from "./sheet.mjs";
import { getSpreadsheet } from "./sheets-client.mjs";
import { readFinancesChannel } from "./discord.mjs";

// ─── Tool definitions ──────────────────────────────────────────────────

const SHEETS_LIB_DOC = `
The \`sheets\` global exposes:

  sheets.sheet(name, { headerRow?: 1 })   → Promise<Sheet>
  sheets.spreadsheetId()                  → string

Sheet API:
  sheet.describe()                        → { sheet, sheetId, headerRow, rowCount, headers: [{index, letter, text}] }
  sheet.insert(record, opts?)             → Promise<{ row, inserted, idempotencyHit }>
        record: { "Header text": value | "=formula" }
        opts:   { idempotencyKey?: string, format?: StyleObject, dryRun?: boolean }
        Atomic — insertDimension + updateCells + idempotency tag in one batchUpdate.
        Formula values can use {row} and {col:HeaderName} placeholders.
        Validation runs against in-sheet data validation rules; configure with setValidation.
  sheet.find(where)                       → Promise<Array<{ row, "Header": value, ... }>>
  sheet.update({ where, set })            → Promise<{ updated, rows }>
  sheet.delete({ where })                 → Promise<{ deleted, rows }>
  sheet.format({ where, set })            → Promise<{ formatted, rows }>
  sheet.setValidation(header, spec)       → set ONE_OF_LIST validation on a column
        spec: { type: "ONE_OF_LIST", values: string[], strict?: boolean }

StyleObject keys:
  backgroundColor: "#hex" | "red" | {red,green,blue}
  horizontalAlignment: "LEFT" | "CENTER" | "RIGHT"
  numberFormat: "DD.MM.YYYY" | "#,##0.00" | "[$$-409]#,##0.00" | { type, pattern }
  textFormat: { bold, italic, fontSize, foregroundColor }
  wrapStrategy: "WRAP" | "OVERFLOW_CELL" | "CLIP"

Records use the sheet's actual header text — no role/schema indirection.
The MCP itself has no knowledge of any specific spreadsheet structure;
sheet-specific knowledge (column lists, categorization rules, formula
templates) lives in CLAUDE.md / SKILL.md.

Use sheets_describe to see headers and sheet metadata before writing scripts.
`.trim();

const tools = [
  {
    name: "sheets_describe",
    description:
      "Describe a Google Sheet — return its sheetId, header row, row count, " +
      "and the actual header texts with their column indices/letters. With no " +
      "argument, lists all sheets in the spreadsheet.",
    inputSchema: {
      type: "object",
      properties: {
        sheet: { type: "string", description: "Sheet name. Omit for spreadsheet-wide list of sheet titles." },
        headerRow: { type: "number", description: "1-based row containing headers (default 1)." },
      },
    },
  },
  {
    name: "sheets_exec",
    description:
      "Run JavaScript against the typed Sheet API.\n\n" + SHEETS_LIB_DOC + "\n\n" +
      "Set dryRun: true to capture the planned batchUpdate request bodies without committing. " +
      "The script is wrapped in `async () => { <your code> }` and awaited; whatever you `return` " +
      "comes back as `result`. Console.log/info/warn/error are captured.",
    inputSchema: {
      type: "object",
      properties: {
        code: { type: "string", description: "JS code to execute. Has access to `sheets`, `console`." },
        dryRun: { type: "boolean", description: "If true, batchUpdate calls are recorded, not sent." },
        timeoutMs: { type: "number", description: "Execution timeout in milliseconds (default 30000)." },
      },
      required: ["code"],
    },
  },
  {
    name: "discord_read_messages",
    description:
      "Read recent messages from the Discord #finances channel (configured via DISCORD_CHANNEL_ID). " +
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

// ─── Handlers ──────────────────────────────────────────────────────────

const handlers = {
  sheets_describe: async (a) => {
    if (!a?.sheet) {
      const meta = await getSpreadsheet();
      return {
        spreadsheetId: meta.spreadsheetId,
        sheets: meta.sheets.map(s => ({ title: s.properties.title, sheetId: s.properties.sheetId })),
      };
    }
    const s = await sheet(a.sheet, { headerRow: a.headerRow });
    return s.describe();
  },

  sheets_exec: async (a) => {
    if (typeof a?.code !== "string") throw new Error("`code` must be a string");
    return exec(a.code, { dryRun: !!a.dryRun, timeoutMs: a.timeoutMs ?? 30000 });
  },

  discord_read_messages: (a) => readFinancesChannel({ limit: a?.limit, after: a?.after }),
};

// ─── Server bootstrap ──────────────────────────────────────────────────

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
