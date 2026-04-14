#!/usr/bin/env node
/**
 * MCP сервер для IOTA-проекта.
 *
 * Предоставляет generic Google Sheets тулы + Discord-интеграцию.
 * В сервере НЕТ никакого хардкода под конкретную структуру таблицы —
 * только базовые API-операции.
 *
 * Тулы:
 *   - sheets_read             — прочитать строки любого листа
 *   - sheets_append           — дописать строку в любой лист
 *   - sheets_update           — обновить одну ячейку
 *   - sheets_info             — список листов, превью первых 3 строк
 *   - discord_read_messages   — прочитать сообщения Discord-канала
 */

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";

import { read, append, update, info, clear } from "./sheets.mjs";
import { readFinancesChannel } from "./discord.mjs";

// ───────────────────────────────────────────────────────────────────────────
// Tool definitions
// ───────────────────────────────────────────────────────────────────────────

const tools = [
  {
    name: "sheets_read",
    description:
      "Read rows from a Google Sheets tab. Can read a specific A1 range or the last N rows.",
    inputSchema: {
      type: "object",
      properties: {
        sheet:  { type: "string", description: "Sheet name" },
        range:  { type: "string", description: "Optional A1 range within the sheet, e.g. 'B1:K100'." },
        last_n: { type: "number", description: "If set, return only the last N rows." },
      },
      required: ["sheet"],
    },
  },
  {
    name: "sheets_append",
    description:
      "Append a single row to a Google Sheets tab. Pure append — no auto-fill, no duplicate detection. " +
      "Pass the target A1 column range and a positional array of cell values. " +
      "If you need auto-fill of derived columns (rate/period/USD) and duplicate detection, use the smart-append helper script instead: " +
      "`node scripts/smart-append.mjs --sheet=<name> --values=<JSON>`.",
    inputSchema: {
      type: "object",
      properties: {
        sheet:  { type: "string", description: "Sheet name" },
        range:  { type: "string", description: "Column range to append to, e.g. 'B:K'" },
        values: {
          type: "array",
          items: {},
          description: "Array of cell values for one row, in column order.",
        },
      },
      required: ["sheet", "range", "values"],
    },
  },
  {
    name: "sheets_update",
    description: "Update a single cell in a Google Sheets tab.",
    inputSchema: {
      type: "object",
      properties: {
        sheet: { type: "string", description: "Sheet name" },
        cell:  { type: "string", description: "Cell address, e.g. 'C2230'" },
        value: { description: "New value (string, number, or formula starting with =)" },
        raw:   { type: "boolean", description: "If true, uses RAW input mode (preserves literal text, no date/formula parsing). Use when Sheets reformats your dates to unwanted format." },
      },
      required: ["sheet", "cell", "value"],
    },
  },
  {
    name: "sheets_clear",
    description: "Clear values in a range (cells become empty, rows stay). Useful to erase bad data before re-writing.",
    inputSchema: {
      type: "object",
      properties: {
        sheet: { type: "string", description: "Sheet name" },
        range: { type: "string", description: "A1 range to clear, e.g. 'B2070:K2113'" },
      },
      required: ["sheet", "range"],
    },
  },
  {
    name: "sheets_info",
    description:
      "Get spreadsheet metadata: list of all sheet names. Optionally preview the first 3 rows of a specific sheet (useful for discovering headers/structure).",
    inputSchema: {
      type: "object",
      properties: {
        sheet: { type: "string", description: "Optional sheet name to preview first 3 rows." },
      },
    },
  },
  {
    name: "discord_read_messages",
    description:
      "Read recent messages from the Discord channel (channel ID configured in .env). " +
      "Downloads image attachments to reports/discord/ for visual inspection via Read.",
    inputSchema: {
      type: "object",
      properties: {
        limit: { type: "number", description: "Number of messages to fetch (default 10, max 100)" },
        after: { type: "string", description: "Only fetch messages after this message ID (for pagination)." },
      },
    },
  },
];

// ───────────────────────────────────────────────────────────────────────────
// Tool dispatch
// ───────────────────────────────────────────────────────────────────────────

const handlers = {
  sheets_read:           (a) => read(a.sheet, a.range, a.last_n),
  sheets_append:         (a) => append(a.sheet, a.range, a.values),
  sheets_update:         (a) => update(a.sheet, a.cell, a.value, a?.raw),
  sheets_clear:          (a) => clear(a.sheet, a.range),
  sheets_info:           (a) => info(a?.sheet),
  discord_read_messages: (a) => readFinancesChannel({ limit: a?.limit, after: a?.after }),
};

// ───────────────────────────────────────────────────────────────────────────
// Server setup
// ───────────────────────────────────────────────────────────────────────────

const server = new Server(
  { name: "sheets", version: "2.0.0" },
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
    return { content: [{ type: "text", text: `ERROR: ${err.message}` }], isError: true };
  }
});

const transport = new StdioServerTransport();
await server.connect(transport);
console.error("[sheets-mcp] Server started, awaiting requests on stdio");
