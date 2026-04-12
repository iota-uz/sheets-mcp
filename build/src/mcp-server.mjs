#!/usr/bin/env node
/**
 * IOTA Sheets MCP Server.
 *
 * Provides generic Google Sheets operations + Discord integration
 * for the IOTA finance tracking workflow.
 *
 * Tools:
 *   - iota_sheets_read      — read rows from any sheet
 *   - iota_sheets_append    — append a row to any sheet
 *   - iota_sheets_update    — update a single cell
 *   - iota_sheets_info      — list sheets + preview headers
 *   - iota_next_iteration   — get next iteration number for a project
 *   - iota_discord_read_finances — read messages from Discord #finances
 */

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";

import { read, append, update, info, getNextIteration } from "./sheets.mjs";
import { readFinancesChannel } from "./discord.mjs";

// ───────────────────────────────────────────────────────────────────────────
// Tool definitions
// ───────────────────────────────────────────────────────────────────────────

const tools = [
  {
    name: "iota_sheets_read",
    description:
      "Read rows from a Google Sheets tab. Can read a specific range or the last N rows.",
    inputSchema: {
      type: "object",
      properties: {
        sheet: { type: "string", description: "Sheet name, e.g. 'расходы', 'платежи', 'ддс', 'итерации'" },
        range: { type: "string", description: "Optional A1 range within the sheet, e.g. 'B1:K100'. Defaults to full sheet." },
        last_n: { type: "number", description: "If set, return only the last N rows (useful for checking recent entries)." },
      },
      required: ["sheet"],
    },
  },
  {
    name: "iota_sheets_append",
    description:
      "Append a single row to a Google Sheets tab. Pass the target column range and an array of cell values.",
    inputSchema: {
      type: "object",
      properties: {
        sheet: { type: "string", description: "Sheet name, e.g. 'расходы'" },
        range: { type: "string", description: "Column range to append to, e.g. 'B:K' for расходы, 'A:K' for платежи, 'A:H' for ддс" },
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
    name: "iota_sheets_update",
    description: "Update a single cell in a Google Sheets tab.",
    inputSchema: {
      type: "object",
      properties: {
        sheet: { type: "string", description: "Sheet name" },
        cell: { type: "string", description: "Cell address, e.g. 'C2230'" },
        value: { description: "New value (string or number)" },
      },
      required: ["sheet", "cell", "value"],
    },
  },
  {
    name: "iota_sheets_info",
    description:
      "Get spreadsheet metadata: list of all sheet names. Optionally preview the first 3 rows of a specific sheet.",
    inputSchema: {
      type: "object",
      properties: {
        sheet: { type: "string", description: "Optional: sheet name to preview first 3 rows of." },
      },
    },
  },
  {
    name: "iota_next_iteration",
    description:
      "Get the next iteration number for a project. Reads the 'итерации' sheet, finds the highest existing number for the given project, and returns { lastNum, nextNum, nextId }.",
    inputSchema: {
      type: "object",
      properties: {
        project: { type: "string", description: "Project name, e.g. 'Granite', 'Takhir', 'EAI_Website'" },
      },
      required: ["project"],
    },
  },
  {
    name: "iota_discord_read_finances",
    description:
      "Read recent messages from the Discord #finances channel. Downloads image attachments to Reports/discord/ for visual inspection via Read tool.",
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
  iota_sheets_read: (a) => read(a.sheet, a.range, a.last_n),
  iota_sheets_append: (a) => append(a.sheet, a.range, a.values),
  iota_sheets_update: (a) => update(a.sheet, a.cell, a.value),
  iota_sheets_info: (a) => info(a?.sheet),
  iota_next_iteration: (a) => getNextIteration(a.project),
  iota_discord_read_finances: (a) => readFinancesChannel({ limit: a?.limit, after: a?.after }),
};

// ───────────────────────────────────────────────────────────────────────────
// Server setup
// ───────────────────────────────────────────────────────────────────────────

const server = new Server(
  { name: "iota-sheets", version: "1.0.0" },
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
console.error("[iota-sheets-mcp] Server started, awaiting requests on stdio");
