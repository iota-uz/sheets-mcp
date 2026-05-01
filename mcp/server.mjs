#!/usr/bin/env node
/**
 * MCP server for IOTA Finances spreadsheet automation.
 *
 * Two sheets-related tools (Playwright-style: agent writes scripts against
 * a typed Table API, runner executes them in a sandbox):
 *   - sheets_describe — return schema(s) for the spreadsheet
 *   - sheets_exec     — run JS using the `sheets` library
 *
 * Plus Discord integration for the #finances channel.
 */

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";

import { exec } from "./runner.mjs";
import { table } from "./table.mjs";
import { listSchemas } from "./schema.mjs";
import { readFinancesChannel } from "./discord.mjs";

// ─── Tool definitions ──────────────────────────────────────────────────

const SHEETS_LIB_DOC = `
The \`sheets\` global exposes:

  sheets.table(name)              → Promise<Table>   (resolves schema + binds columns)
  sheets.listSchemas()            → string[]
  sheets.spreadsheetId()          → string

Table API:
  table.describe()                → schema info (columns, types, formulas, primary key)
  table.insert(record, opts?)     → Promise<{ row, inserted, idempotencyHit? }>
        opts: { idempotencyKey?: string, format?: StyleObject, dryRun?: boolean }
        Atomic: insertDimension + updateCells + idempotency tag in one batchUpdate.
        Computed/formula columns auto-filled — pass only literal fields.
  table.find(where)               → Promise<Array<{row, ...record}>>
  table.update({where, set})      → Promise<{updated, rows}>
  table.format({where, set})      → Promise<{formatted, rows}>

StyleObject keys:
  backgroundColor: "#hex" | "red" | {red,green,blue}
  horizontalAlignment: "LEFT" | "CENTER" | "RIGHT"
  numberFormat: "DD.MM.YYYY" | "#,##0.00" | "[$$-409]#,##0.00" | { type, pattern }
  textFormat: { bold, italic, fontSize, foregroundColor }
  wrapStrategy: "WRAP" | "OVERFLOW_CELL" | "CLIP"

Records use schema roles, not header text. E.g. for "расходы":
  { date, category, amount, currency, account, description }

Use sheets_describe to inspect a sheet's schema before writing scripts.
`.trim();

const tools = [
  {
    name: "sheets_describe",
    description:
      "Describe one or all sheet schemas. Returns column roles, types, headers, " +
      "formulas, primary key, and required vs optional fields. Call this before " +
      "writing a sheets_exec script so you know what fields to pass.",
    inputSchema: {
      type: "object",
      properties: {
        sheet: { type: "string", description: "Sheet name. Omit for list of all schemas." },
      },
    },
  },
  {
    name: "sheets_exec",
    description:
      "Run JavaScript against the typed Table API.\n\n" + SHEETS_LIB_DOC + "\n\n" +
      "Set dryRun: true to capture the planned batchUpdate without committing. " +
      "The script is wrapped in `async () => { <your code> }` and awaited; whatever " +
      "you `return` comes back as `result`. Console.log/info/warn/error are captured.",
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
    if (a?.sheet) {
      const t = await table(a.sheet);
      return t.describe();
    }
    const names = listSchemas();
    const all = {};
    for (const name of names) {
      try {
        const t = await table(name);
        all[name] = t.describe();
      } catch (e) {
        all[name] = { error: e.message };
      }
    }
    return all;
  },

  sheets_exec: async (a) => {
    if (typeof a?.code !== "string") throw new Error("`code` must be a string");
    return exec(a.code, { dryRun: !!a.dryRun, timeoutMs: a.timeoutMs ?? 30000 });
  },

  discord_read_messages: (a) => readFinancesChannel({ limit: a?.limit, after: a?.after }),
};

// ─── Server bootstrap ──────────────────────────────────────────────────

const server = new Server(
  { name: "sheets", version: "3.0.0" },
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
console.error("[sheets-mcp] Server v3 started — sheets_describe + sheets_exec + discord_read_messages");
