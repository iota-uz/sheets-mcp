# IOTA Finances

Generic Google Sheets automation server (MCP) used to keep IOTA's `Finances_2024_2025` spreadsheet in sync with bank statements, Discord screenshots, and mobile-banking screenshots.

The MCP layer is **sheet-agnostic** — it has no knowledge of "расходы", "Категория", "ФОТ", or any other column name from this specific spreadsheet. Records are keyed by the actual header text of the target sheet, read at runtime. The agent provides everything (column names, category lists, formula templates) — sheet-specific knowledge lives in `CLAUDE.md` and the project's `.claude/skills/` markdown files alongside the business logic.

## Architecture

The agent doesn't issue per-cell tool calls. It writes a JS script against a typed `Sheet` API; the runner compiles each script to **one atomic `spreadsheets.batchUpdate`**.

```
agent JS  →  sheets_exec  →  vm sandbox  →  Sheet.insert  →  batchUpdate  →  Sheets API
                                                                    ↓
                                                            (or: dryRun captures
                                                             the request body)
```

`Sheet.insert` produces a single batch with `insertDimension` + `updateCells` (with formulas referencing the new row's actual position) + an optional `createDeveloperMetadata` idempotency token. One round trip, atomic.

## MCP tools

The MCP server (`mcp/server.mjs`) exposes three tools.

### `sheets_describe`

Return the metadata for one sheet (or list all sheets).

| arg         | type     | default | description                                    |
|-------------|----------|---------|------------------------------------------------|
| `sheet`     | `string` | —       | Sheet name. Omit for spreadsheet-wide list.    |
| `headerRow` | `number` | `1`     | 1-based row containing headers.                |

Returns: `{ sheet, sheetId, headerRow, rowCount, headers: [{index, letter, text}] }`. Or, with no `sheet`: `{ spreadsheetId, sheets: [{title, sheetId}] }`.

### `sheets_exec`

Execute JavaScript against the typed Sheet API in a sandboxed `vm` context. The code is wrapped in `async () => { ... }` and awaited. Whatever you `return` comes back as `result`. `console.log/info/warn/error` are captured into `stdout`/`stderr`.

| arg         | type      | default | description                                            |
|-------------|-----------|---------|--------------------------------------------------------|
| `code`      | `string`  | —       | Required. JS to execute. Has access to `sheets`, `console`, JS builtins. |
| `dryRun`    | `boolean` | `false` | If true, `batchUpdate` calls are recorded into `planned[]` instead of being sent. |
| `timeoutMs` | `number`  | `30000` | Execution timeout in milliseconds.                     |

The `sheets` global:

```ts
sheets.sheet(name: string, opts?: { headerRow?: number }): Promise<Sheet>
sheets.spreadsheetId(): string
```

The `Sheet`:

```ts
sheet.describe(): { sheet, sheetId, headerRow, rowCount, headers }

sheet.insertMany(records, opts?): Promise<{ inserted, skipped, rows }>
   records: Array<{ "Header text": value | "=formula" }>
   opts:    { idempotencyKey?: (record, i) => string,
              format?: StyleObject, dryRun?: boolean }
   N records → ONE batchUpdate. Use this for any batch.

sheet.insert(record, opts?): Promise<{ row, inserted, idempotencyHit }>
   Sugar over insertMany for one record.
   opts: { idempotencyKey?: string, format?: StyleObject, dryRun?: boolean }

sheet.find(where): Promise<Array<{ row, "Header": value, ... }>>
sheet.update({ where?, rows?, set }): Promise<{ updated, rows }>
sheet.delete({ where?, rows? }): Promise<{ deleted, rows }>
sheet.format({ where?, rows?, set }): Promise<{ formatted, rows }>
sheet.setValidation(header, { type, values, strict? }): Promise<{ ok }>
```

Pass `rows: [123, 456]` to update/delete/format to skip the find() lookup when the row numbers are already known.

**Records use the sheet's actual header text** — no role/schema indirection.

**Formula values** start with `=` and may use placeholders:

- `{row}` — the new row's actual number, resolved at write time
- `{col:Header}` — the A1 column letter for the column with that exact header

**Idempotency**: pass an opaque `idempotencyKey` as a string. The MCP stores it as `DeveloperMetadata` (`iota:idempotency` namespace, `PROJECT` visibility) on the new row. Re-running the same insert performs a metadata search first and returns the existing row with `inserted: false, idempotencyHit: <row>` instead of duplicating.

**Validation**: pulled from in-sheet data validation rules. Set them once with `sheet.setValidation(header, { type: "ONE_OF_LIST", values, strict })`. After that, both the spreadsheet UI and `sheet.insert/update` reject invalid values pre-flight.

`StyleObject` keys: `backgroundColor`, `horizontalAlignment`, `numberFormat`, `textFormat: { bold, italic, fontSize, foregroundColor }`, `wrapStrategy`.

### `discord_read_messages`

Read recent messages from the Discord `#finances` channel (configured via `DISCORD_CHANNEL_ID`). Downloads image attachments to `reports/discord/` for visual inspection via `Read`.

| arg     | type     | description                                       |
|---------|----------|---------------------------------------------------|
| `limit` | `number` | Number of messages to fetch (default 10, max 100).|
| `after` | `string` | Only fetch messages after this message ID.        |

## Repo layout

```
.
├── README.md                     this file
├── CLAUDE.md                     spreadsheet-specific knowledge
├── mcp/
│   ├── server.mjs                MCP entrypoint (sheets_describe, sheets_exec, discord)
│   ├── runner.mjs                vm sandbox for sheets_exec
│   ├── sheet.mjs                 Sheet API (insert/find/update/delete/format/setValidation)
│   ├── schema.mjs                generic helpers (style/color/colLetter)
│   ├── sheets-client.mjs         low-level Google Sheets v4 client
│   ├── excel.mjs                 .xlsx parser for bank statements
│   ├── discord.mjs               Discord REST wrapper
│   └── auth.mjs                  Google service account auth
├── scripts/
│   └── reconcile.mjs             month-end balance reconciliation
└── .claude/skills/               Claude skills (markdown business logic)
    ├── balance-check/            triggered by "сверь баланс"
    ├── import-bank-statement/    triggered by "обработай выписку"
    └── import-card-transactions/ triggered by "обработай карту/tenge/uzum"
```

## Setup

1. **Service account.** Create one in Google Cloud → IAM, download a JSON key, place at `auth/service-account.json`. Share the spreadsheet with the SA email (Editor).

2. **`.env`.** Copy `.env.example`:
   ```
   SPREADSHEET_ID=...
   GOOGLE_SERVICE_ACCOUNT_KEY_PATH=auth/service-account.json
   DISCORD_BOT_TOKEN=...      # optional
   DISCORD_CHANNEL_ID=...     # optional
   ```

3. **Install deps.**
   ```bash
   npm install
   ```

4. **MCP config** (`.mcp.json` already wired):
   ```json
   { "mcpServers": { "sheets": { "command": "node", "args": ["mcp/server.mjs"] } } }
   ```

5. **Validation (optional, one-off)** — set dropdowns on enum columns so the sheet itself enforces allowed values:
   ```js
   const расходы = await sheets.sheet("расходы");
   await расходы.setValidation("Категория", {
     type: "ONE_OF_LIST",
     values: ["ФОТ", "Налоги", "EAI Google", "Сервера и подписки", ...],
     strict: true,
   });
   ```

## Examples

**Insert one expense (formulas use placeholders):**

```js
const расходы = await sheets.sheet("расходы");

const RATE = '=SWITCH({col:Валюта}{row}; "USD";1; "SOM";IFERROR(INDEX(GOOGLEFINANCE("CURRENCY:USDUZS";"price";{col:Дата}{row});2;2);INDEX(GOOGLEFINANCE("CURRENCY:USDUZS";"price";WORKDAY({col:Дата}{row};-1));2;2)); "EUR";IFERROR(INDEX(GOOGLEFINANCE("CURRENCY:USDEUR";"price";{col:Дата}{row});2;2);INDEX(GOOGLEFINANCE("CURRENCY:USDEUR";"price";WORKDAY({col:Дата}{row};-1));2;2)); NA())';

return расходы.insert({
  "Дата": "06.04.2026",
  "Категория": "Сервера и подписки",
  "Сумма": 292.50,
  "Курс": RATE,
  "Валюта": "USD",
  "Сумма ($)": "={col:Сумма}{row}/{col:Курс}{row}",
  "Расчетный период": "01.04.2026",
  "Исходящий счет": "Tenge Mastercard",
  "Описание": "Blacksmith Software",
}, { idempotencyKey: "06.04.2026|292.5|Blacksmith" });
```

**Find and update:**

```js
const расходы = await sheets.sheet("расходы");
return расходы.update({
  where: { "Исходящий счет": "Uzum UZS" },         // wrong name
  set:   { "Исходящий счет": "Uzum Visa UZS" },    // correct name
});
```

**Cross-sheet (different `headerRow`):**

```js
const ддс    = await sheets.sheet("ддс", { headerRow: 3 });
const счета  = await sheets.sheet("счета", { headerRow: 2 });
return счета.find({ "Краткое название": "Tenge Mastercard" });
```

**Dry-run preview:**

Pass `dryRun: true` to `sheets_exec` to capture the planned `batchUpdate` request bodies without committing.

## Why this design

The earlier MCP tied itself to this specific spreadsheet — schemas, role mappings, formula templates per sheet, primary keys. Adding a column or renaming a sheet meant editing code in `mcp/`. Sheet-specific knowledge ended up duplicated across schema files and skill prompts.

The current MCP is fully generic. It can drive any Google Sheets spreadsheet without changes. All sheet-specific knowledge — column names, category lists, formula templates, validation values — lives in `CLAUDE.md` and `.claude/skills/`, alongside the business logic that uses it. The agent reads that, composes a script, and the MCP records what the agent says without interpreting it.

Tradeoffs:
- Scripts are slightly more verbose (long formulas pasted per-call)
- Validation rules must be set once per sheet via `setValidation` (then enforced in both UI and MCP)
- Header renames require updating `CLAUDE.md`, not code

In return: the MCP layer is reusable across any spreadsheet, and there's exactly one source of truth per type of knowledge.
