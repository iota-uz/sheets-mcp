# @iota-uz/sheets-mcp

Generic Model Context Protocol server for Google Sheets — a typed, sheet-agnostic `Sheet` API that compiles agent-authored scripts to atomic `spreadsheets.batchUpdate` calls with built-in idempotency.

No hardcoded sheet names, columns, or schemas. Records are keyed by the actual header text of the target sheet, read at runtime. All spreadsheet-specific knowledge (column names, category lists, formula templates, validation values) lives in the *consumer's* `CLAUDE.md` / `.claude/skills/`, not in this package.

## Install

In any project's `.mcp.json`:

```json
{
  "mcpServers": {
    "sheets": {
      "command": "npx",
      "args": ["-y", "@iota-uz/sheets-mcp@latest"],
      "env": {
        "GOOGLE_APPLICATION_CREDENTIALS": "/absolute/path/to/service-account.json"
      }
    }
  }
}
```

Optional env (only if you use `discord_read_messages`):

```
DISCORD_BOT_TOKEN
DISCORD_CHANNEL_ID
DISCORD_IMAGES_DIR   # default: <cwd>/reports/discord
```

## Setup (one time)

1. **GCP service account** — In Google Cloud → IAM, create a service account, download a JSON key. No special roles needed; the SA only needs access to spreadsheets you share with it.
2. **Enable APIs** — Enable the Google Sheets API (and Drive API if you'll use Drive features) on the project.
3. **Share each spreadsheet** with the service account's email (e.g. `mcp@project.iam.gserviceaccount.com`), Editor role.
4. **Point `GOOGLE_APPLICATION_CREDENTIALS`** at the JSON key (absolute path).

## Tools

### `sheets_describe`

Return the metadata for one sheet (or list all sheets in a spreadsheet).

| arg             | type     | default | description                                    |
|-----------------|----------|---------|------------------------------------------------|
| `spreadsheetId` | `string` | —       | **Required.** Target Google Sheets ID.         |
| `sheet`         | `string` | —       | Sheet name. Omit for spreadsheet-wide list.    |
| `headerRow`     | `number` | `1`     | 1-based row containing headers.                |

Returns: `{ sheet, sheetId, headerRow, rowCount, headers: [{index, letter, text}] }`. Or, with no `sheet`: `{ spreadsheetId, sheets: [{title, sheetId}] }`.

### `sheets_exec`

Execute JavaScript against the typed Sheet API in a sandboxed `vm` context. The code is wrapped in `async () => { ... }` and awaited. Whatever you `return` comes back as `result`. `console.log/info/warn/error` are captured into `stdout`/`stderr`.

| arg             | type      | default | description                                            |
|-----------------|-----------|---------|--------------------------------------------------------|
| `spreadsheetId` | `string`  | —       | **Required.** Bound to the `sheets` global for this call. |
| `code`          | `string`  | —       | **Required.** JS to execute. Has access to `sheets`, `console`, JS builtins. |
| `dryRun`        | `boolean` | `false` | If true, `batchUpdate` calls are recorded into `planned[]` instead of being sent. |
| `timeoutMs`     | `number`  | `30000` | Execution timeout in milliseconds.                     |

The `sheets` global (bound to the call's `spreadsheetId`):

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

sheet.readRange(a1, { valueRender? }): Promise<any[][]>
sheet.writeRange(a1, values2d, { raw? }): Promise<{ updatedRange, updatedCells, ... }>
```

Pass `rows: [123, 456]` to `update/delete/format` to skip the `find()` lookup when row numbers are already known.

**Records use the sheet's actual header text** — no role/schema indirection.

**Formula values** start with `=` and may use placeholders:

- `{row}` — the new row's actual number, resolved at write time
- `{col:Header}` — the A1 column letter for the column with that exact header

**Idempotency**: pass an opaque `idempotencyKey` as a string. The MCP stores it as `DeveloperMetadata` (`iota:idempotency` namespace, `PROJECT` visibility) on the new row. Re-running the same insert performs a metadata search first and returns the existing row with `inserted: false, idempotencyHit: <row>` instead of duplicating.

**Validation**: pulled from in-sheet data validation rules. Set them once with `sheet.setValidation(header, { type: "ONE_OF_LIST", values, strict })`. After that, both the spreadsheet UI and `sheet.insert/update` reject invalid values pre-flight.

`StyleObject` keys: `backgroundColor`, `horizontalAlignment`, `numberFormat`, `textFormat: { bold, italic, fontSize, foregroundColor }`, `wrapStrategy`.

### `discord_read_messages`

Read recent messages from a Discord channel (configured via `DISCORD_BOT_TOKEN` and `DISCORD_CHANNEL_ID` env vars). Downloads image attachments to `$DISCORD_IMAGES_DIR` (default `<cwd>/reports/discord`) for visual inspection via `Read`.

| arg     | type     | description                                       |
|---------|----------|---------------------------------------------------|
| `limit` | `number` | Number of messages to fetch (default 10, max 100).|
| `after` | `string` | Only fetch messages after this message ID.        |

## Examples

**Insert one row (formulas use placeholders):**

```js
const tx = await sheets.sheet("transactions");

const RATE = '=SWITCH({col:Currency}{row}; "USD";1; "EUR";IFERROR(INDEX(GOOGLEFINANCE("CURRENCY:USDEUR";"price";{col:Date}{row});2;2)); NA())';

return tx.insert({
  "Date": "06.04.2026",
  "Amount": 292.50,
  "Currency": "USD",
  "Rate": RATE,
  "Amount ($)": "={col:Amount}{row}/{col:Rate}{row}",
  "Description": "Server hosting",
}, { idempotencyKey: "2026-04-06|292.5|hosting" });
```

**Find and update:**

```js
const tx = await sheets.sheet("transactions");
return tx.update({
  where: { "Vendor": "Acme Co" },
  set:   { "Vendor": "Acme Corp" },
});
```

**Cross-sheet (different `headerRow`):**

```js
const ledger = await sheets.sheet("ledger", { headerRow: 3 });
const accts  = await sheets.sheet("accounts", { headerRow: 2 });
return accts.find({ "Short name": "Main USD" });
```

**Dry-run preview:**

Pass `dryRun: true` to `sheets_exec` to capture the planned `batchUpdate` request bodies without committing.

## Architecture

The agent doesn't issue per-cell tool calls. It writes a JS script against the typed `Sheet` API; the runner compiles each script to **one atomic `spreadsheets.batchUpdate`**.

```
agent JS  →  sheets_exec  →  vm sandbox  →  Sheet.insert  →  batchUpdate  →  Sheets API
                                                                    ↓
                                                            (or: dryRun captures
                                                             the request body)
```

`Sheet.insert` produces a single batch with `insertDimension` + `updateCells` (with formulas referencing the new row's actual position) + an optional `createDeveloperMetadata` idempotency token. One round trip, atomic.

## Release (maintainers)

```bash
npm version patch  # or minor / major
git push --follow-tags
```

The `publish` GitHub Action picks up the tag, runs `npm ci && npm publish`, and ships to npm using the `NPM_TOKEN` repo secret.
