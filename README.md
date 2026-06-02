# @iota-uz/sheets-mcp

Generic Model Context Protocol server for Google Sheets — a typed, sheet-agnostic `Sheet` API that compiles agent-authored scripts to atomic `spreadsheets.batchUpdate` calls with built-in idempotency.

No hardcoded sheet names, columns, or schemas. Records are keyed by the actual header text of the target sheet, read at runtime. All spreadsheet-specific knowledge (column names, category lists, formula templates, validation values) lives in the *consumer's* `CLAUDE.md` / `.claude/skills/`, not in this package.

## Install

### Claude Code

```bash
claude mcp add sheets \
  --env GOOGLE_APPLICATION_CREDENTIALS=/absolute/path/to/service-account.json \
  -- npx -y @iota-uz/sheets-mcp@latest
```

Add `--scope user` to make it available across all projects (default scope is `local` — this project only). Pass extra `--env KEY=VALUE` flags for the optional Discord vars below.

### Manual (`.mcp.json`)

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
| `dryRun`        | `boolean` | `false` | If true, every mutation is recorded — not sent — into `plannedOps[]`, an ordered log of each intended op (batchUpdate bodies + `writeRange` value writes), tagged with its `kind`. |
| `timeoutMs`     | `number`  | `30000` | Execution timeout in milliseconds.                     |

The `sheets` global (bound to the call's `spreadsheetId`):

```ts
sheets.sheet(name: string, opts?: { headerRow?: number }): Promise<Sheet>
sheets.spreadsheetId(): string

// Structural ops — create / manage tabs (no human in the browser needed)
sheets.addSheet(title, opts?): Promise<{ sheetId, title }>
   opts: { rows?, cols?, index?, tabColor?, frozenRows?, frozenCols? }
sheets.ensureSheet(title, opts?): Promise<{ sheetId, title, existed }>   // idempotent — re-runnable
sheets.deleteSheet(nameOrId): Promise<{ ok }>
sheets.renameSheet(nameOrId, newTitle): Promise<{ ok }>
sheets.duplicateSheet(nameOrId, newTitle?): Promise<{ sheetId }>

// Native Tables — typed columns, dropdowns, banding
sheets.addTable(name, "Sheet!A1:I", { columns }): Promise<{ tableId, name }>
   columns: [{ name, type?, values? }]
   type ∈ TEXT | DOUBLE(NUMBER) | CURRENCY | PERCENT | DATE | TIME | DATE_TIME | BOOLEAN | DROPDOWN
   values: […] ⇒ DROPDOWN column + ONE_OF_LIST rule
sheets.ensureTable(name, "Sheet!A1:I", { columns }): Promise<{ tableId, name, existed }>  // idempotent
sheets.updateTable(nameOrId, { name?, columns?, range? }): Promise<{ ok, tableId }>
sheets.deleteTable(nameOrId, { deleteData? }): Promise<{ ok, tableId, preserved, rows }>
   // removes the table but KEEPS its cells (values/format/notes) by default;
   // pass { deleteData: true } for Google's native behavior, which also clears the range

// Raw escape hatch — full Sheets v4 power for anything not covered by the sugar
sheets.batchUpdate(requests): Promise<...>          // any Request[], dry-run aware
sheets.valuesBatchGet(ranges, opts?): Promise<...>  // multi-range read, one round trip
sheets.getSpreadsheet(opts?): Promise<...>          // metadata: all sheets, named ranges, …
sheets.developerMetadataSearch(filters): Promise<...>
```

The `Sheet`:

```ts
sheet.describe(): { sheet, sheetId, headerRow, rowCount, headers, table? }
   table (when the tab has a native Table): { tableId, name, columns: [{ name, type, letter }] }
sheet.toTable(name, { columns?, rows? }): Promise<{ tableId, name }>   // wrap this tab as a Table

sheet.insertMany(records, opts?): Promise<{ inserted, skipped, rows }>
   records: Array<{ "Header text": value | "=formula" }>
   opts:    { idempotencyKey?: (record, i) => string, format?: StyleObject }
   N records → ONE batchUpdate. Use this for any batch.

sheet.insert(record, opts?): Promise<{ row, inserted, idempotencyHit }>
   Sugar over insertMany for one record.
   opts: { idempotencyKey?: string, format?: StyleObject }

sheet.find(where): Promise<Array<{ row, "Header": value, ... }>>
sheet.update({ where?, rows?, set }): Promise<{ updated, rows }>
sheet.delete({ where?, rows? }): Promise<{ deleted, rows }>

// Formatting & presentation
sheet.format({ where?, rows?, range?, set }): Promise<{ formatted, rows }>
   where/rows style whole rows; pass `range` (A1) to style an arbitrary range.
sheet.formatRange(a1, style): Promise<...>          // any range/column: "B2:D10", "C:C"
sheet.merge(range, type?) / sheet.unmerge(range)    // type: MERGE_ALL|MERGE_COLUMNS|MERGE_ROWS
sheet.setBorders(range, spec)                       // "all" | "outer" | "DASHED" | { top?, …, style?, color? }
sheet.addConditionalFormat(range, rule)             // { condition, format } | { gradient: { min, mid?, max } }
sheet.freeze({ rows?, cols? })
sheet.resizeColumns(range, { width } | { auto: true })
sheet.insertColumns(at, count?) / sheet.deleteColumns(at, count?)   // at: header | letter | index
sheet.sort(range, [{ column, order?: "ASC"|"DESC" }])
sheet.setFilter(range)
sheet.findReplace(find, replace, opts?)             // this sheet by default; opts.range / opts.allSheets
sheet.setNote(cell, text)
sheet.setRowHeight(rowOrA1, px)                     // 22 | "22" | "22:24"
sheet.copyFormat(srcA1, dstA1)                      // copy ONLY formatting: "A18:G18" → "A22:G22"

sheet.setValidation(header, spec): Promise<{ ok }>
   spec.type ∈ ONE_OF_LIST { values } | ONE_OF_RANGE { range } | BOOLEAN {} |
   NUMBER_BETWEEN { min, max } | NUMBER_GREATER/LESS/EQ { value } |
   DATE_BETWEEN { min, max } | DATE_AFTER/BEFORE { value } |
   TEXT_CONTAINS/EQ { value } | CUSTOM_FORMULA { formula } | "clear"
   (+ optional strict?, showCustomUi?)
   On a native-Table typed column, auto-routes to updateTable (raw setDataValidation
   is rejected on typed columns).

sheet.readRange(a1, { valueRender? }): Promise<any[][]>
sheet.readFormatting(a1, opts?): Promise<{ data, merges }>   // formatting, notes, merges, validation
sheet.writeRange(a1, values2d, { raw?, bind? }): Promise<{ updatedRange, updatedCells, ... }>
   // Table structured refs (=…[@[Col]]…) auto-route through updateCells so they
   // bind in row context instead of #ERROR!; bind:true forces it, raw:true skips.
```

Pass `rows: [123, 456]` to `update/delete/format` to skip the `find()` lookup when row numbers are already known. Every mutating method compiles to `batchUpdate` (or a guarded value write) and therefore respects `dryRun`.

**Records use the sheet's actual header text** — no role/schema indirection.

**Formula values** start with `=` and may use placeholders:

- `{row}` — the new row's actual number, resolved at write time
- `{col:Header}` — the A1 column letter for the column with that exact header

**Idempotency**: pass an opaque `idempotencyKey` as a string. The MCP stores it as `DeveloperMetadata` (`sheets-mcp:idempotency` namespace, `PROJECT` visibility) on the new row. Re-running the same insert performs a metadata search first and returns the existing row with `inserted: false, idempotencyHit: <row>` instead of duplicating.

**Validation**: set with `sheet.setValidation(header, spec)` — `ONE_OF_LIST`/`ONE_OF_RANGE` dropdowns, `BOOLEAN` checkboxes, `NUMBER_*`/`DATE_*`/`TEXT_*` conditions, `CUSTOM_FORMULA`, or `"clear"` to remove a rule. `ONE_OF_LIST` rules are also enforced pre-flight by `sheet.insert/update` (both the spreadsheet UI and the MCP reject invalid values).

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

**Create and structure a brand-new tab (no human in the browser):**

```js
await sheets.addSheet("обязательства", { tabColor: "#34a853", frozenRows: 1 });
const s = await sheets.sheet("обязательства");
await s.formatRange("A1:D1", { textFormat: { bold: true }, backgroundColor: "#f1f3f4" });
await s.setBorders("A1:D50", "all");
await s.setValidation("Status", { type: "BOOLEAN" });
await s.freeze({ rows: 1 });
```

**Dry-run preview:**

Pass `dryRun: true` to `sheets_exec` to capture every intended mutation without committing. The response carries `plannedOps[]` — an ordered log of every intended op, each tagged with its `kind` (`"batchUpdate"` with `requests`, or `"valuesUpdate"` with `range`/`values`).

## Architecture

The agent doesn't issue per-cell tool calls. It writes a JS script against the typed `Sheet` API; the runner compiles each script to **one atomic `spreadsheets.batchUpdate`**.

```
agent JS  →  sheets_exec  →  vm sandbox  →  Sheet.insert  →  batchUpdate  →  Sheets API
                                                                    ↓
                                                            (or: dryRun captures
                                                             the request body)
```

`Sheet.insert` produces a single batch with `insertDimension` + `updateCells` (with formulas referencing the new row's actual position) + an optional `createDeveloperMetadata` idempotency token. One round trip, atomic.

The A1↔GridRange math (`mcp/a1.mjs`) and the Sheets v4 request builders (`mcp/requests.mjs`) are pure and isolated, so the structural/presentation surface stays testable without hitting Google.

## Tests

```bash
npm test   # node --test — unit tests for the A1 grammar + request builders
```

No network or credentials needed; the tests cover the pure layers (`a1.mjs`, `requests.mjs`) plus the dry-run wiring (`integration.test.mjs`).

## Isolation & correctness

Each `sheets_exec` call runs against its own client and handle registry, so:

- **Concurrent execs are isolated.** Dry-run captures and the per-sheet handle cache are scoped to one exec — overlapping calls never see each other's writes.
- **Handles track tab identity.** `sheets.renameSheet` updates any handle you're already holding (so it keeps reading/writing the right tab); `sheets.deleteSheet` marks held handles dead — calling one throws `Sheet "<name>" was deleted` instead of silently hitting the wrong tab.
- **Idempotency survives a rename.** Tokens are scoped to a row by `sheetId`, not by tab title.
- **Ambiguous headers fail loud.** If two columns share a header (after case/`ё`→`е` normalization), a write keyed by it throws `Ambiguous header "<h>" — appears in columns B, E` rather than guessing. `describe()` still lists every column so you can spot the dup.

### Known limitations

- **Dry-run can plan writes against a tab/table you created earlier in the same script, but can't *read* it.** Under `dryRun: true`, structural ops (`addSheet`/`duplicateSheet`/`addTable`) are *captured*, not executed, and now return a synthetic **placeholder id** (negative `sheetId`, `dryrun:` `tableId`) so a follow-up `writeRange`/`batchUpdate` referencing it still gets planned. But reads (`find`, `readRange`, `describe`) hit the real spreadsheet, so reading a not-yet-real tab/table — or one you renamed/deleted earlier in the same dry-run — fails. Run such flows for real (or split the preview).
- **GoogleFinance / volatile cells read as `#N/A`.** `GOOGLEFINANCE`, `IMPORT*`, and historical lookups compute only in the browser; the API returns `#N/A` for them, and a nearby write can invalidate a previously-cached value. This is Google's behavior, not a server error — don't treat `#N/A` from these as a failure; verify in the UI.
- **Table structured references must bind in row context.** A formula like `=IF([@[Сумма]]=…)` written through the plain values path renders `#ERROR!`. `writeRange` detects structured refs and auto-routes them through `updateCells` (which binds correctly); `opts.raw: true` skips that and stores the formula verbatim.
- **Table dropdown chip ↔ arrow display is UI-only.** The chip vs. arrow toggle isn't in the Sheets API — native Table dropdown columns always render as chips. An arrow-style dropdown needs a non-Table range plus a manual UI toggle.

## Release (maintainers)

```bash
npm version patch  # or minor / major
git push --follow-tags
```

The `publish` GitHub Action picks up the tag, runs `npm ci && npm publish`, and ships to npm using the `NPM_TOKEN` repo secret.
