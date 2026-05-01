---
name: import-card-transactions
description: Импорт расходов с корпоративных карт IOTA (Tenge Mastercard, Uzum Visa UZS) в Google Sheets. iPhone Mirroring блокирует computer-use, скриншоты приходят напрямую в чат. Классифицирует подписки/Google Ads, отсекает личные траты и failed-транзакции. Используй когда пользователь говорит "обработай карту", "импортируй транзакции с tenge", "обработай uzum", "перенеси расходы с карты".
allowed-tools: mcp__sheets__sheets_describe, mcp__sheets__sheets_exec, Read
---

# Import card transactions

Перенос транзакций с корпоративных карт в лист `расходы`. Пользователь шарит экран мобильного приложения банка через iPhone Mirroring и присылает скриншоты в чат (computer-use не работает с iPhone Mirroring — Apple блокирует захват экрана на уровне композитора, screenshot тулзы тайм-аутят).

## Когда применять

- "обработай карту", "обработай tenge/uzum", "перенеси расходы с карты"
- Пользователь упоминает Tenge Mastercard или Uzum Visa UZS и присылает скрины

## Workflow

### Шаг 1 — найти точку отсчёта

```js
// sheets_exec
const расходы = await sheets.table("расходы");
const last = await расходы.find({ account: "Tenge Mastercard" });  // or "Uzum Visa UZS"
return last.slice(-5);  // последние 5 для анализа точки отсчёта
```

Точка отсчёта = дата последней записи по нужному счёту. Брать транзакции **позже**.

### Шаг 2 — попросить скриншоты

Сказать пользователю стартовую дату и попросить скрины списка операций (newest-first).

### Шаг 3 — распарсить и классифицировать

Распознать каждую видимую транзакцию: дату, время, описание, сумму, валюту, **иконку** (важно для failed).

#### Категории

| Категория | Что попадает |
|-----------|--------------|
| **EAI Google** | `GOOGLE*ADS...` (контекстная реклама проекта EAI), описание = `google` |
| **Сервера и подписки** | Notion Labs, OpenAI/ChatGPT, Claude.ai, Webflow, Replit, Docker, Perplexity, Adobe, Freepik, LiveDune, Blacksmith Software, GitHub, Google Workspace, Cloudflare, Digitalocean и т.п. |

#### ⛔ ПРОПУСКАТЬ (личные траты, не идут в расходы компании)

| Признак | Пример |
|---------|--------|
| Иностранная валюта типа JPY/TRY (не USD/EUR/SOM) | Покупки в Японии/Турции |
| Uber / Uber*trip / Uber*Pending | Все Uber независимо от валюты |
| Apple.com/bill | Личные подписки Apple |
| DoorDash (`Dd *doordash`, `Doordash`) | Еда |
| Aliexpress, Amazon, Ministop, ритейл | Шопинг |
| Wise > LONDON / Wise-переводы | Личные международные переводы |
| Personality.co | Личная подписка |
| Railway (если уточнял пользователь — личное) | Спросить если непонятно |
| Travel-merchants (Narita Airport и т.п.) | Командировки/поездки |
| Cashback (входящий `+`) | Это доход, не расход |
| `UZUMBANK UZCARD to VISAUZ` (входящий `+`) | Перевод между своими счетами |

#### ⛔ FAILED транзакции

В Uzum Bank рядом с failed-операциями иконка `→` с **красной точкой** вместо `−` (минус). **Пропускать** — деньги не списались. В Tenge Bank иконка обычно `←→` (стрелки) — отличать визуально по цвету/значку.

### Шаг 4 — dry-run

Показать пользователю таблицу того что будет записано (отдельно EAI Google и Сервера и подписки). Можно сначала прогнать через `sheets_exec` с `dryRun: true` чтобы увидеть planned batchUpdate.

### Шаг 5 — записать

Один скрипт через `sheets_exec` для всей пачки (один MCP вызов, один атомарный батч):

```js
const расходы = await sheets.table("расходы");
const account = "Tenge Mastercard";  // или "Uzum Visa UZS"

const items = [
  { date: "29.04.2026", category: "Сервера и подписки", amount: 53, currency: "USD", account, description: "Webflow" },
  { date: "27.04.2026", category: "Сервера и подписки", amount: 48, currency: "USD", account, description: "OpenAI ChatGPT" },
  // ...
];

const results = [];
for (const item of items) {
  results.push(await расходы.insert(item));
}
return {
  inserted: results.filter(r => r.inserted).length,
  skipped:  results.filter(r => r.idempotencyHit).length,
  rows: results.map(r => r.row),
};
```

Курс/USD/период заполняются автоматически из формул в схеме. Idempotency дедуп — по primary key (date+amount+category+description) через DeveloperMetadata, повторный запуск того же скрипта ничего не дублирует.

### Шаг 6 — обновить балансы

После всех вставок спросить у пользователя актуальные балансы карт (скрин "By products" в банковском приложении). Обновить через Table API:

```js
const счета = await sheets.table("счета");
await счета.update({ where: { name: "Tenge Mastercard" }, set: { amount: 795637.75 }});
await счета.update({ where: { name: "Uzum Visa UZS" }, set: { amount: 30840294.87 }});
```

Карты `**** 1305` и `**** 1148` (Kapitalbank) — личные, в листе `счета` их нет, пропускать.

## Корпоративные счета (для записи в `account`)

- `Tenge Mastercard` — `**** 3734`, USD/EUR подписки
- `Uzum Visa UZS` — `**** 9407`, Google Ads и SOM-подписки

Точное написание берётся из листа `счета` — `sheets_describe("счета")` показывает список.
