---
name: import-bank-statement
description: Импорт банковской выписки IOTA (.xlsx из reports/) в Google Sheets. Классифицирует каждую строку по смыслу (ФОТ/Налоги/Маркетинг/и т.д. для расходов, проект для платежей), создаёт итерации перед платежами, распознаёт ддс (фин займ). Используй когда пользователь говорит "обработай отчёт", "обработай выписку", "импортируй выписку", или упоминает xlsx в reports/.
allowed-tools: Bash(node *), mcp__sheets__sheets_describe, mcp__sheets__sheets_exec, Read, Glob
argument-hint: "[путь к xlsx] [--dry-run]"
---

# Import bank statement

Разбор банковской выписки в листы Google Sheets. Используется для расчётного счёта (Kapital UZS).

## Когда применять

- "обработай отчёт", "обработай выписку", "импортируй выписку"
- Пользователь указал .xlsx файл в `reports/`
- Упомянул "xlsx" или "банковскую выписку"

## Шаг 1 — прочитать xlsx

```bash
node -e "import('./mcp/excel.mjs').then(m => { const r = m.readBankStatement('reports/<файл>.xlsx'); console.log(JSON.stringify(r, null, 2)); });"
```

Возвращает массив строк с полями: `doc_date`, `debit`, `credit`, `account_name`, `description`, и т.д.

Уточни у пользователя период (обычно один календарный месяц, формат "март 2026" → фильтр `/\.03\.2026$/` по doc_date). Если не сказал — спроси.

## Шаг 2 — узнать схему листов

```
sheets_describe                  # все доступные схемы
sheets_describe sheet="расходы"  # детали одной
```

Целевые листы:
- `расходы` — для трат
- `платежи` — для поступлений (ссылается на `итерации`)
- `итерации` — нужна новая запись ПЕРЕД каждым платежом
- `ддс` — для внутренних переводов (фин помощь)

## Шаг 3 — классифицировать каждую строку

### Определение типа

| Условие | Тип | Куда пишем |
|---------|-----|------------|
| `credit > 0` | Поступление | `платежи` + создать новую `итерация` |
| `debit > 0` + описание содержит "фин займ"/"фин помощь" или контрагент содержит "Транзитный счёт ... Самарканд" | Внутренний перевод | `ддс` |
| `debit > 0` иначе | Расход | `расходы` |

### Расходы — категория (по смыслу, не regex)

- Контрагент `BEFORMED`, "ЗП", "зарплат", "премиальн", "премия", "неиспользованный отпуск" → **ФОТ** (извлеки имя сотрудника в `employee`)
- "обслуживание пассивных счетов", "пенсия", "НДФЛ", узбекские налоговые термины ("жисмоний шахслар даромад", "?ўшилган ?иймат", "импорт услуг", "чет эл юридик") → **Налоги**
- "тариф", "CORPORATE", "SWIFT", "комиссия инобанка", "за документ S=" → **Комиссия банка**
- "реклама", "размещение", "публикация", "пакет сайт", "партнерство ассоциация", "маркетинг" → **Маркетинг и реклама**
- "аренда помещения", "аренда офис" → **Аренда офис**
- "A3 LIST", "ROLLAB", "стенд", "полиграф" → **Офисные припасы**
- "бухгалтерские услуги", "аудит" → **Наемные рабоники вне штата**
- Ничего не подошло → **Другое**

### Платежи — проект (по смыслу + сумма)

- "АИС GRANITE" / "разработка и внедрение автоматизированной системы" → **Granite** / `Разработка ПО под ключ`
- "консалтинговые услуги по цифровизации" + сумма ≈ 101 000 000 → **Takhir** / `Консалтинг и продажа часов`
- "цифровое курирование" / ("консалтинг" + сумма > 300M) → **Reklama EAI** / `Другие источники дохода`
- "техническая поддержка, сопровождение веб" / ("консалтинг" + средняя сумма) → **EAI_Website** / `Поддержка и мелкие доработки`

## Шаг 4 — dry-run

Перед записью **ВСЕГДА** покажи пользователю итоговую классификацию:
- Расходы — разбивка по категориям
- Платежи — по проектам + явно упомяни сколько новых `итерация` будет создано
- ддс — список

Жди подтверждение ("да"/"пиши"/"ок"). Если пользователь сказал `--dry-run` — просто показать и выйти.

Можно прогнать через `sheets_exec` с `dryRun: true` чтобы увидеть planned batchUpdate.

## Шаг 5 — запись (один скрипт через sheets_exec)

Один JS-скрипт делает всё атомарно. Порядок: сначала итерации (платежам нужны их ID), потом платежи / расходы / ддс.

```js
const итерации = await sheets.sheet("итерации");
const платежи  = await sheets.sheet("платежи");
const расходы  = await sheets.sheet("расходы");
const ддс      = await sheets.sheet("ддс", { headerRow: 3 });

const RATE = '=SWITCH({col:Валюта}{row}; "USD";1; "SOM";IFERROR(INDEX(GOOGLEFINANCE("CURRENCY:USDUZS";"price";{col:Дата}{row});2;2);INDEX(GOOGLEFINANCE("CURRENCY:USDUZS";"price";WORKDAY({col:Дата}{row};-1));2;2)); "EUR";IFERROR(INDEX(GOOGLEFINANCE("CURRENCY:USDEUR";"price";{col:Дата}{row});2;2);INDEX(GOOGLEFINANCE("CURRENCY:USDEUR";"price";WORKDAY({col:Дата}{row};-1));2;2)); "RUB";IFERROR(INDEX(GOOGLEFINANCE("CURRENCY:USDRUB";"price";{col:Дата}{row});2;2);INDEX(GOOGLEFINANCE("CURRENCY:USDRUB";"price";WORKDAY({col:Дата}{row};-1));2;2)); NA())';
const USD = '={col:Сумма}{row}/{col:Курс}{row}';
const firstOfMonth = (d) => { const [, mm, yyyy] = d.split("."); return `01.${mm}.${yyyy}`; };

// 1. Найти max # итерации по проекту
const all = await итерации.find({});
const maxByProject = {};
for (const it of all) {
  const n = parseInt(it["# итерации"], 10) || 0;
  const p = it["Проект"];
  if (!maxByProject[p] || n > maxByProject[p]) maxByProject[p] = n;
}

// 2. Создать новые итерации (по одной для каждого платежа)
const newIterations = [];
for (const p of payments) {
  const next = (maxByProject[p.project] ?? 0) + 1;
  maxByProject[p.project] = next;
  const id = `${p.project}-${next}`;
  await итерации.insert({
    "Проект": p.project, "ID итерации": id, "# итерации": next,
    "Всего": p.amount, "Оплачено ($)": p.amount, "Остаток ($)": 0,
    "Валюта": p.currency, "Начало": p.date,
  }, { idempotencyKey: `iteration|${id}` });
  p.iteration = id;
  newIterations.push(id);
}

// 3. Платежи
for (const p of payments) {
  await платежи.insert({
    "Итерация": p.iteration, "Проект": p.project, "Сумма": p.amount,
    "Категория": p.category, "Курс": RATE, "Валюта": p.currency,
    "Сумма ($)": USD, "Дата": p.date, "Рассчетный период": firstOfMonth(p.date),
    "Входящий счет": "Kapital UZS (р/с)", "Комментарий": p.comment ?? "На р/с",
  }, { idempotencyKey: `payment|${p.iteration}|${p.amount}` });
}

// 4. Расходы
for (const e of expenses) {
  await расходы.insert({
    "Дата": e.date, "Категория": e.category, "Сумма": e.amount,
    "Курс": RATE, "Валюта": "SOM", "Сумма ($)": USD,
    "Расчетный период": firstOfMonth(e.date),
    "Исходящий счет": "Kapital UZS (р/с)",
    "Сотрудник": e.employee ?? "", "Описание": e.description,
  }, { idempotencyKey: `expense|${e.date}|${e.amount}|${e.category}|${e.description}` });
}

// 5. ддс
for (const t of transfers) {
  await ддс.insert({
    "Дата": t.date, "Сумма": t.amount,
    "Исходящий счет": t.outAccount, "Входящий счет": t.inAccount,
    "Тип": "P2P", "Рассчетный период": firstOfMonth(t.date),
    "Комментарий": t.comment ?? "",
  }, { idempotencyKey: `transfer|${t.date}|${t.amount}|${t.outAccount}|${t.inAccount}` });
}

return {
  iterations: newIterations,
  payments: payments.length,
  expenses: expenses.length,
  transfers: transfers.length,
};
```

`insert` идемпотентен по primary key через DeveloperMetadata — повторный запуск с теми же данными вернёт `idempotencyHit` и не создаст дублей. Курс/USD-сумма/период заполняются формулами автоматически.

## Шаг 6 — отчёт

После запуска — короткий отчёт:
- Платежей: N (итерации: <список ID>)
- Расходов: N
- ддс: N
- Дубли (idempotencyHit): N — если были

## Чего НЕ делать

- ❌ **Не пиши платёж до итерации.** Скрипт выше делает в правильном порядке — не меняй.
- ❌ **Не записывай без dry-run подтверждения.**
- ❌ **Не пиши руками A1-диапазоны.** Schema-aware API не требует знать колонки.
- ❌ **Не игнорируй ошибки валидации** — если `insert` бросил `validation failed`, исправь данные и повтори.

## Детали xlsx-формата

Kapital Bank экспортит: `№ пп`, `Дата документа`, `Дата обработки`, `№ док`, `Наименование счёта` (контрагент), `ИНН`, `№ счёта`, `МФО`, `Обороты по дебету`, `Обороты по кредиту`, `Назначение платежа`.

Парсер `mcp/excel.mjs` находит заголовок автоматически. Если xlsx из другого банка — формат может отличаться; скажи пользователю что требуется доработка парсера.

## Счёт для расходов

Для расходов из выписки с р/с — счёт всегда `Kapital UZS (р/с)`. Точное написание подсмотри через `sheets_describe sheet="счета"` или через `sheets.sheet("счета", { headerRow: 2 }).find({})`.
