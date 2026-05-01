---
name: extract-sms-transactions
description: Извлекает транзакции с корпоративных карт IOTA (Tenge Mastercard, Uzum Visa UZS) из SMS-уведомлений в macOS Messages (chat.db). Альтернатива скриншотам когда iPhone Mirroring блокируется (например Tenge запрещает захват экрана). Используй когда пользователь говорит "обработай sms", "вытащи транзакции из смс", "импортируй с tenge через смс", "обработай карты через sms".
allowed-tools: Bash, Read, mcp__sheets__sheets_describe, mcp__sheets__sheets_exec
---

# Extract SMS transactions

Парсит SMS от `TENGEBANK` / `UzumBank` из macOS Messages, фильтрует только реальные траты (без OTP, отказов, кэшбэка, входящих переводов), возвращает структурированный список. Дальше идёт по тому же пайплайну что `import-card-transactions` — категоризация → dry-run → запись в `расходы`.

## Когда применять

- Пользователь говорит "обработай sms / смс", "вытащи транзакции из messages", "импортируй tenge через смс"
- Нужны транзакции по корпоративной карте, а скриншоты недоступны (Tenge блокирует screen capture, iPhone Mirroring не работает)

## Предпосылки

- На Mac включён **Messages in iCloud** + **Text Message Forwarding** для этого iPhone — иначе SMS не дойдут до `chat.db`.
- Установлен `pytypedstream` (`python3 -m pip install --user pytypedstream`) — нужен для расшифровки `attributedBody` blobs (Apple хранит тело SMS как NSAttributedString typedstream, не plain text).
- DB живёт здесь: `~/Library/Messages/chat.db`. Терминалу нужен Full Disk Access.

## Workflow

### Шаг 1 — найти точку отсчёта по каждой карте

```js
// sheets_exec
const расходы = await sheets.sheet("расходы");
const tenge = (await расходы.find({ "Исходящий счет": "Tenge Mastercard" })).slice(-3);
const uzum  = (await расходы.find({ "Исходящий счет": "Uzum Visa UZS" })).slice(-3);
return { tenge, uzum };
```

Стартовая дата = max(последняя по Tenge, последняя по Uzum) или -1 день для запаса (idempotency-ключ всё равно отсечёт дубли).

### Шаг 2 — извлечь SMS

```bash
python3 .claude/skills/extract-sms-transactions/scripts/extract_sms.py --since 2026-04-15
# опционально: --bank tenge | --bank uzum | --bank all (default)
```

Выводит JSON-массив записей:
```json
{
  "bank": "tenge" | "uzum",
  "card": "3734",
  "account": "Tenge Mastercard" | "Uzum Visa UZS",
  "amount": 25.35,
  "currency": "USD" | "SOM" | "EUR" | "RUB",
  "date": "30.04.2026",
  "datetime": "2026-04-30 20:44:59",
  "merchant": "GOOGLE*ADS4053006539",  // null для Tenge — нет в SMS
  "raw": "Pokupka 25.35 USD, po karte 518100******3734, ...",
  "sms_ts": "..."
}
```

Скрипт уже отбрасывает: OTP-коды, "Otkaz operatsii" (declined), "Transaction failed", "CASH-BACK" (доход), "Top-up" (входящие), Confirm-the-transfer, опросы, "TENGE24 - One-time code". Реальные траты — только `Pokupka` (Tenge) и `Spisanie/Pokupka/Withdrawal/Purchase` (Uzum).

### Шаг 3 — отбросить личные траты и переводы между своими

Tenge SMS не содержит merchant — ориентируйся на сумму и страну (есть в `raw`). Если непонятно — спросить у пользователя что есть что (даты + суммы).

Uzum SMS содержит merchant. Применяй те же правила что в `import-card-transactions`:

| Признак | Что делать |
|---------|-----------|
| `GOOGLE*ADS...` | Категория = `EAI Google`, описание = `google` |
| Notion / OpenAI / Claude / Webflow / Replit / Adobe / GitHub / Cloudflare / DigitalOcean / Posthog / Ahrefs / Vercel и т.п. | Категория = `Сервера и подписки` |
| `UB VISAUZUM TO UZCARD` или `UZUMBANK UZCARD to VISAUZ` | Перевод между своими счетами — **пропустить** |
| Apple.com/bill, Uber, DoorDash, Aliexpress, Amazon, Wise, Personality.co | Личное — **пропустить** |
| Иностранная валюта типа JPY/TRY (не USD/EUR/SOM) | Личное — **пропустить** |

### Шаг 4 — dry-run

Покажи пользователю две таблицы (Tenge / Uzum), отдельно EAI Google и Сервера и подписки. Подтверди что Tenge-описания угаданы верно (merchant из SMS не приходит).

### Шаг 5 — записать через `sheets_exec`

Используй ту же формулу `RATE_FORMULA` / `USD_FORMULA` и `recordFor` что в `import-card-transactions` — структура листа `расходы` идентична. Idempotency-ключ строй из `${date}|${amount}|${currency}|${description}` чтобы пере-запуск был безопасен.

```js
const items = [
  { date: "30.04.2026", category: "EAI Google", amount: 500, currency: "USD", description: "google", account: "Uzum Visa UZS" },
  { date: "30.04.2026", category: "Сервера и подписки", amount: 25.35, currency: "USD", description: "Replit", account: "Tenge Mastercard" },
  // ...
];
```

Обрати внимание: `account` теперь может быть **разным** в одной пачке (Tenge и Uzum), в отличие от screenshot-импорта где обычно одна карта за раз.

### Шаг 6 — обновить балансы

В отличие от скрин-импорта, **из SMS известен `Dostupno`** в каждом сообщении — последний доступный = текущий баланс. Достаточно взять последнюю строку по каждой карте из вывода скрипта (поле `raw`, парсить балансовое значение) и обновить лист `счета`. Если хочешь — попроси пользователя подтвердить (балансы могут расходиться на сумму pending-транзакций).

## Корпоративные счета (mapping card → account)

- `*3734` → `Tenge Mastercard`
- `*9407` → `Uzum Visa UZS`

Маппинг живёт в `scripts/extract_sms.py` (`CARD_MAP`). Если появится новая карта, скрипт вернёт `account = "card-XXXX"` — добавь её в `CARD_MAP` или в лист `счета`.

## Известные ограничения

- **Tenge SMS не содержит merchant** — только сумма / валюта / дата / страна. Категоризация делается вручную или по-смежности с другими источниками (выписка с р/с, скрины из приложения позже). Для рекуррентных подписок можно вести memo-табличку (если 25.35 USD из США каждый месяц 30-го → Replit).
- **SMS-форвардинг даёт только новые сообщения.** Историю до включения форвардинга нужно тянуть из iCloud Messages (если синк догнал) или из encrypted iPhone backup.
- **chat.db иногда лочится** при активной Messages.app — если SQLite ругается `database is locked`, закрой Messages.app или сделай копию: `cp ~/Library/Messages/chat.db /tmp/chat.db && python3 ... --db /tmp/chat.db` (потребует флаг `--db` если расширишь скрипт).
