/**
 * One-shot batch importer for bank statement Excel files.
 *
 * Reads an .xlsx export from the расчётный счёт, classifies each row, and
 * appends расходы to the Google Sheets workbook. Платежи (credit > 0) and
 * fin-zaim transfers (ddс) are skipped.
 *
 * Usage:
 *   node build/scripts/import-statement.mjs reports/test.xlsx [--rate=12122] [--dry-run]
 */

import path from "path";
import { fileURLToPath } from "url";

import { readBankStatement } from "../src/excel.mjs";
import {
  appendRashod,
  appendDds,
  appendPlatezh,
  appendIteration,
  getNextIteration,
} from "../src/sheets.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = path.resolve(__dirname, "..", "..");

// ───────────────────────────────────────────────────────────────────────────
// CLI args
// ───────────────────────────────────────────────────────────────────────────

const args = process.argv.slice(2);
let inputFile = null;
let rate = 12122;
let dryRun = false;

for (const arg of args) {
  if (arg.startsWith("--rate=")) rate = parseFloat(arg.slice(7));
  else if (arg === "--dry-run") dryRun = true;
  else if (!inputFile) inputFile = arg;
}

if (!inputFile) {
  console.error("Usage: node import-statement.mjs <file.xlsx> [--rate=12122] [--dry-run]");
  process.exit(1);
}

// ───────────────────────────────────────────────────────────────────────────
// Classification rules
// ───────────────────────────────────────────────────────────────────────────

const ACCOUNT = "Kapital UZS (р/с)";

const RULES = [
  {
    name: "ФОТ",
    test: (r) =>
      /beformed/i.test(r.account_name) ||
      /\bз\/?п\b|зарплат|премиальн|премия|неисп.*отпуск/i.test(r.description),
    classify: (r) => ({
      category: "ФОТ",
      employee: extractEmployee(r.description),
    }),
  },
  {
    // Полиграфия (стенды, ролл-апы, листы) — офисные припасы
    name: "Офисные припасы",
    test: (r) => /a3\s*list|rollab|rollup|стенд|полиграф/i.test(r.description),
    classify: () => ({ category: "Офисные припасы", employee: "" }),
  },
  {
    name: "Маркетинг и реклама",
    test: (r) =>
      /реклам|размещени|публикац|пакет.*сайт|пост.*пакет|партнерств.*ассоциац|маркетинг.*ассоциа/i.test(
        r.description
      ),
    classify: () => ({ category: "Маркетинг и реклама", employee: "" }),
  },
  {
    name: "Аренда офис",
    test: (r) => /аренд.*помещени|аренд.*офис/i.test(r.description),
    classify: () => ({ category: "Аренда офис", employee: "" }),
  },
  {
    name: "Наемные рабоники вне штата",
    test: (r) => /бухгалтерск/i.test(r.description),
    classify: () => ({ category: "Наемные рабоники вне штата", employee: "" }),
  },
  {
    // Любые банковские тарифы и комиссии — проверяем ДО налогов,
    // потому что тарифы часто содержат фразу "в том числе НДС".
    name: "Комиссия банка",
    test: (r) =>
      /тариф|swift|комиссия инобанка|за документ.*s=/i.test(r.description),
    classify: () => ({ category: "Комиссия банка", employee: "" }),
  },
  {
    name: "Налоги",
    test: (r) =>
      // Note: узбекские "қ"/"ғ" парсер xlsx подменяет на "?" — поэтому literal \?
      /обслуживание пассивных счетов|пенсия|пенсионн|жисмоний|жам\?ариб|импорт.*услуг|чет эл юридик|жисм|\?ўшилган.*\?иймат|\?ийм.*соли\?|даромад.*соли\?|ндс/i.test(
        r.description
      ),
    classify: () => ({ category: "Налоги", employee: "" }),
  },
];

const DEFAULT = { category: "Другое", employee: "" };

function classify(row) {
  for (const rule of RULES) {
    if (rule.test(row)) return rule.classify(row);
  }
  return DEFAULT;
}

/**
 * Stopwords — Cyrillic / Latin words that look like names but aren't.
 * Used to filter out non-name tokens from candidate matches.
 */
const NAME_STOPWORDS = new Set([
  // Russian short words & prepositions
  "за", "и", "с", "со", "по", "для", "от", "до", "при", "без", "на", "в", "к", "о",
  "у", "из", "как", "что", "это", "так", "не", "ни", "же", "ли", "бы", "то",
  // Currency / units
  "сум", "сума", "сумма", "сум.", "uzs", "usd", "eur", "rub",
  // Time
  "год", "года", "году", "годам", "г", "месяц", "месяца", "м", "число",
  "январ", "феврал", "март", "марта", "апрел", "май", "июн", "июл", "август",
  "сентябр", "октябр", "ноябр", "декабр",
  "января", "февраля", "апреля", "мая", "июня", "июля", "сентября", "октября",
  "ноября", "декабря",
  // Payment vocabulary
  "перечисляется", "перечисление", "перечислено", "перевод",
  "оплата", "оплачено", "оплате", "оплат",
  "сотрудник", "сотруднику", "сотрудника", "сотрудники",
  "премиальные", "премиальная", "премия", "премии", "премиальных", "премиальной",
  "выплат", "выплата", "выплаты",
  "зарплат", "зарплата", "зарплаты", "зп", "з/п",
  "неиспользованный", "неиспользованного", "неиспользованные", "отпуск", "отпуска",
  // Card / account vocabulary
  "данные", "карта", "карту", "карты", "карте", "пк",
  // Document vocabulary
  "приложение", "договор", "договору", "договоров", "согласно", "согл",
  "сог-но", "сог", "организации", "организация", "ооо", "мчж", "мчдж",
]);

function isLikelyNameWord(word) {
  if (!word) return false;
  if (word.length < 3) return false;
  if (/\d/.test(word)) return false;
  return !NAME_STOPWORDS.has(word.toLowerCase());
}

function toTitleCase(word) {
  return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
}

/**
 * Take a sequence of words (possibly with stopwords mixed in) and return
 * the trailing 2-3 name-like words joined and title-cased. Returns null
 * if fewer than 2 valid name words can be found.
 */
function refineNameCandidate(rawWords) {
  const words = rawWords.split(/\s+/).filter(Boolean);
  // Walk from the end, collecting consecutive name-like words.
  const collected = [];
  for (let i = words.length - 1; i >= 0; i--) {
    if (isLikelyNameWord(words[i])) {
      collected.unshift(words[i]);
    } else if (collected.length > 0) {
      // We hit a stopword after starting to collect — stop.
      break;
    }
  }
  if (collected.length < 2) return null;
  // Take at most 3 words (Surname Name Patronymic).
  const trimmed = collected.slice(-3);
  return trimmed.map(toTitleCase).join(" ");
}

/**
 * Extract employee name from a payment description.
 *
 * Strategy:
 *   1) Anchored: capture words right before "на ПК" / "карт*" / "данные карт*"
 *      / "(" — these are very reliable indicators that a name is right
 *      before them. Works even if the name has a lowercase first letter
 *      (like "рустамов Аббосхон").
 *   2) Mixed case Cyrillic — "Иванов Иван [Иванович]"
 *   3) ALL CAPS Cyrillic — "ИВАНОВ ИВАН"
 *   4) ALL CAPS Latin — "IVANOV IVAN"
 *
 * In all strategies, stopwords like "Премиальные", "Выплаты", "Февраль"
 * are filtered out, and the result is normalised to Title Case.
 */
function extractEmployee(desc) {
  // Strip digits/parens/codes, normalise whitespace.
  const text = desc
    .replace(/\d+/g, " ")
    .replace(/[()]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  // ── 1) Anchored extraction (works with lowercase names) ───────────────
  // Capture a run of Cyrillic words (any case) immediately before an anchor.
  // Anchors: "на ПК", "на карту/карте/карты", "карту/карты", "данные карт*"
  //
  // Note: JS regex \b doesn't recognize Cyrillic word boundaries (\w is
  // ASCII-only by default), so we use a negative lookahead instead.
  const NOT_LETTER = "(?![а-яёА-ЯЁa-zA-Z])";
  const anchorRegex = new RegExp(
    `([а-яёА-ЯЁ\\s]+?)\\s+(?:на\\s+(?:ПК|пк|карт[уыеа])|данные\\s+карт[уыеа]|карт[уыеа])${NOT_LETTER}`,
    "i"
  );
  const anchored = text.match(anchorRegex);
  if (anchored) {
    const candidate = refineNameCandidate(anchored[1]);
    if (candidate) return candidate;
  }

  // ── 2) Mixed case Cyrillic — "Иванов Иван [Иванович]" ────────────────
  const cyrMixed = text.match(
    /([А-ЯЁ][а-яё]+(?:\s+[А-ЯЁ][а-яё]+){1,3})/
  );
  if (cyrMixed) {
    const candidate = refineNameCandidate(cyrMixed[1]);
    if (candidate) return candidate;
  }

  // ── 3) ALL CAPS Cyrillic — "ИВАНОВ ИВАН ИВАНОВИЧ" ────────────────────
  const cyrUpper = text.match(/([А-ЯЁ]{2,}(?:\s+[А-ЯЁ]{2,}){1,3})/);
  if (cyrUpper) {
    const candidate = refineNameCandidate(cyrUpper[1]);
    if (candidate) return candidate;
  }

  // ── 4) ALL CAPS Latin — "IVANOV IVAN" ─────────────────────────────────
  const lat = text.match(/\b([A-Z]{2,}(?:\s+[A-Z]{2,}){1,3})\b/);
  if (lat) {
    const candidate = refineNameCandidate(lat[1]);
    if (candidate) return candidate;
  }

  return "";
}

/**
 * Make the description shorter and cleaner before saving.
 */
function shortenDescription(d) {
  let s = d
    // Remove leading 5-digit document codes (00633, 00668, ...)
    .replace(/^\d{5}\s*/g, "")
    // Collapse repeated codes
    .replace(/(00633\s+)+/g, "00633 ")
    .replace(/(00668\s+)+/g, "00668 ")
    .replace(/\s+/g, " ")
    .trim();
  if (s.length > 150) s = s.slice(0, 150) + "…";
  return s;
}

// ───────────────────────────────────────────────────────────────────────────
// Filter rules
// ───────────────────────────────────────────────────────────────────────────

function isPlatezh(row) {
  return row.credit > 0;
}

function isDds(row) {
  // фин займ / фин помощь — переводы через "Транзитный счет ... Самарканд"
  return (
    /фин\s*займ|фин\s*помощь/i.test(row.description) ||
    /транзитный счет.*самарканд/i.test(row.account_name)
  );
}

// ───────────────────────────────────────────────────────────────────────────
// Run
// ───────────────────────────────────────────────────────────────────────────

const inputPath = path.isAbsolute(inputFile)
  ? inputFile
  : path.resolve(PROJECT_ROOT, inputFile);

console.log(`Reading: ${inputPath}`);
console.log(`Rate:    ${rate} UZS/USD`);
console.log(`Mode:    ${dryRun ? "DRY RUN (no writes)" : "LIVE"}\n`);

const result = readBankStatement(inputPath);
console.log(`Parsed ${result.data_rows} data rows from sheet "${result.sheet}"\n`);

const buckets = { платежи: [], ддс: [], расходы: [] };

for (const row of result.rows) {
  if (isPlatezh(row)) buckets.платежи.push(row);
  else if (isDds(row)) buckets.ддс.push(row);
  else buckets.расходы.push(row);
}

console.log(`To append — платежи:     ${buckets.платежи.length}`);
console.log(`To append — ддс:         ${buckets.ддс.length}`);
console.log(`To append — расходы:     ${buckets.расходы.length}\n`);

let success = 0;
let failed = 0;

// ─── Payment project mapping ─────────────────────────────────────────────
// Map keywords in bank statement description to project names + categories.

const PAYMENT_PROJECTS = [
  {
    test: (desc) => /АИС\s*GRANITE|разработк.*внедрени.*автоматизир/i.test(desc),
    project: "Granite",
    category: "Разработка ПО под ключ",
  },
  {
    test: (desc) => /техническ.*поддержк.*сопровожд|функционирован.*веб/i.test(desc),
    project: "EAI_Website",
    category: "Поддержка и мелкие доработки",
  },
  {
    test: (desc) => /консалтинг.*цифровизац/i.test(desc),
    // Differentiate by amount:
    //   101M → Takhir
    //   >300M → Reklama EAI
    //   остальное → EAI_Website
    projectFn: (row) => {
      if (row.credit === 101000000) return "Takhir";
      if (row.credit >= 300000000) return "Reklama EAI";
      return "EAI_Website";
    },
    categoryFn: (row) => {
      if (row.credit === 101000000) return "Консалтинг и продажа часов";
      if (row.credit >= 300000000) return "Другие источники дохода";
      return "Поддержка и мелкие доработки";
    },
  },
  {
    // "цифровое курирование" (с возможной опечаткой "цифрАвому")
    test: (desc) => /цифр[аов].*курирован/i.test(desc),
    project: "Reklama EAI",
    category: "Другие источники дохода",
  },
];

function classifyPayment(row) {
  for (const rule of PAYMENT_PROJECTS) {
    if (rule.test(row.description)) {
      const project = rule.projectFn ? rule.projectFn(row) : rule.project;
      const category = rule.categoryFn ? rule.categoryFn(row) : rule.category;
      return { project, category };
    }
  }
  return { project: "UNKNOWN", category: "UNKNOWN" };
}

// ─── Write payment rows ──────────────────────────────────────────────────
for (const row of buckets.платежи) {
  const cls = classifyPayment(row);

  if (dryRun) {
    console.log(
      `DRY [платёж] #${row.row_num} ${row.doc_date} ${row.credit.toLocaleString("ru-RU")} → ${cls.project} (${cls.category})`
    );
    continue;
  }

  try {
    // Step 1: determine next iteration number
    const iter = await getNextIteration(cls.project);

    // Step 2: add row to "итерации"
    await appendIteration({
      project: cls.project,
      iteration_id: iter.nextId,
      iteration_num: iter.nextNum,
      amount: row.credit,
      currency: "SOM",
    });

    // Step 3: add row to "платежи"
    const res = await appendPlatezh({
      iteration: iter.nextId,
      project: cls.project,
      amount: row.credit,
      category: cls.category,
      rate,
      currency: "SOM",
      date: row.doc_date,
      incoming_account: "Kapital UZS (р/с)",
      comment: "На р/с",
    });

    success++;
    console.log(
      `✓ [платёж] #${row.row_num} ${row.doc_date} ${row.credit.toLocaleString("ru-RU")} → ${cls.project} (${iter.nextId}) @ ${res.range}`
    );
  } catch (err) {
    failed++;
    console.log(`✗ [платёж] #${row.row_num} ${row.doc_date} ${row.credit} — ${err.message}`);
  }
}

// ─── Write DDS rows ──────────────────────────────────────────────────────
for (const row of buckets.ддс) {
  const data = {
    date: row.doc_date,
    amount: row.debit,
    // source_account, target_account, type — use defaults from appendDds
  };

  if (dryRun) {
    console.log(`DRY [ддс] #${row.row_num} ${row.doc_date} ${row.debit.toLocaleString("ru-RU")}`);
    continue;
  }

  try {
    const res = await appendDds(data);
    success++;
    console.log(`✓ [ддс] #${row.row_num} ${row.doc_date} ${row.debit.toLocaleString("ru-RU")} @ ${res.range}`);
  } catch (err) {
    failed++;
    console.log(`✗ [ддс] #${row.row_num} ${row.doc_date} ${row.debit} — ${err.message}`);
  }
}

// ─── Write расходы rows ──────────────────────────────────────────────────

for (const row of buckets.расходы) {
  const cls = classify(row);
  const data = {
    date: row.doc_date,
    category: cls.category,
    amount: row.debit,
    rate,
    currency: "SOM",
    account: ACCOUNT,
    employee: cls.employee,
    description: shortenDescription(row.description),
  };

  if (dryRun) {
    console.log(
      `DRY #${row.row_num} ${row.doc_date} ${row.debit.toLocaleString("ru-RU")} → ${cls.category}` +
        (cls.employee ? ` [${cls.employee}]` : "")
    );
    continue;
  }

  try {
    const res = await appendRashod(data);
    success++;
    console.log(
      `✓ #${row.row_num} ${row.doc_date} ${row.debit.toLocaleString("ru-RU")} → ${cls.category}` +
        (cls.employee ? ` [${cls.employee}]` : "") +
        ` @ ${res.range}`
    );
  } catch (err) {
    failed++;
    console.log(`✗ #${row.row_num} ${row.doc_date} ${row.debit} — ${err.message}`);
  }
}

if (!dryRun) {
  console.log(`\nDone: ${success} success, ${failed} failed`);
}
