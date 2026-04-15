#!/usr/bin/env node
/**
 * Обновляет Google OAuth-токен для доступа к Sheets.
 *
 * Запускается один раз, когда MCP начинает выдавать `invalid_grant` (refresh
 * token протух — Google ограничивает testing-mode приложения 7 днями).
 *
 * Что делает:
 *   1. Читает GOOGLE_CLIENT_ID / GOOGLE_CLIENT_SECRET из .env
 *   2. Поднимает локальный HTTP-сервер на http://localhost:3000
 *   3. Открывает в браузере ссылку авторизации Google
 *   4. После подтверждения получает code через редирект
 *   5. Меняет code на access + refresh токены
 *   6. Сохраняет их в auth/.gdrive-server-credentials.json
 *
 * Usage: node scripts/reauth.mjs
 *
 * После: перезапустить MCP (закрыть/открыть VSCode) — новый токен подхватится.
 */

import { google } from "googleapis";
import http from "http";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import { exec } from "child_process";
import { loadEnv } from "../mcp/env.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "..");
const TOKEN_PATH = path.join(ROOT, "auth", ".gdrive-server-credentials.json");

const PORT = 3000;
const REDIRECT_URI = `http://localhost:${PORT}`;
const SCOPES = ["https://www.googleapis.com/auth/spreadsheets"];

const env = loadEnv();
if (!env.GOOGLE_CLIENT_ID || !env.GOOGLE_CLIENT_SECRET) {
  console.error("Отсутствуют GOOGLE_CLIENT_ID или GOOGLE_CLIENT_SECRET в .env");
  process.exit(1);
}

const client = new google.auth.OAuth2(
  env.GOOGLE_CLIENT_ID,
  env.GOOGLE_CLIENT_SECRET,
  REDIRECT_URI,
);

const authUrl = client.generateAuthUrl({
  access_type: "offline",
  prompt: "consent",       // чтобы всегда выдавало refresh_token
  scope: SCOPES,
});

console.log("\nОткрываю в браузере ссылку авторизации...\n");
console.log(authUrl);
console.log("\nЕсли браузер не открылся сам — скопируй ссылку выше и открой вручную.\n");

// Попытаться открыть в системном браузере (кроссплатформа)
const openCmd = process.platform === "win32" ? `start "" "${authUrl}"`
  : process.platform === "darwin" ? `open "${authUrl}"`
  : `xdg-open "${authUrl}"`;
exec(openCmd, () => {});

// Ждём redirect с code
const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url, REDIRECT_URI);
    const code = url.searchParams.get("code");
    const error = url.searchParams.get("error");

    if (error) {
      res.end(`Ошибка авторизации: ${error}. Можно закрыть вкладку.`);
      console.error(`\n❌ Ошибка от Google: ${error}`);
      server.close();
      process.exit(1);
    }

    if (!code) {
      res.end("Ожидаю code в URL — откройте ссылку выше.");
      return;
    }

    const { tokens } = await client.getToken(code);

    // Сохраняем токен
    fs.mkdirSync(path.dirname(TOKEN_PATH), { recursive: true });
    fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens, null, 2));

    res.end("Готово! Токен сохранён. Можно закрыть вкладку и вернуться в VSCode.");
    console.log(`\n✅ Токен сохранён в ${TOKEN_PATH}`);
    console.log(`   Содержит refresh_token: ${tokens.refresh_token ? "да" : "НЕТ (плохо — повтори с --force)"}`);
    console.log(`\nТеперь закрой и открой VSCode — MCP подхватит новый токен.\n`);

    server.close();
    process.exit(0);
  } catch (e) {
    res.end(`Ошибка: ${e.message}`);
    console.error(e);
    server.close();
    process.exit(1);
  }
});

server.listen(PORT, () => {
  console.log(`Жду редирект на http://localhost:${PORT} ...`);
});
