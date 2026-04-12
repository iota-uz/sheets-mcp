/**
 * Google OAuth2 client loader.
 *
 * Reads client credentials from environment variables (GOOGLE_CLIENT_ID,
 * GOOGLE_CLIENT_SECRET) and the refresh token from `auth/.gdrive-server-credentials.json`.
 *
 * Token refreshes are persisted back to disk automatically.
 */

import { google } from "googleapis";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import { loadEnv } from "./env.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "..");
const TOKEN_PATH = path.join(ROOT, "auth", ".gdrive-server-credentials.json");

let cachedClient = null;

/**
 * Returns a memoized, authenticated Google OAuth2 client.
 */
export function loadAuth() {
  if (cachedClient) return cachedClient;

  const env = loadEnv();
  const tokens = JSON.parse(fs.readFileSync(TOKEN_PATH, "utf-8"));

  const client = new google.auth.OAuth2(
    env.GOOGLE_CLIENT_ID,
    env.GOOGLE_CLIENT_SECRET,
    "http://localhost"
  );
  client.setCredentials(tokens);

  client.on("tokens", (newTokens) => {
    const merged = { ...tokens, ...newTokens };
    fs.writeFileSync(TOKEN_PATH, JSON.stringify(merged, null, 2));
  });

  cachedClient = client;
  return cachedClient;
}
