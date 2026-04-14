/**
 * Minimal .env file loader. Reads key=value pairs from the project root `.env`.
 * No external dependencies needed.
 */

import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ENV_PATH = path.resolve(__dirname, "..", ".env");

let cached = null;

export function loadEnv() {
  if (cached) return cached;

  if (!fs.existsSync(ENV_PATH)) {
    throw new Error(`.env file not found at ${ENV_PATH}. Copy .env.example to .env and fill in values.`);
  }

  const content = fs.readFileSync(ENV_PATH, "utf-8");
  const env = {};

  for (const line of content.split("\n")) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const idx = trimmed.indexOf("=");
    if (idx === -1) continue;
    const key = trimmed.slice(0, idx).trim();
    const value = trimmed.slice(idx + 1).trim();
    env[key] = value;
    // Also set in process.env for libraries that need it
    process.env[key] = value;
  }

  cached = env;
  return cached;
}
