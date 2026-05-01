/**
 * Google service account auth.
 *
 * Reads a service account JSON key from GOOGLE_SERVICE_ACCOUNT_KEY_PATH
 * (defaults to auth/service-account.json) and returns a GoogleAuth client
 * with Sheets + Drive scopes. JWT refresh is handled by googleapis.
 */

import { google } from "googleapis";
import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";
import { loadEnv } from "./env.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "..");

const SCOPES = [
  "https://www.googleapis.com/auth/spreadsheets",
  "https://www.googleapis.com/auth/drive",
];

let cached = null;

export function loadAuth() {
  if (cached) return cached;

  const env = loadEnv();
  const keyPath = path.resolve(
    ROOT,
    env.GOOGLE_SERVICE_ACCOUNT_KEY_PATH || "auth/service-account.json"
  );

  if (!fs.existsSync(keyPath)) {
    throw new Error(
      `Service account key not found at ${keyPath}. ` +
        `Download a JSON key from GCP IAM and save it there, ` +
        `or set GOOGLE_SERVICE_ACCOUNT_KEY_PATH in .env.`
    );
  }

  cached = new google.auth.GoogleAuth({ keyFile: keyPath, scopes: SCOPES });
  return cached;
}
