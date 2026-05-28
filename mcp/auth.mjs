/**
 * Google service account auth.
 *
 * Resolution order for the key path:
 *   1. process.env.GOOGLE_APPLICATION_CREDENTIALS  (Google's standard, absolute path)
 *   2. process.env.GOOGLE_SERVICE_ACCOUNT_KEY_PATH (back-compat, resolved against cwd)
 *
 * Returns a GoogleAuth client with Sheets + Drive scopes. JWT refresh is handled
 * by googleapis.
 */

import { google } from "googleapis";
import path from "path";
import fs from "fs";

const SCOPES = [
  "https://www.googleapis.com/auth/spreadsheets",
  "https://www.googleapis.com/auth/drive",
];

let cached = null;

export function loadAuth() {
  if (cached) return cached;

  const explicit = process.env.GOOGLE_APPLICATION_CREDENTIALS
    || process.env.GOOGLE_SERVICE_ACCOUNT_KEY_PATH;

  if (!explicit) {
    throw new Error(
      "Service account key path not configured. Set GOOGLE_APPLICATION_CREDENTIALS " +
      "(absolute path to your service-account.json) in the MCP env."
    );
  }

  const keyPath = path.isAbsolute(explicit) ? explicit : path.resolve(process.cwd(), explicit);

  if (!fs.existsSync(keyPath)) {
    throw new Error(
      `Service account key not found at ${keyPath}. ` +
      `Download a JSON key from GCP IAM and point GOOGLE_APPLICATION_CREDENTIALS at it.`
    );
  }

  cached = new google.auth.GoogleAuth({ keyFile: keyPath, scopes: SCOPES });
  return cached;
}
