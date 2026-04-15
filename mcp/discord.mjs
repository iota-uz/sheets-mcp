/**
 * Discord REST API client for reading messages from #finances channel.
 *
 * Reads messages via the Discord REST API (no gateway/websocket needed),
 * downloads image attachments to a local folder, and returns structured
 * data that Claude can process.
 *
 * Config is read from environment variables:
 *   DISCORD_BOT_TOKEN, DISCORD_CHANNEL_ID
 */

import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import { loadEnv } from "./env.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = path.resolve(__dirname, "..");
const IMAGES_DIR = path.join(PROJECT_ROOT, "reports", "discord");

const DISCORD_API = "https://discord.com/api/v10";

// ───────────────────────────────────────────────────────────────────────────
// Discord REST API
// ───────────────────────────────────────────────────────────────────────────

async function discordFetch(endpoint, token) {
  const res = await fetch(`${DISCORD_API}${endpoint}`, {
    headers: { Authorization: `Bot ${token}` },
  });

  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Discord API ${res.status}: ${body}`);
  }

  return res.json();
}

async function downloadFile(url, destPath) {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Failed to download ${url}: ${res.status}`);

  const buffer = Buffer.from(await res.arrayBuffer());
  fs.mkdirSync(path.dirname(destPath), { recursive: true });
  fs.writeFileSync(destPath, buffer);
  return destPath;
}

// ───────────────────────────────────────────────────────────────────────────
// Public API
// ───────────────────────────────────────────────────────────────────────────

/**
 * Read the last N messages from the #finances Discord channel.
 * Downloads image attachments to `Reports/discord/`.
 */
export async function readFinancesChannel(options = {}) {
  const env = loadEnv();
  const token = env.DISCORD_BOT_TOKEN;
  const channelId = env.DISCORD_CHANNEL_ID;

  if (!token || !channelId) {
    throw new Error("DISCORD_BOT_TOKEN and DISCORD_CHANNEL_ID must be set in .env");
  }

  const limit = Math.min(options.limit || 10, 100);
  let endpoint = `/channels/${channelId}/messages?limit=${limit}`;
  if (options.after) endpoint += `&after=${options.after}`;

  const messages = await discordFetch(endpoint, token);
  messages.reverse(); // chronological order

  const result = [];

  for (const msg of messages) {
    const imagePaths = [];
    for (const att of msg.attachments || []) {
      if (att.content_type && att.content_type.startsWith("image/")) {
        const ext = path.extname(att.filename) || ".png";
        const localName = `${msg.id}_${att.id}${ext}`;
        const destPath = path.join(IMAGES_DIR, localName);

        try {
          await downloadFile(att.url, destPath);
          imagePaths.push(destPath);
        } catch (err) {
          console.error(`Failed to download ${att.filename}: ${err.message}`);
        }
      }
    }

    result.push({
      id: msg.id,
      author: msg.author?.username || msg.author?.global_name || "unknown",
      date: msg.timestamp ? new Date(msg.timestamp).toLocaleDateString("ru-RU") : "",
      text: msg.content || "",
      images: imagePaths,
    });
  }

  return {
    channel_id: channelId,
    message_count: result.length,
    messages: result,
  };
}
