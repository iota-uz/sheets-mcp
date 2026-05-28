/**
 * Discord REST API client for reading messages from a configured channel.
 *
 * Reads messages via the Discord REST API (no gateway/websocket needed),
 * downloads image attachments to a local folder, and returns structured
 * data that Claude can process.
 *
 * Config read from environment:
 *   DISCORD_BOT_TOKEN    Bot token
 *   DISCORD_CHANNEL_ID   Channel to read from
 *   DISCORD_IMAGES_DIR   (optional) Where to write attachments.
 *                        Default: <cwd>/reports/discord
 */

import fs from "fs";
import path from "path";

const DISCORD_API = "https://discord.com/api/v10";

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

export async function readFinancesChannel(options = {}) {
  const token = process.env.DISCORD_BOT_TOKEN;
  const channelId = process.env.DISCORD_CHANNEL_ID;

  if (!token || !channelId) {
    throw new Error("DISCORD_BOT_TOKEN and DISCORD_CHANNEL_ID must be set in the MCP env.");
  }

  const imagesDir = process.env.DISCORD_IMAGES_DIR
    || path.join(process.cwd(), "reports", "discord");

  const limit = Math.min(options.limit || 10, 100);
  let endpoint = `/channels/${channelId}/messages?limit=${limit}`;
  if (options.after) endpoint += `&after=${options.after}`;

  const messages = await discordFetch(endpoint, token);
  messages.reverse();

  const result = [];

  for (const msg of messages) {
    const imagePaths = [];
    for (const att of msg.attachments || []) {
      if (att.content_type && att.content_type.startsWith("image/")) {
        const ext = path.extname(att.filename) || ".png";
        const localName = `${msg.id}_${att.id}${ext}`;
        const destPath = path.join(imagesDir, localName);

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
