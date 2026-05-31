/**
 * The `sheets` global bound into the sheets_exec sandbox — one per exec.
 *
 * makeSheetsApi(spreadsheetId, client) is the per-exec SESSION. It owns:
 *   - a handle registry (byKey) so repeated sheets.sheet(name) calls reuse a
 *     handle, scoped to this exec (no process-global cache → concurrent execs
 *     don't clobber each other; issue #7 limitation #1).
 *   - the structural ops, which keep the registry coherent: rename updates and
 *     re-keys live handles, delete marks them dead (issue #7 limitation #2).
 *
 * Everything mutating goes through the injected `client`, whose batchUpdate /
 * valuesUpdate honor the dry-run capture. Reads pass through.
 */

import { createSheet, normHeader } from "./sheet.mjs";
import {
  buildAddSheet,
  buildDeleteSheet,
  buildRenameSheet,
  buildDuplicateSheet,
} from "./requests.mjs";

export function makeSheetsApi(spreadsheetId, client) {
  // Per-exec handle registry. Key: `${normHeader(name)}|${headerRow}`.
  const byKey = new Map();

  const keyFor = (name, headerRow) => `${normHeader(name)}|${headerRow}`;

  // After a column op on one handle, refresh sibling handles to the same tab
  // (opened with a different headerRow) so they don't keep stale positions.
  async function reinitSiblings(sheetId, except) {
    const jobs = [];
    for (const h of byKey.values()) {
      if (h !== except && h.sheetId === sheetId && !h._deleted) jobs.push(h._init());
    }
    await Promise.all(jobs);
  }

  async function getSheet(name, opts = {}) {
    const headerRow = opts.headerRow ?? 1;
    const key = keyFor(name, headerRow);
    if (byKey.has(key)) return byKey.get(key);
    const handle = await createSheet(spreadsheetId, name, opts, client);
    handle._onStructureChange = (h) => reinitSiblings(h.sheetId, h);
    byKey.set(key, handle);
    return handle;
  }

  /**
   * Resolve a sheet title OR numeric id to a sheetId.
   *   number → already an id (returned as-is, no round trip)
   *   string → matched by normHeader(title); throws on not-found / ambiguous.
   */
  async function resolveSheetId(nameOrId) {
    if (typeof nameOrId === "number") return nameOrId;
    const name = String(nameOrId);
    const meta = await client.getSpreadsheet(spreadsheetId);
    const target = normHeader(name);
    const matches = (meta.sheets || []).filter(s => normHeader(s.properties?.title) === target);
    if (matches.length === 0) {
      const known = (meta.sheets || []).map(s => s.properties?.title).filter(Boolean).join(", ");
      throw new Error(`Sheet "${name}" not found. Known: ${known}`);
    }
    if (matches.length > 1) {
      throw new Error(`Sheet name "${name}" is ambiguous — matches ${matches.length} tabs. Pass the numeric sheetId.`);
    }
    return matches[0].properties.sheetId;
  }

  return {
    // ── per-sheet handles ──
    sheet: getSheet,
    spreadsheetId: () => spreadsheetId,

    // ── raw escape hatch (all via the client → dry-run aware) ──
    batchUpdate: (requests, opts) => client.batchUpdate(spreadsheetId, requests, opts),
    valuesBatchGet: (ranges, opts) => client.valuesBatchGet(spreadsheetId, ranges, opts),
    // Default fetch returns sheet/dev-metadata; pass { fields } / { ranges } for
    // full control (named ranges, grid data) via the lower-level get.
    getSpreadsheet: (opts = {}) =>
      (opts.fields || opts.ranges)
        ? client.spreadsheetsGet(spreadsheetId, opts)
        : client.getSpreadsheet(spreadsheetId, opts),
    developerMetadataSearch: (filters) => client.developerMetadataSearch(spreadsheetId, filters),

    // ── structural ops ──
    async addSheet(title, opts = {}) {
      const res = await client.batchUpdate(spreadsheetId, buildAddSheet({ title, ...opts }));
      const props = res?.replies?.[0]?.addSheet?.properties;
      if (!props || props.sheetId == null) return { sheetId: null, title, dryRun: true };
      return { sheetId: props.sheetId, title: props.title, index: props.index };
    },

    async deleteSheet(nameOrId) {
      const sheetId = await resolveSheetId(nameOrId);
      await client.batchUpdate(spreadsheetId, buildDeleteSheet(sheetId));
      for (const [k, h] of [...byKey]) {
        if (h.sheetId === sheetId) { h._markDeleted(); byKey.delete(k); }
      }
      return { ok: true, sheetId };
    },

    async renameSheet(nameOrId, newTitle) {
      const sheetId = await resolveSheetId(nameOrId);
      await client.batchUpdate(spreadsheetId, buildRenameSheet(sheetId, newTitle));
      // Update + re-key every cached handle for this tab so handles held in the
      // script keep working after the rename.
      for (const [k, h] of [...byKey]) {
        if (h.sheetId === sheetId) {
          byKey.delete(k);
          h.name = newTitle;
          byKey.set(keyFor(newTitle, h.headerRow), h);
        }
      }
      return { ok: true, sheetId, title: newTitle };
    },

    async duplicateSheet(nameOrId, newTitle) {
      const sourceSheetId = await resolveSheetId(nameOrId);
      const res = await client.batchUpdate(
        spreadsheetId,
        buildDuplicateSheet(sourceSheetId, newTitle != null ? { newTitle } : {}),
      );
      const props = res?.replies?.[0]?.duplicateSheet?.properties;
      if (!props || props.sheetId == null) return { sheetId: null, title: newTitle ?? null, dryRun: true };
      return { sheetId: props.sheetId, title: props.title };
    },
  };
}
