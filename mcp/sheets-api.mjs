/**
 * The `sheets` global bound into the sheets_exec sandbox.
 *
 * Built by makeSheetsApi(spreadsheetId) so the surface is defined in one place
 * (not inlined in the runner) and is importable/testable. It exposes:
 *   - the existing per-sheet handle factory (sheet, spreadsheetId)
 *   - a raw escape hatch over the low-level client (batchUpdate + reads)
 *   - structural spreadsheet ops (add/delete/rename/duplicate a tab)
 *
 * Everything mutating compiles to batchUpdate, so it automatically respects the
 * runner's dry-run capture. Structural ops resolve a sheet name → sheetId with
 * one getSpreadsheet() and clear the Sheet handle cache afterward so a later
 * sheets.sheet() in the same script re-reads fresh metadata.
 */

import { sheet as makeSheet, clearCache, normHeader } from "./sheet.mjs";
import {
  batchUpdate,
  valuesBatchGet,
  getSpreadsheet,
  developerMetadataSearch,
} from "./sheets-client.mjs";
import {
  buildAddSheet,
  buildDeleteSheet,
  buildRenameSheet,
  buildDuplicateSheet,
} from "./requests.mjs";

/**
 * Resolve a sheet title OR numeric id to a sheetId.
 *   number → already an id (returned as-is)
 *   string → matched by normHeader(title); throws on not-found / ambiguous.
 */
async function resolveSheetId(spreadsheetId, nameOrId) {
  if (typeof nameOrId === "number") return nameOrId;
  const name = String(nameOrId);
  const meta = await getSpreadsheet(spreadsheetId);
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

export function makeSheetsApi(spreadsheetId) {
  return {
    // ── existing surface (unchanged) ──
    sheet: (name, opts) => makeSheet(spreadsheetId, name, opts),
    spreadsheetId: () => spreadsheetId,

    // ── raw escape hatch ──
    batchUpdate: (requests, opts) => batchUpdate(spreadsheetId, requests, opts),
    valuesBatchGet: (ranges, opts) => valuesBatchGet(spreadsheetId, ranges, opts),
    getSpreadsheet: (opts) => getSpreadsheet(spreadsheetId, opts),
    developerMetadataSearch: (filters) => developerMetadataSearch(spreadsheetId, filters),

    // ── structural ops ──
    async addSheet(title, opts = {}) {
      const res = await batchUpdate(spreadsheetId, buildAddSheet({ title, ...opts }));
      clearCache();
      const props = res?.replies?.[0]?.addSheet?.properties;
      if (!props || props.sheetId == null) return { sheetId: null, title, dryRun: true };
      return { sheetId: props.sheetId, title: props.title, index: props.index };
    },

    async deleteSheet(nameOrId) {
      const sheetId = await resolveSheetId(spreadsheetId, nameOrId);
      await batchUpdate(spreadsheetId, buildDeleteSheet(sheetId));
      clearCache();
      return { ok: true, sheetId };
    },

    async renameSheet(nameOrId, newTitle) {
      const sheetId = await resolveSheetId(spreadsheetId, nameOrId);
      await batchUpdate(spreadsheetId, buildRenameSheet(sheetId, newTitle));
      clearCache();
      return { ok: true, sheetId, title: newTitle };
    },

    async duplicateSheet(nameOrId, newTitle) {
      const sourceSheetId = await resolveSheetId(spreadsheetId, nameOrId);
      const res = await batchUpdate(
        spreadsheetId,
        buildDuplicateSheet(sourceSheetId, newTitle != null ? { newTitle } : {}),
      );
      clearCache();
      const props = res?.replies?.[0]?.duplicateSheet?.properties;
      if (!props || props.sheetId == null) return { sheetId: null, title: newTitle ?? null, dryRun: true };
      return { sheetId: props.sheetId, title: props.title };
    },
  };
}
