/**
 * Operational Metrics Engine (Google Sheets + Apps Script)
 * Version: 1.0.0
 *
 * Author: Eduardo Sousa
 *
 * License (Dual):
 * - Open Source: GNU AGPLv3 (see LICENSE)
 * - Commercial: See COMMERCIAL_LICENSE.md
 *
 * SUMMARY (commercial):
 * - Converts a raw input dataset (INPUT_DATA) into a team-ready board (TEAM_BOARD)
 * - Persists operational edits into HISTORY by a unique ID
 * - Applies validations, protections, conditional formatting and stable layout
 * - Optional monthly snapshot routine to freeze "previous period" metrics in HISTORY
 *
 * ✅ CUSTOMIZE WITHOUT TOUCHING CORE:
 * - Use CONFIG keys to map input columns and rules:
 *   MAP_ID, MAP_OWNER, MAP_SUBJECT, MAP_CREATED_AT, MAP_PRIORITY
 *   EDITABLE_FIELDS, STATUS_OPTIONS, ESSENTIAL_COLUMNS
 *   ESSENTIAL_BY_HEADER_COLOR, ESSENTIAL_COLOR_HEX, COLOR_TOLERANCE
 *   IGNORE_TEXT, PROTECT_NON_EDITABLE
 *
 * SECURITY NOTE:
 * - Never publish real customer data. Use anonymized/sample datasets only.
 */

/* =========================
 * SHEETS
 * ========================= */
const SHEETS = {
  ABOUT: "ABOUT",
  CONFIG: "CONFIG",
  INPUT: "INPUT_DATA",
  HISTORY: "HISTORY",
  TEAM: "TEAM_BOARD",
  DASH: "DASHBOARD",
  LOGS: "LOGS",
};

/* =========================
 * DEFAULT HISTORY HEADERS (generic)
 * - You can add more columns; the engine will keep them.
 * ========================= */
const HISTORY_HEADERS_DEFAULT = [
  "RECORD_ID",
  "OWNER",
  "STATUS",
  "NEXT_ACTION",
  "ATTEMPTS",
  "CONTACTED_AT",
  "NOTE",
  "VALUE",
  "PREV_PERIOD_FLAG",
  "PREV_PERIOD_TIER",
  "UPDATED_AT",
];

/* =========================
 * MENU
 * ========================= */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("Operational Engine")
      .addItem("1) Install / Ensure Structure", "installTemplate")
      .addSeparator()
      .addItem("2) Monthly Snapshot (optional)", "importMonthlySnapshot")
      .addItem("3) Daily Update", "runDailyUpdate")
      .addToUi();
  } catch (e) {
    console.error(e);
  }
}

/* =========================
 * INSTALL
 * ========================= */
function installTemplate() {
  const ss = SpreadsheetApp.getActive();

  const shAbout = getOrCreateSheet_(ss, SHEETS.ABOUT, 1);
  const shCfg   = getOrCreateSheet_(ss, SHEETS.CONFIG, 2);
  const shIn    = getOrCreateSheet_(ss, SHEETS.INPUT, 3);
  const shHist  = getOrCreateSheet_(ss, SHEETS.HISTORY, 4);
  const shTeam  = getOrCreateSheet_(ss, SHEETS.TEAM, 5);
  const shDash  = getOrCreateSheet_(ss, SHEETS.DASH, 6);
  const shLogs  = getOrCreateSheet_(ss, SHEETS.LOGS, 7);

  ensureAbout_(shAbout);
  ensureConfig_(shCfg);
  ensureHeaders_(shHist, HISTORY_HEADERS_DEFAULT);

  // Optional: seed input header if empty
  if (shIn.getLastRow() === 0) {
    shIn.getRange(1, 1, 1, 5).setValues([["id", "owner", "subject", "created_at", "priority"]]);
    shIn.setFrozenRows(1);
  }

  SpreadsheetApp.getUi().alert("Template structure ensured ✅");
}

/* =========================
 * onEdit:
 * - Sync edits made in TEAM_BOARD back into HISTORY
 * ========================= */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== SHEETS.TEAM) return;

    const ss = e.source;
    const shCfg  = ss.getSheetByName(SHEETS.CONFIG);
    const shHist = ss.getSheetByName(SHEETS.HISTORY);
    if (!shCfg || !shHist) return;

    const cfg = readConfig_(shCfg);

    // Team headers
    const lastCol = sh.getLastColumn();
    if (lastCol < 1) return;

    const teamHeaders = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(x => String(x||"").trim());
    const tIdx = headerIndex_(teamHeaders);

    const idCol = tIdx["RECORD_ID"];
    if (!idCol) return;

    // Editable fields from CONFIG
    const editable = parseList_(cfg.EDITABLE_FIELDS);
    const editableSet = new Set(editable);

    ensureHeaders_(shHist, HISTORY_HEADERS_DEFAULT);
    const hist = loadHistoryIndex_(shHist);

    const r0 = e.range.getRow();
    const c0 = e.range.getColumn();
    const nr = e.range.getNumRows();
    const nc = e.range.getNumColumns();
    if (r0 < 2) return;

    const values = e.range.getValues();

    for (let dr = 0; dr < nr; dr++) {
      const row = r0 + dr;
      const recordId = String(sh.getRange(row, idCol).getValue() || "").trim();
      if (!recordId) continue;

      const histRow = hist.mapIdToRow.get(recordId);
      if (!histRow) continue; // no creation from onEdit (safe)

      for (let dc = 0; dc < nc; dc++) {
        const col = c0 + dc;
        const field = teamHeaders[col - 1];
        if (!editableSet.has(field)) continue;

        const hCol = hist.headerIndex[field];
        if (!hCol) continue;

        shHist.getRange(histRow, hCol).setValue(values[dr][dc]);
      }

      const tsCol = hist.headerIndex["UPDATED_AT"];
      if (tsCol) shHist.getRange(histRow, tsCol).setValue(new Date());
    }
  } catch (_) {}
}

/* =========================================================
 * DAILY UPDATE
 * - Reads INPUT_DATA, maps columns, merges with HISTORY, builds TEAM_BOARD
 * ========================================================= */
function runDailyUpdate() {
  const ss = SpreadsheetApp.getActive();

  const shCfg  = ss.getSheetByName(SHEETS.CONFIG);
  const shIn   = ss.getSheetByName(SHEETS.INPUT);
  const shHist = ss.getSheetByName(SHEETS.HISTORY);
  const shTeam = ss.getSheetByName(SHEETS.TEAM);
  const shDash = ss.getSheetByName(SHEETS.DASH);
  const shLogs = ss.getSheetByName(SHEETS.LOGS);

  if (!shCfg || !shIn || !shHist || !shTeam || !shDash || !shLogs) {
    SpreadsheetApp.getUi().alert("Missing sheets. Run: Install / Ensure Structure");
    return;
  }

  clearLegacyProtectionsAndValidations_(ss);

  const cfg = readConfig_(shCfg);
  ensureHeaders_(shHist, HISTORY_HEADERS_DEFAULT);

  const input = readTable_(shIn);
  if (!input.rows.length) {
    shTeam.clear();
    SpreadsheetApp.getUi().alert("INPUT_DATA is empty.");
    return;
  }

  const mapped = mapInput_(input, cfg); // <- core mapping + fallbacks
  const hist = loadHistoryFull_(shHist);

  // Essentials (extra columns)
  const essentials = resolveEssentialColumns_(input.headers, input.headerBg, cfg);

  // Build TEAM headers (stable core + essentials)
  const teamHeaders = buildTeamHeaders_(essentials);

  // Merge rows
  const outRows = [];
  const newHistRows = [];
  const now = new Date();

  const ignoreText = String(cfg.IGNORE_TEXT || "").trim().toLowerCase();

  for (const item of mapped) {
    if (!item.RECORD_ID) continue;

    if (ignoreText) {
      const joined = Object.values(item).map(v => String(v ?? "")).join(" ").toLowerCase();
      if (joined.includes(ignoreText)) continue;
    }

    // Load history state (persisted fields)
    const histRowNum = hist.mapIdToRow.get(item.RECORD_ID);
    let histObj;

    if (histRowNum) {
      histObj = hist.rowObjects.get(item.RECORD_ID) || defaultHistoryObject_(item, cfg);
      // allow daily owner overwrite if configured
      if (cfg.DAILY_OVERWRITE_OWNER === "YES" && hist.headerIndex["OWNER"]) {
        const pos = histRowNum - 2;
        hist.colOWNER[pos][0] = item.OWNER || histObj.OWNER || "";
      }
    } else {
      // create new history row with defaults
      const newRow = buildNewHistoryRow_(item, hist.headerIndex, hist.headers.length, now);
      newHistRows.push(newRow);
      histObj = defaultHistoryObject_(item, cfg);
    }

    // Compose TEAM row
    const row = [];
    const extraMap = item.__EXTRA || {};

    for (const h of teamHeaders) {
      if (h in histObj) row.push(histObj[h] ?? "");
      else if (h in item) row.push(item[h] ?? "");
      else row.push(extraMap[h] ?? "");
    }
    outRows.push(row);
  }

  // Append new HISTORY rows
  if (newHistRows.length) {
    const start = shHist.getLastRow() + 1;
    ensureSheetSize_(shHist, start + newHistRows.length - 1, newHistRows[0].length);
    clearValidations_(shHist, start, 1, newHistRows.length, newHistRows[0].length);
    shHist.getRange(start, 1, newHistRows.length, newHistRows[0].length).setValues(newHistRows);
  }

  // Update OWNER column if we buffered it
  if (hist.nRows > 0 && hist.headerIndex["OWNER"]) {
    clearValidations_(shHist, 2, hist.headerIndex["OWNER"], hist.nRows, 1);
    shHist.getRange(2, hist.headerIndex["OWNER"], hist.nRows, 1).setValues(hist.colOWNER);
  }

  // Write TEAM_BOARD
  shTeam.clear();
  hardResetSurface_(shTeam);
  shTeam.getRange(1, 1, 1, teamHeaders.length).setValues([teamHeaders]);

  if (outRows.length) {
    clearValidations_(shTeam, 2, 1, outRows.length, teamHeaders.length);
    shTeam.getRange(2, 1, outRows.length, teamHeaders.length).setValues(outRows);
  }

  // Apply styles
  applyNumberFormats_(shTeam, teamHeaders);
  applyPriorityRowColors_(shTeam, teamHeaders);
  applyValidationsAndProtections_(shTeam, teamHeaders, cfg);
  applyLayoutStyles_(shTeam, teamHeaders, { isTeam: true });

  // Minimal DASH (optional – placeholder)
  buildBasicDashboard_(shDash, shTeam, teamHeaders);

  // Log
  appendLog_(shLogs, `Daily update OK. Rows: ${outRows.length}`);

  SpreadsheetApp.getUi().alert("Daily update ✅");
}

/* =========================================================
 * MONTHLY SNAPSHOT (optional)
 * - Freezes prev-period metrics into HISTORY
 * - This is optional and can be adapted to any business meaning
 * ========================================================= */
function importMonthlySnapshot() {
  const ss = SpreadsheetApp.getActive();
  const shCfg  = ss.getSheetByName(SHEETS.CONFIG);
  const shIn   = ss.getSheetByName(SHEETS.INPUT);
  const shHist = ss.getSheetByName(SHEETS.HISTORY);
  const shLogs = ss.getSheetByName(SHEETS.LOGS);

  if (!shCfg || !shIn || !shHist || !shLogs) {
    SpreadsheetApp.getUi().alert("Missing sheets. Run Install first.");
    return;
  }

  const cfg = readConfig_(shCfg);
  ensureHeaders_(shHist, HISTORY_HEADERS_DEFAULT);

  // Using INPUT_DATA as snapshot source by default.
  // ✅ CUSTOMIZE HERE:
  // If you prefer a separate sheet (MONTHLY_SNAPSHOT), create it and read from there.
  const snapshot = readTable_(shIn);
  if (!snapshot.rows.length) {
    SpreadsheetApp.getUi().alert("Snapshot source is empty.");
    return;
  }

  const snapMapped = mapInput_(snapshot, cfg);
  const hist = loadHistoryFull_(shHist);

  const now = new Date();

  const flagField = cfg.PREV_PERIOD_FLAG_FIELD || "PREV_PERIOD_FLAG";
  const tierField = cfg.PREV_PERIOD_TIER_FIELD || "PREV_PERIOD_TIER";

  // Ensure fields exist in HISTORY (append headers if missing)
  ensureHeaders_(shHist, Array.from(new Set(HISTORY_HEADERS_DEFAULT.concat([flagField, tierField]))));

  // Reload after ensure
  const hist2 = loadHistoryFull_(shHist);

  for (const item of snapMapped) {
    const id = item.RECORD_ID;
    if (!id) continue;

    const rowNum = hist2.mapIdToRow.get(id);
    if (!rowNum) continue;

    // ✅ CUSTOMIZE HERE:
    // Define what "prev period flag" means in your operation.
    // Default: if record appears in snapshot => YES.
    const prevFlag = "YES";

    // ✅ CUSTOMIZE HERE:
    // Define how to compute tier. Default: if PRIORITY_SCORE exists, clamp 0-4.
    const prevTier = clamp0to4_(toIntSafe_(item.PRIORITY_SCORE));

    const idx = hist2.headerIndex;
    if (idx[flagField]) shHist.getRange(rowNum, idx[flagField]).setValue(prevFlag);
    if (idx[tierField]) shHist.getRange(rowNum, idx[tierField]).setValue(prevTier);
    if (idx["UPDATED_AT"]) shHist.getRange(rowNum, idx["UPDATED_AT"]).setValue(now);
  }

  appendLog_(shLogs, "Monthly snapshot imported OK.");
  SpreadsheetApp.getUi().alert("Monthly snapshot ✅");
}

/* =========================
 * INPUT MAPPING (core)
 * - Maps any dataset into a standard object:
 *   RECORD_ID, OWNER, SUBJECT_NAME, CREATED_AT, PRIORITY_SCORE
 * - Keeps extras for "essentials"
 * ========================= */
function mapInput_(table, cfg) {
  const idx = headerIndex_(table.headers);

  const idCol = idx[cfg.MAP_ID] || idx["RECORD_ID"] || idx["ID"] || idx["CNPJ"] || idx["ticket_id"] || idx["lead_id"];
  const ownerCol = idx[cfg.MAP_OWNER] || idx["OWNER"] || idx["GERENTE"] || idx["ASSIGNEE"];
  const subjCol = idx[cfg.MAP_SUBJECT] || idx["SUBJECT_NAME"] || idx["NOME_CLIENTE"] || idx["CLIENT_NAME"] || idx["SUBJECT"];
  const createdCol = idx[cfg.MAP_CREATED_AT] || idx["CREATED_AT"] || idx["DT_CONTA_CRIADA"] || idx["created_at"];
  const prioCol = idx[cfg.MAP_PRIORITY] || idx["PRIORITY_SCORE"] || idx["SINALEIRO"] || idx["PRIORITY"];

  const essentials = parseList_(cfg.ESSENTIAL_COLUMNS);

  const out = [];
  for (const r of table.rows) {
    const obj = {
      RECORD_ID: idCol ? String(r[idCol - 1] || "").trim() : "",
      OWNER: ownerCol ? String(r[ownerCol - 1] || "").trim() : "",
      SUBJECT_NAME: subjCol ? String(r[subjCol - 1] || "").trim() : "",
      CREATED_AT: createdCol ? r[createdCol - 1] : "",
      PRIORITY_SCORE: prioCol ? clamp0to4_(toIntSafe_(r[prioCol - 1])) : 0,
      __EXTRA: {},
    };

    // Essentials fixed list
    for (const h of essentials) {
      const col = idx[h];
      if (col) obj.__EXTRA[h] = r[col - 1];
    }

    out.push(obj);
  }
  return out;
}

/* =========================
 * TEAM headers (stable)
 * ========================= */
function buildTeamHeaders_(essentials) {
  return [
    "PRIORITY_SCORE",
    "OWNER",
    "RECORD_ID",
    "SUBJECT_NAME",
    "CREATED_AT",
    "STATUS",
    "NEXT_ACTION",
    "ATTEMPTS",
    "CONTACTED_AT",
    "NOTE",
    "VALUE",
    "PREV_PERIOD_FLAG",
    "PREV_PERIOD_TIER",
    "UPDATED_AT",
  ].concat(essentials);
}

/* =========================
 * HISTORY helpers
 * ========================= */
function ensureHeaders_(sh, headersWanted) {
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headersWanted.length).setValues([headersWanted]);
    sh.setFrozenRows(1);
    return;
  }

  const lc = sh.getLastColumn();
  const cur = sh.getRange(1, 1, 1, lc).getValues()[0].map(v => String(v||"").trim());
  const set = new Set(cur);
  const missing = headersWanted.filter(h => !set.has(h));

  if (missing.length) {
    sh.getRange(1, lc + 1, 1, missing.length).setValues([missing]);
  }
  sh.setFrozenRows(1);
}

function loadHistoryIndex_(shHist) {
  const lr = shHist.getLastRow(), lc = shHist.getLastColumn();
  const headers = shHist.getRange(1, 1, 1, lc).getValues()[0].map(v => String(v||"").trim());
  const hIdx = headerIndex_(headers);
  const idCol = hIdx["RECORD_ID"];
  const map = new Map();

  if (!idCol || lr < 2) return { headers, headerIndex: hIdx, mapIdToRow: map };

  const ids = shHist.getRange(2, idCol, lr - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    const id = String(ids[i][0] || "").trim();
    if (id && !map.has(id)) map.set(id, i + 2);
  }
  return { headers, headerIndex: hIdx, mapIdToRow: map };
}

function loadHistoryFull_(shHist) {
  const lr = shHist.getLastRow(), lc = shHist.getLastColumn();
  const headers = shHist.getRange(1, 1, 1, lc).getValues()[0].map(v => String(v||"").trim());
  const hIdx = headerIndex_(headers);
  const nRows = Math.max(0, lr - 1);
  const values = nRows ? shHist.getRange(2, 1, nRows, lc).getValues() : [];

  const idPos = hIdx["RECORD_ID"] ? hIdx["RECORD_ID"] - 1 : 0;
  const ownerPos = hIdx["OWNER"] ? hIdx["OWNER"] - 1 : null;

  const map = new Map();
  const rowObjects = new Map();
  const colOWNER = [];

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const id = String(row[idPos] || "").trim();
    colOWNER.push([ownerPos != null ? row[ownerPos] : ""]);

    if (!id) continue;
    if (map.has(id)) continue;

    map.set(id, i + 2);

    const obj = {};
    for (let c = 0; c < headers.length; c++) {
      const h = headers[c];
      if (h) obj[h] = row[c];
    }
    rowObjects.set(id, obj);
  }

  return {
    headers,
    headerIndex: hIdx,
    nRows,
    mapIdToRow: map,
    rowObjects,
    colOWNER,
  };
}

function buildNewHistoryRow_(item, hIdx, totalCols, now) {
  const row = new Array(totalCols).fill("");

  if (hIdx["RECORD_ID"]) row[hIdx["RECORD_ID"] - 1] = item.RECORD_ID;
  if (hIdx["OWNER"]) row[hIdx["OWNER"] - 1] = item.OWNER || "";
  if (hIdx["STATUS"]) row[hIdx["STATUS"] - 1] = "";
  if (hIdx["NEXT_ACTION"]) row[hIdx["NEXT_ACTION"] - 1] = "";
  if (hIdx["ATTEMPTS"]) row[hIdx["ATTEMPTS"] - 1] = "";
  if (hIdx["CONTACTED_AT"]) row[hIdx["CONTACTED_AT"] - 1] = "";
  if (hIdx["NOTE"]) row[hIdx["NOTE"] - 1] = "";
  if (hIdx["VALUE"]) row[hIdx["VALUE"] - 1] = "";

  if (hIdx["PREV_PERIOD_FLAG"]) row[hIdx["PREV_PERIOD_FLAG"] - 1] = "NO";
  if (hIdx["PREV_PERIOD_TIER"]) row[hIdx["PREV_PERIOD_TIER"] - 1] = 0;
  if (hIdx["UPDATED_AT"]) row[hIdx["UPDATED_AT"] - 1] = now;

  return row;
}

function defaultHistoryObject_(item) {
  return {
    RECORD_ID: item.RECORD_ID,
    OWNER: item.OWNER || "",
    STATUS: "",
    NEXT_ACTION: "",
    ATTEMPTS: "",
    CONTACTED_AT: "",
    NOTE: "",
    VALUE: "",
    PREV_PERIOD_FLAG: "NO",
    PREV_PERIOD_TIER: 0,
    UPDATED_AT: "",
  };
}

/* =========================
 * Essentials (fixed list + optional header color detection)
 * ========================= */
function resolveEssentialColumns_(headers, headerBg, cfg) {
  const fixed = parseList_(cfg.ESSENTIAL_COLUMNS);
  const byColor = (String(cfg.ESSENTIAL_BY_HEADER_COLOR || "").toUpperCase() === "YES")
    ? detectEssentialByColor_(headers, headerBg, cfg)
    : [];

  return uniqueList_(fixed.concat(byColor))
    .filter(h => !["RECORD_ID","OWNER","SUBJECT_NAME","CREATED_AT","PRIORITY_SCORE"].includes(String(h||"").trim()));
}

function detectEssentialByColor_(headers, headerBg, cfg) {
  const target = normalizeColor_(cfg.ESSENTIAL_COLOR_HEX || "#FFFF00");
  const tol = clampInt_(cfg.COLOR_TOLERANCE || 110, 0, 255);
  const out = [];

  for (let i = 0; i < headers.length; i++) {
    const name = String(headers[i] || "").trim();
    if (!name) continue;
    const bg = normalizeColor_(headerBg[i]);
    if (isYellowTolerant_(bg, target, tol)) out.push(name);
  }
  return out;
}

/* =========================
 * Validations / protections / formatting / layout
 * ========================= */
function applyValidationsAndProtections_(sh, headers, cfg) {
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2) return;

  const idx = headerIndex_(headers);

  // STATUS dropdown
  const statusOptions = parseList_(cfg.STATUS_OPTIONS || "New;In Progress;Waiting;Done;Lost");
  setDropdown_(sh, idx["STATUS"], statusOptions);

  // protections
  const protect = String(cfg.PROTECT_NON_EDITABLE || "YES").toUpperCase() === "YES";
  if (!protect) return;

  // remove old protections
  sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => { try { p.remove(); } catch(_){} });

  const editable = new Set(parseList_(cfg.EDITABLE_FIELDS || "STATUS;NEXT_ACTION;ATTEMPTS;CONTACTED_AT;NOTE;VALUE"));
  // Always allow these basics to be edited
  editable.add("STATUS"); editable.add("NEXT_ACTION"); editable.add("ATTEMPTS"); editable.add("CONTACTED_AT"); editable.add("NOTE"); editable.add("VALUE");

  for (let c = 1; c <= lc; c++) {
    const name = headers[c - 1];
    // lock everything not editable
    if (!editable.has(name)) {
      sh.getRange(2, c, lr - 1, 1).protect().setDescription("Locked: " + name).setWarningOnly(false);
    }
  }

  // Freeze key columns
  sh.setFrozenColumns(4);
}

function applyPriorityRowColors_(sh, headers) {
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return;

  const idx = headerIndex_(headers);
  const pCol = idx["PRIORITY_SCORE"];
  if (!pCol) return;

  const letter = colToLetter_(pCol);
  const range = sh.getRange(2, 1, lr - 1, lc);

  sh.setConditionalFormatRules([]);

  // Example: 0 = attention (red-ish), >0 = ok (green-ish)
  const rule0 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$${letter}2=0`)
    .setBackground("#ffe3e3")
    .setRanges([range])
    .build();

  const ruleG = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$${letter}2>0`)
    .setBackground("#d3f9d8")
    .setRanges([range])
    .build();

  sh.setConditionalFormatRules([rule0, ruleG]);
}

function applyNumberFormats_(sh, headers) {
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return;

  const FMT_TEXT = "@";
  const FMT_INT = "0";
  const FMT_DATE = "dd/MM/yyyy";
  const FMT_DATETIME = "dd/MM/yyyy HH:mm";

  const idx = headerIndex_(headers);

  if (idx["RECORD_ID"]) sh.getRange(2, idx["RECORD_ID"], lr - 1, 1).setNumberFormat(FMT_TEXT);

  for (let c = 1; c <= lc; c++) {
    const h = String(headers[c - 1] || "").trim();
    const r = sh.getRange(2, c, lr - 1, 1);

    if (h === "PRIORITY_SCORE" || h === "PREV_PERIOD_TIER") r.setNumberFormat(FMT_INT);
    if (h.endsWith("_AT") || h === "UPDATED_AT") r.setNumberFormat(FMT_DATETIME);
    if (h === "CREATED_AT") r.setNumberFormat(FMT_DATE);
  }
}

function applyLayoutStyles_(sh, headers, opts) {
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lc < 1) return;

  const header = sh.getRange(1, 1, 1, lc);
  header.setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true)
    .setBackground("#d9f2d9");

  sh.setFrozenRows(1);

  try { const f = sh.getFilter(); if (f) f.remove(); } catch(_) {}
  try { sh.getRange(1, 1, Math.max(1, lr), lc).createFilter(); } catch(_) {}

  // widths (simple)
  const widthByHeader = {
    PRIORITY_SCORE: 110,
    OWNER: 180,
    RECORD_ID: 160,
    SUBJECT_NAME: 320,
    STATUS: 170,
    NEXT_ACTION: 220,
    NOTE: 320,
    CREATED_AT: 140,
    UPDATED_AT: 170,
  };

  for (let c = 1; c <= lc; c++) {
    const h = String(headers[c - 1] || "").trim();
    const w = widthByHeader[h] || 140;
    try { sh.setColumnWidth(c, w); } catch(_) {}
  }

  if (opts && opts.isTeam && lr >= 2) {
    sh.setFrozenColumns(4);
  }
}

/* =========================
 * Dashboard (basic placeholder)
 * - Keep simple; can be expanded later
 * ========================= */
function buildBasicDashboard_(shDash, shTeam, headers) {
  shDash.clear();
  shDash.getRange(1,1).setValue("Basic Dashboard (customizable)");
  shDash.getRange(2,1).setValue("Rows in TEAM_BOARD:");
  shDash.getRange(2,2).setValue(Math.max(0, shTeam.getLastRow() - 1));
}

/* =========================
 * CONFIG
 * ========================= */
function ensureConfig_(sh) {
  sh.getRange(1, 1, 1, 2).setValues([["KEY", "VALUE"]]);

  const current = sh.getDataRange().getValues();
  const have = new Set();
  for (let i = 1; i < current.length; i++) {
    const k = String(current[i][0] || "").trim();
    if (k) have.add(k);
  }

  const defaults = [
    ["MAP_ID", "id"],
    ["MAP_OWNER", "owner"],
    ["MAP_SUBJECT", "subject"],
    ["MAP_CREATED_AT", "created_at"],
    ["MAP_PRIORITY", "priority"],

    ["EDITABLE_FIELDS", "STATUS;NEXT_ACTION;ATTEMPTS;CONTACTED_AT;NOTE;VALUE"],
    ["STATUS_OPTIONS", "New;In Progress;Waiting;Done;Lost"],
    ["PROTECT_NON_EDITABLE", "YES"],
    ["DAILY_OVERWRITE_OWNER", "YES"],

    ["ESSENTIAL_COLUMNS", ""],
    ["ESSENTIAL_BY_HEADER_COLOR", "NO"],
    ["ESSENTIAL_COLOR_HEX", "#FFFF00"],
    ["COLOR_TOLERANCE", "110"],

    ["IGNORE_TEXT", ""],

    ["PREV_PERIOD_FLAG_FIELD", "PREV_PERIOD_FLAG"],
    ["PREV_PERIOD_TIER_FIELD", "PREV_PERIOD_TIER"],
  ];

  const toAppend = defaults.filter(([k]) => !have.has(k));
  if (toAppend.length) sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, 2).setValues(toAppend);

  sh.setFrozenRows(1);
}

function readConfig_(sh) {
  const values = sh.getDataRange().getValues();
  const obj = {};
  for (let i = 1; i < values.length; i++) {
    const k = String(values[i][0] || "").trim();
    const v = String(values[i][1] || "").trim();
    if (k) obj[k] = v;
  }
  return obj;
}

function ensureAbout_(sh) {
  sh.clear();
  sh.getRange(1,1).setValue("Operational Metrics Engine");
  sh.getRange(2,1).setValue("Author:");
  sh.getRange(2,2).setValue("Eduardo Sousa");
  sh.getRange(3,1).setValue("License:");
  sh.getRange(3,2).setValue("AGPLv3 + Commercial (dual license)");
  sh.getRange(5,1).setValue("How to use:");
  sh.getRange(6,1).setValue("1) Paste/import your dataset into INPUT_DATA");
  sh.getRange(7,1).setValue("2) Configure column mappings and rules in CONFIG");
  sh.getRange(8,1).setValue("3) Run: Daily Update");
  sh.getRange(9,1).setValue("4) Team works on TEAM_BOARD (edits sync into HISTORY)");
}

/* =========================
 * TABLE reader
 * ========================= */
function readTable_(sh) {
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return { headers: [], rows: [], headerBg: [] };

  const headers = sh.getRange(1, 1, 1, lc).getValues()[0].map(v => String(v||"").trim());
  const headerBg = sh.getRange(1, 1, 1, lc).getBackgrounds()[0].map(c => normalizeColor_(c));
  const rows = sh.getRange(2, 1, lr - 1, lc).getValues();
  return { headers, rows, headerBg };
}

/* =========================
 * Logs
 * ========================= */
function appendLog_(sh, msg) {
  const now = new Date();
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,2).setValues([["timestamp","message"]]);
  sh.appendRow([now, msg]);
}

/* =========================
 * Cleanup
 * ========================= */
function clearLegacyProtectionsAndValidations_(ss) {
  ss.getSheets().forEach(sh => {
    sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => { try { p.remove(); } catch(_){} });
    const lr = sh.getLastRow(), lc = sh.getLastColumn();
    if (lr >= 2 && lc >= 1) sh.getRange(2, 1, lr - 1, lc).clearDataValidations();
  });
}

function hardResetSurface_(sh) {
  sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => { try { p.remove(); } catch(_){} });
  const mr = sh.getMaxRows(), mc = sh.getMaxColumns();
  if (mr > 0 && mc > 0) sh.getRange(1,1,mr,mc).clearDataValidations();
}

function clearValidations_(sh, r, c, nr, nc) {
  if (nr <= 0 || nc <= 0) return;
  sh.getRange(r, c, nr, nc).clearDataValidations();
}

function ensureSheetSize_(sh, needRows, needCols) {
  const mr = sh.getMaxRows(), mc = sh.getMaxColumns();
  if (needRows > mr) sh.insertRowsAfter(mr, needRows - mr);
  if (needCols > mc) sh.insertColumnsAfter(mc, needCols - mc);
}

/* =========================
 * Utils
 * ========================= */
function getOrCreateSheet_(ss, name, position) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (position && typeof position === "number") {
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(position);
  }
  return sh;
}

function headerIndex_(headers) {
  const idx = {};
  headers.forEach((h, i) => {
    h = String(h || "").trim();
    if (h) idx[h] = i + 1;
  });
  return idx;
}

function parseList_(s) {
  return String(s || "")
    .split(";")
    .map(x => String(x||"").trim())
    .filter(Boolean);
}

function uniqueList_(arr) {
  const out = [];
  const set = new Set();
  for (const x of arr) {
    const k = String(x || "").trim();
    if (!k) continue;
    const kk = k.toUpperCase();
    if (!set.has(kk)) {
      set.add(kk);
      out.push(k);
    }
  }
  return out;
}

function normalizeColor_(c) {
  if (!c) return "";
  c = String(c).trim();
  return c[0] === "#" ? c.toUpperCase() : c.toUpperCase();
}

function clampInt_(v, min, max) {
  const n = parseInt(String(v || "").trim(), 10);
  if (isNaN(n)) return min;
  return Math.max(min, Math.min(max, n));
}

function hexToRgb_(hex) {
  if (!hex) return null;
  hex = String(hex).trim();
  if (hex[0] !== "#") return null;
  if (hex.length === 7) return {
    r: parseInt(hex.slice(1,3),16),
    g: parseInt(hex.slice(3,5),16),
    b: parseInt(hex.slice(5,7),16),
  };
  return null;
}

function isYellowTolerant_(hex, targetHex, tolerance) {
  const rgb = hexToRgb_(hex);
  const tgt = hexToRgb_(targetHex);
  if (rgb && tgt) {
    const dist = Math.abs(rgb.r - tgt.r) + Math.abs(rgb.g - tgt.g) + Math.abs(rgb.b - tgt.b);
    if (dist <= tolerance) return true;
  }
  if (rgb) return rgb.r >= 200 && rgb.g >= 170 && rgb.b <= 170;
  return false;
}

function toIntSafe_(v) {
  if (v === null || v === undefined) return 0;
  if (typeof v === "number") return isFinite(v) ? Math.trunc(v) : 0;
  const s = String(v).trim();
  if (!s) return 0;
  const n = parseInt(s, 10);
  return isNaN(n) ? 0 : n;
}

function clamp0to4_(n) {
  n = toIntSafe_(n);
  if (n < 0) return 0;
  if (n > 4) return 4;
  return n;
}

function colToLetter_(column) {
  let temp = "", letter = "";
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function setDropdown_(sh, col, options) {
  if (!col) return;
  const lr = sh.getLastRow();
  if (lr < 2) return;
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(options, true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(2, col, lr - 1, 1).setDataValidation(rule);
}
