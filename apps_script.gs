/**
 * Cricket Schedule – Google Apps Script Backend
 * ================================================
 * Deploy as: Extensions → Apps Script → Deploy → New deployment
 *   Type: Web app  |  Execute as: Me  |  Who has access: Anyone
 * Copy the deployment URL → paste into the PWA Settings modal.
 *
 * Sheet tabs required (auto-created by scraper setup.py):
 *   Leagues   – league registry
 *   Played    – NEW: per-profile played/not-played records
 */

const SS_ID      = '1YN5cqnw4LbI2NQEDOTxZOdWCF_A1Jid8ZENempkoo_0'; // ← your spreadsheet ID
const TAB_LEAGUES = 'Leagues';
const TAB_PLAYED  = 'Played';
const TAB_REQUESTS = 'LeagueRequests';

// ── CORS helper ─────────────────────────────────────────────────────────────
function cors(output) {
  return ContentService
    .createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Router ───────────────────────────────────────────────────────────────────
function doGet(e) {
  const action  = (e.parameter.action  || '').trim();
  const profile = (e.parameter.profile || 'default').trim().toLowerCase();

  if (action === 'getPlayed') return cors(getPlayed(profile));
  return cors({ error: 'Unknown GET action: ' + action });
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch(_) {}
  const action  = (body.action  || '').trim();
  const profile = (body.profile || 'default').trim().toLowerCase();

  if (action === 'syncPlayed')     return cors(syncPlayed(profile, body.played  || {}));
  if (action === 'requestLeague')  return cors(requestLeague(profile, body.league || {}));
  return cors({ error: 'Unknown POST action: ' + action });
}

// ── Played sync ──────────────────────────────────────────────────────────────
// Row format: profile | game_key | status (played/unplayed) | updated_at
function ensurePlayedTab(ss) {
  let ws;
  try { ws = ss.getSheetByName(TAB_PLAYED); } catch(_) {}
  if (!ws) {
    ws = ss.insertSheet(TAB_PLAYED);
    ws.appendRow(['Profile', 'GameKey', 'Status', 'UpdatedAt']);
    ws.setFrozenRows(1);
  }
  return ws;
}

function getPlayed(profile) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const ws = ensurePlayedTab(ss);
  const data = ws.getDataRange().getValues();
  if (data.length < 2) return { played: {} };
  const played = {};
  for (let i = 1; i < data.length; i++) {
    const [prof, key, status] = data[i];
    if (String(prof).toLowerCase() === profile) {
      played[key] = status === 'played';
    }
  }
  return { played };
}

function syncPlayed(profile, incoming) {
  // incoming = { gameKey: true/false, ... }
  const ss = SpreadsheetApp.openById(SS_ID);
  const ws = ensurePlayedTab(ss);
  const data = ws.getDataRange().getValues();
  const now  = new Date().toISOString();

  // Build index of existing rows for this profile: key → rowIndex (1-based)
  const rowIdx = {};
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === profile) {
      rowIdx[data[i][1]] = i + 1; // sheet rows are 1-based
    }
  }

  Object.entries(incoming).forEach(([key, isPlayed]) => {
    const status = isPlayed ? 'played' : 'unplayed';
    if (rowIdx[key]) {
      // Update existing row
      ws.getRange(rowIdx[key], 3, 1, 2).setValues([[status, now]]);
    } else {
      // Append new row
      ws.appendRow([profile, key, status, now]);
    }
  });

  return { ok: true, synced: Object.keys(incoming).length };
}

// ── League request ────────────────────────────────────────────────────────────
// Writes to a LeagueRequests tab; admin reviews and manually moves to Leagues
function ensureRequestsTab(ss) {
  let ws;
  try { ws = ss.getSheetByName(TAB_REQUESTS); } catch(_) {}
  if (!ws) {
    ws = ss.insertSheet(TAB_REQUESTS);
    ws.appendRow(['Submitted At', 'Profile', 'League Name', 'Sport', 'Source URL', 'Source URL Past', 'Notes', 'Status']);
    ws.setFrozenRows(1);
  }
  return ws;
}

function requestLeague(profile, league) {
  const name = (league.name || '').trim();
  if (!name) return { error: 'League name is required' };

  const ss = SpreadsheetApp.openById(SS_ID);
  const ws = ensureRequestsTab(ss);
  ws.appendRow([
    new Date().toISOString(),
    profile,
    name,
    (league.sport || 'Cricket').trim(),
    (league.sourceUrl || '').trim(),
    (league.sourceUrlPast || '').trim(),
    (league.notes || '').trim(),
    'Requested'
  ]);
  return { ok: true, message: 'League request submitted. Admin will review and approve.' };
}
