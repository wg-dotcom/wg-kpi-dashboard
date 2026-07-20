/**
 * WG KPI Dashboard — Google Sheet → GitHub sync
 *
 * Reads the "White Glove" tab of the tracker sheet and pushes wg-data.json
 * to the wg-kpi-dashboard repo so the live dashboard stays fresh.
 *
 * SETUP (one time):
 *  1. Go to https://script.google.com/home → New project.
 *  2. Delete the default code, paste ALL of this, Save.
 *  3. Fill GITHUB_TOKEN below (ghp_… with "repo" scope). Save again.
 *  4. Run pushToGitHub() once — Apps Script will ask you to authorize
 *     access to Sheets + external URL fetch. Approve.
 *  5. Run installTrigger() once — installs the hourly time-driven trigger.
 *  6. In the project sidebar, hit ⚙ Project Settings → set "Notify me
 *     immediately" for failure-notification. This is what saves us next
 *     time Google silently disables the trigger.
 *
 * If it stops working: run pushToGitHub() manually. If it succeeds,
 * run installTrigger() again to reinstate the schedule. If it errors on
 * auth, revoke the app in your Google Account settings and re-authorize
 * on next run.
 *
 * The full source is also in Claude Server/wg-kpi-dashboard/apps-script-sync.gs
 * — do not delete it.
 */

// ─── CONFIG ──────────────────────────────────────────────────────────
const GITHUB_TOKEN = 'PASTE_YOUR_PAT_HERE';                 // ghp_… (repo scope)
const REPO         = 'wg-dotcom/wg-kpi-dashboard';
const FILE_PATH    = 'wg-data.json';
const BRANCH       = 'main';

const TRACKER_ID   = '1pKvIJlav4FvYQa4HqOusb0yhfinI-1XaQbESZ-vjMfI'; // same tracker as recruiting-dashboard
const WG_GID       = 1818912442;                                     // "White Glove" tab
// ─────────────────────────────────────────────────────────────────────

function sheetByGid(ss, gid) {
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === gid) return sheets[i];
  }
  throw new Error('WG tab (gid=' + gid + ') not found. Someone renamed or deleted it.');
}

function readRows() {
  const ss = SpreadsheetApp.openById(TRACKER_ID);
  const sh = sheetByGid(ss, WG_GID);
  // getDisplayValues() returns cells exactly as shown ("$5,400", "30%", "August 22, 2025")
  // so the dashboard doesn't need to re-format anything.
  const values = sh.getDataRange().getDisplayValues();
  if (values.length < 2) return [];
  const headers = values[0].map(h => String(h).trim());
  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    // Skip fully-empty rows (spreadsheets often have trailing blanks)
    if (row.every(c => String(c).trim() === '')) continue;
    const obj = {};
    for (let c = 0; c < headers.length; c++) {
      if (!headers[c]) continue;                 // skip unnamed columns
      obj[headers[c]] = String(row[c] == null ? '' : row[c]).trim();
    }
    rows.push(obj);
  }
  return rows;
}

function buildPayload() {
  return {
    updated: new Date().toISOString(),
    rows: readRows(),
  };
}

function pushToGitHub() {
  if (!GITHUB_TOKEN || GITHUB_TOKEN.indexOf('ghp_') !== 0) {
    throw new Error('Set GITHUB_TOKEN at the top of this script (needs "repo" scope).');
  }
  const payload = buildPayload();
  const body = JSON.stringify(payload, null, 2);

  const api = 'https://api.github.com/repos/' + REPO + '/contents/' + FILE_PATH;

  // Get current SHA (required for updates; omitted for first commit)
  let sha = null;
  const getResp = UrlFetchApp.fetch(api + '?ref=' + BRANCH, {
    method: 'get',
    headers: { Authorization: 'token ' + GITHUB_TOKEN, Accept: 'application/vnd.github+json' },
    muteHttpExceptions: true,
  });
  if (getResp.getResponseCode() === 200) {
    sha = JSON.parse(getResp.getContentText()).sha;
  } else if (getResp.getResponseCode() !== 404) {
    throw new Error('GET failed: ' + getResp.getResponseCode() + ' ' + getResp.getContentText().slice(0, 200));
  }

  const timestamp = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'M/d/yyyy, h:mm:ss a');
  const commitBody = {
    message: 'Auto-sync White Glove: ' + timestamp,
    content: Utilities.base64Encode(body, Utilities.Charset.UTF_8),
    branch: BRANCH,
  };
  if (sha) commitBody.sha = sha;

  const putResp = UrlFetchApp.fetch(api, {
    method: 'put',
    headers: { Authorization: 'token ' + GITHUB_TOKEN, Accept: 'application/vnd.github+json' },
    contentType: 'application/json',
    payload: JSON.stringify(commitBody),
    muteHttpExceptions: true,
  });
  const code = putResp.getResponseCode();
  if (code !== 200 && code !== 201) {
    throw new Error('PUT failed: ' + code + ' ' + putResp.getContentText().slice(0, 300));
  }
  console.log('Synced ' + payload.rows.length + ' rows at ' + timestamp);
  return payload.rows.length;
}

function installTrigger() {
  // Clear any orphaned triggers first
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'pushToGitHub') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('pushToGitHub').timeBased().everyHours(1).create();
  console.log('Hourly trigger installed. Next run within the hour.');
}

function removeAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  console.log('All triggers removed.');
}
