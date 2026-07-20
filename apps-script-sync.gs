/**
 * Tracker sheet — Apps Script (recruiting endpoint + WG KPI sync)
 *
 * TWO things live in this one script:
 *  1. doGet() — Web App that returns the recruiting-command-center payload as JSON.
 *  2. pushWgKpiToGitHub() — scheduled push of the WG tab to wg-kpi-dashboard/wg-data.json.
 *
 * Paste the WHOLE file into the tracker's bound Apps Script
 * (Extensions → Apps Script), save, then:
 *
 *   A) Recruiting endpoint (already deployed — do this only on FIRST setup or if broken):
 *      Deploy → New deployment → Web app → Execute as: Me, Access: Anyone → Deploy.
 *
 *   B) WG KPI sync (this is the piece that just broke):
 *      1. Fill GITHUB_TOKEN below (ghp_… with `repo` scope).
 *      2. Run pushWgKpiToGitHub once — authorize when prompted.
 *      3. Run installWgKpiTrigger once — hourly time-driven trigger.
 *      4. Project Settings ⚙ → "Notify me: Immediately" for failure notifications.
 *
 * The full source is versioned at Claude Server/wg-kpi-dashboard/apps-script-sync.gs.
 * Do not delete it — recovery from a lost Google project is otherwise from scratch.
 */

// ─── SHARED CONFIG (used by both features) ────────────────────────────
var SHEET_ID = '1pKvIJlav4FvYQa4HqOusb0yhfinI-1XaQbESZ-vjMfI';
var WG_GID = 1818912442;
var CORE_GID = 1378398655;

// ─── WG KPI SYNC CONFIG ───────────────────────────────────────────────
var GITHUB_TOKEN = 'PASTE_YOUR_PAT_HERE';         // ghp_… with `repo` scope
var REPO         = 'wg-dotcom/wg-kpi-dashboard';
var FILE_PATH    = 'wg-data.json';
var BRANCH       = 'main';
// ──────────────────────────────────────────────────────────────────────

var PLACED = 'Placed / Hired', CLOSED = 'Closed / Lost';
var STAGE_ORDER = ['KickOff / New','Candidates Presented','Interviewing','Offer / Invoice','Placed / Hired','Closed / Lost'];

/* ═══════════════════════════════════════════════════════════════════════
   1) RECRUITING COMMAND CENTER — live JSON endpoint (unchanged from before)
   ═══════════════════════════════════════════════════════════════════════ */

function doGet(e) {
  var payload = buildPayload();
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheetByGid(ss, gid) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) if (sheets[i].getSheetId() === gid) return sheets[i];
  return null;
}

// getDisplayValues() returns cells exactly as shown ("April 6, 2026", "1,800", "18%") — matches CSV.
function rowsOf(ss, gid) {
  var sh = sheetByGid(ss, gid);
  if (!sh) return [];
  return sh.getDataRange().getDisplayValues();
}

function stageFromTracker(status) {
  var n = parseInt(status, 10);
  if (n === 1) return 'KickOff / New';
  if (n === 2) return 'Candidates Presented';
  if (n === 3 || n === 4 || n === 5) return 'Interviewing';
  if (n === 6 || n === 7) return 'Offer / Invoice';
  if (n === 8) return PLACED;
  if (n === 9) return CLOSED;
  return 'KickOff / New';
}
function money(s) { if (s == null || s === '') return null; var n = parseFloat(String(s).replace(/[^0-9.\-]/g, '')); return isFinite(n) ? n : null; }
function hrDigits(s) { var m = String(s == null ? '' : s).match(/(\d{3,})/); return m ? m[1] : null; }
function daysSince(s, now) { var t = Date.parse(s); return isFinite(t) ? Math.round((now.getTime() - t) / 864e5) : null; }

function parseTab(grid, line, now) {
  if (!grid.length) return [];
  var hdr = grid[0].map(function (h) { return String(h).trim().toLowerCase(); });
  function col() {
    var names = Array.prototype.slice.call(arguments);
    for (var i = 0; i < hdr.length; i++) for (var j = 0; j < names.length; j++)
      if (hdr[i] === names[j] || hdr[i].indexOf(names[j]) >= 0) return i;
    return -1;
  }
  var C = {
    hr: col('hiring request', 'hr #', 'hr'), status: col('status'), source: col('source'),
    advisor: col('advisor', 'responsible'), name: col('name'), company: col('company name', 'company'),
    role: col('role', 'position'), deal: col('deal value'), budget: col('budget'), wgpct: col('%wg'),
    repl: col('is replacement'), start: col('creation date', 'assigned date', 'date'),
    kickoff: col('kickoff', 'alignment'), placed: col('date placed'),
    invoiceDate: col('invoice date'), paymentDate: col('payment date'), totalPaid: col('total paid'),
    dfp: col('days for placement'), notes: col('notes')
  };
  var out = [];
  for (var i = 1; i < grid.length; i++) {
    var row = grid[i];
    if (!row || row.join('').trim() === '') continue;
    var g = function (idx) { return (idx >= 0 && idx < row.length) ? String(row[idx]).trim() : ''; };
    var status = g(C.status);
    if (!status && !g(C.hr) && !g(C.company)) continue;
    var stage = stageFromTracker(status);
    var startDate = g(C.start) || g(C.kickoff);
    var terminal = (stage === PLACED || stage === CLOSED);
    var daysOpen = terminal ? null : daysSince(startDate, now);
    var flags = [];
    if (/true|yes/i.test(g(C.repl))) flags.push('Replacement');
    if (stage === CLOSED) flags.push('Closed/Paused');
    if (daysOpen != null && daysOpen > 90) flags.push('Aging');
    out.push({
      id: 'TR-' + line.replace(/\s/g, '') + '-' + i, source: 'tracker', serviceLine: line,
      hr: hrDigits(g(C.hr)), rawStatus: status, stage: stage,
      role: g(C.role), company: g(C.company), advisor: g(C.advisor), candidate: g(C.name),
      daysOpen: daysOpen, dealValue: money(g(C.deal)), budget: money(g(C.budget)), wgPct: g(C.wgpct) || null,
      flags: flags, notes: g(C.notes) || '', startDate: startDate || null, placedDate: g(C.placed) || null,
      closeReason: stage === CLOSED ? 'Closed / Paused / M.I.A.' : null,
      bucket: stage === PLACED ? 'placed' : (stage === CLOSED ? 'cancelled' : 'active'),
      link: 'https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit#gid=' + (line === 'White Glove' ? WG_GID : CORE_GID),
      invoiceDate: g(C.invoiceDate) || null, paymentDate: g(C.paymentDate) || null, totalPaid: money(g(C.totalPaid)),
      daysForPlacement: money(g(C.dfp))
    });
  }
  return out;
}

function buildPayload() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var tz = ss.getSpreadsheetTimeZone();
  var now = new Date();
  var board = parseTab(rowsOf(ss, WG_GID), 'White Glove', now).concat(parseTab(rowsOf(ss, CORE_GID), 'Core', now));
  return {
    generatedAt: now.toISOString(),
    snapshotMonth: Utilities.formatDate(now, tz, 'yyyy-MM'),
    source: 'Live from tracker (Apps Script web app)',
    count: 0, rows: [], board: board, stageOrder: STAGE_ORDER,
    trackerCount: board.length, revenue: {}
  };
}

/* ═══════════════════════════════════════════════════════════════════════
   2) WG KPI DASHBOARD — scheduled snapshot pushed to wg-kpi-dashboard repo
   ═══════════════════════════════════════════════════════════════════════ */

// Reads the WG tab and returns rows keyed by the sheet's own column headers,
// preserving displayValues ("$5,400", "30%", "August 22, 2025") so the
// dashboard doesn't need to re-format anything.
function readWgKpiRows() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = sheetByGid(ss, WG_GID);
  if (!sh) throw new Error('WG tab (gid=' + WG_GID + ') not found. Someone renamed or deleted it.');
  var values = sh.getDataRange().getDisplayValues();
  if (values.length < 2) return [];
  var headers = values[0].map(function (h) { return String(h).trim(); });
  var rows = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (row.every(function (c) { return String(c).trim() === ''; })) continue;
    var obj = {};
    for (var c = 0; c < headers.length; c++) {
      if (!headers[c]) continue;
      obj[headers[c]] = String(row[c] == null ? '' : row[c]).trim();
    }
    rows.push(obj);
  }
  return rows;
}

function buildWgKpiPayload() {
  return {
    updated: new Date().toISOString(),
    rows: readWgKpiRows(),
  };
}

function pushWgKpiToGitHub() {
  if (!GITHUB_TOKEN || GITHUB_TOKEN.indexOf('ghp_') !== 0) {
    throw new Error('Set GITHUB_TOKEN at the top of this script (needs "repo" scope).');
  }
  var payload = buildWgKpiPayload();
  var body = JSON.stringify(payload, null, 2);

  var api = 'https://api.github.com/repos/' + REPO + '/contents/' + FILE_PATH;
  var authHeaders = { Authorization: 'token ' + GITHUB_TOKEN, Accept: 'application/vnd.github+json' };

  // Get current SHA (required for updates; a 404 means file doesn't exist yet)
  var sha = null;
  var getResp = UrlFetchApp.fetch(api + '?ref=' + BRANCH, {
    method: 'get', headers: authHeaders, muteHttpExceptions: true,
  });
  if (getResp.getResponseCode() === 200) {
    sha = JSON.parse(getResp.getContentText()).sha;
  } else if (getResp.getResponseCode() !== 404) {
    throw new Error('GET failed: ' + getResp.getResponseCode() + ' ' + getResp.getContentText().slice(0, 200));
  }

  var timestamp = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'M/d/yyyy, h:mm:ss a');
  var commitBody = {
    message: 'Auto-sync White Glove: ' + timestamp,
    content: Utilities.base64Encode(body, Utilities.Charset.UTF_8),
    branch: BRANCH,
  };
  if (sha) commitBody.sha = sha;

  var putResp = UrlFetchApp.fetch(api, {
    method: 'put', headers: authHeaders,
    contentType: 'application/json', payload: JSON.stringify(commitBody),
    muteHttpExceptions: true,
  });
  var code = putResp.getResponseCode();
  if (code !== 200 && code !== 201) {
    throw new Error('PUT failed: ' + code + ' ' + putResp.getContentText().slice(0, 300));
  }
  console.log('Synced ' + payload.rows.length + ' WG rows at ' + timestamp);
  return payload.rows.length;
}

function installWgKpiTrigger() {
  // Clear any duplicate triggers for this handler
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'pushWgKpiToGitHub') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('pushWgKpiToGitHub').timeBased().everyHours(1).create();
  console.log('Hourly WG KPI trigger installed. Next run within the hour.');
}

function removeWgKpiTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'pushWgKpiToGitHub') ScriptApp.deleteTrigger(t);
  });
  console.log('WG KPI trigger removed.');
}
