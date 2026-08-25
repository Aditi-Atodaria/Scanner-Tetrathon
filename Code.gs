/**
 * QR Entry Scan — event guest entry tracking system
 *
 * Sheets used (auto-created by setupSheets()):
 *   Guests  — one row per guest. Each guest's own QR code is emailed
 *             straight to their Email column so they can show it (on their
 *             phone or printed) to be scanned once at entry.
 *   Config  — optional entry window (start/end) + role→prefix map (edit
 *             here, no code changes needed)
 *   ScanLog — full audit trail of every scan attempt (success/duplicate/invalid)
 *
 * Deploy as Web App (Extensions > Deploy > New deployment > Web app,
 * execute as "Me", access "Anyone with the link"). Copy the /exec URL into
 * Scanner.html (the external camera page) after deploying.
 *
 * NOTE for sheets set up BEFORE this version: setupSheets() only writes
 * placeholder data the first time a sheet is created, so re-running it on an
 * existing spreadsheet will NOT retroactively add the "Guest" role/prefix
 * row to your Config sheet. If your Config sheet's Role/Prefix table doesn't
 * already have a "Guest" row, add one manually (Role: Guest, Prefix: G) —
 * otherwise new guests added without a role will fall back to a generic "X"
 * prefix, which still works fine but is less tidy.
 */

const GUESTS_SHEET = 'Guests';
const CONFIG_SHEET = 'Config';
const LOG_SHEET = 'ScanLog';

// Access code the scanner's login screen checks against. Change this before
// your event, then redeploy — anyone with this code (not just the link) can
// open the scanner.
const AUTH_CODE = 'ENTRY2026';

// Fixed columns in the Guests sheet. "Entry Time" is set the moment a guest
// is scanned in — empty means not yet arrived. This is a single scan-once
// check-in system: no per-slot columns.
const GUEST_FIXED_COLS = ['GuestID', 'Name', 'Email', 'Role', 'QR Code', 'Entry Time'];
const ENTRY_TIME_COL_INDEX = GUEST_FIXED_COLS.length; // 1-based column index of "Entry Time"

// Role used for guests added via the simplified "Name, Email" bulk/CSV/single
// intake, which no longer asks for a role. Still tracked internally (ID
// prefix, dashboard column) but no longer required from the person entering
// guests.
const DEFAULT_GUEST_ROLE = 'Guest';

const DEFAULT_ROLE_PREFIXES = [
  { role: 'Guest',        prefix: 'G' },
  { role: 'Participant',  prefix: 'P' },
  { role: 'Volunteer',    prefix: 'V' },
  { role: 'Faculty',      prefix: 'F' },
  { role: 'Club Member',  prefix: 'C' }
];

// ---------- One-time setup ----------

// Run this once from the Apps Script editor (select setupSheets, click Run)
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let guests = ss.getSheetByName(GUESTS_SHEET);
  if (!guests) guests = ss.insertSheet(GUESTS_SHEET);
  guests.getRange(1, 1, 1, GUEST_FIXED_COLS.length).setValues([GUEST_FIXED_COLS]);
  guests.setFrozenRows(1);

  let config = ss.getSheetByName(CONFIG_SHEET);
  if (!config) config = ss.insertSheet(CONFIG_SHEET);
  config.clear();
  config.getRange(1, 1, 1, 2).setValues([['EntryStart', 'EntryEnd']]);
  // Placeholder entry window, anchored to "now" — EDIT THESE in the Config
  // sheet to your real event start/end before going live. Leave BOTH blank
  // to allow scanning at any time (no window restriction).
  const start = new Date();
  start.setMinutes(0, 0, 0);
  const end = new Date(start.getTime() + 12 * 3600000);
  config.getRange(2, 1, 1, 2).setValues([[start, end]]);
  config.getRange(2, 1, 1, 2).setNumberFormat('yyyy-mm-dd hh:mm');

  const roleStart = 4;
  config.getRange(roleStart, 1, 1, 2).setValues([['Role', 'Prefix']]);
  config.getRange(roleStart + 1, 1, DEFAULT_ROLE_PREFIXES.length, 2).setValues(
    DEFAULT_ROLE_PREFIXES.map(r => [r.role, r.prefix])
  );
  config.setFrozenRows(1);

  let log = ss.getSheetByName(LOG_SHEET);
  if (!log) log = ss.insertSheet(LOG_SHEET);
  log.getRange(1, 1, 1, 6).setValues([[
    'Timestamp', 'GuestID', 'Name', 'Role', 'Station Note', 'Result'
  ]]);
  log.setFrozenRows(1);

  SpreadsheetApp.getUi().alert(
    'Setup complete: Guests, Config, and ScanLog sheets are ready.\n\n' +
    'IMPORTANT: The Config sheet has a placeholder entry window (start/end), ' +
    'anchored to right now. Edit those dates/times to match your real event, ' +
    'or clear BOTH cells to allow scanning at any time.'
  );
}

// ---------- Web app entry points ----------

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || '';
  const action = (e && e.parameter && e.parameter.action) || '';

  // JSON API for the externally-hosted Scanner.html (camera page can't run
  // inside the Apps Script iframe — see README).
  if (action === 'state') {
    return jsonResponse_(Object.assign({ status: 'ok' }, getEventState_()));
  }
  // Manual fallback lookup: staff types a guest's name when the QR code
  // won't scan (damaged phone screen, printout issue, etc.) and picks the
  // right person from matches instead of typing a Guest ID blind.
  if (action === 'search') {
    const q = (e.parameter.q || '').toString().trim();
    if (!q) return jsonResponse_({ status: 'ok', guests: [] });
    return jsonResponse_({ status: 'ok', guests: searchGuestsByName_(q) });
  }

  if (page === 'generator') {
    const t = HtmlService.createTemplateFromFile('Generator');
    t.scriptUrl = ScriptApp.getService().getUrl();
    return t.evaluate()
      .setTitle('QR Entry Scan — Guest & QR Generator')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === 'dashboard') {
    const t = HtmlService.createTemplateFromFile('Dashboard');
    t.scriptUrl = ScriptApp.getService().getUrl();
    return t.evaluate()
      .setTitle('QR Entry Scan — Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  const homeTemplate = HtmlService.createTemplateFromFile('Home');
  homeTemplate.scriptUrl = ScriptApp.getService().getUrl();
  return homeTemplate.evaluate()
    .setTitle('QR Entry Scan')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Receives a scan from the external Scanner.html page.
// Body: { guestId: "P-A3F9", station: "Gate A" (optional) }
// This is a single scan-once check-in: the first successful scan marks the
// guest as arrived; any further scan of the same QR code comes back as
// "duplicate" rather than being logged again.
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // avoid two simultaneous scans double-crediting entry

    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('No data received');
    }
    const data = JSON.parse(e.postData.contents);

    // Scanner login check — handled first so it never falls through to the
    // guestId-required scan logic below.
    if (data.action === 'login') {
      if (data.code === AUTH_CODE) {
        return jsonResponse_({ status: 'success' });
      }
      return jsonResponse_({ status: 'invalid', message: 'Incorrect access code.' });
    }

    const guestId = (data.guestId || '').toString().trim().toUpperCase();
    const station = (data.station || '').toString().trim();

    if (!guestId) throw new Error('guestId missing');

    const eventState = getEventState_();
    if (eventState.state !== 'active') {
      logScan_(guestId, '', '', station, 'Scan outside window: ' + eventState.message);
      return jsonResponse_({ status: 'closed', message: eventState.message });
    }

    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GUESTS_SHEET);
    const found = findGuestRow_(sh, guestId);
    if (!found) {
      logScan_(guestId, '', '', station, 'Unknown guest ID');
      return jsonResponse_({ status: 'invalid', message: 'QR not recognized: ' + guestId });
    }

    const { rowIndex, name, role, entryTime } = found;

    if (entryTime) {
      logScan_(guestId, name, role, station, 'Duplicate');
      return jsonResponse_({
        status: 'duplicate',
        message: name + ' already checked in at ' + formatDateTime_(entryTime),
        name, role,
        entryTime: entryTime.toISOString()
      });
    }

    const ts = new Date();
    sh.getRange(rowIndex, ENTRY_TIME_COL_INDEX).setValue(ts);

    logScan_(guestId, name, role, station, 'Success');

    return jsonResponse_({
      status: 'success',
      name, role,
      timestamp: ts.toISOString()
    });
  } catch (err) {
    return jsonResponse_({ status: 'error', message: err.message });
  } finally {
    lock.releaseLock();
  }
}

// ---------- Called from Generator.html ----------

// guestList: [{name, email, role}, ...] — role is optional and defaults to
// DEFAULT_GUEST_ROLE, since the standard intake format is now just
// "Name, Email" (no role column). Each guest needs a valid email since their
// QR code is emailed directly to them for scanning at entry.
// Guests with a name that already exists (case-insensitive, trimmed) are
// skipped rather than added again — both against existing sheet rows and
// against earlier entries in the same batch.
function bulkGenerateGuests(guestList) {
  if (!guestList || !guestList.length) return { created: [], skipped: [] };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(GUESTS_SHEET);
  const rolePrefixes = getRolePrefixMap_();

  const existingIds = new Set(
    sh.getLastRow() > 1
      ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat().filter(String)
      : []
  );
  const existingNames = new Set(
    sh.getLastRow() > 1
      ? sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues().flat()
          .filter(String).map(n => n.toString().trim().toLowerCase())
      : []
  );
  // Maps GuestID -> "name|role" for every current row, used only to tell a
  // genuine hash collision (two different names landing on the same ID)
  // apart from a returning guest whose ID is SUPPOSED to match a past one.
  const idToNameRole = new Map();
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(r => {
      if (r[0]) idToNameRole.set(r[0], r[1].toString().trim().toLowerCase() + '|' + r[3].toString().trim().toLowerCase());
    });
  }

  const rows = [];
  const created = [];
  const skipped = [];
  guestList.forEach(g => {
    const name = (g.name || '').toString().trim();
    const email = (g.email || '').toString().trim();
    const role = (g.role || '').toString().trim() || DEFAULT_GUEST_ROLE;
    if (!name) return;

    if (!email || !isValidEmail_(email)) {
      skipped.push({ name, role, reason: 'Missing or invalid email address for "' + name + '".' });
      return;
    }

    const nameKey = name.toLowerCase();
    if (existingNames.has(nameKey)) {
      skipped.push({ name, role, reason: 'A guest named "' + name + '" already exists.' });
      return;
    }
    existingNames.add(nameKey); // also catches duplicates within this same pasted batch

    const prefix = rolePrefixes[role] || 'X';
    const guestId = generateDeterministicId_(name, role, prefix, existingIds, idToNameRole);
    existingIds.add(guestId);
    idToNameRole.set(guestId, nameKey + '|' + role.toLowerCase());

    const qrFormula = '=IMAGE("https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=' + guestId + '")';
    const row = [guestId, name, email, role, qrFormula, '']; // '' = not yet checked in
    rows.push(row);
    created.push({ guestId, name, email, role });
  });

  if (rows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
  return { created, skipped };
}

function addSingleGuest(name, email, role) {
  name = (name || '').toString().trim();
  const result = bulkGenerateGuests([{ name, email, role }]);
  if (result.skipped.length) {
    throw new Error(result.skipped[0].reason);
  }
  return result.created[0] || null;
}

// Returns every guest currently in the sheet, for reprinting badges later
// (not just ones generated in the current browser session).
function getAllGuestsForPrint() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GUESTS_SHEET);
  if (sh.getLastRow() < 2) return [];
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues(); // GuestID, Name, Email, Role
  return values
    .filter(r => r[0])
    .map(r => ({ guestId: r[0], name: r[1], email: r[2], role: r[3] }));
}

// Edits a guest's name/email/role in place. The Guest ID (and printed QR)
// never changes — only the details attached to it — so an already-printed
// badge or already-sent email keeps working after an edit. If the email is
// changed, the guest's QR code is NOT auto-resent to the new address — use
// the "Resend Email" action for that guest to send it there.
function updateGuest(guestId, name, email, role) {
  name = (name || '').toString().trim();
  email = (email || '').toString().trim();
  role = (role || '').toString().trim() || DEFAULT_GUEST_ROLE;
  if (!guestId) throw new Error('guestId missing');
  if (!name) throw new Error('Name cannot be empty');
  if (!email || !isValidEmail_(email)) throw new Error('Enter a valid email address.');

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GUESTS_SHEET);
  const rowIndex = findGuestRowIndex_(sh, guestId);
  if (!rowIndex) throw new Error('Guest not found: ' + guestId);

  const nameKey = name.toLowerCase();
  const allIds = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues(); // GuestID, Name
  const collision = allIds.some(r => r[0] !== guestId && r[1] && r[1].toString().trim().toLowerCase() === nameKey);
  if (collision) throw new Error('A guest named "' + name + '" already exists.');

  sh.getRange(rowIndex, 2, 1, 3).setValues([[name, email, role]]); // Name, Email, Role columns
  return { guestId, name, email, role };
}

// Permanently removes a guest's row (their entry record for the event goes
// with it). ScanLog audit history is left untouched.
function deleteGuest(guestId) {
  if (!guestId) throw new Error('guestId missing');
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GUESTS_SHEET);
  const rowIndex = findGuestRowIndex_(sh, guestId);
  if (!rowIndex) throw new Error('Guest not found: ' + guestId);

  const name = sh.getRange(rowIndex, 2).getValue();
  const role = sh.getRange(rowIndex, 4).getValue();
  sh.deleteRow(rowIndex);
  logScan_(guestId, name, role, '', 'Guest deleted');
  return { guestId, deleted: true };
}

function findGuestRowIndex_(sh, guestId) {
  if (sh.getLastRow() < 2) return null;
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === guestId) return i + 2;
  }
  return null;
}

// Removes multiple guests at once (checkbox selection from the Generator page).
// Rewrites the remaining rows in one pass rather than deleting rows one at a
// time, since deleting rows individually shifts every later row index.
function deleteGuests(guestIds) {
  if (!guestIds || !guestIds.length) return { deletedCount: 0 };
  const idsToRemove = new Set(guestIds);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GUESTS_SHEET);
  if (sh.getLastRow() < 2) return { deletedCount: 0 };

  const width = GUEST_FIXED_COLS.length;
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues();

  const removed = values.filter(r => idsToRemove.has(r[0]));
  const remaining = values.filter(r => !idsToRemove.has(r[0]));

  sh.getRange(2, 1, values.length, width).clearContent();
  if (remaining.length) {
    sh.getRange(2, 1, remaining.length, width).setValues(remaining);
  }

  removed.forEach(r => logScan_(r[0], r[1], r[3], '', 'Guest deleted (bulk)'));
  return { deletedCount: removed.length };
}

// Wipes every guest from the sheet. Entry records go with them.
// ScanLog audit trail is left untouched, with one summary entry recorded.
function deleteAllGuests() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GUESTS_SHEET);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { deletedCount: 0 };

  const width = GUEST_FIXED_COLS.length;
  const count = lastRow - 1;
  sh.getRange(2, 1, count, width).clearContent();
  logScan_('ALL', '', '', '', 'All guests deleted (' + count + ')');
  return { deletedCount: count };
}

// ---------- Called from Dashboard.html ----------

function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(GUESTS_SHEET);

  let guests = [];
  let checkedIn = 0;
  if (sh.getLastRow() > 1) {
    const values = sh.getRange(2, 1, sh.getLastRow() - 1, GUEST_FIXED_COLS.length).getValues();
    guests = values.map(r => {
      const entryTime = r[ENTRY_TIME_COL_INDEX - 1];
      const entered = !!entryTime;
      if (entered) checkedIn++;
      return {
        guestId: r[0], name: r[1], email: r[2], role: r[3],
        entered,
        entryTime: entered ? entryTime : null
      };
    });
  }
  return { summary: { checkedIn, totalGuests: guests.length }, guests, totalGuests: guests.length };
}

function getRecentScans(limit) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  const n = limit || 15;
  const last = sh.getLastRow();
  const start = Math.max(2, last - n + 1);
  return sh.getRange(start, 1, last - start + 1, 6).getValues().reverse();
}

// ---------- Internal helpers ----------

// Determines whether entry scanning is open right now, from the Config
// sheet's EntryStart/EntryEnd cells. If either is blank, scanning is always
// allowed (no window restriction) — this lets the scanner "just work" for
// events that don't need a strict entry window.
function getEventState_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
  const row = sh.getRange(2, 1, 1, 2).getValues()[0];
  const start = row[0] instanceof Date ? row[0] : (row[0] ? new Date(row[0]) : null);
  const end = row[1] instanceof Date ? row[1] : (row[1] ? new Date(row[1]) : null);

  if (!start || !end) {
    return { state: 'active' };
  }

  const now = new Date();
  if (now < start) {
    return { state: 'closed', message: 'Not open yet. Entry opens at ' + formatDateTime_(start) };
  }
  if (now > end) {
    return { state: 'closed', message: 'Entry window has closed (ended ' + formatDateTime_(end) + ').' };
  }
  return { state: 'active' };
}

function formatDateTime_(date) {
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  return Utilities.formatDate(date, tz, 'MMM d, h:mm a');
}

function getRolePrefixMap_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
  const values = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  const startIdx = values.findIndex(r => r[0] === 'Role' && r[1] === 'Prefix');
  const map = {};
  if (startIdx === -1) return map;
  for (let i = startIdx + 1; i < values.length; i++) {
    const [role, prefix] = values[i];
    if (!role) break;
    map[role.toString().trim()] = prefix.toString().trim();
  }
  return map;
}

// ---------- Emailing QR codes directly to guests ----------

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test((email || '').toString().trim());
}

// Builds and sends one guest's QR code straight to their own inbox — shown
// inline in the email body (cid: reference) and also attached as a PNG so
// they can save/print it. This is the badge they show once to be scanned
// at entry.
function sendGuestQrEmail_(guestId, name, email) {
  const qrUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=400x400&data=' + encodeURIComponent(guestId);
  const qrBlob = UrlFetchApp.fetch(qrUrl).getBlob().setName(guestId + '_QR.png');

  const htmlBody =
    '<p>Hi ' + name + ',</p>' +
    '<p>Here is your personal QR code for entry check-in at the event. Please have it ready ' +
    '(on your phone, or printed) — it will be scanned once at the entrance.</p>' +
    '<p><img src="cid:qrImage" alt="Your QR code" width="220" height="220"></p>' +
    '<p style="color:#64748b;font-size:13px;">Guest ID: ' + guestId + '</p>' +
    '<p>See you there!</p>';

  MailApp.sendEmail({
    to: email,
    subject: 'Your Entry QR Code',
    htmlBody: htmlBody,
    inlineImages: { qrImage: qrBlob },
    attachments: [qrBlob]
  });
}

// Emails every guest in the sheet their own QR code, one email each, sent
// directly to the Email column on their row. Guests with a missing or
// invalid email are skipped and reported back rather than failing the whole
// batch. Because this sends one email per guest (not a shared PDF), very
// large guest lists may approach Apps Script's execution time limit (~6
// min) — for 500+ guests, consider running this in batches.
function emailQrToAllGuests() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GUESTS_SHEET);
  if (sh.getLastRow() < 2) throw new Error('No guests in the sheet yet.');

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues(); // GuestID, Name, Email, Role
  let sentCount = 0;
  const skipped = [];

  values.forEach(r => {
    const guestId = r[0], name = r[1], email = (r[2] || '').toString().trim();
    if (!guestId) return;
    if (!email || !isValidEmail_(email)) {
      skipped.push({ guestId, name, reason: 'Missing or invalid email address.' });
      return;
    }
    sendGuestQrEmail_(guestId, name, email);
    sentCount++;
  });

  return { sentCount, skippedCount: skipped.length, skipped };
}

// Resends (or sends for the first time) one guest's QR code — used by the
// per-row "Resend Email" button, e.g. after fixing a typo'd address.
function emailQrToSingleGuest(guestId) {
  guestId = (guestId || '').toString().trim().toUpperCase();
  if (!guestId) throw new Error('guestId missing');

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GUESTS_SHEET);
  const rowIndex = findGuestRowIndex_(sh, guestId);
  if (!rowIndex) throw new Error('Guest not found: ' + guestId);

  const name = sh.getRange(rowIndex, 2).getValue();
  const email = (sh.getRange(rowIndex, 3).getValue() || '').toString().trim();
  if (!email || !isValidEmail_(email)) throw new Error('This guest has no valid email on file — edit their row to add one.');

  sendGuestQrEmail_(guestId, name, email);
  return { guestId, name, email, sent: true };
}

// Emails every guest's QR badge to a leader's inbox as downloadable PDF
// attachment(s). Kept as an optional admin/backup tool (e.g. a printed
// master list at the check-in desk) — guests now get their own QR code
// directly via emailQrToAllGuests() above, so this is no longer the primary
// distribution method. If there are more than MAX_PER_PDF guests, they're
// split across multiple emails (one PDF each, labeled "Part X of Y") since a
// single PDF/email with hundreds of fetched QR images gets slow and risks
// hitting Gmail's attachment size limits.
function emailGuestBundleToLeader(leaderEmail, leaderName) {
  leaderEmail = (leaderEmail || '').toString().trim();
  leaderName = (leaderName || '').toString().trim();

  if (!leaderEmail || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(leaderEmail)) {
    throw new Error('Enter a valid email address.');
  }

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GUESTS_SHEET);
  if (sh.getLastRow() < 2) throw new Error('No guests in the sheet yet.');

  const allGuests = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues() // GuestID, Name, Email, Role
    .filter(r => r[0])
    .map(r => [r[0], r[1], r[3]]); // buildGuestBadgePdf_ expects [GuestID, Name, Role]
  if (!allGuests.length) throw new Error('No guests found.');

  const MAX_PER_PDF = 150;
  const batches = [];
  for (let i = 0; i < allGuests.length; i += MAX_PER_PDF) {
    batches.push(allGuests.slice(i, i + MAX_PER_PDF));
  }
  const totalBatches = batches.length;

  batches.forEach((batch, idx) => {
    const partLabel = totalBatches > 1 ? 'All Guests (Part ' + (idx + 1) + ' of ' + totalBatches + ')' : 'All Guests';
    const pdfBlob = buildGuestBadgePdf_(batch, partLabel);

    const greeting = leaderName ? 'Hi ' + leaderName + ',' : 'Hi,';
    const partNote = totalBatches > 1
      ? '<p>This is part ' + (idx + 1) + ' of ' + totalBatches + ' (guest list split because there are ' +
        allGuests.length + ' guests in total).</p>'
      : '';
    const htmlBody =
      '<p>' + greeting + '</p>' +
      '<p>Attached is a printable PDF with QR entry badges for ' + batch.length + ' guest(s).</p>' +
      partNote;

    MailApp.sendEmail({
      to: leaderEmail,
      subject: 'QR Entry Badges' + (totalBatches > 1 ? ' — Part ' + (idx + 1) + ' of ' + totalBatches : '') +
        ' (' + batch.length + ' guests)',
      htmlBody: htmlBody,
      attachments: [pdfBlob]
    });
  });

  return { sentCount: allGuests.length, emailCount: totalBatches, leaderEmail };
}

// Builds a printable PDF of QR badges (image + name + ID per guest, 3 per
// row) by assembling a temporary Google Doc and exporting it, since Apps
// Script has no direct "HTML string with images -> PDF" conversion that
// reliably handles embedded images. The temp Doc is deleted immediately
// after the PDF is extracted so it doesn't clutter Drive.
function buildGuestBadgePdf_(matching, role) {
  const doc = DocumentApp.create('QR Badges — ' + role + ' — ' + new Date().toISOString());
  const docId = doc.getId();
  const body = doc.getBody();
  body.setMarginTop(24).setMarginBottom(24).setMarginLeft(24).setMarginRight(24);

  body.appendParagraph('QR Entry Badges — ' + role)
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(matching.length + ' guest(s) — generated ' + new Date().toLocaleString());
  body.appendParagraph('');

  const COLS = 3;
  for (let i = 0; i < matching.length; i += COLS) {
    const rowGuests = matching.slice(i, i + COLS);
    const cellsData = rowGuests.concat(Array(COLS - rowGuests.length).fill(null));
    // Seed each cell with the guest's name as its initial text (rather than
    // an empty string + cell.clear() + appendParagraph, which was silently
    // dropping the name in some cases) — this guarantees the name paragraph
    // actually exists in the cell.
    const table = body.appendTable([cellsData.map(g => (g ? g[1] : ''))]);
    for (let c = 0; c < COLS; c++) {
      const g = cellsData[c];
      if (!g) continue;
      const guestId = g[0];
      const cell = table.getCell(0, c);

      cell.editAsText().setBold(true); // bold the name text already in the cell
      cell.appendParagraph(guestId);

      const qrUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=200x200&data=' + encodeURIComponent(guestId);
      const imgBlob = UrlFetchApp.fetch(qrUrl).getBlob();
      const img = cell.insertImage(0, imgBlob); // insert image ABOVE the name/ID text
      img.setWidth(110);
      img.setHeight(110);
    }
    body.appendParagraph('');
  }

  doc.saveAndClose();
  const pdfBlob = DriveApp.getFileById(docId).getAs('application/pdf').setName('QR_Badges_' + role.replace(/\s+/g, '_') + '.pdf');
  DriveApp.getFileById(docId).setTrashed(true); // clean up the temp Doc
  return pdfBlob;
}

// Deterministic Guest ID: derived from a hash of (name + role), NOT random.
// The same name+role always produces the exact same ID — so deleting a
// guest and re-adding them later (even after wiping the whole sheet)
// regenerates the identical QR code. A badge printed once stays valid
// forever, with no need to "reset" instead of delete.
function generateDeterministicId_(name, role, prefix, existingIds, idToNameRole) {
  const normalizedKey = name.trim().toLowerCase() + '|' + role.trim().toLowerCase();
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, normalizedKey);
  const hex = digest.map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
  let id = prefix + '-' + hex.substring(0, 6).toUpperCase();

  // Extremely unlikely, but guard against a hash collision between two
  // DIFFERENT names landing on the same ID — extend with more hash
  // characters rather than silently overwriting someone else's badge.
  let extra = 6;
  while (existingIds.has(id) && idToNameRole.get(id) !== normalizedKey) {
    extra += 2;
    id = prefix + '-' + hex.substring(0, extra).toUpperCase();
  }
  return id;
}

// Manual fallback for the scanner: staff types part of a guest's name (case
// insensitive, matches anywhere in the name) and gets back a short list to
// pick the right person from, instead of guessing a Guest ID. Capped at 15
// results so a one-letter search doesn't dump the whole guest list.
function searchGuestsByName_(query) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GUESTS_SHEET);
  if (sh.getLastRow() < 2) return [];
  const q = query.toLowerCase();
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, GUEST_FIXED_COLS.length).getValues();
  const matches = [];
  for (let i = 0; i < values.length; i++) {
    const r = values[i];
    if (!r[0]) continue;
    if (r[1].toString().toLowerCase().includes(q)) {
      const entryTime = r[ENTRY_TIME_COL_INDEX - 1];
      matches.push({
        guestId: r[0], name: r[1], role: r[3],
        entered: !!entryTime,
        entryTime: entryTime ? (entryTime instanceof Date ? entryTime.toISOString() : new Date(entryTime).toISOString()) : null
      });
      if (matches.length >= 15) break;
    }
  }
  return matches;
}

function findGuestRow_(sh, guestId) {
  if (sh.getLastRow() < 2) return null;
  const width = GUEST_FIXED_COLS.length;
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === guestId) {
      const rawEntryTime = values[i][ENTRY_TIME_COL_INDEX - 1];
      return {
        rowIndex: i + 2,
        name: values[i][1],
        role: values[i][3],
        entryTime: rawEntryTime ? (rawEntryTime instanceof Date ? rawEntryTime : new Date(rawEntryTime)) : null
      };
    }
  }
  return null;
}

function logScan_(guestId, name, role, station, result) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
  sh.appendRow([new Date(), guestId, name, role, station, result]);
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
