/**
 * QR Meal Scan — event guest & meal tracking system
 *
 * Sheets used (auto-created by setupSheets()):
 *   Guests  — one row per guest, one column per meal slot
 *   Config  — meal slot list + role→prefix map (edit here, no code changes needed)
 *   ScanLog — full audit trail of every scan attempt (success/duplicate/invalid)
 *
 * Deploy as Web App (Extensions > Deploy > New deployment > Web app,
 * execute as "Me", access "Anyone with the link"). Copy the /exec URL into
 * Scanner.html (the external camera page) after deploying.
 */

const GUESTS_SHEET = 'Guests';
const CONFIG_SHEET = 'Config';
const LOG_SHEET = 'ScanLog';

// Access code the scanner's login screen checks against. Change this before
// your event, then redeploy — anyone with this code (not just the link) can
// open the scanner.
const AUTH_CODE = 'MEAL2026';

// Fixed columns before the per-meal columns start in Guests sheet
const GUEST_FIXED_COLS = ['GuestID', 'Name', 'Role', 'QR Code'];

// Default 8-slot schedule — only used the first time setupSheets() runs.
// After that, edit the Config sheet directly; this list is not read again.
// StartDateTime/EndDateTime below are PLACEHOLDERS — edit them in the Config
// sheet to match your real event dates/times before going live.
const DEFAULT_MEAL_SLOTS = [
  { key: 'CheckIn',     label: 'Check-in',                       startOffsetHr: 0,  durationHr: 1.5 },
  { key: 'Refresh1',    label: 'Refreshments (Day 1, AM)',       startOffsetHr: 2,  durationHr: 1 },
  { key: 'Lunch1',      label: 'Lunch (Day 1)',                  startOffsetHr: 5,  durationHr: 1.5 },
  { key: 'Refresh2',    label: 'Refreshments (Day 1, PM)',       startOffsetHr: 8,  durationHr: 1 },
  { key: 'Dinner1',     label: 'Dinner (Day 1)',                 startOffsetHr: 11, durationHr: 1.5 },
  { key: 'Breakfast2',  label: 'Breakfast (Day 2)',               startOffsetHr: 22, durationHr: 1.5 },
  { key: 'Refresh3',    label: 'Refreshments (Day 2)',            startOffsetHr: 25, durationHr: 1 },
  { key: 'Lunch2',      label: 'Lunch (Day 2)',                   startOffsetHr: 28, durationHr: 1.5 }
];

const DEFAULT_ROLE_PREFIXES = [
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
  const slots = DEFAULT_MEAL_SLOTS.map(s => s.key);
  const headers = GUEST_FIXED_COLS.concat(slots, ['Meals Completed']);
  guests.getRange(1, 1, 1, headers.length).setValues([headers]);
  guests.setFrozenRows(1);

  let config = ss.getSheetByName(CONFIG_SHEET);
  if (!config) config = ss.insertSheet(CONFIG_SHEET);
  config.clear();
  config.getRange(1, 1, 1, 4).setValues([['SlotKey', 'SlotLabel', 'StartDateTime', 'EndDateTime']]);

  // Placeholder schedule anchored to "now" — EDIT THESE in the Config sheet
  // to your real event dates/times before going live.
  const anchor = new Date();
  anchor.setMinutes(0, 0, 0);
  const scheduleRows = DEFAULT_MEAL_SLOTS.map(s => {
    const start = new Date(anchor.getTime() + s.startOffsetHr * 3600000);
    const end = new Date(start.getTime() + s.durationHr * 3600000);
    return [s.key, s.label, start, end];
  });
  config.getRange(2, 1, scheduleRows.length, 4).setValues(scheduleRows);
  config.getRange(2, 3, scheduleRows.length, 2).setNumberFormat('yyyy-mm-dd hh:mm');

  const roleStart = DEFAULT_MEAL_SLOTS.length + 3;
  config.getRange(roleStart, 1, 1, 2).setValues([['Role', 'Prefix']]);
  config.getRange(roleStart + 1, 1, DEFAULT_ROLE_PREFIXES.length, 2).setValues(
    DEFAULT_ROLE_PREFIXES.map(r => [r.role, r.prefix])
  );
  config.setFrozenRows(1);

  let log = ss.getSheetByName(LOG_SHEET);
  if (!log) log = ss.insertSheet(LOG_SHEET);
  log.getRange(1, 1, 1, 7).setValues([[
    'Timestamp', 'GuestID', 'Name', 'Role', 'Meal Slot', 'Station Note', 'Result'
  ]]);
  log.setFrozenRows(1);

  SpreadsheetApp.getUi().alert(
    'Setup complete: Guests, Config, and ScanLog sheets are ready.\n\n' +
    'IMPORTANT: The Config sheet has placeholder Start/End times for each ' +
    'meal slot, anchored to right now. Edit those dates/times to match your ' +
    'real event schedule before going live — the scanner uses them to decide ' +
    'what\'s currently open.'
  );
}

// ---------- Web app entry points ----------

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || '';
  const action = (e && e.parameter && e.parameter.action) || '';

  // JSON API for the externally-hosted Scanner.html (camera page can't run
  // inside the Apps Script iframe — see README).
  if (action === 'slots') {
    return jsonResponse_({ status: 'ok', slots: getMealSlotsConfig_().map(s => ({ key: s.key, label: s.label })) });
  }
  if (action === 'state') {
    return jsonResponse_(Object.assign({ status: 'ok' }, getEventState_()));
  }

  if (page === 'generator') {
    const t = HtmlService.createTemplateFromFile('Generator');
    t.scriptUrl = ScriptApp.getService().getUrl();
    return t.evaluate()
      .setTitle('QR Meal Scan — Guest & QR Generator')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === 'dashboard') {
    const t = HtmlService.createTemplateFromFile('Dashboard');
    t.scriptUrl = ScriptApp.getService().getUrl();
    return t.evaluate()
      .setTitle('QR Meal Scan — Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  const homeTemplate = HtmlService.createTemplateFromFile('Home');
  homeTemplate.scriptUrl = ScriptApp.getService().getUrl();
  return homeTemplate.evaluate()
    .setTitle('QR Meal Scan')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Receives a scan from the external Scanner.html page.
// Body: { guestId: "P-A3F9", station: "Gate A" (optional) }
// mealSlot is normally NOT sent — the server figures out what's currently
// open from the Config sheet schedule. A client MAY still send an explicit
// mealSlot to override auto-detection (e.g. a staff-picked dropdown variant);
// if present, it's honored instead of the schedule.
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // avoid two simultaneous scans double-crediting a meal

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
    const explicitMealSlot = (data.mealSlot || '').toString().trim();

    if (!guestId) throw new Error('guestId missing');

    const slots = getMealSlotsConfig_();
    let mealSlot, validSlot;

    if (explicitMealSlot) {
      mealSlot = explicitMealSlot;
      validSlot = slots.find(s => s.key === mealSlot);
      if (!validSlot) {
        logScan_(guestId, '', '', mealSlot, station, 'Invalid meal slot');
        return jsonResponse_({ status: 'error', message: 'Unknown meal slot: ' + mealSlot });
      }
    } else {
      const eventState = getEventState_();
      if (eventState.state !== 'active') {
        logScan_(guestId, '', '', '', station, 'Scan outside schedule: ' + eventState.message);
        return jsonResponse_({ status: 'closed', message: eventState.message });
      }
      mealSlot = eventState.activeSlot.key;
      validSlot = { key: eventState.activeSlot.key, label: eventState.activeSlot.label };
    }

    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GUESTS_SHEET);
    const found = findGuestRow_(sh, guestId, slots, mealSlot);
    if (!found) {
      logScan_(guestId, '', '', mealSlot, station, 'Unknown guest ID');
      return jsonResponse_({ status: 'invalid', message: 'QR not recognized: ' + guestId });
    }

    const { rowIndex, name, role, mealColIndex, totalSlots, mealsCompleted } = found;

    const existingVal = sh.getRange(rowIndex, mealColIndex).getValue();
    if (existingVal) {
      logScan_(guestId, name, role, mealSlot, station, 'Duplicate');
      return jsonResponse_({
        status: 'duplicate',
        message: name + ' already logged for ' + validSlot.label,
        name, role, mealLabel: validSlot.label,
        mealsCompleted, totalSlots
      });
    }

    const ts = new Date();
    sh.getRange(rowIndex, mealColIndex).setValue(ts);
    const newCompleted = mealsCompleted + 1;
    const completedColIndex = GUEST_FIXED_COLS.length + totalSlots + 1;
    sh.getRange(rowIndex, completedColIndex).setValue(newCompleted);

    logScan_(guestId, name, role, mealSlot, station, 'Success');

    return jsonResponse_({
      status: 'success',
      name, role,
      mealLabel: validSlot.label,
      mealsCompleted: newCompleted,
      totalSlots,
      timestamp: ts.toISOString()
    });
  } catch (err) {
    return jsonResponse_({ status: 'error', message: err.message });
  } finally {
    lock.releaseLock();
  }
}

// ---------- Called from Generator.html ----------

// guestList: [{name, role}, ...]
// Guests with a name that already exists (case-insensitive, trimmed) are
// skipped rather than added again — both against existing sheet rows and
// against earlier entries in the same batch.
function bulkGenerateGuests(guestList) {
  if (!guestList || !guestList.length) return { created: [], skipped: [] };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(GUESTS_SHEET);
  const slots = getMealSlotsConfig_();
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
    sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(r => {
      if (r[0]) idToNameRole.set(r[0], r[1].toString().trim().toLowerCase() + '|' + r[2].toString().trim().toLowerCase());
    });
  }

  const rows = [];
  const created = [];
  const skipped = [];
  guestList.forEach(g => {
    const name = (g.name || '').toString().trim();
    const role = (g.role || '').toString().trim();
    if (!name) return;

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
    const row = [guestId, name, role, qrFormula].concat(slots.map(() => ''), [0]);
    rows.push(row);
    created.push({ guestId, name, role });
  });

  if (rows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
  return { created, skipped };
}

function addSingleGuest(name, role) {
  name = (name || '').toString().trim();
  const result = bulkGenerateGuests([{ name, role }]);
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
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues(); // GuestID, Name, Role
  return values
    .filter(r => r[0])
    .map(r => ({ guestId: r[0], name: r[1], role: r[2] }));
}

// Edits a guest's name/role in place. The Guest ID (and printed QR) never
// changes — only the details attached to it — so an already-printed badge
// keeps working after an edit.
function updateGuest(guestId, name, role) {
  name = (name || '').toString().trim();
  role = (role || '').toString().trim();
  if (!guestId) throw new Error('guestId missing');
  if (!name) throw new Error('Name cannot be empty');

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GUESTS_SHEET);
  const rowIndex = findGuestRowIndex_(sh, guestId);
  if (!rowIndex) throw new Error('Guest not found: ' + guestId);

  const nameKey = name.toLowerCase();
  const allIds = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues(); // GuestID, Name
  const collision = allIds.some(r => r[0] !== guestId && r[1] && r[1].toString().trim().toLowerCase() === nameKey);
  if (collision) throw new Error('A guest named "' + name + '" already exists.');

  sh.getRange(rowIndex, 2, 1, 2).setValues([[name, role]]); // Name, Role columns
  return { guestId, name, role };
}

// Permanently removes a guest's row (their meal history for the event goes
// with it). ScanLog audit history is left untouched.
function deleteGuest(guestId) {
  if (!guestId) throw new Error('guestId missing');
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GUESTS_SHEET);
  const rowIndex = findGuestRowIndex_(sh, guestId);
  if (!rowIndex) throw new Error('Guest not found: ' + guestId);

  const name = sh.getRange(rowIndex, 2).getValue();
  const role = sh.getRange(rowIndex, 3).getValue();
  sh.deleteRow(rowIndex);
  logScan_(guestId, name, role, '', '', 'Guest deleted');
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

  const width = GUEST_FIXED_COLS.length + getMealSlotsConfig_().length + 1;
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues();

  const removed = values.filter(r => idsToRemove.has(r[0]));
  const remaining = values.filter(r => !idsToRemove.has(r[0]));

  sh.getRange(2, 1, values.length, width).clearContent();
  if (remaining.length) {
    sh.getRange(2, 1, remaining.length, width).setValues(remaining);
  }

  removed.forEach(r => logScan_(r[0], r[1], r[2], '', '', 'Guest deleted (bulk)'));
  return { deletedCount: removed.length };
}

// Wipes every guest from the sheet. Meal-tracking history goes with them.
// ScanLog audit trail is left untouched, with one summary entry recorded.
function deleteAllGuests() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GUESTS_SHEET);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { deletedCount: 0 };

  const width = GUEST_FIXED_COLS.length + getMealSlotsConfig_().length + 1;
  const count = lastRow - 1;
  sh.getRange(2, 1, count, width).clearContent();
  logScan_('ALL', '', '', '', '', 'All guests deleted (' + count + ')');
  return { deletedCount: count };
}

// ---------- Called from Dashboard.html ----------

function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(GUESTS_SHEET);
  const slots = getMealSlotsConfig_();
  const summary = {};
  slots.forEach(s => (summary[s.key] = 0));

  let guests = [];
  if (sh.getLastRow() > 1) {
    const values = sh.getRange(2, 1, sh.getLastRow() - 1, GUEST_FIXED_COLS.length + slots.length + 1).getValues();
    guests = values.map(r => {
      const obj = { guestId: r[0], name: r[1], role: r[2], mealsCompleted: r[r.length - 1] };
      slots.forEach((s, i) => {
        const val = r[GUEST_FIXED_COLS.length + i];
        obj[s.key] = !!val;
        if (val) summary[s.key]++;
      });
      return obj;
    });
  }
  return { slots, summary, guests, totalGuests: guests.length };
}

function getRecentScans(limit) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  const n = limit || 15;
  const last = sh.getLastRow();
  const start = Math.max(2, last - n + 1);
  return sh.getRange(start, 1, last - start + 1, 7).getValues().reverse();
}

function getMealSlotsForClient() {
  return getMealSlotsConfig_();
}

// ---------- Internal helpers ----------

function getMealSlotsConfig_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  const slots = [];
  for (const [key, label, start, end] of values) {
    if (!key) break; // stop at the blank row before the Role/Prefix section
    slots.push({
      key: key.toString().trim(),
      label: label.toString().trim(),
      start: start instanceof Date ? start : (start ? new Date(start) : null),
      end: end instanceof Date ? end : (end ? new Date(end) : null)
    });
  }
  return slots;
}

// Determines what's open right now, purely from the Config sheet schedule —
// this is what lets the scanner auto-detect the active meal instead of
// staff picking one from a dropdown.
function getEventState_() {
  const slots = getMealSlotsConfig_();
  const now = new Date();

  const active = slots.find(s => s.start && s.end && now >= s.start && now <= s.end);
  if (active) {
    return { state: 'active', activeSlot: { key: active.key, label: active.label } };
  }

  const upcoming = slots
    .filter(s => s.start && now < s.start)
    .sort((a, b) => a.start - b.start)[0];
  if (upcoming) {
    return {
      state: 'closed',
      message: 'Not open yet. Next: ' + upcoming.label + ' starts at ' + formatDateTime_(upcoming.start)
    };
  }

  const anyScheduled = slots.some(s => s.start && s.end);
  if (!anyScheduled) {
    return { state: 'closed', message: 'No meal schedule configured yet — check the Config sheet.' };
  }
  return { state: 'closed', message: 'The event has ended — all meal slots are closed.' };
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

// Emails every guest's QR badge to a leader's inbox as downloadable PDF
// attachment(s). If there are more than MAX_PER_PDF guests, they're split
// across multiple emails (one PDF each, labeled "Part X of Y") since a
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

  const allGuests = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues() // GuestID, Name, Role
    .filter(r => r[0]);
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
      '<p>Attached is a printable PDF with QR meal badges for ' + batch.length + ' guest(s).</p>' +
      partNote;

    MailApp.sendEmail({
      to: leaderEmail,
      subject: 'QR Meal Badges' + (totalBatches > 1 ? ' — Part ' + (idx + 1) + ' of ' + totalBatches : '') +
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

  body.appendParagraph('QR Meal Badges — ' + role)
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

function findGuestRow_(sh, guestId, slots, mealSlot) {
  if (sh.getLastRow() < 2) return null;
  const width = GUEST_FIXED_COLS.length + slots.length + 1;
  const slotIndex = slots.findIndex(s => s.key === mealSlot);
  if (slotIndex === -1) return null;
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === guestId) {
      return {
        rowIndex: i + 2,
        name: values[i][1],
        role: values[i][2],
        mealColIndex: GUEST_FIXED_COLS.length + slotIndex + 1,
        totalSlots: slots.length,
        mealsCompleted: values[i][width - 1] || 0
      };
    }
  }
  return null;
}

function logScan_(guestId, name, role, mealSlot, station, result) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
  sh.appendRow([new Date(), guestId, name, role, mealSlot, station, result]);
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
