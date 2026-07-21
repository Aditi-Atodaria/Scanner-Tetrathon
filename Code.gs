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

// Fixed columns before the per-meal columns start in Guests sheet
const GUEST_FIXED_COLS = ['GuestID', 'Name', 'Role', 'QR Code'];

// Default 8-slot schedule — only used the first time setupSheets() runs.
// After that, edit the Config sheet directly; this list is not read again.
const DEFAULT_MEAL_SLOTS = [
  { key: 'CheckIn',     label: 'Check-in' },
  { key: 'Refresh1',    label: 'Refreshments (Day 1, AM)' },
  { key: 'Lunch1',      label: 'Lunch (Day 1)' },
  { key: 'Refresh2',    label: 'Refreshments (Day 1, PM)' },
  { key: 'Dinner1',     label: 'Dinner (Day 1)' },
  { key: 'Breakfast2',  label: 'Breakfast (Day 2)' },
  { key: 'Refresh3',    label: 'Refreshments (Day 2)' },
  { key: 'Lunch2',      label: 'Lunch (Day 2)' }
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
  config.getRange(1, 1, 1, 2).setValues([['SlotKey', 'SlotLabel']]);
  config.getRange(2, 1, DEFAULT_MEAL_SLOTS.length, 2).setValues(
    DEFAULT_MEAL_SLOTS.map(s => [s.key, s.label])
  );
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

  SpreadsheetApp.getUi().alert('Setup complete: Guests, Config, and ScanLog sheets are ready.');
}

// ---------- Web app entry points ----------

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || '';
  const action = (e && e.parameter && e.parameter.action) || '';

  // JSON API for the externally-hosted Scanner.html (camera page can't run
  // inside the Apps Script iframe — see README).
  if (action === 'slots') {
    return jsonResponse_({ status: 'ok', slots: getMealSlotsConfig_() });
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
// Body: { guestId: "P-A3F9", mealSlot: "Lunch1", station: "Gate A" (optional) }
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // avoid two simultaneous scans double-crediting a meal

    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('No data received');
    }
    const data = JSON.parse(e.postData.contents);
    const guestId = (data.guestId || '').toString().trim().toUpperCase();
    const mealSlot = (data.mealSlot || '').toString().trim();
    const station = (data.station || '').toString().trim();

    if (!guestId) throw new Error('guestId missing');
    if (!mealSlot) throw new Error('mealSlot missing');

    const slots = getMealSlotsConfig_();
    const validSlot = slots.find(s => s.key === mealSlot);
    if (!validSlot) {
      logScan_(guestId, '', '', mealSlot, station, 'Invalid meal slot');
      return jsonResponse_({ status: 'error', message: 'Unknown meal slot: ' + mealSlot });
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
function bulkGenerateGuests(guestList) {
  if (!guestList || !guestList.length) return [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(GUESTS_SHEET);
  const slots = getMealSlotsConfig_();
  const rolePrefixes = getRolePrefixMap_();

  const existingIds = new Set(
    sh.getLastRow() > 1
      ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat().filter(String)
      : []
  );

  const rows = [];
  const created = [];
  guestList.forEach(g => {
    const name = (g.name || '').toString().trim();
    const role = (g.role || '').toString().trim();
    if (!name) return;
    const prefix = rolePrefixes[role] || 'X';
    const guestId = generateUniqueId_(prefix, existingIds);
    existingIds.add(guestId);

    const qrFormula = '=IMAGE("https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=' + guestId + '")';
    const row = [guestId, name, role, qrFormula].concat(slots.map(() => ''), [0]);
    rows.push(row);
    created.push({ guestId, name, role });
  });

  if (rows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
  return created;
}

function addSingleGuest(name, role) {
  return bulkGenerateGuests([{ name, role }])[0] || null;
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
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  const slots = [];
  for (const [key, label] of values) {
    if (!key) break; // stop at the blank row before the Role/Prefix section
    slots.push({ key: key.toString().trim(), label: label.toString().trim() });
  }
  return slots;
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

function generateUniqueId_(prefix, existingIds) {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // no 0/O/1/I ambiguity
  let id;
  do {
    let suffix = '';
    for (let i = 0; i < 4; i++) suffix += chars.charAt(Math.floor(Math.random() * chars.length));
    id = prefix + '-' + suffix;
  } while (existingIds.has(id));
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
