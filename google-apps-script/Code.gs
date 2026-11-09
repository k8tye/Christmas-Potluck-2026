/**
 * Ugly Sweater Christmas Potluck: RSVP backend
 * ---------------------------------------------
 * Paste this whole file into Extensions > Apps Script of your Google Sheet.
 * See README.md for the 5-minute setup.
 *
 *  - Every RSVP is saved as a row in the "RSVPs" tab.
 *  - If someone RSVPs again with the same name, their row is updated (not duplicated).
 *  - You get an email for every RSVP.
 *  - The RSVP page reads the public guest list: names, party sizes, dishes and drinks.
 *    Dietary needs and notes are never sent to the page.
 *
 * Columns are found by their HEADER NAME in row 1, so you can delete, add, rename-back
 * or reorder columns freely. Missing columns are simply skipped (and the essential
 * ones are re-added at the end if they are ever deleted).
 */

// Who gets the notification emails. Blank = the Google account that owns this script.
// Add more addresses separated by commas, e.g. "you@gmail.com, cohost@lsu.edu"
const NOTIFY_EMAIL = '';

const SHEET_NAME = 'RSVPs';

// Columns used for a brand-new sheet, in this order.
const HEADERS = [
  'First submitted', 'Last updated', 'Name', 'Attending',
  'Party size', 'Kids', 'Dish type', 'Dish(es)', 'Drinks',
  'Dietary needs', 'Dietary notes', 'Note'
];
// These must exist; if one is deleted it is added back at the end of row 1.
const REQUIRED = ['Name', 'Attending', 'Party size', 'Dish type', 'Dish(es)', 'Drinks'];

/** Run this once from the editor (select "setup" > Run) to create the tab and grant permissions. */
function setup() {
  const sheet = getSheet_();
  const n = sheet.getLastColumn();
  sheet.getRange(1, 1, 1, n).setFontWeight('bold').setBackground('#a5121c').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, n, 150);
  MailApp.getRemainingDailyQuota(); // triggers the email permission prompt
}

/**
 * Run this from the editor (select "testEmail" > Run) to check notifications.
 * It sends one test email and shows in the Execution log exactly where it went,
 * or the exact error if Google blocked it.
 */
function testEmail() {
  const to = NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();
  console.log('Sending a test email to: ' + to);
  MailApp.sendEmail(to, '🎄 Potluck RSVP test email', 'Notifications are working! RSVP emails will arrive at this address.\n\nSheet: ' +
    SpreadsheetApp.getActiveSpreadsheet().getUrl());
  console.log('Sent. Emails left today: ' + MailApp.getRemainingDailyQuota());
}

/** Receives an RSVP sent as a POST (kept for compatibility). */
function doPost(e) {
  try {
    return json_(saveRsvp_(JSON.parse((e && e.postData && e.postData.contents) || '{}')));
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** Saves one RSVP: adds a row, or updates the row with the same name. */
function saveRsvp_(d) {
  if (d.website) return { ok: true, saved: true }; // spam bot filled the hidden field

  const name = clean_(d.name, 100);
  const attending = ['Yes', 'Maybe', 'No'].indexOf(d.attending) >= 0 ? d.attending : '';
  if (!name || !attending) return { ok: false, error: 'Name and attendance are required.' };

  const coming = attending !== 'No';
  const party = coming ? Math.max(1, Math.min(20, parseInt(d.party, 10) || 1)) : 0;
  const kids = coming ? Math.max(0, Math.min(party, parseInt(d.kids, 10) || 0)) : 0;
  const now = new Date();

  // Values by column header. Headers not present in the sheet are skipped.
  const rec = {
    'First submitted': now, 'Last updated': now, 'Name': name, 'Attending': attending,
    'Party size': party, 'Kids': kids,
    'Dish type': clean_(d.dishType, 1000), 'Dish(es)': clean_(d.dishes, 3000), 'Drinks': clean_(d.drinks, 3000),
    'Dietary needs': clean_(d.diet, 200), 'Dietary notes': clean_(d.dietNotes, 300), 'Note': clean_(d.note, 500),
    'Phone / email': clean_(d.contact, 120), 'Karaoke song': clean_(d.song, 120)
  };

  const lock = LockService.getScriptLock();
  let existing;
  try {
    lock.waitLock(15000);
    const sheet = getSheet_();
    const col = columns_(sheet);
    const width = sheet.getLastColumn();
    existing = findRowByName_(sheet, col, name);
    // Start from the existing row so any columns you added yourself are kept as-is.
    const row = existing ? sheet.getRange(existing, 1, 1, width).getValues()[0] : new Array(width).fill('');
    if (existing && col['First submitted'] !== undefined && row[col['First submitted']]) {
      rec['First submitted'] = row[col['First submitted']]; // keep the original date
    }
    Object.keys(rec).forEach(function (h) { if (col[h] !== undefined) row[col[h]] = rec[h]; });
    if (existing) sheet.getRange(existing, 1, 1, width).setValues([row]);
    else sheet.appendRow(row);
  } finally {
    lock.releaseLock();
  }

  // The RSVP is saved at this point; a failed email shouldn't make the guest retry.
  let emailed = true;
  try { notify_(rec, !!existing); } catch (err) { emailed = false; console.error('Email failed: ' + err); }

  return { ok: true, saved: true, updated: !!existing, emailed: emailed };
}

/** Saves an RSVP (?rsvp=...) or returns the public guest list: names, party sizes, dishes and drinks.
 *  Contact info, dietary needs and notes are never sent to the page. */
function doGet(e) {
  // The RSVP page sends answers as ?rsvp={...}
  if (e && e.parameter && e.parameter.rsvp) {
    try {
      return json_(saveRsvp_(JSON.parse(e.parameter.rsvp)));
    } catch (err) {
      return json_({ ok: false, error: String(err) });
    }
  }
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  const out = { ok: true, guests: 0, households: 0, people: [], dishes: [], drinks: [], drinkNames: [] };
  if (last < 2) return json_(out);

  const col = columns_(sheet);
  const get = function (r, h) { return col[h] === undefined ? '' : r[col[h]]; };
  const txt = function (r, h) { return String(get(r, h) || '').replace(/^'/, '').trim(); };
  const rows = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
  rows.forEach(function (r) {
    const attending = txt(r, 'Attending');
    if (attending !== 'Yes' && attending !== 'Maybe') return;
    const name = txt(r, 'Name'), party = Number(get(r, 'Party size')) || 0;
    const type = txt(r, 'Dish type'), dish = txt(r, 'Dish(es)'), drink = txt(r, 'Drinks');
    if (attending === 'Yes') { out.guests += party; out.households++; }
    out.people.push({ name: name, attending: attending, party: party });
    if (dish || type) out.dishes.push({ type: type, dish: dish || ('A ' + type.split(', ')[0].toLowerCase() + ' (TBD)'), name: name });
    if (drink) { out.drinks.push(drink); out.drinkNames.push(name); }
  });
  return json_(out);
}

/* ---------------- helpers ---------------- */

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
  }
  return sheet;
}

/** Map of header name -> 0-based column index, read from row 1. Re-adds required headers if missing. */
function columns_(sheet) {
  let width = Math.max(sheet.getLastColumn(), 1);
  const head = sheet.getRange(1, 1, 1, width).getValues()[0];
  const col = {};
  head.forEach(function (h, i) { const k = String(h || '').trim(); if (k && col[k] === undefined) col[k] = i; });
  REQUIRED.forEach(function (h) {
    if (col[h] === undefined) {
      width += 1;
      sheet.getRange(1, width).setValue(h);
      col[h] = width - 1;
    }
  });
  return col;
}

function findRowByName_(sheet, col, name) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const key = normalize_(name);
  const names = sheet.getRange(2, col['Name'] + 1, last - 1, 1).getValues();
  for (let i = 0; i < names.length; i++) {
    if (normalize_(names[i][0]) === key) return i + 2;
  }
  return 0;
}

function normalize_(s) {
  return String(s || '').replace(/^'/, '').trim().toLowerCase().replace(/\s+/g, ' ');
}

/** Trim, cap length, and stop text from being treated as a spreadsheet formula. */
function clean_(v, max) {
  let s = String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function notify_(rec, updated) {
  const to = NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();
  if (!to) return;
  const v = function (h) { return String(rec[h] == null ? '' : rec[h]).replace(/^'/, ''); };
  const coming = v('Attending') !== 'No';
  const subject = (updated ? '🔁 Updated RSVP: ' : '🎄 New RSVP: ') + v('Name') + ' (' + v('Attending') +
    (coming ? ', party of ' + v('Party size') : '') + ')';
  const lines = ['Name: ' + v('Name'), 'Attending: ' + v('Attending')];
  if (coming) {
    lines.push('Party size: ' + v('Party size') + (Number(rec['Kids']) ? ' (' + v('Kids') + ' kids)' : ''));
    lines.push('Dish: ' + [v('Dish type'), v('Dish(es)')].filter(String).join(': '));
    lines.push('Drinks: ' + v('Drinks'));
    lines.push('Dietary: ' + [v('Dietary needs'), v('Dietary notes')].filter(String).join(' / '));
  }
  if (v('Note')) lines.push('Note: ' + v('Note'));
  lines.push('', 'All RSVPs: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl());
  MailApp.sendEmail(to, subject, lines.join('\n'));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
