/**
 * Ugly Sweater Christmas Potluck: RSVP backend
 * ---------------------------------------------
 * Paste this whole file into Extensions > Apps Script of your Google Sheet.
 * See README.md for the 5-minute setup.
 *
 *  - Every RSVP is saved as a row in the "RSVPs" tab.
 *  - If someone RSVPs again with the same name, their row is updated (not duplicated).
 *  - You get an email for every RSVP.
 *  - The RSVP page reads a public dish/drink list (no names, contacts or dietary info).
 */

// Who gets the notification emails. Blank = the Google account that owns this script.
// Add more addresses separated by commas, e.g. "you@gmail.com, cohost@lsu.edu"
const NOTIFY_EMAIL = '';

const SHEET_NAME = 'RSVPs';
const HEADERS = [
  'First submitted', 'Last updated', 'Name', 'Phone / email', 'Attending',
  'Party size', 'Kids', 'Dish type', 'Dish(es)', 'Drinks',
  'Dietary needs', 'Dietary notes', 'Karaoke song', 'Note'
];

/** Run this once from the editor (select "setup" > Run) to create the tab and grant permissions. */
function setup() {
  const sheet = getSheet_();
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#a5121c').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, HEADERS.length, 150);
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

  const row = [
    now, now, name, clean_(d.contact, 120), attending,
    party, kids, clean_(d.dishType, 1000), clean_(d.dishes, 3000), clean_(d.drinks, 3000),
    clean_(d.diet, 200), clean_(d.dietNotes, 300), clean_(d.song, 120), clean_(d.note, 500)
  ];

  const lock = LockService.getScriptLock();
  let existing;
  try {
    lock.waitLock(15000);
    const sheet = getSheet_();
    existing = findRowByName_(sheet, name);
    if (existing) {
      row[0] = sheet.getRange(existing, 1).getValue() || now; // keep the original date
      sheet.getRange(existing, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
  } finally {
    lock.releaseLock();
  }

  // The RSVP is saved at this point; a failed email shouldn't make the guest retry.
  let emailed = true;
  try { notify_(row, !!existing); } catch (err) { emailed = false; console.error('Email failed: ' + err); }

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

  const rows = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
  rows.forEach(function (r) {
    const attending = r[4], party = Number(r[5]) || 0;
    const type = String(r[7] || ''), dish = String(r[8] || '').replace(/^'/, ''), drink = String(r[9] || '').replace(/^'/, '');
    if (attending !== 'Yes' && attending !== 'Maybe') return;
    const name = String(r[2] || '').replace(/^'/, '').trim();
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

function findRowByName_(sheet, name) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const key = normalize_(name);
  const names = sheet.getRange(2, 3, last - 1, 1).getValues();
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

function notify_(row, updated) {
  const to = NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();
  if (!to) return;
  const v = function (i) { return String(row[i] || '').replace(/^'/, ''); };
  const coming = v(4) !== 'No';
  const subject = (updated ? '🔁 Updated RSVP: ' : '🎄 New RSVP: ') + v(2) + ' (' + v(4) +
    (coming ? ', party of ' + v(5) : '') + ')';
  const lines = [
    'Name: ' + v(2),
    'Attending: ' + v(4)
  ];
  if (coming) {
    lines.push('Party size: ' + v(5) + (Number(row[6]) ? ' (' + v(6) + ' kids)' : ''));
    lines.push('Dish: ' + [v(7), v(8)].filter(String).join(': '));
    lines.push('Drinks: ' + v(9));
    lines.push('Dietary: ' + [v(10), v(11)].filter(String).join(' / '));
    if (v(12)) lines.push('Karaoke song: ' + v(12));
  }
  if (v(3)) lines.push('Contact: ' + v(3));
  if (v(13)) lines.push('Note: ' + v(13));
  lines.push('', 'All RSVPs: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl());
  MailApp.sendEmail(to, subject, lines.join('\n'));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
