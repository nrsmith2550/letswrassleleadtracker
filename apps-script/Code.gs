/**
 * Let's Wrassle Leads: approve or deny wholesale leads right in Google Sheets.
 *
 * Tabs:
 *   Leads     your review queue. Set "Decision" to Approve or Deny.
 *   Sent      leads you've emailed (follow-ups get queued back into Leads).
 *   Denied    leads you said no to. Kept so they never come back.
 *   Settings  from address, signature, follow-up timing, incoming folder.
 *
 * Approve → sends the email from your Gmail (using your support@ alias) and moves the row to Sent.
 * Deny    → moves the row to Denied.
 * New leads arrive as CSV files in a Drive folder (the weekly Claude run drops them there)
 * and are imported every hour, with duplicates skipped. Imported files are trashed (or renamed "[imported] …").
 */

// Drive folder the weekly Claude run drops new-lead CSVs into (owned by nathansmith2550@gmail.com,
// shared with the Gmail account that runs this script). Leave '' to have Setup create a new one.
const INCOMING_FOLDER_ID = '1-CjSNMzBOOC7uxM2CJiEig8fGLPY83uW';

const TAB = { LEADS: 'Leads', SENT: 'Sent', DENIED: 'Denied', SETTINGS: 'Settings' };

const COLS = [
  'Decision', 'Lead', 'Email', 'Subject', 'Body', 'Instagram', 'DM draft',
  'Type', 'Location', 'Contact name', 'Phone', 'Status', 'Notes', 'Source',
  'Date added', 'Date sent', 'Sent to', 'Thread ID', 'Follow-up sent'
];
const C = Object.fromEntries(COLS.map((name, i) => [name, i + 1])); // 1-based column numbers

const DEFAULT_SETTINGS = [
  ['From address', 'support@letswrassleapparel.com', 'Must be a "Send mail as" alias on the Gmail account that owns this script.'],
  ['From name', "Let's Wrassle Apparel Co.", 'Name recipients see.'],
  ['Signature', "--\nNate Smith\nLet's Wrassle Apparel Co.\nletswrassleapparel.com | @lets_wrassle\n407 Nike Drive, Pittsburgh, PA",
    'Added to every email. Anti-spam law (CAN-SPAM) requires a physical mailing address; a PO box is fine.'],
  ['Follow-up after (days)', 5, 'Leads with no reply after this many days get one follow-up queued back into Leads.'],
  ['Max emails per day', 50, 'Safety cap. Consumer Gmail allows about 100 recipients a day from scripts.'],
  ['Incoming folder ID', INCOMING_FOLDER_ID, 'The weekly Claude run drops CSV files of new leads here. Setup creates a folder if this is blank.'],
];

// ---------- Menu & setup ----------

function onOpen() {
  SpreadsheetApp.getUi().createMenu("Let's Wrassle")
    .addItem('Run setup (first time)', 'setup')
    .addSeparator()
    .addItem('Import new leads now', 'importIncoming')
    .addItem('Check for follow-ups now', 'queueFollowUps')
    .addItem('Process Approve/Deny now', 'processAllDecisions')
    .addItem('Tidy row heights', 'tidySheets')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActive();
  [TAB.LEADS, TAB.SENT, TAB.DENIED].forEach(name => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold')
      .setBackground('#1F3A5F').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.setColumnWidth(C['Decision'], 110);
    sh.setColumnWidth(C['Lead'], 220);
    sh.setColumnWidth(C['Email'], 240);
    sh.setColumnWidth(C['Subject'], 260);
    sh.setColumnWidth(C['Body'], 360);
    sh.setColumnWidth(C['DM draft'], 300);
    sh.setColumnWidth(C['Notes'], 280);
    sh.getRange(2, C['Body'], sh.getMaxRows() - 1, 1).setWrap(false);
  });

  const leads = ss.getSheetByName(TAB.LEADS);
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(['Approve', 'Deny'], true).setAllowInvalid(false).build();
  leads.getRange(2, C['Decision'], leads.getMaxRows() - 1, 1).setDataValidation(rule);
  leads.setFrozenColumns(2);

  // Settings tab
  let st = ss.getSheetByName(TAB.SETTINGS);
  if (!st) {
    st = ss.insertSheet(TAB.SETTINGS);
    st.getRange(1, 1, 1, 3).setValues([['Setting', 'Value', 'Notes']]).setFontWeight('bold');
    st.getRange(2, 1, DEFAULT_SETTINGS.length, 3).setValues(DEFAULT_SETTINGS);
    st.setColumnWidth(1, 200); st.setColumnWidth(2, 380); st.setColumnWidth(3, 480);
    st.getRange(2, 2, DEFAULT_SETTINGS.length, 1).setWrap(true).setBackground('#FFF2CC');
  }

  // Incoming folder
  let folderId = getSetting_('Incoming folder ID');
  if (!folderId) {
    const folder = DriveApp.createFolder("Let's Wrassle – Incoming Leads");
    folderId = folder.getId();
    setSetting_('Incoming folder ID', folderId);
  }

  // Triggers: replace any existing ones from this project
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('handleEdit').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('importIncoming').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('queueFollowUps').timeBased().everyDays(1).atHour(8).create();

  // Remove the default empty "Sheet1" if it's still there
  const s1 = ss.getSheetByName('Sheet1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s1);

  importIncoming();

  tidySheets();
  const aliasOk = checkAlias_();
  notify_(
    "Let's Wrassle setup complete ✅\n\n" +
    'Incoming leads folder ID:\n' + folderId + '\n(share this with Claude for the weekly run)\n\n' +
    (aliasOk ? 'From address verified.' : '⚠️ "' + getSetting_('From address') + '" is not a Send-mail-as alias on ' +
      Session.getEffectiveUser().getEmail() + '. Add it in Gmail → Settings → Accounts, or emails will fail.') +
    '\n\nNext: put your mailing address in Settings → Signature.'
  );
}

/** Shows a popup when run from the sheet menu; just logs when run from the script editor. */
function notify_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

/** Keeps every row one line tall (long emails are clipped, not wrapped) so the sheet is easy to scroll on a phone. */
function tidySheets() {
  const ss = SpreadsheetApp.getActive();
  [TAB.LEADS, TAB.SENT, TAB.DENIED].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return;
    const n = sh.getLastRow() - 1;
    sh.getRange(2, 1, n, COLS.length).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setVerticalAlignment('middle');
    sh.setRowHeightsForced(2, n, 24);
  });
}

// ---------- Approve / Deny ----------

function handleEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== TAB.LEADS) return;
  const col = e.range.getColumn();
  const lastCol = col + e.range.getNumColumns() - 1;
  if (col > C['Decision'] || lastCol < C['Decision']) return;
  processDecisions_(sh, e.range.getRow(), e.range.getLastRow());
}

function processAllDecisions() {
  const sh = SpreadsheetApp.getActive().getSheetByName(TAB.LEADS);
  if (sh.getLastRow() >= 2) processDecisions_(sh, 2, sh.getLastRow());
}

function processDecisions_(sh, firstRow, lastRow) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return;
  try {
    // bottom-up so deleting rows doesn't shift the ones we still have to handle
    for (let r = Math.min(lastRow, sh.getLastRow()); r >= Math.max(2, firstRow); r--) {
      const decision = String(sh.getRange(r, C['Decision']).getValue()).trim().toLowerCase();
      if (decision === 'approve') approveRow_(sh, r);
      else if (decision === 'deny') moveRow_(sh, r, TAB.DENIED, { 'Status': 'Not a fit', 'Decision': '' });
    }
  } finally {
    lock.releaseLock();
  }
}

function approveRow_(sh, r) {
  const rec = readRow_(sh, r);
  const to = String(rec['Email'] || '').trim();
  const fail = msg => {
    sh.getRange(r, C['Decision']).setValue('');
    sh.getRange(r, C['Status']).setValue('⚠️ ' + msg);
  };

  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(to)) return fail('Add a valid email, then Approve again');
  if (!rec['Body']) return fail('Body is empty');
  if (!underDailyCap_()) return fail('Daily send cap reached. Try again tomorrow, or raise it in Settings');

  const from = getSetting_('From address');
  const opts = { name: getSetting_('From name') };
  if (from && checkAlias_()) opts.from = from;
  else if (from) return fail(from + ' is not a Gmail alias on this account (see Settings)');

  const body = String(rec['Body']).trim() + '\n\n' + String(getSetting_('Signature') || '').trim();
  const isFollowUp = !!rec['Thread ID'];

  try {
    let threadId = rec['Thread ID'];
    if (isFollowUp) {
      // Sent as a new message with the same "Re:" subject, so Gmail threads it with the original.
      // (GmailThread.reply() would go to the last sender, which is you.)
      GmailApp.sendEmail(to, rec['Subject'], body, opts);
    } else {
      const subject = rec['Subject'] || ("Let's Wrassle apparel for " + rec['Lead']);
      GmailApp.sendEmail(to, subject, body, opts);
      Utilities.sleep(1500);
      const found = GmailApp.search('in:sent to:' + to + ' subject:"' + subject.replace(/"/g, '') + '"', 0, 1);
      threadId = found.length ? found[0].getId() : '';
    }
    countSend_();
    const now = new Date();
    const updates = isFollowUp
      ? { 'Decision': '', 'Status': 'Followed up', 'Follow-up sent': now }
      : { 'Decision': '', 'Status': 'Contacted', 'Date sent': now, 'Sent to': to, 'Thread ID': threadId };
    if (isFollowUp) {
      // the original row is still on Sent: update it there instead of adding a second one
      const sent = SpreadsheetApp.getActive().getSheetByName(TAB.SENT);
      const sr = findRowByThread_(sent, threadId);
      if (sr) { writeFields_(sent, sr, updates); sh.deleteRow(r); return; }
    }
    moveRow_(sh, r, TAB.SENT, updates);
  } catch (err) {
    fail('Send failed: ' + err.message);
  }
}

// ---------- Follow-ups ----------

/** Marks replies, and queues ONE follow-up per unanswered lead back into the Leads tab for approval. */
function queueFollowUps() {
  const ss = SpreadsheetApp.getActive();
  const sent = ss.getSheetByName(TAB.SENT), leads = ss.getSheetByName(TAB.LEADS);
  const days = Number(getSetting_('Follow-up after (days)')) || 5;
  if (sent.getLastRow() < 2) return;
  const rows = sent.getRange(2, 1, sent.getLastRow() - 1, COLS.length).getValues();
  const queued = new Set(colValues_(leads, C['Thread ID']));
  const me = Session.getEffectiveUser().getEmail().toLowerCase();

  rows.forEach((row, i) => {
    const r = i + 2, rec = toRec_(row);
    const threadId = rec['Thread ID'];
    if (!threadId || rec['Follow-up sent'] || rec['Status'] === 'Replied' || queued.has(threadId)) return;

    let thread;
    try { thread = GmailApp.getThreadById(threadId); } catch (e) { return; }
    if (!thread) return;
    const to = String(rec['Sent to'] || rec['Email']).toLowerCase();
    const replied = thread.getMessages().some(m => {
      const f = m.getFrom().toLowerCase();
      return f.indexOf(to) !== -1 || (f.indexOf(me) === -1 && f.indexOf('letswrassleapparel.com') === -1);
    });
    if (replied) { sent.getRange(r, C['Status']).setValue('Replied'); return; }

    const sentAt = rec['Date sent'] instanceof Date ? rec['Date sent'] : new Date(rec['Date sent']);
    if (isNaN(sentAt) || (Date.now() - sentAt.getTime()) < days * 864e5) return;

    const first = String(rec['Contact name'] || '').split(/[ (]/)[0];
    const fu = Object.assign({}, rec, {
      'Decision': '',
      'Subject': 'Re: ' + (thread.getFirstMessageSubject() || rec['Subject']),
      'Body': (first ? 'Hi ' + first + ',' : 'Hi there,') +
        '\n\nJust bumping this up in case it got buried. I\'d love to drop off a sample Let\'s Wrassle shirt for ' +
        rec['Lead'] + ' so you can see the quality. Would a quick call or a visit this week work?' +
        "\n\nIf the timing isn't right, no worries at all.\n\nThanks,\nNate",
      'Status': 'Follow-up due',
      'Date added': new Date()
    });
    leads.appendRow(COLS.map(c => fu[c] === undefined ? '' : fu[c]));
    queued.add(threadId);
  });
}

// ---------- Import new leads from Drive ----------

/** Imports every CSV in the incoming folder, skips duplicates, then trashes the file. */
function importIncoming() {
  const folderId = getSetting_('Incoming folder ID');
  if (!folderId) return;
  const ss = SpreadsheetApp.getActive();
  const leads = ss.getSheetByName(TAB.LEADS);
  const seen = existingKeys_();
  const files = DriveApp.getFolderById(folderId).getFiles();
  let added = 0;

  while (files.hasNext()) {
    const f = files.next();
    if (/^\[imported\]/.test(f.getName())) continue;
    const isCsv = f.getMimeType() === MimeType.CSV || /\.csv$/i.test(f.getName()) || f.getMimeType() === MimeType.PLAIN_TEXT;
    if (!isCsv) continue;
    const data = Utilities.parseCsv(f.getBlob().getDataAsString('UTF-8'));
    if (data.length < 2) { markImported_(f); continue; }
    const header = data[0].map(h => normHeader_(h));
    const rows = [];
    data.slice(1).forEach(line => {
      const rec = {};
      header.forEach((h, i) => { if (h) rec[h] = line[i]; });
      if (!rec['Lead']) return;
      const key = leadKey_(rec);
      if (seen.has(key.name) || (key.email && seen.has(key.email))) return;
      seen.add(key.name); if (key.email) seen.add(key.email);
      rec['Decision'] = '';
      rec['Status'] = rec['Status'] || (rec['Email'] ? 'Ready to send' : rec['Instagram'] ? 'DM only' : 'Find contact');
      rec['Date added'] = new Date();
      rows.push(COLS.map(c => rec[c] === undefined ? '' : rec[c]));
    });
    if (rows.length) {
      leads.getRange(leads.getLastRow() + 1, 1, rows.length, COLS.length).setValues(rows);
      added += rows.length;
    }
    markImported_(f);
  }
  if (added) { tidySheets(); ss.toast(added + ' new lead(s) imported', "Let's Wrassle", 5); }
}

// ---------- Helpers ----------

/** Trash the file if we own it; otherwise (shared folder) rename it so it's skipped next time. */
function markImported_(f) {
  try { f.setTrashed(true); }
  catch (e) { f.setName('[imported] ' + f.getName()); }
}

const HEADER_ALIASES = {
  'lead': 'Lead', 'name': 'Lead', 'lead name': 'Lead', 'email': 'Email', 'subject': 'Subject', 'body': 'Body',
  'instagram': 'Instagram', 'dm': 'DM draft', 'dm draft': 'DM draft', 'type': 'Type', 'location': 'Location',
  'contact name': 'Contact name', 'contactname': 'Contact name', 'contact': 'Contact name', 'phone': 'Phone',
  'status': 'Status', 'notes': 'Notes', 'source': 'Source'
};
function normHeader_(h) { return HEADER_ALIASES[String(h).trim().toLowerCase()] || null; }

function leadKey_(rec) {
  return {
    name: 'n:' + String(rec['Lead'] || '').toLowerCase().replace(/[^a-z0-9]/g, ''),
    email: rec['Email'] ? 'e:' + String(rec['Email']).trim().toLowerCase() : ''
  };
}

function existingKeys_() {
  const ss = SpreadsheetApp.getActive(), keys = new Set();
  [TAB.LEADS, TAB.SENT, TAB.DENIED].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return;
    sh.getRange(2, 1, sh.getLastRow() - 1, COLS.length).getValues().forEach(row => {
      const k = leadKey_(toRec_(row)); keys.add(k.name); if (k.email) keys.add(k.email);
    });
  });
  return keys;
}

function toRec_(row) { const o = {}; COLS.forEach((c, i) => o[c] = row[i]); return o; }
function readRow_(sh, r) { return toRec_(sh.getRange(r, 1, 1, COLS.length).getValues()[0]); }
function colValues_(sh, col) { return sh.getLastRow() < 2 ? [] : sh.getRange(2, col, sh.getLastRow() - 1, 1).getValues().map(v => v[0]).filter(String); }

function writeFields_(sh, r, fields) {
  Object.keys(fields).forEach(k => { if (C[k]) sh.getRange(r, C[k]).setValue(fields[k]); });
}

function moveRow_(sh, r, destName, updates) {
  const rec = Object.assign(readRow_(sh, r), updates || {});
  const dest = SpreadsheetApp.getActive().getSheetByName(destName);
  dest.appendRow(COLS.map(c => rec[c] === undefined ? '' : rec[c]));
  sh.deleteRow(r);
}

function findRowByThread_(sh, threadId) {
  const vals = sh.getLastRow() < 2 ? [] : sh.getRange(2, C['Thread ID'], sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) if (vals[i][0] === threadId) return i + 2;
  return 0;
}

function getSetting_(key) {
  const st = SpreadsheetApp.getActive().getSheetByName(TAB.SETTINGS);
  if (!st) { const d = DEFAULT_SETTINGS.find(s => s[0] === key); return d ? d[1] : ''; }
  const vals = st.getRange(2, 1, Math.max(st.getLastRow() - 1, 1), 2).getValues();
  const hit = vals.find(v => v[0] === key);
  return hit ? hit[1] : '';
}

function setSetting_(key, value) {
  const st = SpreadsheetApp.getActive().getSheetByName(TAB.SETTINGS);
  const vals = st.getRange(2, 1, Math.max(st.getLastRow() - 1, 1), 1).getValues();
  const i = vals.findIndex(v => v[0] === key);
  if (i >= 0) st.getRange(i + 2, 2).setValue(value);
  else st.appendRow([key, value, '']);
}

function checkAlias_() {
  const from = String(getSetting_('From address') || '').toLowerCase();
  if (!from) return true;
  if (Session.getEffectiveUser().getEmail().toLowerCase() === from) return true;
  return GmailApp.getAliases().map(a => a.toLowerCase()).indexOf(from) !== -1;
}

function underDailyCap_() {
  const p = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const n = p.getProperty('sent:' + today);
  return Number(n || 0) < (Number(getSetting_('Max emails per day')) || 50);
}

function countSend_() {
  const p = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  p.setProperty('sent:' + today, String(Number(p.getProperty('sent:' + today) || 0) + 1));
}
