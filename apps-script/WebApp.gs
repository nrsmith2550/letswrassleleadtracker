/**
 * Phone-friendly review app for Let's Wrassle leads.
 * Deploy: Apps Script → Deploy → New deployment → Web app
 *         Execute as: Me · Who has access: Only myself
 * Then open the /exec URL on your iPhone in Safari → Share → Add to Home Screen.
 *
 * The app reads and writes the same sheet tabs (Leads / Sent / Denied), so the
 * hourly import, follow-ups and the sheet itself keep working exactly as before.
 */

/** Pages in the hamburger menu: ?page=quote opens the quote calculator, anything else the leads app. */
const PAGES = {
  leads: { file: 'Index', title: "Let's Wrassle Leads" },
  quote: { file: 'Quote', title: "Let's Wrassle Quote" }
};

function doGet(e) {
  const page = PAGES[e && e.parameter && e.parameter.page] ? e.parameter.page : 'leads';
  const t = HtmlService.createTemplateFromFile(PAGES[page].file);
  t.appUrl = ScriptApp.getService().getUrl();
  t.page = page;
  return t.evaluate()
    .setTitle(PAGES[page].title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .addMetaTag('apple-mobile-web-app-capable', 'yes');
}

/** Pulls a shared HTML file (like Menu.html) into a page: <?!= include('Menu', {appUrl, page}) ?> */
function include(name, vars) {
  const t = HtmlService.createTemplateFromFile(name);
  Object.assign(t, vars || {});
  return t.evaluate().getContent();
}

/** Everything the app shows: the review queue plus recent sent/denied. */
function apiList() {
  const ss = SpreadsheetApp.getActive();
  const read = (name, limit) => {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return [];
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, COLS.length).getValues().map(toRec_);
    const out = rows.map(r => ({
      lead: String(r['Lead'] || ''), email: String(r['Email'] || ''), subject: String(r['Subject'] || ''),
      body: String(r['Body'] || ''), instagram: String(r['Instagram'] || ''), dm: String(r['DM draft'] || ''),
      type: String(r['Type'] || ''), location: String(r['Location'] || ''), contact: String(r['Contact name'] || ''),
      phone: String(r['Phone'] || ''), status: String(r['Status'] || ''), notes: String(r['Notes'] || ''),
      followUp: !!r['Thread ID'] && name === TAB.LEADS,
      when: r['Date sent'] instanceof Date ? r['Date sent'].toISOString() : (r['Date added'] instanceof Date ? r['Date added'].toISOString() : ''),
      sentTo: String(r['Sent to'] || '')
    })).filter(r => r.lead);
    return limit ? out.slice(-limit).reverse() : out;
  };
  return {
    review: read(TAB.LEADS),
    sent: read(TAB.SENT, 50),
    denied: read(TAB.DENIED, 50),
    signatureReady: String(getSetting_('Signature') || '').indexOf('[ADD YOUR MAILING ADDRESS') === -1
  };
}

/** Approve one lead (optionally with edited email/subject/body). Returns {ok, message}. */
function apiApprove(lead, edits) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return { ok: false, message: 'Busy, try again in a moment.' };
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(TAB.LEADS);
    const r = findRowByLead_(sh, lead);
    if (!r) return { ok: false, message: 'That lead is no longer in the queue.' };
    edits = edits || {};
    ['Email', 'Phone', 'Subject', 'Body'].forEach(k => {
      const v = edits[k.toLowerCase()];
      if (typeof v === 'string') sh.getRange(r, C[k]).setValue(v);
    });
    approveRow_(sh, r);
    // approveRow_ either moved the row to Sent (success) or left it with a ⚠️ status
    const still = findRowByLead_(sh, lead);
    if (still) return { ok: false, message: String(sh.getRange(still, C['Status']).getValue()).replace(/^⚠️\s*/, '') };
    return { ok: true, message: 'Sent' };
  } finally {
    lock.releaseLock();
  }
}

/** Save edits (email, phone, subject, body) without sending. */
function apiUpdate(lead, fields) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return { ok: false, message: 'Busy, try again in a moment.' };
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(TAB.LEADS);
    const r = findRowByLead_(sh, lead);
    if (!r) return { ok: false, message: 'Lead not found.' };
    fields = fields || {};
    ['Email', 'Phone', 'Subject', 'Body'].forEach(k => {
      const v = fields[k.toLowerCase()];
      if (typeof v === 'string') sh.getRange(r, C[k]).setValue(v);
    });
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function apiDeny(lead) {
  return moveByLead_(TAB.LEADS, TAB.DENIED, lead, { 'Status': 'Not a fit', 'Decision': '' });
}

function apiRestore(lead) {
  return moveByLead_(TAB.DENIED, TAB.LEADS, lead, { 'Status': 'Restored', 'Decision': '' });
}

function moveByLead_(from, to, lead, updates) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return { ok: false, message: 'Busy, try again in a moment.' };
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(from);
    const r = findRowByLead_(sh, lead);
    if (!r) return { ok: false, message: 'Lead not found.' };
    moveRow_(sh, r, to, updates);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function findRowByLead_(sh, lead) {
  if (!sh || sh.getLastRow() < 2) return 0;
  const vals = sh.getRange(2, C['Lead'], sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]) === String(lead)) return i + 2;
  return 0;
}
