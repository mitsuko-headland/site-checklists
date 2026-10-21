/**
 * Headland Site Walks – Google Apps Script backend.
 * Paste this into Code.gs in the Apps Script editor of the "Site Walks" Google Sheet.
 * The sheet must have these tabs (headers in row 1):
 *   Users      : Name | Email | Active | Admin   (optional – see sign-in notes below)
 *   Items      : ItemID | Walk | Trade | Area | Hold | Item | Division | Order | Active
 *   Jobs       : JobID | Job Name | Address | Created By | Created At | Active
 *   Job Items  : JobID | ItemID | Type | Walk | Trade | Area | Hold | Item | Added By | Added At | Removed
 *   Check Log  : JobID | ItemID | Walk | Status | Name | Timestamp
 */

const TZ = 'America/Vancouver';
const DOMAIN = 'headlandconstruction.com';   // only Google accounts on this domain can use the app

/* ---------------- web app entry ---------------- */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Headland Site Walks')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/* ---------------- sheet helpers ---------------- */
function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('Missing tab "' + name + '" in the Google Sheet.');
  return sh;
}
function rows_(name) {
  const sh = sheet_(name);
  const values = sh.getDataRange().getValues();
  const head = values.shift().map(h => String(h).trim());
  return values
    .map((r, i) => { const o = { _row: i + 2 }; head.forEach((h, j) => o[h] = r[j]); return o; })
    .filter(o => head.some(h => o[h] !== '' && o[h] !== null));
}
function append_(name, obj) {
  const sh = sheet_(name);
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
  sh.appendRow(head.map(h => (obj[h] === undefined ? '' : obj[h])));
}
function yes_(v) { return String(v).trim().toUpperCase().startsWith('Y'); }
function fmt_(d) { return d instanceof Date ? Utilities.formatDate(d, TZ, 'MMM d, h:mm a') : String(d || ''); }
function newId_(prefix) { return prefix + Date.now().toString(36).toUpperCase() + Math.floor(Math.random() * 1296).toString(36).toUpperCase(); }
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------------- sign-in (Google account) ----------------
 * Deploy as: Execute as "Me", Who has access "Anyone within headlandconstruction.com".
 * Google signs the person in; this script only checks the email domain and the Users tab.
 * Anyone with a Headland email can use it. The Users tab is optional and is used to:
 *   - give someone a display name (otherwise it is made from their email, e.g. jane.doe -> Jane Doe)
 *   - block someone (Active = N)
 *   - make someone an admin (Admin = Y) so they can add items to the master list
 */
function user_() {
  const email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (!email || !email.endsWith('@' + DOMAIN)) throw new Error('AUTH');
  const row = rows_('Users').find(r => String(r['Email']).trim().toLowerCase() === email);
  if (row && String(row['Active']).trim() && !yes_(row['Active'])) throw new Error('BLOCKED');
  const owner = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  const name = (row && String(row['Name']).trim()) || email.split('@')[0].split(/[._-]+/).filter(Boolean).map(w => w[0].toUpperCase() + w.slice(1)).join(' ');
  return { email: email, name: name, admin: email === owner || !!(row && yes_(row['Admin'])) };
}

/** Called on app start: who is signed in. */
function whoAmI() { const u = user_(); return { name: u.name, email: u.email, admin: u.admin }; }

/* ---------------- jobs ---------------- */
function getJobs(token) {
  user_(token);
  return rows_('Jobs')
    .filter(j => !String(j['Active']).trim() || yes_(j['Active']))
    .map(j => ({ id: String(j['JobID']), name: String(j['Job Name']), address: String(j['Address'] || ''), created: j['Created At'] instanceof Date ? j['Created At'].getTime() : 0 }))
    .sort((a, b) => b.created - a.created);
}
function createJob(token, name, address) {
  const u = user_(token);
  name = String(name || '').trim();
  if (!name) throw new Error('Enter a job name.');
  return withLock_(() => {
    const id = newId_('J');
    append_('Jobs', { 'JobID': id, 'Job Name': name, 'Address': String(address || '').trim(), 'Created By': u.name, 'Created At': new Date(), 'Active': 'Y' });
    return { id: id, name: name, address: String(address || '').trim() };
  });
}

/** Everything for one job: all walks' items plus the latest status of each. */
function getJob(token, jobId) {
  user_(token);
  jobId = String(jobId);
  const job = rows_('Jobs').find(j => String(j['JobID']) === jobId);
  if (!job) throw new Error('That job no longer exists.');
  const extras = rows_('Job Items').filter(r => String(r['JobID']) === jobId && !yes_(r['Removed']));
  const hidden = new Set(extras.filter(r => String(r['Type']).toUpperCase() === 'HIDE').map(r => String(r['ItemID'])));
  const items = [];
  rows_('Items').forEach(r => {
    const id = String(r['ItemID']);
    if (!id || hidden.has(id) || (String(r['Active']).trim() && !yes_(r['Active']))) return;
    items.push({ id: id, src: 'm', walk: Number(r['Walk']), trade: String(r['Trade'] || 'General'), area: String(r['Area'] || ''), hold: String(r['Hold']).toUpperCase().indexOf('HOLD') >= 0, text: String(r['Item']), order: Number(r['Order']) || 0 });
  });
  extras.filter(r => String(r['Type']).toUpperCase() !== 'HIDE').forEach(r => {
    items.push({ id: String(r['ItemID']), src: 'j', walk: Number(r['Walk']), trade: String(r['Trade'] || 'General'), area: String(r['Area'] || ''), hold: String(r['Hold']).toUpperCase().indexOf('HOLD') >= 0, text: String(r['Item']), order: 100000 + (r['Added At'] instanceof Date ? r['Added At'].getTime() / 1e9 : 0) });
  });
  const checks = {};
  rows_('Check Log').forEach(r => {           // later rows win, so this is the current status
    if (String(r['JobID']) !== jobId) return;
    const st = String(r['Status']);
    checks[String(r['ItemID'])] = st === 'Cleared' ? null : { status: st, name: String(r['Name']), at: fmt_(r['Timestamp']) };
  });
  return { job: { id: jobId, name: String(job['Job Name']), address: String(job['Address'] || '') }, items: items, checks: checks };
}

/** status: 'Done', 'N/A' or 'Cleared'. Every change is a new row in Check Log. */
function setCheck(token, jobId, itemId, walk, status) {
  const u = user_(token);
  if (['Done', 'N/A', 'Cleared'].indexOf(status) < 0) throw new Error('Unknown status.');
  return withLock_(() => {
    const now = new Date();
    append_('Check Log', { 'JobID': String(jobId), 'ItemID': String(itemId), 'Walk': Number(walk), 'Status': status, 'Name': u.name, 'Timestamp': now });
    return status === 'Cleared' ? null : { status: status, name: u.name, at: fmt_(now) };
  });
}

/** rows: [{walk, trade, area, hold, text}], scope: 'job' or 'all' (master list – Admin only). */
function addItems(token, jobId, scope, rows) {
  const u = user_(token);
  rows = (rows || []).filter(r => r && String(r.text || '').trim() && Number(r.walk) >= 1 && Number(r.walk) <= 4);
  if (!rows.length) throw new Error('Nothing to add.');
  if (scope === 'all' && !u.admin) throw new Error('Only admins can add to the master list. Choose "This job only" or ask the office.');
  return withLock_(() => {
    if (scope === 'all') {
      let order = Math.max(0, ...rows_('Items').map(r => Number(r['Order']) || 0));
      rows.forEach(r => append_('Items', { 'ItemID': newId_('M'), 'Walk': Number(r.walk), 'Trade': String(r.trade || 'General').trim(), 'Area': String(r.area || '').trim(), 'Hold': r.hold ? 'HOLD' : '', 'Item': String(r.text).trim(), 'Division': '', 'Order': ++order, 'Active': 'Y' }));
    } else {
      rows.forEach(r => append_('Job Items', { 'JobID': String(jobId), 'ItemID': newId_('X'), 'Type': 'Item', 'Walk': Number(r.walk), 'Trade': String(r.trade || 'General').trim(), 'Area': String(r.area || '').trim(), 'Hold': r.hold ? 'HOLD' : '', 'Item': String(r.text).trim(), 'Added By': u.name, 'Added At': new Date(), 'Removed': '' }));
    }
    return rows.length;
  });
}

/** Removes an item from this job only. Master-list items are hidden for this job, not deleted. */
function removeItem(token, jobId, itemId, src, walk) {
  const u = user_(token);
  return withLock_(() => {
    if (src === 'm') {
      append_('Job Items', { 'JobID': String(jobId), 'ItemID': String(itemId), 'Type': 'HIDE', 'Walk': Number(walk), 'Added By': u.name, 'Added At': new Date(), 'Removed': '' });
    } else {
      const sh = sheet_('Job Items');
      const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
      const r = rows_('Job Items').find(x => String(x['JobID']) === String(jobId) && String(x['ItemID']) === String(itemId));
      if (r) sh.getRange(r._row, head.indexOf('Removed') + 1).setValue('Y');
    }
    return true;
  });
}
