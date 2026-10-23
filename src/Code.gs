/**
 * Headland Site Checklists – Google Apps Script backend.
 * Lives in the "Site Checklists" Google Sheet; published automatically from GitHub (mitsuko-headland/site-checklists).
 * Tabs (headers in row 1):
 *   Users      : Name | PIN | Active | Admin
 *   Items      : ItemID | Checklist | Trade | Area | Critical | Item | Division | Order | Active
 *   Jobs       : JobID | Job Name | Address | Created By | Created At | Active
 *   Job Items  : JobID | ItemID | Type | Checklist | Trade | Area | Critical | Item | Added By | Added At | Removed
 *                Type = Item (task added to one job) | HIDE (master task removed from one job) | EDIT (master task changed for one job)
 *                       | LIST (whole checklist hidden on one job by an admin; Removed = Y means shown again)
 *   Check Log  : JobID | ItemID | Checklist | Status | Name | Timestamp
 *   Edit Log   : Timestamp | Name | JobID | ItemID | Field | Old Value | New Value   (created automatically)
 *   Notes      : JobID | ItemID | Checklist | Note | Name | Timestamp                 (created automatically)
 *   Trades     : Code | Trade   – heading codes and sort order                      (created automatically)
 * Older sheets that still say "Walk" / "Hold" are upgraded automatically on first use (see upgrade_).
 */

const TZ = 'America/Vancouver';
const SESSION_DAYS = 30;
const MAX_FAILS = 5;            // wrong PINs allowed per name …
const LOCK_MINUTES = 15;        // … before that name is locked for this long
const CHECKLISTS = 6;           // 1 Project Startup & Site Protection … 6 Finishing
const SETUP_TEXT_VERSION = '6';
const NOTE_MAX = 2000;          // longest note accepted, in characters  // bump to rewrite the Setup tab after instructions change
/* Trade codes (from Headland's divisions list). Seeds the Trades tab the first time; after that the tab is the source. */
const TRADES = [['1036', 'Mobilize & Startup Site Protection'], ['1050', 'Site Supervisor'], ['2001', 'Hazardous Demolition'], ['2006', 'Demolition'],
  ['3601', 'Concrete'], ['5160', 'Metalwork'], ['6101', 'Framing'], ['6715', 'Exterior Carpentry'], ['6815', 'Finish Carpentry'], ['7201', 'Roofing'],
  ['7501', 'Insulation'], ['8102', 'Windows'], ['8202', 'Doors'], ['8501', 'Glass & Mirror'], ['9001', 'Drywall'], ['9210', 'Painting'], ['9215', 'Wallpaper'],
  ['9401', 'Flooring'], ['9601', 'Tile'], ['10400', 'Fireplace'], ['11500', 'Appliance Installer'], ['12350', 'Millwork'], ['12370', 'Stone Countertops'],
  ['15201', 'Gas Fitter'], ['15201', 'Plumbing'], ['15319', 'HVAC'], ['15555', 'Sprinklers'], ['16500', 'Electrical'], ['16700', 'Low Voltage / AV'], ['21100', 'Landscaping']];
/* One-time trade moves (old trade, lower case -> new trade), done with the trade codes change. */
const TRADE_MOVES = { 'energy consultant': 'Site Supervisor', 'eavestrough': 'Metalwork', 'siding': 'Exterior Carpentry', 'surveyor': 'Site Supervisor',
  'cleaning': 'Site Supervisor', 'masonry (brick)': 'Exterior Carpentry', 'masonry (stone)': 'Exterior Carpentry', 'deck carpentry': 'Exterior Carpentry',
  'parging': 'Concrete', 'closet shelving': 'Finish Carpentry' };

/* ---------------- web app entry ---------------- */
function doGet(e) {
  const t = HtmlService.createTemplateFromFile('Index');
  t.jobParam = String((e && e.parameter && e.parameter.job) || '').replace(/[^A-Za-z0-9]/g, '').slice(0, 40);
  t.appUrl = ScriptApp.getService().getUrl();
  return t.evaluate()
    .setTitle('Site Checklists')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/* ---------------- sheet helpers ---------------- */
function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('Missing tab "' + name + '" in the Google Sheet.');
  return sh;
}
function head_(sh) { return sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(h => String(h).trim()); }
function rows_(name) {
  upgrade_();
  const sh = sheet_(name);
  const values = sh.getDataRange().getValues();
  const head = values.shift().map(h => String(h).trim());
  return values
    .map((r, i) => { const o = { _row: i + 2 }; head.forEach((h, j) => o[h] = r[j]); return o; })
    .filter(o => head.some(h => o[h] !== '' && o[h] !== null));
}
function append_(name, obj) {
  const sh = sheet_(name);
  sh.appendRow(head_(sh).map(h => (obj[h] === undefined ? '' : obj[h])));
}
function yes_(v) { return String(v).trim().toUpperCase().startsWith('Y'); }
function critical_(v) { return /CRIT|HOLD|^Y/i.test(String(v || '').trim()); }
function fmt_(d) { return d instanceof Date ? Utilities.formatDate(d, TZ, 'MMM d, h:mm a') : String(d || ''); }
function day_(d) { return d instanceof Date ? Utilities.formatDate(d, TZ, 'MMM d') : ''; }
function newId_(prefix) { return prefix + Date.now().toString(36).toUpperCase() + Math.floor(Math.random() * 1296).toString(36).toUpperCase(); }
let inLock_ = false;              // the script lock is taken once per request; nested calls reuse it
function withLock_(fn) {
  if (inLock_) return fn();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  inLock_ = true;
  try { return fn(); } finally { inLock_ = false; lock.releaseLock(); }
}

/* ---------------- one-time sheet upgrade ----------------
 * Runs automatically the first time the new version is used (or run setup() from the editor).
 *  - Users: makes sure Name | PIN | Active | Admin columns exist.
 *  - Items / Job Items / Check Log: renames Walk -> Checklist and Hold -> Critical,
 *    shifts existing checklist numbers 1-4 to 3-6 (two new checklists go first),
 *    and adds the Project Startup & Site Protection items.
 * Safe to run more than once: it only upgrades a tab that still has a "Walk" column.
 */
function setup() { CacheService.getScriptCache().remove('schema'); upgrade_(); return 'Sheet is up to date.'; }
function upgrade_() {
  const cache = CacheService.getScriptCache();
  if (cache.get('schema') === '7') return;
  withLock_(() => {
    const ss = SpreadsheetApp.getActive();
    // Users
    const us = ss.getSheetByName('Users');
    if (us) {
      const h = head_(us).filter(Boolean);
      if (!h.length) us.getRange(1, 1, 1, 4).setValues([['Name', 'PIN', 'Active', 'Admin']]);
      else ['Name', 'PIN', 'Active', 'Admin'].forEach(c => { if (h.indexOf(c) < 0) { us.getRange(1, us.getLastColumn() + 1).setValue(c); h.push(c); } });
      const pinCol = head_(us).indexOf('PIN') + 1;
      if (pinCol > 0) us.getRange(1, pinCol, us.getMaxRows(), 1).setNumberFormat('@');
    }
    // Items, Job Items, Check Log
    let itemsUpgraded = false;
    ['Items', 'Job Items', 'Check Log'].forEach(name => {
      const sh = ss.getSheetByName(name);
      if (!sh) return;
      const h = head_(sh);
      const w = h.indexOf('Walk');
      if (w < 0) return;                                  // already upgraded
      sh.getRange(1, w + 1).setValue('Checklist');
      const hc = h.indexOf('Hold');
      if (hc >= 0) sh.getRange(1, hc + 1).setValue('Critical');
      const n = sh.getLastRow() - 1;
      if (n > 0) {
        const col = sh.getRange(2, w + 1, n, 1);
        col.setValues(col.getValues().map(r => [r[0] === '' || isNaN(Number(r[0])) ? r[0] : Number(r[0]) + 2]));
        if (hc >= 0) {
          const cc = sh.getRange(2, hc + 1, n, 1);
          cc.setValues(cc.getValues().map(r => [critical_(r[0]) ? 'CRITICAL' : '']));
        }
      }
      if (name === 'Items') itemsUpgraded = true;
    });
    // Startup items: added once (tracked by a script property, so a failed run simply retries).
    const props = PropertiesService.getScriptProperties();
    const sh = ss.getSheetByName('Items');
    if (sh && !props.getProperty('STARTUP_ITEMS') && head_(sh).indexOf('Checklist') >= 0) {
      const hd = head_(sh);
      const ids = sh.getLastRow() > 1 ? sh.getRange(2, hd.indexOf('ItemID') + 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
      if (ids.indexOf('S001') < 0) {
        const orderCol = hd.indexOf('Order');
        let order = 0;
        if (orderCol >= 0 && sh.getLastRow() > 1) order = Math.max(0, ...sh.getRange(2, orderCol + 1, sh.getLastRow() - 1, 1).getValues().map(r => Number(r[0]) || 0));
        ['Site signage', 'Trade rules signage posted', 'Garbage bins, brooms, garbage bags, ladders, etc. on site', 'Areas outside the scope protected for demolition and work']
          .forEach((t, i) => sh.appendRow(hd.map(c => ({ 'ItemID': 'S' + String(i + 1).padStart(3, '0'), 'Checklist': 1, 'Trade': 'Site Supervisor', 'Area': '', 'Critical': '', 'Item': t, 'Division': 'Div 01 General', 'Order': order + i + 1, 'Active': 'Y' })[c] ?? '')));
      }
      props.setProperty('STARTUP_ITEMS', 'added');
    }
    tradesUpgrade_(ss, props);
    // runs once; marked first so a failure never runs it twice (the error is kept for Claude to read back)
    if (!props.getProperty('PARK_CLEAN_V1')) { props.setProperty('PARK_CLEAN_V1', 'started');
      try { cleanupRun_(ss, 'Park', 'Mitsuko', 'Park Cleanup'); props.setProperty('PARK_CLEAN_V1', 'done'); } catch (e) { props.setProperty('PARK_CLEAN_V1', 'error: ' + e.message); }
      props.setProperty('PARK_REPORT_V1', 'done'); }
    if (!props.getProperty('RESTORE_T112_V1')) { try { restoreHidden_(ss, 'JMVICS4598Q', 't112', 'Mitsuko'); } catch (e) { } props.setProperty('RESTORE_T112_V1', 'done'); }
    if (!props.getProperty('PARK_REPORT_V1')) { try { cleanupReport_(ss, 'Park', 'Mitsuko', 'Park Cleanup'); } catch (e) { /* report is optional */ } props.setProperty('PARK_REPORT_V1', 'done'); }
    if (itemsUpgraded || props.getProperty('SETUP_TEXT') !== SETUP_TEXT_VERSION) {
      try { writeSetupTab_(ss); props.setProperty('SETUP_TEXT', SETUP_TEXT_VERSION); } catch (e) { /* instructions tab is optional */ }
    }
    cache.put('schema', '7', 21600);
  });
}
/* Trade codes: creates the Trades tab, then (once) moves items off retired trades.
 * Every move is written to Edit Log; a short report is kept for the Setup tab. */
function tradesUpgrade_(ss, props) {
  if (props.getProperty('TRADES_V1')) return;
  if (!ss.getSheetByName('Trades')) {
    const tr = ss.insertSheet('Trades');
    tr.getRange(1, 1, 1, 2).setValues([['Code', 'Trade']]);
    tr.getRange(2, 1, TRADES.length, 1).setNumberFormat('@');
    tr.getRange(2, 1, TRADES.length, 2).setValues(TRADES);
    tr.setFrozenRows(1);
  }
  const now = new Date(), moved = {}, logs = [], seen = {};
  const key = t => String(t || '').replace(/^\d+\s*-\s*/, '').trim().toLowerCase();
  ['Items', 'Job Items'].forEach(tab => {
    const sh = ss.getSheetByName(tab);
    if (!sh || sh.getLastRow() < 2) return;
    const h = head_(sh), tc = h.indexOf('Trade'), ic = h.indexOf('ItemID'), cc = h.indexOf('Checklist'), jc = h.indexOf('JobID'), ty = h.indexOf('Type');
    if (tc < 0) return;
    const v = sh.getRange(2, 1, sh.getLastRow() - 1, h.length).getValues();
    let changed = false;
    v.forEach(r => {
      if (ty >= 0 && String(r[ty]).toUpperCase() === 'HIDE') return;
      const old = String(r[tc] || '').trim(), k = key(old);
      let to = TRADE_MOVES[k];
      if (!to && tab === 'Items' && /^S00[1-4]$/.test(String(r[ic])) && Number(r[cc]) === 1 && k === 'site supervisor') to = 'Mobilize & Startup Site Protection';
      if (!to) { if (old) seen[k] = old; return; }
      r[tc] = to; changed = true; seen[key(to)] = to;
      const m = old + ' → ' + to; moved[m] = (moved[m] || 0) + 1;
      logs.push({ 'Timestamp': now, 'Name': 'Setup (trade codes)', 'JobID': tab === 'Items' ? '(all jobs)' : String(r[jc]), 'ItemID': String(r[ic]), 'Field': 'Trade', 'Old Value': old, 'New Value': to });
    });
    if (changed) sh.getRange(2, tc + 1, v.length, 1).setValues(v.map(r => [r[tc]]));
  });
  if (logs.length) { ensureEditLog_(); logs.forEach(l => append_('Edit Log', l)); }
  const coded = {};
  const tr = ss.getSheetByName('Trades');
  if (tr && tr.getLastRow() > 1) tr.getRange(2, 1, tr.getLastRow() - 1, 2).getValues().forEach(r => { if (String(r[1]).trim()) coded[key(r[1])] = 1; });
  const uncoded = Object.keys(seen).filter(k => !coded[k]).map(k => seen[k]).sort();
  props.setProperty('TRADES_V1_REPORT', JSON.stringify({ at: fmt_(now), moved: moved, uncoded: uncoded }));
  props.setProperty('TRADES_V1', 'done');
}
/* ---------------- test clean-up report (read-only) ----------------
 * Lists everything one person did on one job, so the owner can review it before anything is cleared.
 * Writes (or rewrites) the tab `tabName`. Changes nothing else.
 */
function cleanupReport_(ss, jobName, person, tabName, preface) {
  const lc = s => String(s || '').trim().toLowerCase();
  const me = lc(person);
  // read directly (rows_ would re-enter upgrade_, which is what calls this)
  const get = name => { const sh = ss.getSheetByName(name); if (!sh || sh.getLastRow() < 2) return [];
    const v = sh.getDataRange().getValues(), hd = v.shift().map(x => String(x).trim());
    return v.map((r, i) => { const o = { _row: i + 2 }; hd.forEach((c, j) => o[c] = r[j]); return o; }).filter(o => hd.some(c => o[c] !== '' && o[c] !== null)); };
  const jobsAll = get('Jobs');
  let jobs = jobsAll.filter(j => lc(j['Job Name']) === lc(jobName));
  if (!jobs.length) jobs = jobsAll.filter(j => lc(j['Job Name']).indexOf(lc(jobName)) >= 0);
  const ids = new Set(jobs.map(j => String(j['JobID'])));
  const jobLabel = id => { const j = jobsAll.find(x => String(x['JobID']) === id); return j ? String(j['Job Name']) : id; };
  const master = {}; get('Items').forEach(r => { master[String(r['ItemID'])] = r; });
  const jobItems = get('Job Items');
  const extra = {}; jobItems.forEach(r => { if (String(r['Type']).trim().toUpperCase() === 'ITEM') extra[String(r['JobID']) + '|' + String(r['ItemID'])] = r; });
  const task = (jobId, itemId) => { const r = extra[jobId + '|' + itemId] || master[itemId]; return r ? [Number(r['Checklist']) || '', String(r['Trade'] || ''), String(r['Item'] || '')] : ['', '', '(task no longer exists: ' + itemId + ')']; };
  const when = d => d instanceof Date ? Utilities.formatDate(d, TZ, 'MMM d, yyyy h:mm a') : String(d || '');
  const checks = get('Check Log'), notes = get('Notes'), edits = get('Edit Log');
  const out = [['PARK CLEAN-UP – SUMMARY ONLY (nothing has been changed or deleted)', '', '', '', '', '', ''],
    ['Person: ' + person + '   ·   Job(s): ' + (jobs.length ? jobs.map(j => String(j['Job Name']) + ' (' + j['JobID'] + ')').join(', ') : 'NONE FOUND named "' + jobName + '"') + '   ·   Made: ' + when(new Date()), '', '', '', '', '', ''],
    ['', '', '', '', '', '', '']];
  if (preface) { out[0][0] = 'PARK CLEAN-UP – DONE. What is left below was kept on purpose, or is master-list only.'; preface.forEach(l => out.push([l, '', '', '', '', '', ''])); out.push(['', '', '', '', '', '', '']); }
  const section = (title, head, rows) => {
    out.push([title + ' – ' + rows.length, '', '', '', '', '', '']);
    if (rows.length) { out.push(head); rows.forEach(r => out.push(r)); } else out.push(['(none)', '', '', '', '', '', '']);
    out.push(['', '', '', '', '', '', '']);
  };
  const H = ['Job', 'Checklist', 'Trade', 'Task', 'What', 'When', 'Note'];
  // 1. ticks and unticks
  const mine = checks.filter(r => ids.has(String(r['JobID'])) && lc(r['Name']) === me);
  section('1. YOUR TICKS AND UNTICKS (will be deleted from Check Log)', H, mine.map(r => {
    const j = String(r['JobID']), it = String(r['ItemID']);
    const others = checks.filter(x => String(x['JobID']) === j && String(x['ItemID']) === it && lc(x['Name']) !== me);
    const last = others[others.length - 1];
    const after = !last || String(last['Status']) === 'Cleared' ? 'Afterwards: not done' : 'Afterwards: ' + last['Status'] + ' by ' + last['Name'] + ' (' + when(last['Timestamp']) + ')';
    return [jobLabel(j)].concat(task(j, it), [String(r['Status']), when(r['Timestamp']), after]);
  }));
  // 2. notes
  section('2. YOUR NOTES (will be deleted)', H, notes.filter(r => ids.has(String(r['JobID'])) && lc(r['Name']) === me)
    .map(r => [jobLabel(String(r['JobID']))].concat(task(String(r['JobID']), String(r['ItemID'])), ['Note: ' + String(r['Note']), when(r['Timestamp']), ''])));
  // 3. tasks added
  section('3. TASKS YOU ADDED TO THIS JOB (will be removed)', H, jobItems.filter(r => ids.has(String(r['JobID'])) && String(r['Type']).trim().toUpperCase() === 'ITEM' && lc(r['Added By']) === me)
    .map(r => {
      const j = String(r['JobID']), it = String(r['ItemID']);
      const by = {};
      checks.filter(x => String(x['JobID']) === j && String(x['ItemID']) === it && lc(x['Name']) !== me).forEach(x => by[String(x['Name'])] = (by[String(x['Name'])] || '') + 'ticked ');
      notes.filter(x => String(x['JobID']) === j && String(x['ItemID']) === it && lc(x['Name']) !== me).forEach(x => by[String(x['Name'])] = (by[String(x['Name'])] || '') + 'noted ');
      const flag = Object.keys(by).map(n => n + ' ' + by[n].trim().replace(' ', ' + ')).join('; ');
      return [jobLabel(j), Number(r['Checklist']) || '', String(r['Trade'] || ''), String(r['Item'] || ''), yes_(r['Removed']) ? 'Added (already removed)' : 'Added', when(r['Added At']), flag ? 'CHECK: ' + flag + ' this task' : ''];
    }));
  // 4. tasks removed
  const hides = jobItems.filter(r => ids.has(String(r['JobID'])) && String(r['Type']).trim().toUpperCase() === 'HIDE' && lc(r['Added By']) === me)
    .map(r => [jobLabel(String(r['JobID']))].concat(task(String(r['JobID']), String(r['ItemID'])), ['Removed from this job', when(r['Added At']), 'Will come back']));
  const remLog = edits.filter(r => ids.has(String(r['JobID'])) && lc(r['Name']) === me && String(r['Field']) === 'Removed from job' && !master[String(r['ItemID'])])
    .map(r => [jobLabel(String(r['JobID']))].concat(task(String(r['JobID']), String(r['ItemID'])), ['Removed from this job', when(r['Timestamp']), 'Will come back']));
  section('4. TASKS YOU REMOVED FROM THIS JOB (will come back)', H, hides.concat(remLog));
  // 5. edits on this job
  section('5. YOUR EDITS ON THIS JOB (will be undone)', H, edits.filter(r => ids.has(String(r['JobID'])) && lc(r['Name']) === me && String(r['Field']) !== 'Removed from job')
    .map(r => [jobLabel(String(r['JobID']))].concat(task(String(r['JobID']), String(r['ItemID'])), [String(r['Field']) + ': "' + r['Old Value'] + '" → "' + r['New Value'] + '"', when(r['Timestamp']), ''])));
  // 6. master-list edits (changed every job)
  section('6. YOUR MASTER-LIST EDITS – these changed EVERY job (listed only, not undone unless you ask)', ['Job', 'Checklist', 'Trade', 'Task (as it reads now)', 'What', 'When', 'Note'],
    edits.filter(r => String(r['JobID']) === '(all jobs)' && lc(r['Name']) === me)
      .map(r => { const m = master[String(r['ItemID'])]; return ['(all jobs)', m ? Number(m['Checklist']) || '' : '', m ? String(m['Trade'] || '') : '', m ? String(m['Item'] || '') : '(task no longer exists)', String(r['Field']) + ': "' + r['Old Value'] + '" → "' + r['New Value'] + '"', when(r['Timestamp']), 'ItemID ' + r['ItemID']]; }));
  // 7. master-list tasks added through the app (who added isn't recorded)
  section('7. MASTER-LIST TASKS ADDED THROUGH THE APP – who added them is not recorded; tell Claude which are yours', ['Job', 'Checklist', 'Trade', 'Task', 'What', 'When', 'Note'],
    Object.keys(master).filter(id => /^M[0-9A-Z]{9,10}$/.test(id)).map(id => {
      const m = master[id], t = parseInt(id.slice(1, 9), 36);
      return ['(all jobs)', Number(m['Checklist']) || '', String(m['Trade'] || ''), String(m['Item'] || ''), 'Added to master list', isNaN(t) ? '' : when(new Date(t)), 'ItemID ' + id];
    }).sort((a, b) => String(a[5]).localeCompare(String(b[5]))));
  let sh = ss.getSheetByName(tabName);
  if (!sh) sh = ss.insertSheet(tabName); else sh.clear();
  sh.getRange(1, 1, out.length, 7).setValues(out.map(r => r.map(c => c === undefined || c === null ? '' : c)));
  try { sh.getRange(1, 1).setFontWeight('bold'); sh.setFrozenRows(2); } catch (e) { }
  return out.length;
}
/* ---------------- test clean-up (owner approved 2026-10-23) ----------------
 * Removes one person's changes on one job, keeping everyone else's work:
 *  - deletes their Check Log, Notes and Edit Log rows for the job (owner chose to remove them from the audit trail)
 *  - deletes tasks they added (kept if someone else ticked or noted them), their HIDE rows (tasks come back),
 *    and their EDIT rows (master wording comes back) unless someone else also edited that task
 *  - puts back job-only tasks they removed, and undoes their edits to other people's job-only tasks
 *    (only where the value is still what they set)
 *  - copies every deleted row to a backup tab first; master-list items are never touched
 * Then rewrites the summary tab with the result. Reads sheets directly (called from upgrade_).
 */
function cleanupRun_(ss, jobName, person, tabName) {
  const lc = s => String(s || '').trim().toLowerCase();
  const me = lc(person);
  const read = name => { const sh = ss.getSheetByName(name); if (!sh || sh.getLastRow() < 2) return { sh: sh, head: sh ? head_(sh) : [], rows: [] };
    const v = sh.getDataRange().getValues(), hd = v.shift().map(x => String(x).trim());
    return { sh: sh, head: hd, rows: v.map((r, i) => { const o = { _row: i + 2, _vals: r }; hd.forEach((c, j) => o[c] = r[j]); return o; }).filter(o => hd.some(c => o[c] !== '' && o[c] !== null)) }; };
  const J = read('Jobs');
  let jobs = J.rows.filter(j => lc(j['Job Name']) === lc(jobName));
  if (!jobs.length) jobs = J.rows.filter(j => lc(j['Job Name']).indexOf(lc(jobName)) >= 0);
  if (!jobs.length) return null;
  const ids = new Set(jobs.map(j => String(j['JobID'])));
  const inJob = r => ids.has(String(r['JobID']));
  const C = read('Check Log'), N = read('Notes'), X = read('Job Items'), E = read('Edit Log');
  const type = r => String(r['Type']).trim().toUpperCase();
  const key = r => String(r['JobID']) + '|' + String(r['ItemID']);
  const othersTouched = {};
  C.rows.concat(N.rows).forEach(r => { if (inJob(r) && lc(r['Name']) !== me) othersTouched[key(r)] = (othersTouched[key(r)] || []).concat(String(r['Name'])); });
  const othersEdited = {};
  E.rows.forEach(r => { if (inJob(r) && lc(r['Name']) !== me) othersEdited[key(r)] = true; });
  const del = { 'Check Log': [], 'Notes': [], 'Job Items': [], 'Edit Log': [] }, kept = [], counts = { ticks: 0, notes: 0, added: 0, unhidden: 0, unedited: 0, restored: 0, reverted: 0, logs: 0 };
  C.rows.forEach(r => { if (inJob(r) && lc(r['Name']) === me) { del['Check Log'].push(r); counts.ticks++; } });
  N.rows.forEach(r => { if (inJob(r) && lc(r['Name']) === me) { del['Notes'].push(r); counts.notes++; } });
  const myTasks = new Set();
  X.rows.forEach(r => {
    if (!inJob(r) || lc(r['Added By']) !== me) return;
    const t = type(r);
    if (t === 'ITEM') {
      myTasks.add(key(r));
      if (othersTouched[key(r)]) kept.push('Kept your task "' + r['Item'] + '" – ' + [...new Set(othersTouched[key(r)])].join(', ') + ' ticked or noted it.');
      else { del['Job Items'].push(r); counts.added++; }
    } else if (t === 'HIDE') { del['Job Items'].push(r); counts.unhidden++; }
    else if (t === 'EDIT') {
      if (othersEdited[key(r)]) kept.push('Kept the changed wording "' + r['Item'] + '" – someone else also edited this task on this job.');
      else { del['Job Items'].push(r); counts.unedited++; }
    }
  });
  // put back job-only tasks (someone else's) that this person removed, and undo their edits to those tasks
  const xHead = X.head, setCell = (r, field, v) => { const c = xHead.indexOf(field); if (c >= 0) X.sh.getRange(r._row, c + 1).setValue(v); };
  const mineE = E.rows.filter(r => inJob(r) && lc(r['Name']) === me);
  mineE.slice().reverse().forEach(r => {            // newest first, so several edits unwind in order
    const task = X.rows.find(x => key(x) === key(r) && type(x) === 'ITEM');
    if (!task || myTasks.has(key(task))) return;     // master tasks are handled by EDIT/HIDE rows; own tasks are deleted
    const field = String(r['Field']);
    if (field === 'Removed from job') { if (yes_(task['Removed'])) { setCell(task, 'Removed', ''); task['Removed'] = ''; counts.restored++; } return; }
    const col = field === 'Wording' ? 'Item' : field;
    const now = col === 'Critical' ? (critical_(task[col]) ? 'CRITICAL' : '') : String(task[col] === undefined ? '' : task[col]).trim();
    if (now === String(r['New Value'])) { setCell(task, col, r['Old Value']); task[col] = r['Old Value']; counts.reverted++; }
    else kept.push('Did not undo your ' + field + ' change on "' + task['Item'] + '" – it has been changed again since.');
  });
  mineE.forEach(r => { del['Edit Log'].push(r); counts.logs++; });
  // backup, then delete (bottom up, so row numbers stay right)
  const backupName = tabName + ' – Removed rows';
  let bk = ss.getSheetByName(backupName); if (!bk) bk = ss.insertSheet(backupName);
  const width = Math.max(1, ...Object.keys(del).map(t => (read(t).head || []).length));
  const bkRows = [['Copied from tab', 'Row values (in that tab\'s column order) – safe to delete this tab once you are happy'].concat(Array(width - 1).fill(''))];
  Object.keys(del).forEach(t => del[t].forEach(r => bkRows.push([t].concat(r._vals, Array(width - r._vals.length).fill('')))));
  if (bkRows.length > 1) bk.getRange(bk.getLastRow() + 1, 1, bkRows.length, width + 1).setValues(bkRows.map(r => r.slice(0, width + 1)));
  Object.keys(del).forEach(t => {
    const sh = ss.getSheetByName(t); if (!sh) return;
    const rowsDesc = del[t].map(r => r._row).sort((a, b) => b - a);
    let i = 0;
    while (i < rowsDesc.length) {                    // merge neighbouring rows into one delete
      let end = rowsDesc[i], start = end; i++;
      while (i < rowsDesc.length && rowsDesc[i] === start - 1) { start--; i++; }
      sh.deleteRows(start, end - start + 1);
    }
  });
  const pre = ['Cleared ' + Utilities.formatDate(new Date(), TZ, 'MMM d, yyyy h:mm a') + ' for ' + person + ' on ' + jobs.map(j => String(j['Job Name'])).join(', ') + '. Nothing anyone else did was changed.',
    'Removed: ' + counts.ticks + ' ticks/unticks · ' + counts.notes + ' notes · ' + counts.added + ' tasks you added · ' + counts.logs + ' Edit Log rows',
    'Undone: ' + counts.unhidden + ' removed tasks are back · ' + counts.unedited + ' changed master tasks back to master wording · ' + counts.restored + ' job-only tasks put back · ' + counts.reverted + ' job-only edits reversed',
    'Every deleted row was copied to the "' + backupName + '" tab first.'].concat(kept.length ? kept : ['Nothing of yours had to be kept.']);
  cleanupReport_(ss, jobName, person, tabName, pre);
  return counts;
}
/* One-time fix (owner request 2026-10-23): put back a master task removed from one job by mistake.
 * Deletes that person's HIDE row for the task on that job and logs the restore. Reads the sheet directly. */
function restoreHidden_(ss, jobId, itemId, person) {
  const sh = ss.getSheetByName('Job Items');
  if (!sh || sh.getLastRow() < 2) return 0;
  const v = sh.getDataRange().getValues(), hd = v[0].map(x => String(x).trim());
  const c = n => hd.indexOf(n);
  let n = 0;
  for (let i = v.length - 1; i >= 1; i--) {
    const r = v[i];
    if (String(r[c('JobID')]) === jobId && String(r[c('ItemID')]) === itemId && String(r[c('Type')]).trim().toUpperCase() === 'HIDE' && String(r[c('Added By')]).trim().toLowerCase() === person.toLowerCase()) { sh.deleteRow(i + 1); n++; }
  }
  if (n) { ensureEditLog_(); append_('Edit Log', { 'Timestamp': new Date(), 'Name': 'Claude (asked by ' + person + ')', 'JobID': jobId, 'ItemID': itemId, 'Field': 'Restored to job', 'Old Value': 'Removed', 'New Value': 'Back on this job' }); }
  return n;
}
function writeSetupTab_(ss) {
  const s = ss.getSheetByName('Setup');
  if (!s) return;
  const lines = ['HEADLAND SITE CHECKLISTS – BACKEND', '',
    'This Google Sheet is the backend for the Site Checklists app. Keep it private. The app code lives in GitHub (mitsuko-headland/site-checklists) and is published automatically.', '',
    'DO NOT RENAME',
    'Tab names (Users, Items, Jobs, Job Items, Check Log, Edit Log, Notes) or the column headings in row 1. The app finds them by name. Renaming the file, reordering columns, formatting, and adding extra columns to the right are all fine.', '',
    'WHO CAN SIGN IN',
    'People on the Users tab, with their Name and PIN. Active must start with Y to allow sign-in (blank or N blocks, immediately). Admin starting with Y lets that person add items to the master list from the app. Everyone can edit or remove any task in the app, but only for the job they are on; every change is in Edit Log. To change a master-list task on every job, edit it here on the Items tab. A name is locked for 15 minutes after 5 wrong PINs. Sign-ins last 30 days on each phone.', '',
    'CHECKLISTS (the Checklist column on Items)',
    '1 Project Startup & Site Protection · 2 Demolition · 3 Pre-Rough In · 4 Rough In & Pre-Drywall · 5 Drywall · 6 Finishing', '',
    'TABS',
    'Users – the authorised list (Name | PIN | Active | Admin).',
    'Items – the master checklist every job uses. Critical = CRITICAL marks a stop-point item. Set Active = N to retire an item. Do not change ItemID.',
    'Jobs – created from the app. Set Active = N to hide a finished job.',
    'Job Items – tasks added to, removed from (HIDE) or changed for (EDIT) a single job from the app. Do not edit by hand.',
    'Check Log – every tick and untick, with name and date/time. Never edit or sort; it is the audit trail.',
    'Edit Log – every change and removal made to a task in the app: which job, old value, new value, who and when. Created automatically on the first edit.',
    'Notes – notes people add to a task on one job, with name and date/time. Created automatically on the first note. Never edit or sort; the app shows them in this order.',
    'Trades – the code shown on each trade heading, and the order trades are sorted in (lowest code first). Change a code here any time. A trade name must match the Trade column on Items exactly (capitals don\'t matter). Trades with no code show at the end.'];
  let rep = null;
  try { rep = JSON.parse(PropertiesService.getScriptProperties().getProperty('TRADES_V1_REPORT') || 'null'); } catch (e) { rep = null; }
  if (rep) {
    lines.push('', 'TRADE CODES UPDATE (' + rep.at + ')');
    const m = Object.keys(rep.moved || {});
    if (m.length) m.forEach(k => lines.push('Moved ' + rep.moved[k] + ' item' + (rep.moved[k] > 1 ? 's' : '') + ': ' + k)); else lines.push('No items needed moving.');
    lines.push(rep.uncoded && rep.uncoded.length ? 'Trades with no code (shown at the end of their checklist): ' + rep.uncoded.join(', ') : 'Every trade in use has a code.');
  }
  s.clear();
  s.getRange(1, 1, lines.length, 1).setValues(lines.map(l => [l])).setWrap(true);
  s.getRange(1, 1).setFontSize(14).setFontWeight('bold');
}

/* ---------------- sign-in (name + PIN) ---------------- */
function secret_() {
  const p = PropertiesService.getScriptProperties();
  let s = p.getProperty('SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('SECRET', s); }
  return s;
}
function sig_(payload) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(payload, secret_())); }
function findUser_(name) {
  const key = String(name || '').trim().toLowerCase();
  if (!key) return null;
  const u = rows_('Users').find(r => String(r['Name']).trim().toLowerCase() === key);
  return u ? { name: String(u['Name']).trim(), pin: String(u['PIN'] === undefined ? '' : u['PIN']).trim(), active: yes_(u['Active']), admin: yes_(u['Admin']) } : null;
}
function user_(token) {
  try {
    const [p, s] = String(token || '').split('.');
    const payload = Utilities.newBlob(Utilities.base64DecodeWebSafe(p)).getDataAsString();
    if (sig_(payload) !== s) throw 0;
    const [name, exp] = payload.split('|');
    if (Date.now() > Number(exp)) throw 0;
    const u = findUser_(name);
    if (!u || !u.active) throw 0;
    return u;
  } catch (e) { throw new Error('AUTH'); }
}

/** Called from the sign-in screen. */
function login(name, pin) {
  const cache = CacheService.getScriptCache();
  const failKey = 'fail:' + String(name || '').trim().toLowerCase();
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= MAX_FAILS) throw new Error('Too many wrong PINs. Try again in ' + LOCK_MINUTES + ' minutes.');
  const u = findUser_(name);
  const typed = String(pin || '').trim();
  const pinOk = u && u.pin && (u.pin === typed || (/^\d+$/.test(u.pin) && /^\d+$/.test(typed) && Number(u.pin) === Number(typed)));
  if (!u || !u.active || !pinOk) {
    cache.put(failKey, String(fails + 1), LOCK_MINUTES * 60);
    throw new Error('Name or PIN not recognised. Check with the office if you need access.');
  }
  cache.remove(failKey);
  const payload = u.name + '|' + (Date.now() + SESSION_DAYS * 86400000);
  return { token: Utilities.base64EncodeWebSafe(payload) + '.' + sig_(payload), name: u.name, admin: u.admin };
}

/** Confirms a saved session on app start. */
function whoAmI(token) { const u = user_(token); return { name: u.name, admin: u.admin }; }

/** App start-up in ONE round trip (each call to Google costs a second or more):
 *  who is signed in, the job list and, if asked, one job's checklists. */
function startup(token, jobId) {
  const u = user_(token);
  const jobs = jobs_();
  let job = null;
  jobId = String(jobId || '');
  if (jobId && jobs.some(j => j.id === jobId)) { try { job = job_(jobId); } catch (e) { job = null; } }
  return { me: { name: u.name, admin: u.admin }, jobs: jobs, job: job };
}

/* ---------------- jobs ---------------- */
function getJobs(token) { user_(token); return jobs_(); }
function jobs_() {
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

/** Everything for one job: all checklists' items plus the latest status of each. */
function getJob(token, jobId) { user_(token); return job_(jobId); }
function job_(jobId) {
  jobId = String(jobId);
  const job = rows_('Jobs').find(j => String(j['JobID']) === jobId);
  if (!job) throw new Error('That job no longer exists.');
  const extras = rows_('Job Items').filter(r => String(r['JobID']) === jobId && !yes_(r['Removed']));
  const type = r => String(r['Type']).trim().toUpperCase();
  const hidden = new Set(extras.filter(r => type(r) === 'HIDE').map(r => String(r['ItemID'])));
  const edits = {};                            // master task changed for this job only (latest row wins)
  extras.filter(r => type(r) === 'EDIT').forEach(r => { edits[String(r['ItemID'])] = r; });
  const items = [];
  rows_('Items').forEach(r => {
    const id = String(r['ItemID']);
    if (!id || hidden.has(id) || (String(r['Active']).trim() && !yes_(r['Active']))) return;
    const it = { id: id, src: 'm', list: Number(r['Checklist']), trade: String(r['Trade'] || 'General'), area: String(r['Area'] || ''), critical: critical_(r['Critical']), text: String(r['Item']), order: Number(r['Order']) || 0 };
    const e = edits[id];
    if (e) { it.text = String(e['Item']); it.trade = String(e['Trade'] || 'General'); it.area = String(e['Area'] || ''); it.critical = critical_(e['Critical']); it.edited = true; }
    items.push(it);
  });
  // Tasks added to this job go at the top of their trade: newest batch first, each batch in the order it was typed.
  const at = r => (r['Added At'] instanceof Date ? r['Added At'].getTime() : 0);
  extras.filter(r => type(r) === 'ITEM' || !['HIDE', 'EDIT', 'LIST'].includes(type(r)))
    .sort((a, b) => (at(b) - at(a)) || (a._row - b._row))
    .forEach((r, k) => {
      // job-only tasks carry who added them and the day (master-list tasks never do)
      items.push({ id: String(r['ItemID']), src: 'j', list: Number(r['Checklist']), trade: String(r['Trade'] || 'General'), area: String(r['Area'] || ''), critical: critical_(r['Critical']), text: String(r['Item']), order: -1e6 + k, addedBy: String(r['Added By'] || '').trim(), addedOn: day_(r['Added At']) });
    });
  const checks = {};
  rows_('Check Log').forEach(r => {           // later rows win, so this is the current status
    if (String(r['JobID']) !== jobId) return;
    const st = String(r['Status']);
    checks[String(r['ItemID'])] = st === 'Cleared' ? null : { status: st, name: String(r['Name']), at: fmt_(r['Timestamp']) };
  });
  const notes = {};
  if (SpreadsheetApp.getActive().getSheetByName('Notes')) {
    rows_('Notes').forEach(r => {             // oldest first
      if (String(r['JobID']) !== jobId || !String(r['Note']).trim()) return;
      const id = String(r['ItemID']);
      (notes[id] = notes[id] || []).push({ text: String(r['Note']), name: String(r['Name']), at: fmt_(r['Timestamp']) });
    });
  }
  const trades = [];
  if (SpreadsheetApp.getActive().getSheetByName('Trades')) {
    rows_('Trades').forEach(r => { const n = String(r['Trade'] || '').trim(); if (n) trades.push({ name: n, code: String(r['Code'] === undefined ? '' : r['Code']).trim() }); });
  }
  const hiddenLists = [...new Set(extras.filter(r => type(r) === 'LIST').map(r => Number(r['Checklist'])))].filter(n => n >= 1 && n <= CHECKLISTS);
  return { job: { id: jobId, name: String(job['Job Name']), address: String(job['Address'] || '') }, hiddenLists: hiddenLists, items: items, checks: checks, notes: notes, trades: trades };
}

/** Anyone signed in: add a note to a task on this job. Notes are a running log (one row each, never changed). */
function addNote(token, jobId, itemId, list, text) {
  const u = user_(token);
  text = String(text || '').trim();
  if (!text) throw new Error('Type a note first.');
  if (text.length > NOTE_MAX) throw new Error('That note is too long (' + NOTE_MAX + ' characters at most).');
  // Adding a row is safe on its own, so notes don't queue behind other people's requests;
  // the lock is only needed the first time, to create the Notes tab.
  if (!SpreadsheetApp.getActive().getSheetByName('Notes')) withLock_(ensureNotes_);
  const now = new Date();
  append_('Notes', { 'JobID': String(jobId), 'ItemID': String(itemId), 'Checklist': Number(list) || '', 'Note': text, 'Name': u.name, 'Timestamp': now });
  return { text: text, name: u.name, at: fmt_(now) };
}
function ensureNotes_() {
  const ss = SpreadsheetApp.getActive();
  if (ss.getSheetByName('Notes')) return;
  const sh = ss.insertSheet('Notes');
  sh.appendRow(['JobID', 'ItemID', 'Checklist', 'Note', 'Name', 'Timestamp']);
  sh.setFrozenRows(1);
}

/** status: 'Done', 'N/A' or 'Cleared'. Every change is a new row in Check Log. */
function setCheck(token, jobId, itemId, list, status) {
  const u = user_(token);
  if (['Done', 'N/A', 'Cleared'].indexOf(status) < 0) throw new Error('Unknown status.');
  // One appended row per tick: safe without the lock, so ticks never wait behind other requests.
  const now = new Date();
  append_('Check Log', { 'JobID': String(jobId), 'ItemID': String(itemId), 'Checklist': Number(list), 'Status': status, 'Name': u.name, 'Timestamp': now });
  return status === 'Cleared' ? null : { status: status, name: u.name, at: fmt_(now) };
}

/** rows: [{list, trade, area, critical, text}], scope: 'job' or 'all' (master list – Admin only). */
function addItems(token, jobId, scope, rows) {
  const u = user_(token);
  rows = (rows || []).filter(r => r && String(r.text || '').trim() && Number(r.list) >= 1 && Number(r.list) <= CHECKLISTS);
  if (!rows.length) throw new Error('Nothing to add.');
  if (scope === 'all' && !u.admin) throw new Error('Only admins can add to the master list. Choose "This job only" or ask the office.');
  return withLock_(() => {
    if (scope === 'all') {
      let order = Math.max(0, ...rows_('Items').map(r => Number(r['Order']) || 0));
      rows.forEach(r => append_('Items', { 'ItemID': newId_('M'), 'Checklist': Number(r.list), 'Trade': String(r.trade || 'General').trim(), 'Area': String(r.area || '').trim(), 'Critical': r.critical ? 'CRITICAL' : '', 'Item': String(r.text).trim(), 'Division': '', 'Order': ++order, 'Active': 'Y' }));
    } else {
      const now = new Date();       // one time for the whole batch, so the batch keeps its typed order
      rows.forEach(r => append_('Job Items', { 'JobID': String(jobId), 'ItemID': newId_('X'), 'Type': 'Item', 'Checklist': Number(r.list), 'Trade': String(r.trade || 'General').trim(), 'Area': String(r.area || '').trim(), 'Critical': r.critical ? 'CRITICAL' : '', 'Item': String(r.text).trim(), 'Added By': u.name, 'Added At': now, 'Removed': '' }));
    }
    return rows.length;
  });
}

/** Anyone signed in: change a task's wording, trade, area or critical flag ON THIS JOB ONLY.
 *  Job-only tasks are changed in place; a master-list task gets an EDIT row in Job Items for this job,
 *  so other jobs keep the master wording (change the Items tab to change it everywhere).
 *  Every change is written to Edit Log with who and when. */
function editItem(token, jobId, itemId, src, changes) {
  const u = user_(token);
  changes = changes || {};
  const text = String(changes.text || '').trim();
  if (!text) throw new Error('The item wording cannot be empty.');
  jobId = String(jobId); itemId = String(itemId);
  const next = { 'Item': text, 'Trade': String(changes.trade || 'General').trim(), 'Area': String(changes.area || '').trim(), 'Critical': changes.critical ? 'CRITICAL' : '' };
  const val = (r, f) => f === 'Critical' ? (critical_(r[f]) ? 'CRITICAL' : '') : String(r[f] === undefined ? '' : r[f]).trim();
  return withLock_(() => {
    const sh = sheet_('Job Items');
    const head = head_(sh);
    const jobRows = rows_('Job Items').filter(x => String(x['JobID']) === jobId && !yes_(x['Removed']));
    let row = null, current = null;
    if (src === 'm') {
      const m = rows_('Items').find(x => String(x['ItemID']) === itemId);
      if (!m) throw new Error('That item no longer exists.');
      const ed = jobRows.filter(x => String(x['ItemID']) === itemId && String(x['Type']).trim().toUpperCase() === 'EDIT');
      row = ed.length ? ed[ed.length - 1] : null;
      current = row || m;
    } else {
      row = jobRows.find(x => String(x['ItemID']) === itemId && String(x['Type']).trim().toUpperCase() !== 'EDIT');
      if (!row) throw new Error('That item no longer exists.');
      current = row;
    }
    const fields = Object.keys(next).filter(f => val(current, f) !== next[f]);
    if (!fields.length) return 0;
    const now = new Date();
    if (src === 'm' && !row) {
      const m = rows_('Items').find(x => String(x['ItemID']) === itemId);
      append_('Job Items', { 'JobID': jobId, 'ItemID': itemId, 'Type': 'EDIT', 'Checklist': Number(m['Checklist']), 'Trade': next['Trade'], 'Area': next['Area'], 'Critical': next['Critical'], 'Item': next['Item'], 'Added By': u.name, 'Added At': now, 'Removed': '' });
    } else {
      fields.forEach(f => { const col = head.indexOf(f); if (col >= 0) sh.getRange(row._row, col + 1).setValue(next[f]); });
    }
    ensureEditLog_();
    fields.forEach(f => append_('Edit Log', { 'Timestamp': now, 'Name': u.name, 'JobID': jobId, 'ItemID': itemId, 'Field': f === 'Item' ? 'Wording' : f, 'Old Value': val(current, f), 'New Value': next[f] }));
    return fields.length;
  });
}
function ensureEditLog_() {
  const ss = SpreadsheetApp.getActive();
  if (ss.getSheetByName('Edit Log')) return;
  const sh = ss.insertSheet('Edit Log');
  sh.appendRow(['Timestamp', 'Name', 'JobID', 'ItemID', 'Field', 'Old Value', 'New Value']);
  sh.setFrozenRows(1);
}

/** Admin only: hide (or show again) a whole checklist on ONE job. Non-admins don't see a hidden checklist;
 *  admins see it greyed out. Nothing is deleted. Logged in Edit Log. */
function setListHidden(token, jobId, list, hidden) {
  const u = user_(token);
  if (!u.admin) throw new Error('Only admins can hide or show checklists.');
  list = Number(list); jobId = String(jobId);
  if (!(list >= 1 && list <= CHECKLISTS)) throw new Error('Unknown checklist.');
  return withLock_(() => {
    const sh = sheet_('Job Items'), head = head_(sh);
    const rows = rows_('Job Items').filter(r => String(r['JobID']) === jobId && String(r['Type']).trim().toUpperCase() === 'LIST' && Number(r['Checklist']) === list && !yes_(r['Removed']));
    const now = new Date();
    if (hidden && !rows.length) append_('Job Items', { 'JobID': jobId, 'ItemID': 'L' + list, 'Type': 'LIST', 'Checklist': list, 'Added By': u.name, 'Added At': now, 'Removed': '' });
    else if (!hidden && rows.length) rows.forEach(r => sh.getRange(r._row, head.indexOf('Removed') + 1).setValue('Y'));
    else return !!hidden;                     // already in that state – nothing to log
    ensureEditLog_();
    append_('Edit Log', { 'Timestamp': now, 'Name': u.name, 'JobID': jobId, 'ItemID': 'Checklist ' + list, 'Field': 'Checklist visibility', 'Old Value': hidden ? 'Shown' : 'Hidden', 'New Value': hidden ? 'Hidden' : 'Shown' });
    return !!hidden;
  });
}

/** Removes an item from this job only. Master-list items are hidden for this job, not deleted. */
function removeItem(token, jobId, itemId, src, list) {
  const u = user_(token);
  return withLock_(() => {
    const now = new Date();
    let text = '';
    if (src === 'm') { const m = rows_('Items').find(x => String(x['ItemID']) === String(itemId)); text = m ? String(m['Item']) : ''; }
    else { const j = rows_('Job Items').find(x => String(x['JobID']) === String(jobId) && String(x['ItemID']) === String(itemId)); text = j ? String(j['Item']) : ''; }
    ensureEditLog_();
    append_('Edit Log', { 'Timestamp': now, 'Name': u.name, 'JobID': String(jobId), 'ItemID': String(itemId), 'Field': 'Removed from job', 'Old Value': text, 'New Value': '' });
    if (src === 'm') {
      append_('Job Items', { 'JobID': String(jobId), 'ItemID': String(itemId), 'Type': 'HIDE', 'Checklist': Number(list), 'Added By': u.name, 'Added At': now, 'Removed': '' });
    } else {
      const sh = sheet_('Job Items');
      const head = head_(sh);
      const r = rows_('Job Items').find(x => String(x['JobID']) === String(jobId) && String(x['ItemID']) === String(itemId) && String(x['Type']).trim().toUpperCase() !== 'EDIT');
      if (r) sh.getRange(r._row, head.indexOf('Removed') + 1).setValue('Y');
    }
    return true;
  });
}
