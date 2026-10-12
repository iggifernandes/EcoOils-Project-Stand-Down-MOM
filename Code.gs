/**
 * EcoOils Mentari — one Google Apps Script backend  (v2.0.0)
 *
 * Serves BOTH apps from one Google Sheet:
 *   - Incident Management app (IMS):  list?project=, save {project, incidents}, kvGet, kvSet, exportXlsx
 *   - Stand Down MOM app:             list (no project), get?id=, save {id, data, photos}
 * No access key. Photos for the MOMs are stored in a Drive folder in the same Google account.
 *
 * SETUP (one time, signed in to the Google account that should OWN the data):
 *  1. Create a new Google Sheet, e.g. "EcoOils Mentari Data".
 *  2. Extensions -> Apps Script -> delete the sample code -> paste this whole file -> Save.
 *  3. Pick the function  setup  in the toolbar -> Run -> Review permissions -> Allow.
 *  4. Deploy -> New deployment -> gear icon -> Web app ->
 *       Execute as: Me      Who has access: Anyone      -> Deploy.
 *  5. Copy the Web app URL (ends in /exec) and send it to Claude / paste it into the apps.
 *
 * AFTER EDITING THIS FILE LATER: Deploy -> Manage deployments -> pencil -> Version: New version -> Deploy.
 * (The URL stays the same.)
 */
const VERSION     = '2.0.0';
const SS_ID       = '';                      // leave empty = the sheet this script is attached to
const MOM_FOLDER  = 'StandDownMOM_Photos';
const CHUNK       = 40000;                   // Sheets cell limit is 50,000 characters
const MAXCH       = 12;                      // up to 12 chunks (~480k chars) per record

const INC_TAB = 'Incidents';
const KV_TAB  = 'KV';
const MOM_TAB = 'StandDownMOM';
const INC_HEAD = ['Project', 'ID', 'EOSM', 'Title', 'Date', 'Updated'].concat(range_(MAXCH).map(i => 'JSON' + (i + 1)));
const KV_HEAD  = ['Key', 'Updated'].concat(range_(MAXCH).map(i => 'V' + (i + 1)));
const MOM_HEAD = ['ID', 'Date', 'ChairedBy', 'Updated', 'Photo1', 'Photo2', 'JSON'];

function setup() { tab_(INC_TAB, INC_HEAD); tab_(KV_TAB, KV_HEAD); tab_(MOM_TAB, MOM_HEAD); folder_(); }

/* ---------------- web app entry points ---------------- */
function doGet(e) {
  try {
    const p = e.parameter || {}, a = p.action;
    if (a === 'list') {
      if (p.project) return out_({ ok: true, incidents: listInc_(p.project) });   // IMS
      return out_({ ok: true, items: momList_() });                               // MOM
    }
    if (a === 'get')        return out_(Object.assign({ ok: true }, momGet_(p.id)));
    if (a === 'kvGet')      return out_({ ok: true, value: kvGet_(p.key) });
    if (a === 'exportXlsx') return out_(exportXlsx_());
    return out_({ ok: true, service: 'EcoOils Mentari backend', version: VERSION });
  } catch (err) { return out_({ ok: false, error: String(err.message || err) }); }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const b = JSON.parse(e.postData.contents);
    lock.waitLock(25000);
    if (b.action === 'kvSet') { kvSet_(b.key, b.value); return out_({ ok: true }); }
    if (b.action === 'save') {
      if (Array.isArray(b.incidents)) { saveInc_(String(b.project || ''), b.incidents); return out_({ ok: true }); }  // IMS
      return out_({ ok: true, id: momSave_(b) });                                                                      // MOM
    }
    throw new Error('Unknown action');
  } catch (err) { return out_({ ok: false, error: String(err.message || err) }); }
  finally { try { lock.releaseLock(); } catch (x) {} }
}

/* ---------------- shared helpers ---------------- */
function range_(n) { const a = []; for (let i = 0; i < n; i++) a.push(i); return a; }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function ss_() { return SS_ID ? SpreadsheetApp.openById(SS_ID) : SpreadsheetApp.getActiveSpreadsheet(); }
function tab_(name, headers) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function folder_() {
  const it = DriveApp.getFoldersByName(MOM_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(MOM_FOLDER);
}
function split_(s) {
  const out = [];
  for (let i = 0; i < s.length; i += CHUNK) out.push(s.substr(i, CHUNK));
  if (out.length > MAXCH) throw new Error('Record too large to store');
  while (out.length < MAXCH) out.push('');
  return out;
}
function writeRows_(sh, width, keepRows, newRows) {
  const all = keepRows.concat(newRows);
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, width).clearContent();
  if (all.length) {
    const r = sh.getRange(2, 1, all.length, width);
    r.setNumberFormat('@');          // store everything as plain text (no auto number/date conversion)
    r.setValues(all);
  }
}

/* ---------------- IMS: incidents ---------------- */
function listInc_(project) {
  const sh = tab_(INC_TAB, INC_HEAD), last = sh.getLastRow();
  if (last < 2) return [];
  const rows = sh.getRange(2, 1, last - 1, INC_HEAD.length).getValues();
  const out = [];
  rows.forEach(r => {
    if (String(r[0]) !== project) return;
    const json = r.slice(6).join('');
    if (json) { try { out.push(JSON.parse(json)); } catch (e) {} }
  });
  return out;
}
function saveInc_(project, incidents) {
  const sh = tab_(INC_TAB, INC_HEAD), last = sh.getLastRow(), w = INC_HEAD.length;
  const keep = last > 1 ? sh.getRange(2, 1, last - 1, w).getValues().filter(r => String(r[0]) !== project) : [];
  const now = new Date().toISOString();
  const add = incidents.map(i => [project, i.id || '', i.eosm || '', i.title || '', i.date || '', now].concat(split_(JSON.stringify(i))));
  writeRows_(sh, w, keep, add);
}
function allInc_() {
  const sh = tab_(INC_TAB, INC_HEAD), last = sh.getLastRow();
  if (last < 2) return [];
  const out = [];
  sh.getRange(2, 1, last - 1, INC_HEAD.length).getValues().forEach(r => {
    const json = r.slice(6).join('');
    if (json) { try { out.push({ project: String(r[0]), rec: JSON.parse(json) }); } catch (e) {} }
  });
  return out;
}

/* ---------------- IMS: key/value store (projects, settings...) ---------------- */
function kvGet_(key) {
  const sh = tab_(KV_TAB, KV_HEAD), last = sh.getLastRow();
  if (last < 2) return null;
  const rows = sh.getRange(2, 1, last - 1, KV_HEAD.length).getValues();
  for (let i = 0; i < rows.length; i++) if (String(rows[i][0]) === key) return rows[i].slice(2).join('');
  return null;
}
function kvSet_(key, value) {
  const sh = tab_(KV_TAB, KV_HEAD), last = sh.getLastRow(), w = KV_HEAD.length;
  const keep = last > 1 ? sh.getRange(2, 1, last - 1, w).getValues().filter(r => String(r[0]) !== String(key)) : [];
  const row = [String(key), new Date().toISOString()].concat(split_(String(value == null ? '' : value)));
  writeRows_(sh, w, keep, [row]);
}

/* ---------------- IMS: Excel export ---------------- */
function exportXlsx_() {
  const all = allInc_();
  const pref = ['project', 'eosm', 'title', 'date', 'time', 'location', 'contractor', 'type', 'classification', 'riskLevel', 'status', 'description', 'immediateActions', 'corrective', 'recommendation'];
  const seen = {};
  all.forEach(x => Object.keys(x.rec).forEach(k => { const v = x.rec[k]; if (v === null || typeof v !== 'object') seen[k] = true; }));
  const cols = ['project'].concat(pref.filter(k => k !== 'project' && seen[k]), Object.keys(seen).filter(k => k !== 'project' && pref.indexOf(k) < 0 && k !== 'id'));
  const cell = v => (v === null || v === undefined) ? '' : (typeof v === 'object' ? JSON.stringify(v) : String(v).substr(0, 49000));
  const data = [cols].concat(all.map(x => cols.map(c => c === 'project' ? x.project : cell(x.rec[c]))));
  const tmp = SpreadsheetApp.create('HSSE Incident Data (export)');
  try {
    const sh = tmp.getSheets()[0];
    sh.getRange(1, 1, data.length, cols.length).setNumberFormat('@').setValues(data);
    sh.getRange(1, 1, 1, cols.length).setFontWeight('bold');
    sh.setFrozenRows(1);
    SpreadsheetApp.flush();
    const url = 'https://docs.google.com/spreadsheets/d/' + tmp.getId() + '/export?format=xlsx';
    const res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } });
    return { ok: true, filename: 'HSSE Incident Data.xlsx', base64: Utilities.base64Encode(res.getBlob().getBytes()) };
  } finally { DriveApp.getFileById(tmp.getId()).setTrashed(true); }
}

/* ---------------- Stand Down MOM ---------------- */
function momFind_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  return -1;
}
function momList_() {
  const sh = tab_(MOM_TAB, MOM_HEAD), last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 4).getValues()
    .map(r => ({ id: r[0], date: r[1], chair: r[2], updated: r[3] }))
    .sort((a, b) => String(b.updated).localeCompare(String(a.updated)));
}
function momGet_(id) {
  const sh = tab_(MOM_TAB, MOM_HEAD), row = momFind_(sh, id);
  if (row < 0) throw new Error('Not found: ' + id);
  const v = sh.getRange(row, 1, 1, MOM_HEAD.length).getValues()[0];
  const photos = [v[4], v[5]].map(fid => {
    if (!fid) return null;
    try {
      const f = DriveApp.getFileById(fid);
      return 'data:' + f.getMimeType() + ';base64,' + Utilities.base64Encode(f.getBlob().getBytes());
    } catch (e) { return null; }
  });
  return { data: JSON.parse(v[6] || '{}'), photos: photos };
}
function momSave_(b) {
  const sh = tab_(MOM_TAB, MOM_HEAD), id = b.id;
  if (!id) throw new Error('Missing id');
  const row = momFind_(sh, id);
  const old = row > 0 ? sh.getRange(row, 5, 1, 2).getValues()[0] : ['', ''];
  const ids = [0, 1].map(i => {
    const p = (b.photos || [])[i];
    if (p === 'keep') return old[i];
    if (old[i]) { try { DriveApp.getFileById(old[i]).setTrashed(true); } catch (e) {} }
    if (!p) return '';
    const m = /^data:(.+?);base64,(.*)$/.exec(p);
    if (!m) return '';
    return folder_().createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], id + '_photo' + (i + 1) + '.jpg')).getId();
  });
  const vals = [id, b.data.date || '', b.data.chair || '', new Date().toISOString(), ids[0], ids[1], JSON.stringify(b.data)];
  if (row < 0) sh.appendRow(vals);
  else sh.getRange(row, 1, 1, MOM_HEAD.length).setValues([vals]);
  return id;
}
