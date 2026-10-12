/**
 * SETTINGS  (file: SettingsCode.gs)
 *  - Owner-only "Settings" menu that opens Settings.html as a dialog.
 *  - All values are stored in Script Properties (key: SETTINGS) except the dropdown lists,
 *    which stay in the "Dropdown List" sheet.
 *  - Every function here checks assertOwner_(), so non-owners cannot use them.
 *
 * NOTE: top-level constants here must not use COL (defined in Code.gs), because Apps Script may load
 * files in any order. That is why listEnqCol_() is a function.
 */

const SETTINGS_KEY = 'SETTINGS';
const LOCKED_STATUSES = ['won', 'lost', 'closed'];
const LIST_KEYS = ['category', 'source', 'rep', 'status', 'priority'];
const LIST_LABELS = { category: 'Category', source: 'Source', rep: 'Representative', status: 'Status', priority: 'Priority' };
const LIST_HIST_COL = { category: 5, rep: 7, status: 8 };      // columns in Follow-Up History

const SETTINGS_DEFAULTS = {
  wonVerification: true,               // ON: a bill must be submitted before an enquiry can be marked Won
  resetVerificationOnReopen: false,    // ON: moving an enquiry away from Won clears its Verified flag
  adminOnlyStatuses: ['Closed'],       // only admins can set / update these statuses
  billFolderName: 'Sales Bills',
  maxFileMB: 5,
  billSharing: 'anyone'                // 'anyone' (anyone with link can view) or 'private'
};

/** Which fields are required in each form. Edited from Settings > Required Fields. */
const REQUIRED_DEFAULTS = {
  enq:  { date: true, customer: true, contact: false, phone: true, mail: false, category: true, source: true,
          qlink: true, amount: true, rep: true, status: true, priority: true, followUp: false, expected: false, remarks: false },
  fu:   { category: true, rep: true, followUp: true, expected: false, remark: false },
  bill: { billNo: true, billAmount: true, billImage: false }
};
/** Fields that can never be switched off (the system needs them). */
const LOCKED_REQUIRED = { enq: ['date', 'customer', 'status'], fu: [], bill: [] };

function listEnqCol_() {
  return { category: COL.CAT, source: COL.SRC, rep: COL.REP, status: COL.STATUS, priority: COL.PRIO };
}

/* ------------------------------------------------------------------ */
/*  READ SETTINGS  (used by Code.gs, runs for every user)              */
/* ------------------------------------------------------------------ */
function getSettings_() {
  let saved = {};
  try { saved = JSON.parse(PropertiesService.getScriptProperties().getProperty(SETTINGS_KEY) || '{}'); } catch (e) { saved = {}; }
  const s = Object.assign({}, SETTINGS_DEFAULTS, saved);
  s.adminOnlyStatuses = (s.adminOnlyStatuses || []).slice();
  const req = {};
  Object.keys(REQUIRED_DEFAULTS).forEach(f => {
    req[f] = Object.assign({}, REQUIRED_DEFAULTS[f], (saved.required && saved.required[f]) || {});
    (LOCKED_REQUIRED[f] || []).forEach(k => { req[f][k] = true; });
  });
  s.required = req;
  return s;
}

function isAdminOnlyStatus_(status) {
  const x = String(status || '').trim().toLowerCase();
  if (!x) return false;
  return getSettings_().adminOnlyStatuses.some(a => String(a).trim().toLowerCase() === x);
}

/* ------------------------------------------------------------------ */
/*  OWNER CHECK                                                        */
/* ------------------------------------------------------------------ */
function ownerEmail_() {
  let o = '';
  try { o = (ss_().getOwner().getEmail() || '').toLowerCase(); } catch (e) { /* shared drive: no owner */ }
  return o || (PropertiesService.getScriptProperties().getProperty('OWNER_EMAIL') || '').toLowerCase();
}

function isOwner_() {
  let me = '';
  try { me = (Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '').toLowerCase(); } catch (e) { me = ''; }
  const owner = ownerEmail_();
  return !!me && !!owner && me === owner;
}

function assertOwner_() {
  if (!isOwner_()) throw new Error('Only the owner can open or change Settings.');
}

/* ------------------------------------------------------------------ */
/*  OPEN DIALOG                                                        */
/* ------------------------------------------------------------------ */
function showSettings() {
  try { assertOwner_(); } catch (e) { SpreadsheetApp.getUi().alert(e.message); return; }
  const html = HtmlService.createHtmlOutputFromFile('Settings').setWidth(940).setHeight(660);
  SpreadsheetApp.getUi().showModalDialog(html, 'Settings');
}

/* ------------------------------------------------------------------ */
/*  DATA FOR THE DIALOG                                                */
/* ------------------------------------------------------------------ */
function readLists_() {
  const dd = ss_().getSheetByName(SHEETS.DD);
  const last = Math.max(dd.getLastRow(), 2);
  const v = dd.getRange(2, 1, last - 1, 5).getValues();
  const out = {};
  LIST_KEYS.forEach((k, i) => { out[k] = v.map(r => String(r[i]).trim()).filter(x => x); });
  return out;
}

function getUsage_() {
  const out = {};
  LIST_KEYS.forEach(k => { out[k] = {}; });
  const sh = ss_().getSheetByName(SHEETS.ENQ);
  const last = sh.getLastRow();
  if (last < 2) return out;
  const cols = listEnqCol_();
  sh.getRange(2, 1, last - 1, 20).getValues().forEach(r => {
    if (!r[0]) return;
    LIST_KEYS.forEach(k => {
      const x = String(r[cols[k] - 1]).trim();
      if (x) out[k][x] = (out[k][x] || 0) + 1;
    });
  });
  return out;
}

function getSettingsData() {
  assertOwner_();
  const props = PropertiesService.getScriptProperties();
  return {
    settings: getSettings_(),
    lists: readLists_(),
    usage: getUsage_(),
    labels: LIST_LABELS,
    locked: LOCKED_STATUSES,
    admins: (props.getProperty('ADMINS') || '').split(',').filter(String),
    owner: ownerEmail_(),
    webAppUrl: props.getProperty('WEBAPP_URL') || ''
  };
}

/* ------------------------------------------------------------------ */
/*  SAVE: DROPDOWN LISTS                                               */
/*  p = { category:[{orig,name}], source:[...], rep:[...], status:[...], priority:[...] }
/*  A renamed value is also updated in existing enquiries + history.   */
/* ------------------------------------------------------------------ */
function saveDropdowns(p) {
  assertOwner_();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = ss_();
    const dd = ss.getSheetByName(SHEETS.DD);
    const cur = readLists_();
    const names = {}, maps = {};

    LIST_KEYS.forEach(k => {
      const seen = {};
      names[k] = []; maps[k] = {};
      ((p && p[k]) || []).forEach(it => {
        const name = String(it.name || '').replace(/\s+/g, ' ').trim();
        if (!name) return;
        const low = name.toLowerCase();
        if (seen[low]) throw new Error(LIST_LABELS[k] + ': "' + name + '" is entered twice.');
        seen[low] = 1;
        names[k].push(name);
        const orig = String(it.orig || '').trim();
        if (orig && orig !== name) maps[k][orig] = name;
      });
      if (names[k].length > 299) throw new Error(LIST_LABELS[k] + ': too many values (max 299).');
    });

    // Status rules: Won and Lost must exist; locked statuses cannot be renamed or deleted
    const newLow = names.status.map(s => s.toLowerCase());
    ['won', 'lost'].forEach(x => {
      if (newLow.indexOf(x) < 0) throw new Error('Status must keep "Won" and "Lost".');
    });
    cur.status.forEach(s => {
      if (LOCKED_STATUSES.indexOf(s.toLowerCase()) > -1 && names.status.indexOf(s) < 0)
        throw new Error('Status "' + s + '" is required by the system and cannot be renamed or deleted.');
    });

    // Update existing enquiries and history for renamed values
    let changed = 0;
    const tr = ss.getSheetByName(SHEETS.ENQ), hi = ss.getSheetByName(SHEETS.HIST);
    const apply = (sh, col, map) => {
      const last = sh.getLastRow();
      if (last < 2 || !Object.keys(map).length) return;
      const rng = sh.getRange(2, col, last - 1, 1);
      rng.setValues(rng.getValues().map(r => {
        const v = String(r[0]).trim();
        if (Object.prototype.hasOwnProperty.call(map, v)) { changed++; return [map[v]]; }
        return [r[0]];
      }));
    };
    const cols = listEnqCol_();
    LIST_KEYS.forEach(k => {
      apply(tr, cols[k], maps[k]);
      if (LIST_HIST_COL[k]) apply(hi, LIST_HIST_COL[k], maps[k]);
    });

    // Write the lists back to the Dropdown List sheet
    const n = Math.max.apply(null, [1].concat(LIST_KEYS.map(k => names[k].length)));
    const clearRows = Math.min(299, dd.getMaxRows() - 1);
    dd.getRange(2, 1, clearRows, 5).clearContent();
    const out = [];
    for (let i = 0; i < n; i++) out.push(LIST_KEYS.map(k => names[k][i] || ''));
    dd.getRange(2, 1, n, 5).setValues(out);
    SpreadsheetApp.flush();

    applyValidation_();
    refreshToday_();
    refreshDashboard_();
    return { renamedCells: changed };
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------------------------------------------ */
/*  SAVE: STATUS RULES + BILL UPLOAD                                   */
/* ------------------------------------------------------------------ */
function saveGeneralSettings(s) {
  assertOwner_();
  s = s || {};
  const old = getSettings_();
  const n = {
    wonVerification: !!s.wonVerification,
    resetVerificationOnReopen: !!s.resetVerificationOnReopen,
    adminOnlyStatuses: (s.adminOnlyStatuses || []).map(x => String(x).trim()).filter(String),
    billFolderName: String(s.billFolderName || '').trim(),
    maxFileMB: Number(s.maxFileMB),
    billSharing: s.billSharing === 'private' ? 'private' : 'anyone',
    required: old.required                       // keep the Required Fields settings untouched
  };
  if (!n.billFolderName || n.billFolderName.length > 100 || /[\\\/:*?"<>|]/.test(n.billFolderName))
    throw new Error('Folder name is required (max 100 characters, none of \\ / : * ? " < > |).');
  if (!(n.maxFileMB >= 1 && n.maxFileMB <= 20)) throw new Error('Max file size must be between 1 and 20 MB.');

  if (n.billFolderName !== old.billFolderName) {
    try { getBillFolder_().setName(n.billFolderName); } catch (e) { /* a new folder is created on next upload */ }
    PropertiesService.getScriptProperties().deleteProperty('BILL_FOLDER_ID');
  }
  PropertiesService.getScriptProperties().setProperty(SETTINGS_KEY, JSON.stringify(n));
  if (n.billFolderName !== old.billFolderName) getBillFolder_();   // re-links the folder id
  return getSettings_();
}

/* ------------------------------------------------------------------ */
/*  SAVE: REQUIRED FIELDS                                              */
/*  p = { enq:{date:true,...}, fu:{...}, bill:{...} }                  */
/* ------------------------------------------------------------------ */
function saveRequiredFields(p) {
  assertOwner_();
  p = p || {};
  const cur = getSettings_();
  const req = {};
  Object.keys(REQUIRED_DEFAULTS).forEach(f => {
    req[f] = {};
    Object.keys(REQUIRED_DEFAULTS[f]).forEach(k => {
      req[f][k] = !!(p[f] && p[f][k]);
    });
    (LOCKED_REQUIRED[f] || []).forEach(k => { req[f][k] = true; });
  });
  const full = Object.assign({}, cur, { required: req });
  PropertiesService.getScriptProperties().setProperty(SETTINGS_KEY, JSON.stringify(full));
  return getSettings_().required;
}

/* ------------------------------------------------------------------ */
/*  SAVE: ADMINS                                                       */
/* ------------------------------------------------------------------ */
function saveAdmins(text) {
  assertOwner_();
  const list = String(text || '').split(/[,\s;]+/).map(s => s.trim().toLowerCase()).filter(Boolean);
  const bad = list.filter(s => !/^\S+@\S+\.\S+$/.test(s));
  if (bad.length) throw new Error('Not valid email address(es): ' + bad.join(', '));
  const uniq = list.filter((x, i) => list.indexOf(x) === i);
  const props = PropertiesService.getScriptProperties();
  props.setProperty('OWNER_EMAIL', ownerEmail_());
  props.setProperty('ADMINS', uniq.join(','));
  return uniq;
}

/* ------------------------------------------------------------------ */
/*  WEB APP & SECURITY                                                 */
/* ------------------------------------------------------------------ */
function saveWebAppUrlValue(url) {
  assertOwner_();
  const props = PropertiesService.getScriptProperties();
  url = String(url || '').trim();
  if (!url) {
    props.deleteProperty('WEBAPP_URL');
    return 'Web App URL removed. Forms now save directly (other users will not be able to write to protected sheets).';
  }
  if (!/^https:\/\/script\.google\.com\/.*\/exec$/.test(url))
    throw new Error('That does not look like a Web app URL. It must start with https://script.google.com and end with /exec');
  const previous = props.getProperty('WEBAPP_URL');
  props.setProperty('WEBAPP_URL', url);
  try {
    if (callOwner_('ping', {}) !== 'pong') throw new Error('Unexpected reply from Web App.');
  } catch (e) {
    if (previous) props.setProperty('WEBAPP_URL', previous); else props.deleteProperty('WEBAPP_URL');
    throw new Error('Test failed, URL not saved. ' + e.message);
  }
  return 'Connected. The forms now save through the owner-run Web App.';
}

function testWebApp() {
  assertOwner_();
  if (!PropertiesService.getScriptProperties().getProperty('WEBAPP_URL')) return 'No Web App URL saved (forms save directly).';
  return callOwner_('ping', {}) === 'pong' ? 'Connected.' : 'Unexpected reply from Web App.';
}

function regenerateApiKey() {
  assertOwner_();
  PropertiesService.getScriptProperties().setProperty('API_KEY', Utilities.getUuid() + Utilities.getUuid());
  return 'New API key generated. The Web App reads the same key, so no redeploy is needed.';
}

function reapplyProtection() {
  assertOwner_();
  protectSheets_();
  return 'Sheet protection re-applied.';
}

function refreshFromSettings() {
  assertOwner_();
  refreshImpl_();
  return 'All sheets refreshed.';
}

function doGet(e) {
  if (e && e.parameter && e.parameter.api === 'dashboard') {
    return ContentService.createTextOutput(JSON.stringify(getPortalSummary()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput('Sales Tracker is running.');
}

function getPortalSummary() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.ENQ);
  const values = sheet.getDataRange().getValues();
  const h = values.findIndex(r => r.some(c => String(c).trim() === 'Enquiry ID'));
  if (h < 0) return { error: 'Header row not found' };
  const head = values[h].map(c => String(c).trim());
  const iId = head.indexOf('Enquiry ID');
  const iStatus = head.indexOf('Status');
  const iAmt = head.indexOf('Quotation Amount (₹)');
  const rows = values.slice(h + 1).filter(r => String(r[iId]).trim() !== '');
  let won = 0, lost = 0, quoteValue = 0;
  rows.forEach(r => {
    const s = String(r[iStatus]).toLowerCase();
    if (s.indexOf('won') > -1) won++;
    else if (s.indexOf('lost') > -1) lost++;
    quoteValue += Number(r[iAmt]) || 0;
  });
  return {
    totalEnquiries: rows.length,
    open: rows.length - won - lost,
    won: won,
    lost: lost,
    totalQuotationValue: quoteValue
  };
}
