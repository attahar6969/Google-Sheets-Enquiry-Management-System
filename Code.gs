/**
 * SALES ENQUIRY TRACKER  -  Google Apps Script (Code.gs)  -  v2
 *
 * Menu: Sales Tracker > New Enquiry | Follow-Up Update | Upload Bill | Refresh | Admin
 *
 * PROTECTION MODEL
 *  - Enquiry Tracker, Customer Master, Follow-Up History, Today's Follow-Up and Dashboard
 *    are protected: only the owner can type in them.
 *  - Every menu form saves through a Web App that runs AS THE OWNER (doPost), so other
 *    users can add data ONLY through the menu forms.
 */

const SHEETS = {
  ENQ: 'Enquiry Tracker',
  CUST: 'Customer Master',
  HIST: 'Follow-Up History',
  TODAY: "Today's Follow-Up",
  DD: 'Dropdown List',
  DASH: 'Dashboard'
};

const ENQ_HEADERS = ['Enquiry ID', 'Date of Enquiry', 'Customer Name', 'Contact Person Name', 'Phone No.',
  'Mail ID', 'Category', 'Quotation Link', 'Quotation Amount (₹)', 'Source', 'Representative', 'Status',
  'Priority', 'Follow-Up Date', 'Expected Closing Date', 'Remarks', 'Bill No.', 'Bill Image',
  'Bill Amount (₹)', 'Won Verification'];

const COL = { ID: 1, DATE: 2, CUST: 3, CONTACT: 4, PHONE: 5, MAIL: 6, CAT: 7, QLINK: 8, QAMT: 9, SRC: 10,
  REP: 11, STATUS: 12, PRIO: 13, FUP: 14, EXP: 15, REM: 16, BILLNO: 17, BILLIMG: 18, BILLAMT: 19, WONV: 20 };

const MAXROW = 5000;
const DATE_FMT = 'dd-mmm-yyyy';
const INR_FORMAT = '[>=10000000]"₹"##\\,##\\,##\\,##0;[>=100000]"₹"##\\,##\\,##0;"₹"##,##0';
const BILL_FOLDER_NAME = 'Sales Bills';
const PHONE_RE = /^[6-9]\d{9}$/;            // 10 digits, 6000000000 to 9999999999
const LINK_OPTIONAL_SOURCES = ['by meeting', 'just dail', 'just dial', 'india mart', 'cold call'];


/* ------------------------------------------------------------------ */
/*  MENU                                                               */
/* ------------------------------------------------------------------ */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Sales Tracker')
    .addItem('➕ New Enquiry', 'showNewEnquiry')
    .addItem('📞 Follow-Up Update', 'showFollowUp')
    .addItem('🧾 Upload Bill', 'showUploadBill')
    .addItem('📊 Dashboard', 'showDashboard')
    .addSeparator()
    .addItem('🔄 Refresh', 'refreshAll')
    .addSeparator()
    //.addSubMenu(SpreadsheetApp.getUi().createMenu('⚙ Admin (owner only)')
      //.addItem('1. Setup (run once)', 'setupSheet')
      //.addItem('2. Save Web App URL', 'saveWebAppUrl')
      //.addItem('Re-apply sheet protection', 'protectSheets_'))
      //.addItem('Set Admin Emails', 'setAdmins')
    .addToUi();

  
  // Settings menu: shown only to the owner (server side also enforces this)
  const em = (Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!em || isOwner_()) {
    SpreadsheetApp.getUi().createMenu('Settings')
      .addItem('⚙ Open Settings', 'showSettings')
      .addToUi();
  }
}


function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActive();
}

function showDialog_(file, title, w, h) {
  if (!ss_().getSheetByName(SHEETS.ENQ)) {
    SpreadsheetApp.getUi().alert('Owner must run Sales Tracker > Admin > Setup first.');
    return;
  }
  const html = HtmlService.createTemplateFromFile(file).evaluate().setWidth(w).setHeight(h);
  SpreadsheetApp.getUi().showModalDialog(html, title);
}
function showNewEnquiry() { showDialog_('NewEnquiry', 'New Enquiry', 620, 680); }
function showFollowUp()   { showDialog_('FollowUp', 'Follow-Up Update', 620, 700); }
function showUploadBill() { showDialog_('UploadBill', 'Upload Bill', 520, 520); }

/* ------------------------------------------------------------------ */
/*  OWNER-RUN WEB APP  (all writes go through here)                    */
/* ------------------------------------------------------------------ */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const key = PropertiesService.getScriptProperties().getProperty('API_KEY');
    if (!key || body.key !== key) throw new Error('Unauthorized request.');
    return json_({ ok: true, result: dispatch_(body.action, body.data, body.user) });
  } catch (err) {
    return json_({ ok: false, error: err.message });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function dispatch_(action, data, user) {
  switch (action) {
    case 'ping':       return 'pong';
    case 'newEnquiry': return newEnquiryImpl_(data);
    case 'followUp':   return followUpImpl_(data, user);
    case 'bill':       return billImpl_(data);
    case 'refresh':    return refreshImpl_();
    default: throw new Error('Unknown action: ' + action);
  }
}

function userEmail_() {
  try { return Session.getActiveUser().getEmail() || ''; } catch (e) { return ''; }
}

/** Runs the action as the owner via the Web App (or directly if no Web App URL is saved yet). */
function callOwner_(action, data) {
  const props = PropertiesService.getScriptProperties();
  const url = props.getProperty('WEBAPP_URL');
  const user = userEmail_();
  if (!url) return dispatch_(action, data, user);

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: JSON.stringify({ key: props.getProperty('API_KEY'), action: action, data: data, user: user }),
    muteHttpExceptions: true,
    followRedirects: true
  });
  let out;
  try { out = JSON.parse(res.getContentText()); }
  catch (e) {
    throw new Error('Could not reach the Web App (HTTP ' + res.getResponseCode() + '). Ask the owner to check the Web App deployment (Execute as: Me) and redeploy the latest version.');
  }
  if (!out.ok) throw new Error(out.error);
  return out.result;
}

// Public functions called by the popups
function saveNewEnquiry(d) { return callOwner_('newEnquiry', d); }
function saveFollowUp(d)   { return callOwner_('followUp', d); }
function saveBill(d)       { return callOwner_('bill', d); }

/* ------------------------------------------------------------------ */
/*  SETUP  (owner)                                                     */
/* ------------------------------------------------------------------ */
function setupSheet() {
  const ss = SpreadsheetApp.getActive();
  const props = PropertiesService.getScriptProperties();
  props.setProperty('SS_ID', ss.getId());
  if (!props.getProperty('API_KEY')) props.setProperty('API_KEY', Utilities.getUuid() + Utilities.getUuid());

  const names = [SHEETS.ENQ, SHEETS.CUST, SHEETS.HIST, SHEETS.TODAY, SHEETS.DD, SHEETS.DASH];
  names.forEach((n, i) => { if (!ss.getSheetByName(n)) ss.insertSheet(n, i); });
  const def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);

  // Enquiry Tracker
  const tr = ss.getSheetByName(SHEETS.ENQ);
  tr.getRange(1, 1, 1, ENQ_HEADERS.length).setValues([ENQ_HEADERS]);
  styleHeader_(tr, ENQ_HEADERS.length);
  tr.getRange(2, COL.DATE, MAXROW, 1).setNumberFormat(DATE_FMT);
  tr.getRange(2, COL.FUP, MAXROW, 1).setNumberFormat(DATE_FMT);
  tr.getRange(2, COL.EXP, MAXROW, 1).setNumberFormat(DATE_FMT);
  tr.getRange(2, COL.PHONE, MAXROW, 1).setNumberFormat('@');
  tr.getRange(2, COL.QAMT, MAXROW, 1).setNumberFormat(INR_FORMAT);
  tr.getRange(2, COL.BILLAMT, MAXROW, 1).setNumberFormat(INR_FORMAT);
  tr.setColumnWidths(1, ENQ_HEADERS.length, 130);
  tr.setColumnWidth(COL.CUST, 200);

  // Customer Master
  const cm = ss.getSheetByName(SHEETS.CUST);
  const cmH = ['Customer ID', 'Customer Name', 'Contact Person', 'Phone No.', 'Mail ID', 'Last Enquiry Date', 'Last Enquiry No.'];
  cm.getRange(1, 1, 1, cmH.length).setValues([cmH]);
  styleHeader_(cm, cmH.length);
  cm.getRange(2, 4, MAXROW, 1).setNumberFormat('@');
  cm.getRange(2, 6, MAXROW, 1).setNumberFormat(DATE_FMT);
  cm.setColumnWidths(1, cmH.length, 150);

  // Follow-Up History  (Expected Closing Date added as the LAST column, so older data stays aligned)
  const hi = ss.getSheetByName(SHEETS.HIST);
  const hiH = ['Timestamp', 'Enquiry ID', 'Customer Name', 'Phone No.', 'Category', 'Follow-Up Date',
    'Representative', 'Status', 'Latest Remark', 'Updated By', 'Expected Closing Date'];
  hi.getRange(1, 1, 1, hiH.length).setValues([hiH]);
  styleHeader_(hi, hiH.length);
  hi.getRange(2, 1, MAXROW, 1).setNumberFormat('dd-mmm-yyyy hh:mm');
  hi.getRange(2, 4, MAXROW, 1).setNumberFormat('@');
  hi.getRange(2, 6, MAXROW, 1).setNumberFormat(DATE_FMT);
  hi.getRange(2, 11, MAXROW, 1).setNumberFormat(DATE_FMT);
  hi.setColumnWidths(1, hiH.length, 140);
  hi.setColumnWidth(9, 300);

  // Today's Follow-Up
  const td = ss.getSheetByName(SHEETS.TODAY);
  const tdH = ['Enquiry ID', 'Customer Name', 'Phone No.', 'Quotation Link', 'Source', 'Follow-Up Date',
    'Representative', 'Status', 'Remarks'];
  td.getRange(1, 1, 1, tdH.length).setValues([tdH]);
  styleHeader_(td, tdH.length);
  td.getRange(2, 3, MAXROW, 1).setNumberFormat('@');
  td.getRange(2, 6, MAXROW, 1).setNumberFormat(DATE_FMT);
  td.setColumnWidths(1, tdH.length, 140);
  td.setColumnWidth(9, 300);
  td.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($F2<>"",$F2<TODAY())')
      .setFontColor('#d93025').setRanges([td.getRange('F2:F' + MAXROW)]).build()
  ]);

  // Dropdown List (seed only when empty)
  const dd = ss.getSheetByName(SHEETS.DD);
  dd.getRange(1, 1, 1, 5).setValues([['Category', 'Source', 'Representative', 'Status', 'Priority']]);
  styleHeader_(dd, 5);
  dd.setColumnWidths(1, 5, 170);
  if (dd.getLastRow() < 2) {
    const seed = [
      ['Product', 'Website', 'Representative 1', 'New', 'High'],
      ['Service', 'Phone Call', 'Representative 2', 'Contacted', 'Medium'],
      ['AMC', 'Walk-in', 'Representative 3', 'Quotation Sent', 'Low'],
      ['Project', 'Reference', '', 'Negotiation', ''],
      ['Spare Parts', 'Social Media', '', 'On Hold', ''],
      ['Other', 'Email', '', 'Won', ''],
      ['', 'Exhibition', '', 'Lost', '']
    ];
    dd.getRange(2, 1, seed.length, 5).setValues(seed);
  }
  dd.getRange('G1').setValue('Edit / add / delete values in columns A to E. Forms and dropdowns pick up changes automatically. Keep the names "Won" and "Lost" in Status.')
    .setFontStyle('italic').setFontColor('#5f6368');

  // Bill folder + triggers
  getBillFolder_();
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['handleEdit', 'dailyRefresh'].indexOf(t.getHandlerFunction()) > -1) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('handleEdit').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('dailyRefresh').timeBased().everyDays(1).atHour(6).create();

  refreshImpl_();
  protectSheets_();
  ss.setActiveSheet(ss.getSheetByName(SHEETS.DASH));

  SpreadsheetApp.getUi().alert(
    'Setup complete and sheets are protected.\n\n' +
    'NEXT (needed so other users can use the forms):\n' +
    '1. Extensions > Apps Script > Deploy > New deployment > type "Web app"\n' +
    '2. Execute as: Me   |   Who has access: Anyone\n' +
    '3. Click Deploy, authorise, and copy the Web app URL\n' +
    '4. Back here: Sales Tracker > Admin > 2. Save Web App URL');
}

function styleHeader_(sh, n) {
  sh.getRange(1, 1, 1, n).setBackground('#1a4e8a').setFontColor('#ffffff').setFontWeight('bold')
    .setVerticalAlignment('middle').setWrap(true);
  sh.setRowHeight(1, 36);
  sh.setFrozenRows(1);
}

function saveWebAppUrl() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Save Web App URL', 'Paste the Web app URL (ends with /exec):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const url = r.getResponseText().trim();
  if (!/^https:\/\/script\.google\.com\/.*\/exec$/.test(url)) {
    ui.alert('That does not look like a Web app URL. It must start with https://script.google.com and end with /exec');
    return;
  }
  const props = PropertiesService.getScriptProperties();
  props.setProperty('WEBAPP_URL', url);
  try {
    const out = callOwner_('ping', {});
    ui.alert(out === 'pong' ? 'Connected. The forms now save through the owner-run Web App.' : 'Unexpected reply from Web App.');
  } catch (e) {
    props.deleteProperty('WEBAPP_URL');
    ui.alert('Test failed, URL not saved.\n\n' + e.message);
  }
}

/** Only the owner can edit protected sheets; everyone else must use the menu forms. */
function protectSheets_() {
  const ss = ss_();
  const me = Session.getEffectiveUser();
  [SHEETS.ENQ, SHEETS.CUST, SHEETS.HIST, SHEETS.TODAY, SHEETS.DASH].forEach(n => {
    const sh = ss.getSheetByName(n);
    if (!sh) return;
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
    const p = sh.protect().setDescription('Protected. Use the Sales Tracker menu to add or update data.');
    p.addEditor(me);
    p.removeEditors(p.getEditors().filter(u => u.getEmail() !== me.getEmail()));
    if (p.canDomainEdit()) p.setDomainEdit(false);
  });
}

function applyValidation_() {
  const ss = ss_();
  const tr = ss.getSheetByName(SHEETS.ENQ), dd = ss.getSheetByName(SHEETS.DD);
  const map = [[COL.CAT, 'A'], [COL.SRC, 'B'], [COL.REP, 'C'], [COL.STATUS, 'D'], [COL.PRIO, 'E']];
  map.forEach(m => {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(dd.getRange(m[1] + '2:' + m[1] + '300'), true)
      .setAllowInvalid(false).build();
    tr.getRange(2, m[0], MAXROW - 1, 1).setDataValidation(rule);
  });
}

/* ------------------------------------------------------------------ */
/*  TRIGGERS  (installable triggers run as the owner)                  */
/* ------------------------------------------------------------------ */
function handleEdit(e) {
  if (!e || !e.range) return;
  const name = e.range.getSheet().getName();
  if (name === SHEETS.ENQ) {
    rebuildCustomerMaster_();
    refreshToday_();
  } else if (name === SHEETS.DD) {
    applyValidation_();
    refreshDashboard_();
  }
}
function dailyRefresh() { refreshToday_(); }

/* ------------------------------------------------------------------ */
/*  REFRESH                                                            */
/* ------------------------------------------------------------------ */
function refreshAll() {
  const ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName(SHEETS.ENQ)) {
    SpreadsheetApp.getUi().alert('Owner must run Sales Tracker > Admin > Setup first.');
    return;
  }
  try {
    callOwner_('refresh', {});
    ss.toast('All sheets refreshed.', 'Sales Tracker', 4);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Refresh failed: ' + e.message);
  }
}

function refreshImpl_() {
  applyValidation_();
  rebuildCustomerMaster_();
  refreshToday_();
  refreshDashboard_();
  return 'ok';
}

/* ------------------------------------------------------------------ */
/*  HELPERS                                                            */
/* ------------------------------------------------------------------ */
function pad_(n) { return String(n).padStart(3, '0'); }

function nextId_(sh, prefix) {
  const last = sh.getLastRow();
  let max = 0;
  if (last > 1) {
    sh.getRange(2, 1, last - 1, 1).getValues().forEach(r => {
      const m = String(r[0]).match(/(\d+)$/);
      if (m) max = Math.max(max, parseInt(m[1], 10));
    });
  }
  return prefix + '-' + pad_(max + 1);
}

function parseDate_(s) {
  if (!s) return '';
  const p = String(s).split('-');
  if (p.length !== 3) return '';
  return new Date(+p[0], +p[1] - 1, +p[2]);
}

/** Names are saved in CAPITAL LETTERS. For Title Case (Ravi Kumar) change toUpperCase() to the line shown in the comment. */
function cap_(v) {
  return String(v || '').trim().replace(/\s+/g, ' ').toUpperCase();
  // Title Case alternative:
  // return String(v || '').trim().replace(/\s+/g, ' ').toLowerCase().replace(/(^|\s)\S/g, c => c.toUpperCase());
}

/** Remarks in CAPITAL LETTERS (line breaks kept). */
function capText_(v) {
  return String(v || '').trim().toUpperCase();
}

function toTime_(v) { return v instanceof Date ? v.getTime() : 0; }

function isClosed_(s) {
  const x = String(s || '').trim().toLowerCase();
  return x === 'won' || x === 'lost'|| x === 'closed';
}

function setLink_(range, url, label) {
  if (!url) { range.clearContent(); return; }
  url = String(url).trim();
  if (!/^https?:\/\//i.test(url)) url = 'https://' + url;
  range.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(label).setLinkUrl(url).build());
}

function getDropdowns_() {
  const dd = ss_().getSheetByName(SHEETS.DD);
  const last = Math.max(dd.getLastRow(), 2);
  const v = dd.getRange(2, 1, last - 1, 5).getValues();
  const col = i => v.map(r => String(r[i]).trim()).filter(x => x);
  return { categories: col(0), sources: col(1), reps: col(2), statuses: col(3), priorities: col(4) };
}

function todayStr_() {
  return Utilities.formatDate(new Date(), ss_().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
}

function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0]).trim());
  const i = ids.indexOf(String(id).trim());
  return i < 0 ? -1 : i + 2;
}

function getEnquiryOptions_() {
  const ss = ss_();
  const sh = ss.getSheetByName(SHEETS.ENQ);
  const tz = ss.getSpreadsheetTimeZone();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const v = sh.getRange(2, 1, last - 1, 20).getValues();
  const rt = sh.getRange(2, COL.QLINK, last - 1, 1).getRichTextValues();
  return v.map((r, i) => ({
    id: String(r[0]),
    customer: String(r[2]),
    phone: String(r[4]),
    link: rt[i][0].getLinkUrl() || (/^https?:/i.test(String(r[7])) ? String(r[7]) : ''),
    category: String(r[6]),
    rep: String(r[10]),
    status: String(r[11]),
    remark: String(r[15] || ''),        // <-- ADD THIS LINE
    expected: r[14] instanceof Date ? Utilities.formatDate(r[14], tz, 'yyyy-MM-dd') : '',
    verified: String(r[19]).toLowerCase() === 'verified',
    billNo: String(r[16])
  })).filter(o => o.id).reverse();
}

function getBillFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('BILL_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* recreate below */ }
  }
    const it = DriveApp.getFoldersByName(getSettings_().billFolderName);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(getSettings_().billFolderName);
  props.setProperty('BILL_FOLDER_ID', folder.getId());
  return folder;
}

/* ------------------------------------------------------------------ */
/*  DATA FOR FORMS  (read only, run as the user)                       */
/* ------------------------------------------------------------------ */
function getNewEnquiryData() {
  const d = getDropdowns_();
  d.statuses = d.statuses.filter(s => !isClosed_(s));   // Won / Lost hidden in New Enquiry
  d.nextId = nextId_(ss_().getSheetByName(SHEETS.ENQ), 'ENQ');
  d.today = todayStr_();
  d.required = getSettings_().required;
  return d;
}

function getFollowUpData() {
  const d = getDropdowns_();
  const admin = isAdmin_(userEmail_());
  if (!admin) d.statuses = d.statuses.filter(s => !isAdminOnlyStatus_(s));
   const S = getSettings_();
  d.wonVerification = S.wonVerification;
  d.maxFileMB = S.maxFileMB;
  d.required = S.required;
  d.enquiries = getEnquiryOptions_();
  return d;
}

function getBillFormData() {
  const S = getSettings_();
  return { enquiries: getEnquiryOptions_(), maxFileMB: S.maxFileMB, required: S.required };
}
/* ------------------------------------------------------------------ */
/*  IMPLEMENTATIONS  (run as OWNER through the Web App)                */
/* ------------------------------------------------------------------ */
function newEnquiryImpl_(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const prev = seen_(d.requestId); if (prev) return prev;
    const R = getSettings_().required.enq;
    const blank = v => !String(v == null ? '' : v).trim();
    const phone = String(d.phone || '').trim();
    const link = String(d.quotationLink || '').trim();
    const amt = blank(d.amount) ? NaN : Number(d.amount);

    const vals = { date: d.date, customer: d.customer, contact: d.contact, phone: phone, mail: d.mail,
      category: d.category, source: d.source, qlink: link, amount: d.amount, rep: d.rep, status: d.status,
      priority: d.priority, followUp: d.followUp, expected: d.expected, remarks: d.remarks };
    const names = { date: 'Date of Enquiry', customer: 'Customer Name', contact: 'Contact Person Name',
      phone: 'Phone No.', mail: 'Mail ID', category: 'Category', source: 'Source', qlink: 'Quotation Link',
      amount: 'Quotation Amount', rep: 'Representative', status: 'Status', priority: 'Priority',
      followUp: 'Follow-Up Date', expected: 'Expected Closing Date', remarks: 'Remarks' };
    const linkOptional = LINK_OPTIONAL_SOURCES.indexOf(String(d.source || '').trim().toLowerCase()) > -1;
    Object.keys(vals).forEach(k => {
      if (k === 'qlink' && linkOptional) return;   // optional for these sources
      if (R[k] && blank(vals[k])) throw new Error(names[k] + ' is required.');
    });

    if (blank(d.customer) && blank(phone)) throw new Error('Enter at least a Customer Name or a Phone No.');
    if (phone && !PHONE_RE.test(phone)) throw new Error('Phone No. must be exactly 10 digits, between 6000000000 and 9999999999.');
    if (link && !/^https?:\/\/\S+\.\S+/i.test(link)) throw new Error('Quotation Link must be a valid link starting with http:// or https://');
    if (!blank(d.amount) && (isNaN(amt) || amt <= 0)) throw new Error('Quotation Amount must be greater than 0.');
    if (isClosed_(d.status)) throw new Error('Won / Lost cannot be selected in a new enquiry.');

    const sh = ss_().getSheetByName(SHEETS.ENQ);
    const id = nextId_(sh, 'ENQ');
    const row = sh.getLastRow() + 1;

    sh.getRange(row, 1, 1, 20).setValues([[
      id,
      parseDate_(d.date),
      cap_(d.customer),
      cap_(d.contact),
      phone,
      String(d.mail || '').trim(),
      d.category || '',
      '',
      isNaN(amt) ? '' : amt,
      d.source || '',
      d.rep || '',
      d.status,
      d.priority || '',
      parseDate_(d.followUp),
      parseDate_(d.expected),
      capText_(d.remarks),
      '',
      '',
      '',
      ''
    ]]);
    if (link) setLink_(sh.getRange(row, COL.QLINK), link, 'View Quotation');
    SpreadsheetApp.flush();

    queueRefresh_();                 // see Part 2
    const res = { id: id };
    remember_(d.requestId, res);
    return res;
  } 
  finally 
   {
    lock.releaseLock();
   }
}

function billImpl_(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const prev = seen_(d.requestId); if (prev) return prev;
    const S = getSettings_();
    const R = S.required.bill;
    const billNo = String(d.billNo || '').trim();
    const hasAmt = !(d.billAmount === '' || d.billAmount == null);

    if (!d.enquiryId) throw new Error('Enquiry ID is required.');
    if (R.billNo && !billNo) throw new Error('Bill No. is required.');
    if (R.billAmount && !hasAmt) throw new Error('Bill Amount is required.');
    if (R.billImage && !d.base64) throw new Error('Bill Image is required.');
    if (!billNo && !hasAmt && !d.base64) throw new Error('Enter at least a Bill No., Bill Amount or Bill Image.');
    if (hasAmt && isNaN(Number(d.billAmount))) throw new Error('Bill Amount must be a number.');
    if (d.base64 && String(d.base64).length > S.maxFileMB * 1024 * 1024 * 1.34)
      throw new Error('File is larger than ' + S.maxFileMB + ' MB.');

    const sh = ss_().getSheetByName(SHEETS.ENQ);
    const row = findRow_(sh, d.enquiryId);
    if (row < 0) throw new Error('Enquiry ID not found: ' + d.enquiryId);

    let fileUrl = '';
    if (d.base64) {
      const safeName = (d.enquiryId + '_' + (billNo || 'bill') + '_' + d.fileName).replace(/[\\\/:*?"<>|]/g, '-');
      const blob = Utilities.newBlob(Utilities.base64Decode(d.base64), d.mimeType, safeName);
      const file = getBillFolder_().createFile(blob);
      if (S.billSharing === 'anyone') {
        try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* domain policy may block */ }
      }
      fileUrl = file.getUrl();
    }

    if (billNo) sh.getRange(row, COL.BILLNO).setValue(billNo);
    if (fileUrl) setLink_(sh.getRange(row, COL.BILLIMG), fileUrl, 'View Bill');
    if (hasAmt) sh.getRange(row, COL.BILLAMT).setValue(Number(d.billAmount));
    sh.getRange(row, COL.WONV).setValue('Verified');
    SpreadsheetApp.flush();
    const res = { url: fileUrl };
    remember_(d.requestId, res);
    return res;
  } finally {
    lock.releaseLock();
  }
}
function followUpImpl_(d, user) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const prev = seen_(d.requestId); if (prev) return prev;
    const ss = ss_();
    const sh = ss.getSheetByName(SHEETS.ENQ);
    const row = findRow_(sh, d.enquiryId);
    if (row < 0) throw new Error('Enquiry ID not found.');
    const S = getSettings_();
    const R = S.required.fu;
    const blank = v => !String(v == null ? '' : v).trim();

    if (!d.status) throw new Error('Please select a status.');
    if (R.category && blank(d.category)) throw new Error('Category is required.');
    if (R.rep && blank(d.rep)) throw new Error('Representative is required.');
    if (R.expected && blank(d.expected)) throw new Error('Expected Closing Date is required.');
    if (R.remark && blank(d.remark)) throw new Error('Latest Remark is required.');

    const curStatus = String(sh.getRange(row, COL.STATUS).getValue()).trim();
    if (isAdminOnlyStatus_(curStatus) && !isAdmin_(user)) throw new Error('This enquiry is ' + curStatus + '. Only an admin can update it.');
    if (isAdminOnlyStatus_(d.status) && !isAdmin_(user)) throw new Error('Only an admin can set the status to ' + d.status + '.');

    const won = String(d.status).trim().toLowerCase() === 'won';
    if (S.wonVerification && won && String(sh.getRange(row, COL.WONV).getValue()).toLowerCase() !== 'verified')
      throw new Error('Please submit the Bill Details first (required for Won).');
    if (R.followUp && !isClosed_(d.status) && blank(d.followUp))
      throw new Error('Follow-Up Date is required unless the status is Won / Lost.');

    const fup = parseDate_(d.followUp);
    const exp = parseDate_(d.expected);
    sh.getRange(row, COL.CAT).setValue(d.category || '');
    sh.getRange(row, COL.REP).setValue(d.rep || '');
    sh.getRange(row, COL.STATUS).setValue(d.status);
    if (S.resetVerificationOnReopen && curStatus.toLowerCase() === 'won' && !won) sh.getRange(row, COL.WONV).clearContent();
    sh.getRange(row, COL.FUP).setValue(fup);
    sh.getRange(row, COL.EXP).setValue(exp);
    const remark = capText_(d.remark);
    sh.getRange(row, COL.REM).setValue(remark);

    const cust = sh.getRange(row, COL.CUST).getValue();
    const phone = String(sh.getRange(row, COL.PHONE).getValue());
    const hi = ss.getSheetByName(SHEETS.HIST);
    hi.getRange(hi.getLastRow() + 1, 1, 1, 11).setValues([[
      new Date(), d.enquiryId, cust, phone, d.category || '', fup, d.rep || '', d.status, remark, user || '', exp
    ]]);
    SpreadsheetApp.flush();
    queueRefresh_();
    const res = { ok: true };
    remember_(d.requestId, res);
    return res;
  } 
  finally
   {
    lock.releaseLock();
   }
}

/** Owner: run once from the Apps Script editor to convert names already saved in small letters. */
function fixExistingNames() {
  const ss = ss_();
  const sh = ss.getSheetByName(SHEETS.ENQ);
  const last = sh.getLastRow();
  if (last > 1) {
    const names = sh.getRange(2, COL.CUST, last - 1, 2);            // Customer Name + Contact Person
    names.setValues(names.getValues().map(r => [cap_(r[0]), cap_(r[1])]));
    const rem = sh.getRange(2, COL.REM, last - 1, 1);               // Remarks
    rem.setValues(rem.getValues().map(r => [capText_(r[0])]));
  }
  const hi = ss.getSheetByName(SHEETS.HIST);
  const hl = hi.getLastRow();
  if (hl > 1) {
    const h1 = hi.getRange(2, 3, hl - 1, 1);                        // Customer Name
    h1.setValues(h1.getValues().map(r => [cap_(r[0])]));
    const h2 = hi.getRange(2, 9, hl - 1, 1);                        // Latest Remark
    h2.setValues(h2.getValues().map(r => [capText_(r[0])]));
  }
  refreshImpl_();
}

/* ------------------------------------------------------------------ */
/*  CUSTOMER MASTER (rebuilt from Enquiry Tracker, matched by phone)   */
/* ------------------------------------------------------------------ */
function custKey_(phone, name) {
  const digits = String(phone || '').replace(/\D/g, '').slice(-10);
  if (digits) return digits;
  const n = String(name || '').trim().toLowerCase();
  return n ? 'n:' + n : '';
}

function rebuildCustomerMaster_() {
  const ss = ss_();
  const tr = ss.getSheetByName(SHEETS.ENQ), cm = ss.getSheetByName(SHEETS.CUST);
  const idByKey = {};
  let maxNum = 0;
  const cl = cm.getLastRow();
  if (cl > 1) {
    cm.getRange(2, 1, cl - 1, 7).getValues().forEach(r => {
      if (!r[0]) return;
      idByKey[custKey_(r[3], r[1])] = r[0];
      const m = String(r[0]).match(/(\d+)$/);
      if (m) maxNum = Math.max(maxNum, +m[1]);
    });
  }
  const tl = tr.getLastRow();
  const map = {}, order = [];
  if (tl > 1) {
    tr.getRange(2, 1, tl - 1, 20).getValues().forEach(r => {
      if (!r[0]) return;
      const key = custKey_(r[COL.PHONE - 1], r[COL.CUST - 1]);
      if (!key) return;
      const t = toTime_(r[1]);
      const cur = map[key];
      if (!cur) order.push(key);
      if (!cur || t >= cur.t) {
        map[key] = { t: t, name: r[2], contact: r[3], phone: String(r[4]), mail: r[5], date: r[1], enq: r[0] };
      }
    });
  }
  const rows = order.map(k => {
    if (!idByKey[k]) idByKey[k] = 'CUST-' + pad_(++maxNum);
    const c = map[k];
    return [idByKey[k], c.name, c.contact, c.phone, c.mail, c.date, c.enq];
  }).sort((a, b) => String(a[0]).localeCompare(String(b[0])));

  if (cl > 1) cm.getRange(2, 1, cl - 1, 7).clearContent();
  if (rows.length) cm.getRange(2, 1, rows.length, 7).setValues(rows);
}

/* ------------------------------------------------------------------ */
/*  TODAY'S FOLLOW-UP                                                  */
/*  Rule: follow-up date <= today OR blank, status not Won / Lost      */
/* ------------------------------------------------------------------ */
function refreshToday_() {
  const ss = ss_();
  const tr = ss.getSheetByName(SHEETS.ENQ), td = ss.getSheetByName(SHEETS.TODAY);
  const old = td.getLastRow();
  if (old > 1) td.getRange(2, 1, old - 1, 9).clearContent();
  const tl = tr.getLastRow();
  if (tl < 2) return;

  const v = tr.getRange(2, 1, tl - 1, 20).getValues();
  const rt = tr.getRange(2, COL.QLINK, tl - 1, 1).getRichTextValues();
  const end = new Date(); end.setHours(23, 59, 59, 999);

  const items = [];
  v.forEach((r, i) => {
    if (!r[0] || isClosed_(r[COL.STATUS - 1])) return;
    const fu = r[COL.FUP - 1];
    if (fu instanceof Date && fu.getTime() > end.getTime()) return;
    items.push({ r: r, rt: rt[i][0], t: fu instanceof Date ? fu.getTime() : 0 });
  });
  if (!items.length) return;
  items.sort((a, b) => a.t - b.t);

  const vals = items.map(o => [o.r[0], o.r[2], String(o.r[4]), '', o.r[9], o.r[13] || '', o.r[10], o.r[11], o.r[15]]);
  const rich = items.map(o => [o.rt]);
  td.getRange(2, 1, vals.length, 9).setValues(vals);
  td.getRange(2, 4, rich.length, 1).setRichTextValues(rich);
}

/* ------------------------------------------------------------------ */
/*  DASHBOARD                                                          */
/* ------------------------------------------------------------------ */

/** Owner: save the list of admin emails. */
function setAdmins() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const me = (Session.getActiveUser().getEmail() || '').toLowerCase();
  let owner = '';
  try { owner = (SpreadsheetApp.getActive().getOwner().getEmail() || '').toLowerCase(); } catch (e) {}
  if (owner && me !== owner) { ui.alert('Only the owner can change the admin list.'); return; }
  props.setProperty('OWNER_EMAIL', owner || me);
  const r = ui.prompt('Admin emails',
    'Type admin email addresses separated by commas (the owner is always an admin).\nCurrent: ' + (props.getProperty('ADMINS') || 'none'),
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const list = r.getResponseText().split(/[,\s;]+/).map(s => s.trim().toLowerCase()).filter(s => /^\S+@\S+\.\S+$/.test(s));
  props.setProperty('ADMINS', list.join(','));
  ui.alert('Admins saved: ' + (list.join(', ') || 'none (owner only)'));
}

function isAdmin_(email) {
  const e = String(email || '').trim().toLowerCase();
  if (!e) return false;
  const props = PropertiesService.getScriptProperties();
  if (e === (props.getProperty('OWNER_EMAIL') || '')) return true;
  return (props.getProperty('ADMINS') || '').split(',').indexOf(e) > -1;
}

function seen_(rid) {
  if (!rid) return null;
  const v = CacheService.getScriptCache().get('req:' + rid);
  return v ? JSON.parse(v) : null;
}
function remember_(rid, res) {
  if (rid) CacheService.getScriptCache().put('req:' + rid, JSON.stringify(res), 21600);
}

function queueRefresh_() {
  const has = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'deferredRefresh');
  if (!has) ScriptApp.newTrigger('deferredRefresh').timeBased().after(60 * 1000).create();
}

function deferredRefresh() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'deferredRefresh') ScriptApp.deleteTrigger(t);
  });
  rebuildCustomerMaster_();
  refreshToday_();
}
