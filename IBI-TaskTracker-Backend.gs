/**
 * ============================================================
 *  IBI Task Target & Completion — Google Sheet Backend
 *  iINTELLIGENCEi · India Business International
 *
 *  WHAT THIS DOES
 *  Stores ALL app data (tasks, staff, settings) in a Google
 *  Sheet that lives in your Google Drive. The web app talks to
 *  this script; this script reads/writes the Sheet.
 *
 *  ONE-TIME SETUP
 *  1. Go to https://script.google.com  ->  New project
 *  2. Delete the sample code, paste THIS whole file, and Save.
 *  3. (Optional) Set API_TOKEN below to a secret word for light
 *     protection, then enter the same word in the app's
 *     "Cloud Sync" settings. Leave "" for easiest setup.
 *  4. Click  Deploy  ->  New deployment.
 *       - Type: Web app
 *       - Execute as: Me
 *       - Who has access: Anyone
 *  5. Authorise when prompted (it will ask to access your Drive
 *     so it can create the Sheet).
 *  6. Copy the Web app URL (it ends in /exec) and paste it into
 *     the app  ->  Settings  ->  Cloud Sync  ->  Connect Google Sheet.
 *
 *  The Sheet ("IBI Task Tracker DB") is created automatically in
 *  your Drive the first time the app connects. Use "Open Google
 *  Sheet" in the app's settings to view it any time.
 *
 *  WHEN YOU UPDATE THIS SCRIPT later, Deploy -> Manage
 *  deployments -> edit -> New version, so the same URL keeps
 *  working.
 * ============================================================
 */

var DB_NAME    = 'IBI Task Tracker DB';
var API_TOKEN  = '';   // optional shared secret; "" = no check

var TASK_COLS  = ['id','title','description','assignedTo','assignedBy','category',
                  'priority','targetDate','createdDate','status','progress',
                  'completedDate','staffNotes','ceoFeedback'];
var STAFF_COLS = ['id','name','role','active'];

/* ---------- spreadsheet bootstrap ---------- */
function getSS() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SHEET_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* recreate below */ }
  }
  var ss = SpreadsheetApp.create(DB_NAME);
  props.setProperty('SHEET_ID', ss.getId());
  initSheets(ss);
  return ss;
}

function initSheets(ss) {
  var tasks = ss.getSheetByName('Tasks') || ss.insertSheet('Tasks');
  if (tasks.getLastRow() === 0) {
    tasks.appendRow(TASK_COLS);
    // keep date/id columns as plain text to avoid locale date corruption
    ['A', 'H', 'I', 'L'].forEach(function (c) {
      tasks.getRange(c + '2:' + c + '2000').setNumberFormat('@');
    });
  }
  var staff = ss.getSheetByName('Staff') || ss.insertSheet('Staff');
  if (staff.getLastRow() === 0) {
    staff.appendRow(STAFF_COLS);
    staff.appendRow(['s_aswin', 'S. Aswin', 'Packaging Manager', true]);
    staff.appendRow(['m_ajay', 'M. Ajay', 'Catalogue Manager', true]);
  }
  var set = ss.getSheetByName('Settings') || ss.insertSheet('Settings');
  if (set.getLastRow() === 0) {
    set.appendRow(['key', 'value']);
    set.appendRow(['ceoName', 'Dr. T. Sasimurugan']);
    set.appendRow(['ceoPin', '']);
  }
  var def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
}

/* ---------- read helpers ---------- */
function rowsToObjects(sheet, cols) {
  var values = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === '' || values[i][0] === null) continue;
    var o = {};
    for (var j = 0; j < cols.length; j++) o[cols[j]] = values[i][j];
    out.push(o);
  }
  return out;
}

function verifyPin(pin) {
  var ss = getSS();
  var data = ss.getSheetByName('Settings').getDataRange().getValues();
  var stored = '';
  for (var i = 1; i < data.length; i++) { if (String(data[i][0]) === 'ceoPin') stored = String(data[i][1]); }
  if (stored === '') return true;            // no PIN set -> nothing to verify
  return String(pin) === stored;
}

function loadDB() {
  var ss = getSS();
  var tasks = rowsToObjects(ss.getSheetByName('Tasks'), TASK_COLS).map(function (t) {
    t.id = String(t.id);
    t.progress = Number(t.progress) || 0;
    t.targetDate = t.targetDate ? String(t.targetDate) : '';
    t.completedDate = t.completedDate ? String(t.completedDate) : null;
    return t;
  });
  var staff = rowsToObjects(ss.getSheetByName('Staff'), STAFF_COLS).map(function (s) {
    s.id = String(s.id);
    s.active = (s.active === true || String(s.active).toLowerCase() === 'true');
    return s;
  });
  var settings = {};
  rowsToObjects(ss.getSheetByName('Settings'), ['key', 'value']).forEach(function (r) {
    settings[r.key] = r.value;
  });
  return {
    ceoName: settings.ceoName || 'Dr. T. Sasimurugan',
    ceoPinSet: settings.ceoPin ? true : false,
    staff: staff,
    tasks: tasks
  };
}

/* ---------- write helpers ---------- */
function findRowById(sheet, id) {
  var ids = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), 1).getValues();
  for (var i = 1; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

function upsertTask(task) {
  var sh = getSS().getSheetByName('Tasks');
  var row = TASK_COLS.map(function (c) { return task[c] == null ? '' : task[c]; });
  var r = findRowById(sh, task.id);
  if (r > 0) sh.getRange(r, 1, 1, TASK_COLS.length).setValues([row]);
  else sh.appendRow(row);
}

function deleteTask(id) {
  var sh = getSS().getSheetByName('Tasks');
  var r = findRowById(sh, id);
  if (r > 0) sh.deleteRow(r);
}

function upsertStaff(staff) {
  var sh = getSS().getSheetByName('Staff');
  var row = STAFF_COLS.map(function (c) { return staff[c] == null ? '' : staff[c]; });
  var r = findRowById(sh, staff.id);
  if (r > 0) sh.getRange(r, 1, 1, STAFF_COLS.length).setValues([row]);
  else sh.appendRow(row);
}

function setSetting(key, value) {
  var sh = getSS().getSheetByName('Settings');
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(key)) {
      sh.getRange(i + 1, 2).setValue(value == null ? '' : value);
      return;
    }
  }
  sh.appendRow([key, value == null ? '' : value]);
}

function clearTasks() {
  var sh = getSS().getSheetByName('Tasks');
  if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
}

function replaceAll(db) {
  var ss = getSS();
  // tasks
  var t = ss.getSheetByName('Tasks');
  t.clearContents();
  t.appendRow(TASK_COLS);
  ['A', 'H', 'I', 'L'].forEach(function (c) { t.getRange(c + '2:' + c + '2000').setNumberFormat('@'); });
  (db.tasks || []).forEach(function (task) {
    t.appendRow(TASK_COLS.map(function (c) { return task[c] == null ? '' : task[c]; }));
  });
  // staff
  var s = ss.getSheetByName('Staff');
  s.clearContents();
  s.appendRow(STAFF_COLS);
  (db.staff || []).forEach(function (st) {
    s.appendRow(STAFF_COLS.map(function (c) { return st[c] == null ? '' : st[c]; }));
  });
  // settings — preserve existing PIN unless the client explicitly sends one
  var existingPin = '';
  var setData = ss.getSheetByName('Settings').getDataRange().getValues();
  for (var k = 1; k < setData.length; k++) { if (String(setData[k][0]) === 'ceoPin') existingPin = setData[k][1]; }
  var g = ss.getSheetByName('Settings');
  g.clearContents();
  g.appendRow(['key', 'value']);
  g.appendRow(['ceoName', db.ceoName || 'Dr. T. Sasimurugan']);
  g.appendRow(['ceoPin', (db.ceoPin != null ? db.ceoPin : existingPin)]);
}

/* ---------- HTTP entry points ---------- */
function doGet(e)  { return handle((e && e.parameter) || {}); }
function doPost(e) {
  var p = {};
  try { p = JSON.parse(e.postData.contents); }
  catch (err) { p = (e && e.parameter) || {}; }
  return handle(p);
}

function handle(p) {
  if (API_TOKEN && String(p.token || '') !== API_TOKEN) {
    return json({ ok: false, error: 'unauthorized' });
  }
  var action = p.action || 'load';
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    if (action !== 'load' && action !== 'ping') { lock.waitLock(20000); locked = true; }
    switch (action) {
      case 'ping':        return json({ ok: true, pong: true });
      case 'load':        return json({ ok: true, db: loadDB(), sheetUrl: getSS().getUrl() });
      case 'verifyPin':   return json({ ok: true, valid: verifyPin(p.pin) });
      case 'upsertTask':  upsertTask(p.task);            return json({ ok: true });
      case 'deleteTask':  deleteTask(p.id);              return json({ ok: true });
      case 'upsertStaff': upsertStaff(p.staff);          return json({ ok: true });
      case 'setSetting':  setSetting(p.key, p.value);    return json({ ok: true });
      case 'clearTasks':  clearTasks();                  return json({ ok: true });
      case 'replaceAll':  replaceAll(p.db);              return json({ ok: true });
      default:            return json({ ok: false, error: 'unknown action: ' + action });
    }
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (e2) {} }
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
