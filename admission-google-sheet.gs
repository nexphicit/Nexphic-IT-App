/**
 * Admission forms -> Google Sheet + Drive folder   (works for BOTH websites with ONE web app URL)
 *
 *  - Nexphic IT Institute site  (sends JSON)         -> sheet tab "Nexphic Admissions"
 *  - Muhammad Raj site          (sends form fields)  -> sheet tab "Admission Forms"
 *
 * SETUP (one time)
 * 1. Open your Google Sheet > Extensions > Apps Script. Replace ALL code with this file.
 * 2. Optional: put an existing Drive folder ID in FOLDER_ID (the last part of the folder link).
 *    Leave it empty to auto-create the folder "Admission Form Copies".
 * 3. Deploy > Manage deployments > pencil icon > Version: New version > Deploy.
 *    (Same deployment = same URL, so the websites need no change.)
 *    First time: Type = Web app, Execute as = Me, Who has access = Anyone.
 * 4. Approve the Sheets + Drive permissions when asked.
 *
 * What is saved for every admission
 *  - a row in the sheet (with links to the photo and the A4 form copy)
 *  - "<StudentID> - <Name> (photo).jpg" and the A4 form copy ("<StudentID> - <Name>.png / .pdf") in the folder
 */
var FOLDER_ID = '';
var FOLDER_NAME = 'Admission Form Copies';
var NX_SHEET = 'Nexphic Admissions';
var RAJ_SHEET = 'Admission Forms';
var NX_HEADERS = ['Submitted At','Student ID','Student Name','Father/Guardian','Mobile','Gender','Education','Current Address','Course','Device','Classes','Duration','Fee','Payment Method','Signature','Photo','Form Copy','Submission ID'];
var RAJ_HEADERS = ['Submitted At','Student ID','Student Name','Father/Guardian Name','Mobile','Gender','Education','Current Address','Course','Fee','Payment Method','Academic Year','Terms Accepted','Signature','Form PDF','Photo','Site'];

/* ---------- entry points ---------- */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var type = String((e.postData && e.postData.type) || '');
    if (type.indexOf('x-www-form-urlencoded') > -1) return text_(rajSave_(e.parameter || {}));   // Muhammad Raj site
    var data = JSON.parse((e.postData && e.postData.contents) || '{}');                          // Nexphic site
    if (data.action === 'slip') return json_(nxSlip_(data));
    return json_(nxSave_(data));
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action === 'lookup') return json_(nxLookup_(String(p.id || '').toUpperCase()));
  return text_('Admission endpoint is running');
}

/* ---------- Nexphic site ---------- */
function nxSave_(d) {
  var sh = sheet_(NX_SHEET, NX_HEADERS);
  var id = String(d.studentId || '');
  if (!/^NXP-\d{2}-[A-Z0-9]{6}$/.test(id)) return { ok: false, error: 'bad id' };
  var found = findRow_(sh, 2, id);
  if (found) return { ok: true, studentId: id, duplicate: true };
  var folder = folder_();
  var safe = safeName_(id, d.name);
  var photoUrl = '';
  var b64 = stripData_(d.photo);
  if (b64) photoUrl = folder.createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', safe + ' (photo).jpg')).getUrl();
  var row = [d.submittedAt || new Date().toISOString(), id, d.name || '', d.guardian || '', d.phone || '', d.gender || '', d.education || '',
             d.address || '', d.course || '', d.device || '', d.classes || '', d.duration || '', d.fee || '', d.payment || '', d.signature || '',
             photoUrl, '', d.submissionId || ''];
  appendText_(sh, row);
  return { ok: true, studentId: id };
}

function nxSlip_(d) {                                   // the A4 form image sent right after saving
  var sh = sheet_(NX_SHEET, NX_HEADERS);
  var id = String(d.studentId || '');
  var rowNo = findRow_(sh, 2, id);
  if (!rowNo || !d.slipImage) return { ok: false };
  var folder = folder_();
  var safe = safeName_(id, d.name);
  var png = Utilities.newBlob(Utilities.base64Decode(d.slipImage), 'image/png', safe + '.png');
  var url = folder.createFile(png).getUrl();
  try {                                                  // also keep a PDF copy
    var html = '<html><body style="margin:0"><img src="data:image/png;base64,' + d.slipImage + '" style="width:100%"></body></html>';
    var pdf = Utilities.newBlob(html, 'text/html', safe + '.html').getAs('application/pdf').setName(safe + '.pdf');
    url = folder.createFile(pdf).getUrl();
  } catch (err) {}
  sh.getRange(rowNo, NX_HEADERS.indexOf('Form Copy') + 1).setValue(url);
  return { ok: true };
}

function nxLookup_(id) {
  var sh = sheet_(NX_SHEET, NX_HEADERS);
  var rowNo = findRow_(sh, 2, id);
  if (!rowNo) return { ok: false };
  var v = sh.getRange(rowNo, 1, 1, NX_HEADERS.length).getValues()[0];
  var adm = { submittedAt: v[0], studentId: v[1], name: v[2], guardian: v[3], phone: v[4], gender: v[5], education: v[6], address: v[7],
              course: v[8], device: v[9], classes: v[10], duration: v[11], fee: v[12], payment: v[13], signature: v[14], photo: '' };
  var m = String(v[15]).match(/[-\w]{25,}/);
  if (m) {
    try { adm.photo = 'data:image/jpeg;base64,' + Utilities.base64Encode(DriveApp.getFileById(m[0]).getBlob().getBytes()); } catch (err) {}
  }
  return { ok: true, admission: adm };
}

/* ---------- Muhammad Raj site ---------- */
function rajSave_(p) {
  var sh = sheet_(RAJ_SHEET, RAJ_HEADERS);
  var id = String(p.studentId || '');
  if (id && findRow_(sh, 2, id)) return 'duplicate';
  var folder = folder_();
  var safe = safeName_(id, p.name);
  var photoB64 = stripData_(p.photo), photoUrl = '', pdfUrl = '';
  if (photoB64) photoUrl = folder.createFile(Utilities.newBlob(Utilities.base64Decode(photoB64), 'image/jpeg', safe + ' (photo).jpg')).getUrl();
  try {
    var pdf = Utilities.newBlob(rajHtml_(p, photoB64), 'text/html', safe + '.html').getAs('application/pdf').setName(safe + '.pdf');
    pdfUrl = folder.createFile(pdf).getUrl();
  } catch (err) { pdfUrl = 'PDF failed: ' + err; }
  appendText_(sh, [p.submittedAt || new Date().toISOString(), id, p.name || '', p.father || '', p.phone || '', p.gender || '', p.education || '',
    p.address || '', p.course || '', p.fee || '', p.payment || '', p.academicYear || '', p.termsAccepted || '', p.signature || '', pdfUrl, photoUrl, p.site || '']);
  return 'ok';
}

function rajHtml_(p, photoB64) {
  function box(l, v, w) { return '<td style="border:1px solid #888;padding:7px 9px;font-size:12px;' + (w ? 'width:' + w + ';' : '') + '"><span style="color:#777">' + esc_(l) + '</span> <b>' + esc_(v) + '</b></td>'; }
  function full(l, v) { return '<table style="width:100%;border-collapse:separate;border-spacing:0 8px"><tr>' + box(l, v) + '</tr></table>'; }
  var photo = photoB64 ? '<img src="data:image/jpeg;base64,' + photoB64 + '" width="110" height="135" style="border:1px dashed #777">' : '<div style="width:110px;height:135px;border:1px dashed #777"></div>';
  return '<html><body style="font-family:Arial,Helvetica,sans-serif;color:#16202a;margin:0;padding:28px 36px">' +
    '<div style="height:10px;background:#1b2a35;margin:0 -36px 4px -36px;width:75%"></div><div style="height:3px;background:#e0242b;margin:0 0 22px -36px;width:80%"></div>' +
    '<table style="width:100%"><tr><td style="vertical-align:top"><div style="font-size:24px;font-weight:bold;border-bottom:2px solid #16202a;display:inline-block">Admission Form</div>' +
    '<table style="width:100%;border-collapse:separate;border-spacing:0 10px;margin-top:14px"><tr>' + box('Date:', p.formDate, '45%') + box('Form no:', p.studentId) + '</tr></table></td>' +
    '<td style="width:120px;text-align:right;vertical-align:top">' + photo + '</td></tr></table>' +
    full("Student's Full Name:", p.name) + full("Father/Guardian's Name:", p.father) + full('Mobile Number:', p.phone) + full('Gender:', p.gender) +
    full('Education:', p.education) + full('Current Address:', p.address) + full('Course:', p.course + (p.fee ? ' (Tk ' + p.fee + ')' : '')) + full('Payment Method:', p.payment) +
    '<div style="font-size:11px;color:#555;margin:6px 0">Agreed to Terms &amp; Conditions: ' + esc_(p.termsAccepted) + '</div>' +
    '<table style="width:100%;margin-top:26px"><tr><td style="width:46%;text-align:center;border-bottom:1px solid #444;height:56px;vertical-align:bottom;font-family:\'Brush Script MT\',cursive;font-style:italic;font-size:26px">' + esc_(p.signature) + '</td><td></td>' +
    '<td style="width:46%;text-align:center;border-bottom:1px solid #444;vertical-align:bottom;font-style:italic;color:#888">Muhammad Raj</td></tr>' +
    '<tr><td style="text-align:center;font-size:11px;color:#444;padding-top:4px">Student\'s signature</td><td></td><td style="text-align:center;font-size:11px;color:#444;padding-top:4px">Signature of Director</td></tr></table>' +
    '</body></html>';
}

/* ---------- helpers ---------- */
function sheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1b1035').setFontColor('#ffffff');
  }
  return sh;
}
function findRow_(sh, col, value) {
  if (sh.getLastRow() < 2) return 0;
  var vals = sh.getRange(2, col, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) if (String(vals[i][0]) === value) return i + 2;
  return 0;
}
function appendText_(sh, row) {
  var r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, row.length).setNumberFormat('@').setValues([row]);   // keep phone numbers as text
}
function folder_() {
  if (FOLDER_ID) return DriveApp.getFolderById(FOLDER_ID);
  var it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}
function safeName_(id, name) { return (id + ' - ' + String(name || 'student')).replace(/[\\\/:*?"<>|]/g, '').substring(0, 90); }
function stripData_(s) { return String(s || '').replace(/^data:[^,]*,/, ''); }
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function text_(s) { return ContentService.createTextOutput(s); }
