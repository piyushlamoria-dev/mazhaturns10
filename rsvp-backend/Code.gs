/**
 * RSVP backend for mazhaturns10.com
 *
 * Saves each RSVP as a row in the "RSVPs" tab of the Google Sheet this script
 * is attached to, and emails the sheet owner. The public site only ever gets
 * back a head count, never names or contact details.
 *
 * Setup: open the Sheet → Extensions → Apps Script → paste this file →
 * Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone.
 */

const SHEET_NAME = "RSVPs";
const EMAIL_ME_ON_EACH_RSVP = true;
const HEADERS = ["Submitted", "Child", "Coming?", "Parent / guardian", "Phone or email", "Notes"];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    let d;
    try { d = JSON.parse((e && e.postData && e.postData.contents) || "{}"); }
    catch (err) { return json_({ ok: false, error: "bad_request" }); }

    if (d.website) return json_({ ok: true, count: count_() }); // spam bot filled the hidden field

    const child = clean_(d.child, 80), parent = clean_(d.parent, 80);
    const contact = clean_(d.contact, 120), notes = clean_(d.notes, 300);
    if (!child || !parent || !contact || (d.coming !== "yes" && d.coming !== "no")) {
      return json_({ ok: false, error: "missing_fields" });
    }
    const coming = d.coming === "yes" ? "Yes" : "No";
    sheet_().appendRow([new Date(), child, coming, parent, contact, notes]);

    if (EMAIL_ME_ON_EACH_RSVP) {
      MailApp.sendEmail(
        Session.getEffectiveUser().getEmail(),
        "RSVP: " + child + (coming === "Yes" ? " is coming" : " can't make it"),
        "Child: " + child + "\nComing: " + coming + "\nParent: " + parent +
        "\nContact: " + contact + (notes ? "\nNotes: " + notes : "") +
        "\n\nAll RSVPs: " + SpreadsheetApp.getActiveSpreadsheet().getUrl()
      );
    }
    return json_({ ok: true, count: count_() });
  } finally {
    lock.releaseLock();
  }
}

// Public head count for the site: number of children whose latest answer is Yes.
function doGet() {
  return json_({ count: count_() });
}

function count_() {
  const rows = sheet_().getDataRange().getValues().slice(1);
  const latest = {};
  rows.forEach(r => { const k = String(r[1]).trim().toLowerCase(); if (k) latest[k] = r[2]; });
  return Object.keys(latest).filter(k => latest[k] === "Yes").length;
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
  }
  return sh;
}

// Trim, cap length, and stop text being read as a spreadsheet formula.
function clean_(v, max) {
  const s = String(v || "").trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
