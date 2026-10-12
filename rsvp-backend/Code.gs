/**
 * RSVP backend for mazhaturns10.com
 *
 * Google is storage only: each RSVP becomes a row in the "RSVPs" tab of the
 * Sheet this script is attached to. The site reads it two ways:
 *   - public:  head count + first names of kids who are coming
 *   - host:    every reply with contacts and notes, only with the private host key
 *
 * The host key is created on first use and emailed to the Sheet owner as a
 * link to mazhaturns10.com/host. It lives in Script Properties, never in code.
 *
 * Setup: open the Sheet → Extensions → Apps Script → paste this file →
 * Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone.
 */

const SHEET_NAME = "RSVPs";
const SITE = "https://mazhaturns10.com";
const EMAIL_ME_ON_EACH_RSVP = true;
const HEADERS = ["Submitted", "Child", "Coming?", "Parent / guardian", "Phone or email", "Notes"];

function doPost(e) {
  let d;
  try { d = JSON.parse((e && e.postData && e.postData.contents) || "{}"); }
  catch (err) { return json_({ ok: false, error: "bad_request" }); }

  if (d.website) return json_(public_()); // spam bot filled the hidden field

  const child = clean_(d.child, 80), parent = clean_(d.parent, 80);
  const contact = clean_(d.contact, 120), notes = clean_(d.notes, 300);
  if (!child || !parent || !contact || (d.coming !== "yes" && d.coming !== "no")) {
    return json_({ ok: false, error: "missing_fields" });
  }
  const coming = d.coming === "yes" ? "Yes" : "No";

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    sheet_().appendRow([new Date(), child, coming, parent, contact, notes]);
  } finally {
    lock.releaseLock();
  }

  if (EMAIL_ME_ON_EACH_RSVP) {
    MailApp.sendEmail(
      owner_(),
      "RSVP: " + child + (coming === "Yes" ? " is coming" : " can't make it"),
      "Child: " + child + "\nComing: " + coming + "\nParent: " + parent +
      "\nContact: " + contact + (notes ? "\nNotes: " + notes : "") +
      "\n\nAll replies: " + hostLink_()
    );
  }
  return json_(public_());
}

function doGet(e) {
  const key = e && e.parameter && e.parameter.host;
  if (key) {
    if (key !== hostKey_()) return json_({ ok: false, error: "bad_key" });
    return json_({ ok: true, replies: latest_() });
  }
  return json_(public_());
}

// Optional: run once from the editor to get the host link emailed right away.
function setup() {
  sheet_();
  Logger.log("Host page: " + hostLink_());
}

// ---- data ----

// One entry per child (matched on name), keeping their latest answer.
function latest_() {
  const rows = sheet_().getDataRange().getValues().slice(1);
  const byChild = {};
  rows.forEach(r => {
    const k = String(r[1]).trim().toLowerCase();
    if (!k) return;
    const prev = byChild[k];
    byChild[k] = {
      child: String(r[1]), coming: r[2], parent: String(r[3]), contact: String(r[4]), notes: String(r[5]),
      updated: r[0], firstReplied: prev ? prev.firstReplied : r[0], changes: prev ? prev.changes + 1 : 0
    };
  });
  return Object.keys(byChild).map(k => byChild[k])
    .sort((a, b) => new Date(a.firstReplied) - new Date(b.firstReplied));
}

function public_() {
  const yes = latest_().filter(r => r.coming === "Yes");
  return { ok: true, count: yes.length, names: yes.map(r => firstName_(r.child)) };
}

function firstName_(s) {
  const f = String(s).trim().split(/\s+/)[0] || "";
  return f.charAt(0).toUpperCase() + f.slice(1);
}

// ---- host key ----

function hostKey_() {
  const props = PropertiesService.getScriptProperties();
  let k = props.getProperty("HOST_KEY");
  if (k) return k;
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    k = props.getProperty("HOST_KEY");
    if (!k) {
      k = Utilities.getUuid().replace(/-/g, "") + Utilities.getUuid().replace(/-/g, "").slice(0, 12);
      props.setProperty("HOST_KEY", k);
      MailApp.sendEmail(owner_(), "Your private RSVP page for Mazha's party",
        "Bookmark this link to see every RSVP, with contacts and notes:\n\n" + SITE + "/host#" + k +
        "\n\nAnyone with this link can see all replies, so keep it to yourself.");
    }
    return k;
  } finally {
    lock.releaseLock();
  }
}

function hostLink_() { return SITE + "/host#" + hostKey_(); }

// ---- helpers ----

function owner_() { return Session.getEffectiveUser().getEmail(); }

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
