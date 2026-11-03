// Sleep Journal backend for Google Sheets.
// 1. Create a Google Sheet. Copy its ID from the URL (the long part between /d/ and /edit).
// 2. Paste the ID and a long random phrase below. Use the same phrase as SYNC_KEY in index.html.
// 3. Deploy as a Web app (Execute as: Me, Who has access: Anyone). Create a NEW deployment after every change.

var SHEET_ID = "PASTE_SHEET_ID_HERE";
var KEY = "PICK-A-LONG-RANDOM-PHRASE";
var TAB = "SleepData";

var HEAD = ["date", "quality", "minutes_asleep", "bed", "wake", "minutes_to_fall_asleep", "minutes_awake_at_night",
  "melatonin", "melatonin_mg", "stress", "caffeine_after_noon", "alcohol", "problems", "factors", "notes", "updated", "json"];
var JSON_COL = HEAD.length; // json is the last column

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function tab_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var s = ss.getSheetByName(TAB);
  if (!s) {
    s = ss.insertSheet(TAB);
    s.getRange("A:A").setNumberFormat("@");
    s.appendRow(HEAD);
    s.setFrozenRows(1);
  }
  return s;
}

function norm_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return String(v);
}

function findRow_(s, date) {
  var last = s.getLastRow();
  if (last < 2) return -1;
  var vals = s.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (norm_(vals[i][0]) === date) return i + 2;
  }
  return -1;
}

function rowFor_(e) {
  return [e.date, e.quality || "", typeof e.minutes === "number" ? e.minutes : "", e.bed || "", e.wake || "",
    e.latency === null || e.latency === undefined ? "" : e.latency,
    e.awake === null || e.awake === undefined ? "" : e.awake,
    e.melatonin ? "yes" : "no", e.dose || "", e.stress || "", e.caffeine || 0, e.alcohol || 0,
    (e.issues || []).join("; "), (e.factors || []).join("; "), e.notes || "", e.updated || "", JSON.stringify(e)];
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.key !== KEY) return out_({ ok: false, error: "bad key" });
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var s = tab_();
    if (p.action === "list") {
      var last = s.getLastRow();
      var list = [];
      if (last > 1) {
        var vals = s.getRange(2, JSON_COL, last - 1, 1).getValues();
        for (var i = 0; i < vals.length; i++) {
          try { list.push(JSON.parse(vals[i][0])); } catch (x) {}
        }
      }
      return out_({ ok: true, entries: list });
    }
    if (p.action === "save") {
      var en = JSON.parse(p.data);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(en.date)) return out_({ ok: false, error: "bad date" });
      var row = findRow_(s, en.date);
      var cells = rowFor_(en);
      if (row < 0) {
        row = s.getLastRow() + 1;
        s.getRange(row, 1).setNumberFormat("@");
      }
      s.getRange(row, 1, 1, cells.length).setValues([cells]);
      return out_({ ok: true });
    }
    if (p.action === "delete") {
      var r = findRow_(s, String(p.date));
      if (r > 0) s.deleteRow(r);
      return out_({ ok: true });
    }
    return out_({ ok: false, error: "unknown action" });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
