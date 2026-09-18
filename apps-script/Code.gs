// Backend des Templiers : stocke l'état partagé dans l'onglet "state" de la Google Sheet.
// À coller dans Extensions > Apps Script de la Sheet, puis Déployer > Application web.

const CODE = "CHANGE-MOI"; // code d'accès partagé avec les colocs (à modifier avant de déployer)

const SEED = {
  v: 1,
  house: "Les Templiers",
  people: ["Judicaël"],
  joinRequests: [],
  admin: "Judicaël",
  anchor: "2026-09-19",
  weeks: {},
  choreDebts: [],
  ledger: []
};

const CHUNK = 40000; // une cellule Google Sheet est limitée à 50 000 caractères

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("state");
  if (!sh) {
    sh = ss.insertSheet("state");
    save_(SEED, 1, sh);
  }
  return sh;
}

function load_() {
  const sh = sheet_();
  const rev = Number(sh.getRange("A1").getValue()) || 0;
  const last = sh.getLastRow();
  const parts = last >= 1 ? sh.getRange(1, 2, last, 1).getValues().map(function (r) { return r[0]; }) : [];
  return { rev: rev, state: JSON.parse(parts.join("") || "null") };
}

function save_(state, rev, sh) {
  sh = sh || sheet_();
  const json = JSON.stringify(state);
  const chunks = [];
  for (let i = 0; i < json.length; i += CHUNK) chunks.push([json.substring(i, i + CHUNK)]);
  sh.getRange("B:B").clearContent();
  sh.getRange(1, 2, chunks.length, 1).setValues(chunks);
  sh.getRange("A1").setValue(rev);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.code !== CODE) return json_({ ok: false, error: "unauthorized" });
  const s = load_();
  return json_({ ok: true, rev: s.rev, state: s.state });
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: "bad_request" });
  }
  if (body.code !== CODE) return json_({ ok: false, error: "unauthorized" });
  const st = body.state;
  if (!st || !Array.isArray(st.people) || !st.anchor || JSON.stringify(st).length > 400000) {
    return json_({ ok: false, error: "bad_request" });
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const s = load_();
    if (body.baseRev !== s.rev) return json_({ ok: false, error: "conflict", rev: s.rev, state: s.state });
    save_(st, s.rev + 1);
    return json_({ ok: true, rev: s.rev + 1 });
  } finally {
    lock.releaseLock();
  }
}
