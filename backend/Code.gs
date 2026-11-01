/**
 * GM·prépa — account server (Google Apps Script bound to a Google Sheet).
 *
 * The spreadsheet is the database. Sheets are created automatically:
 *   users    — login, salt, pass_hash, token, created_at, last_login
 *   progress — one row per user: summary columns + the full progress JSON split into data_1..data_8
 *   answers  — one row per user and text question: the written answer and its score
 * Any sheet can be downloaded as CSV (File → Download → .csv).
 *
 * Deploy: Deploy → New deployment → Web app → Execute as: Me → Who has access: Anyone.
 */
const CHUNK = 45000;   // a cell holds at most 50 000 characters
const MAX_CHUNKS = 8;
const HEAD = {
  users: ["login", "salt", "pass_hash", "token", "created_at", "last_login"],
  progress: ["login", "updated_at", "xp", "answers", "accuracy_pct", "best_exam", "best_marathon"]
    .concat(Array.from({ length: MAX_CHUNKS }, (_, i) => "data_" + (i + 1))),
  answers: ["login", "text_id", "question", "answer", "score_pct", "updated_at"]
};

function doGet() {
  return out({ ok: true, service: "gm-prepa", time: new Date().toISOString() });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out({ ok: false, error: "bad_json" }); }
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    return out(handle(req));
  } catch (err) {
    return out({ ok: false, error: "server", detail: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function handle(r) {
  const login = String(r.login || "").trim().toLowerCase();
  if (!/^[a-z0-9._-]{3,30}$/.test(login)) return { ok: false, error: "bad_login" };
  const users = sheet("users");
  const row = findRow(users, login);

  if (r.action === "register") {
    if (String(r.password || "").length < 6) return { ok: false, error: "bad_password" };
    if (row) return { ok: false, error: "exists" };
    const salt = Utilities.getUuid();
    const token = newToken();
    const now = new Date().toISOString();
    users.appendRow([login, salt, hash(salt, String(r.password)), token, now, now]);
    return { ok: true, login: login, token: token, data: null };
  }

  if (!row) return { ok: false, error: "no_user" };
  const u = users.getRange(row, 1, 1, HEAD.users.length).getValues()[0];

  if (r.action === "login") {
    if (hash(u[1], String(r.password || "")) !== u[2]) {
      Utilities.sleep(800); // slows down password guessing
      return { ok: false, error: "wrong_password" };
    }
    let token = String(u[3] || "");
    if (!token) { token = newToken(); users.getRange(row, 4).setValue(token); }
    users.getRange(row, 6).setValue(new Date().toISOString());
    const p = readProgress(login);
    return { ok: true, login: login, token: token, data: p.data, updated: p.updated };
  }

  if (!r.token || String(r.token) !== String(u[3])) return { ok: false, error: "auth" };

  if (r.action === "load") {
    const p = readProgress(login);
    return { ok: true, data: p.data, updated: p.updated };
  }

  if (r.action === "save") {
    const json = JSON.stringify(r.data || {});
    if (json.length > CHUNK * MAX_CHUNKS) return { ok: false, error: "too_big" };
    const now = new Date().toISOString();
    const s = r.summary || {};
    const chunks = [];
    // "J" prefix keeps a chunk from ever being read by Sheets as a formula or a number
    for (let i = 0; i < MAX_CHUNKS; i++) {
      const part = json.slice(i * CHUNK, (i + 1) * CHUNK);
      chunks.push(part ? "J" + part : "");
    }
    const vals = [login, now, num(s.xp), num(s.answers), num(s.accuracy),
      s.bestExam == null ? "" : num(s.bestExam), num(s.bestMara)].concat(chunks);
    const ps = sheet("progress");
    const pr = findRow(ps, login);
    if (pr) ps.getRange(pr, 1, 1, vals.length).setValues([vals]);
    else ps.appendRow(vals);
    if (Array.isArray(r.answers)) writeAnswers(login, r.answers, now);
    return { ok: true, updated: now };
  }

  return { ok: false, error: "bad_action" };
}

function readProgress(login) {
  const ps = sheet("progress");
  const pr = findRow(ps, login);
  if (!pr) return { data: null, updated: null };
  const v = ps.getRange(pr, 1, 1, HEAD.progress.length).getValues()[0];
  const json = v.slice(7).map(c => String(c || "")).map(c => c.charAt(0) === "J" ? c.slice(1) : c).join("");
  let data = null;
  try { data = json ? JSON.parse(json) : null; } catch (e) { data = null; }
  return { data: data, updated: v[1] ? String(v[1]) : null };
}

function writeAnswers(login, list, now) {
  const sh = sheet("answers");
  const last = sh.getLastRow();
  const index = {};
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, 3).getValues().forEach((row, i) => {
      index[String(row[0]).toLowerCase() + "|" + row[1] + "|" + row[2]] = i + 2;
    });
  }
  const append = [];
  list.slice(0, 200).forEach(a => {
    const textId = String(a.text_id || "").slice(0, 40);
    const q = num(a.question);
    if (!textId || !q) return;
    const vals = [login, textId, q, safeText(a.answer), num(a.score), now];
    const at = index[login + "|" + textId + "|" + q];
    if (at) sh.getRange(at, 1, 1, vals.length).setValues([vals]);
    else append.push(vals);
  });
  if (append.length) sh.getRange(sh.getLastRow() + 1, 1, append.length, append[0].length).setValues(append);
}

function sheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, HEAD[name].length).setValues([HEAD[name]]).setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.getRange("A:A").setNumberFormat("@");
  }
  return sh;
}

function findRow(sh, login) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const vals = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]).toLowerCase() === login) return i + 2;
  return 0;
}

function hash(salt, password) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + password, Utilities.Charset.UTF_8);
  return "sha256:" + bytes.map(b => ((b + 256) % 256).toString(16).padStart(2, "0")).join("");
}

function newToken() { return Utilities.getUuid() + Utilities.getUuid(); }
function num(v) { const n = Number(v); return isFinite(n) ? n : 0; }
function safeText(v) {
  const s = String(v || "").slice(0, 5000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;   // never let an answer become a formula
}
function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** Run once from the editor to create the three sheets and check permissions. */
function setup() { ["users", "progress", "answers"].forEach(sheet); }
