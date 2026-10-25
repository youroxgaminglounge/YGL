/**
 * YGL sync server (Google Apps Script)
 * ------------------------------------------------------------
 * Keeps each User ID's bookings, payments and revenue in this
 * Google Sheet so every device that logs in with the same
 * User ID and password sees the same data.
 *
 * SETUP (once):
 *  1. Open a new Google Sheet. Menu: Extensions > Apps Script.
 *  2. Delete the sample code, paste this whole file, press Save.
 *  3. Deploy > New deployment > type "Web app".
 *       Execute as: Me
 *       Who has access: Anyone
 *     Press Deploy and allow the permissions it asks for.
 *     (If Google says "unverified app", choose Advanced > Go to project.)
 *  4. Copy the Web app URL (it ends with /exec) and paste it into the
 *     YGL app when you tap "Sign in" for the first time.
 *
 * Keep the Web app URL private. Data is protected by each user's
 * password. Passwords are stored only as salted hashes.
 *
 * To change this code later: Deploy > Manage deployments > pencil icon
 * > Version: New version > Deploy.
 *
 * CUSTOMER BOOKING (new): customers open the booking page (book.html), register
 * with their mobile number and a PIN, and book slots. Their bookings are written
 * into the staff account's data, so they appear in the staff app after it syncs.
 * Also set the script time zone: Project settings > Time zone.
 */

// Optional: set a code that people must type to create a new account.
// Leave it empty ("") to let anyone with the URL create an account.
var SIGNUP_CODE = "";

var MAX_FAILS = 5;       // wrong passwords before a temporary lock
var LOCK_MINUTES = 10;   // length of that lock
var MAX_TOKENS = 12;     // devices that can stay signed in per user

// Which staff User ID owns the lounge data that customers book into.
// Leave empty ("") to use the first staff account that was created.
var LOUNGE_USER = "";
var CUSTOMER_MAX_UPCOMING = 3;   // upcoming bookings one customer may hold

function doGet() {
  return json_({ ok: true, msg: "YGL sync server is running" });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    var req = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    return json_(route_(req));
  } catch (err) {
    return json_({ ok: false, error: "Server error: " + (err && err.message ? err.message : err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function route_(req) {
  switch (req.action) {
    case "register": return register_(req);
    case "login":    return login_(req);
    case "logout":   return logout_(req);
    case "sync":     return sync_(req);
    case "ping":     return { ok: true };
    case "c_info":     return cInfo_(req);
    case "c_register": return cRegister_(req);
    case "c_login":    return cLogin_(req);
    case "c_logout":   return cLogout_(req);
    case "c_avail":    return cAvail_(req);
    case "c_book":     return cBook_(req);
    case "c_my":       return cMy_(req);
    case "c_cancel":   return cCancel_(req);
    default:         return { ok: false, error: "Unknown request" };
  }
}

/* ---------- users ---------- */
function users_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Users");
  if (!sh) {
    sh = ss.insertSheet("Users");
    sh.appendRow(["user", "salt", "passHash", "tokens", "created", "fails", "lockUntil"]);
    sh.setFrozenRows(1);
  }
  return sh;
}
function findUser_(sh, user) {
  var n = sh.getLastRow();
  if (n < 2) return null;
  var vals = sh.getRange(2, 1, n - 1, 7).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === user) return { row: i + 2, v: vals[i] };
  }
  return null;
}
function hash_(s) {
  var d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8);
  return d.map(function (b) { var h = (b < 0 ? b + 256 : b).toString(16); return h.length < 2 ? "0" + h : h; }).join("");
}
function cleanUser_(u) {
  u = String(u || "").trim().toLowerCase();
  return /^[a-z0-9_]{3,20}$/.test(u) ? u : null;
}
function parseTokens_(s) {
  try { var t = JSON.parse(s); return Array.isArray(t) ? t : []; } catch (e) { return []; }
}
function newToken_() { return Utilities.getUuid() + Utilities.getUuid(); }

function register_(req) {
  var user = cleanUser_(req.user);
  if (!user) return { ok: false, error: "User ID must be 3 to 20 letters, numbers or underscores." };
  var pass = String(req.pass || "");
  if (pass.length < 6) return { ok: false, error: "Password must be at least 6 characters." };
  if (SIGNUP_CODE && String(req.code || "") !== SIGNUP_CODE) return { ok: false, error: "Wrong sign-up code." };
  var sh = users_();
  if (findUser_(sh, user)) return { ok: false, error: "That User ID is already taken." };
  var salt = Utilities.getUuid();
  var token = newToken_();
  sh.appendRow([user, salt, hash_(salt + ":" + pass), JSON.stringify([hash_(token)]), Date.now(), 0, 0]);
  return { ok: true, token: token, user: user };
}

function login_(req) {
  var user = cleanUser_(req.user);
  var sh = users_();
  var f = user ? findUser_(sh, user) : null;
  if (!f) return { ok: false, error: "Wrong User ID or password." };
  var v = f.v, now = Date.now();
  if (Number(v[6]) > now) return { ok: false, error: "Too many wrong attempts. Try again in a few minutes." };
  var good = hash_(String(v[1]) + ":" + String(req.pass || "")) === String(v[2]);
  if (!good) {
    var fails = Number(v[5]) + 1, lockUntil = 0;
    if (fails >= MAX_FAILS) { lockUntil = now + LOCK_MINUTES * 60000; fails = 0; }
    sh.getRange(f.row, 6, 1, 2).setValues([[fails, lockUntil]]);
    return { ok: false, error: "Wrong User ID or password." };
  }
  var tokens = parseTokens_(v[3]);
  var token = newToken_();
  tokens.push(hash_(token));
  while (tokens.length > MAX_TOKENS) tokens.shift();
  sh.getRange(f.row, 4).setValue(JSON.stringify(tokens));
  sh.getRange(f.row, 6, 1, 2).setValues([[0, 0]]);
  return { ok: true, token: token, user: user };
}

function auth_(req) {
  var user = cleanUser_(req.user);
  if (!user) return null;
  var sh = users_();
  var f = findUser_(sh, user);
  if (!f) return null;
  var tokens = parseTokens_(f.v[3]);
  var h = hash_(String(req.token || ""));
  return tokens.indexOf(h) >= 0 ? { user: user, sh: sh, f: f, tokens: tokens } : null;
}

function logout_(req) {
  var a = auth_(req);
  if (!a) return { ok: true };
  var h = hash_(String(req.token || ""));
  var left = a.tokens.filter(function (t) { return t !== h; });
  a.sh.getRange(a.f.row, 4).setValue(JSON.stringify(left));
  return { ok: true };
}

/* ---------- data sync ---------- */
function sync_(req) {
  var a = auth_(req);
  if (!a) return { ok: false, auth: true, error: "Please log in again." };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = "d_" + a.user;
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(["key", "updated", "stored", "json"]);
    sh.setFrozenRows(1);
  }
  var n = sh.getLastRow();
  var vals = n > 1 ? sh.getRange(2, 1, n - 1, 4).getValues() : [];
  var idx = {};
  for (var i = 0; i < vals.length; i++) idx[String(vals[i][0])] = i;

  var now = Date.now(), won = {}, dirty = false;
  var ch = req.changes || [];
  for (var k = 0; k < ch.length; k++) {
    var c = ch[k];
    if (!c || typeof c.key !== "string" || !/^[slcu]:[A-Za-z0-9_\-]{1,60}$/.test(c.key)) continue;
    var u = Number(c.u) || 0;
    var json = JSON.stringify(c.rec === undefined ? null : c.rec);
    if (json.length > 45000) continue;
    var at = idx[c.key];
    if (at === undefined) {
      vals.push([c.key, u, now, json]);
      idx[c.key] = vals.length - 1;
      won[c.key] = 1; dirty = true;
    } else if (u > Number(vals[at][1])) {
      vals[at] = [c.key, u, now, json];
      won[c.key] = 1; dirty = true;
    }
  }
  if (dirty) sh.getRange(2, 1, vals.length, 4).setValues(vals);

  var since = Number(req.since) || 0, out = [];
  for (var j = 0; j < vals.length; j++) {
    var r = vals[j];
    if (won[String(r[0])]) continue;
    if (Number(r[2]) >= since) {
      try { out.push({ key: String(r[0]), u: Number(r[1]), rec: JSON.parse(r[3]) }); } catch (e) {}
    }
  }
  return { ok: true, now: now, records: out };
}


/* ================= customer booking ================= */
function ownerUser_() {
  var u = cleanUser_(LOUNGE_USER);
  if (u) return u;
  var sh = users_();
  if (sh.getLastRow() < 2) return null;
  return cleanUser_(sh.getRange(2, 1, 1, 1).getValues()[0][0]);
}
function customers_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Customers");
  if (!sh) {
    sh = ss.insertSheet("Customers");
    sh.appendRow(["phone", "name", "salt", "pinHash", "tokens", "created", "fails", "lockUntil"]);
    sh.setFrozenRows(1);
  }
  return sh;
}
function findCustomer_(sh, phone) {
  var n = sh.getLastRow();
  if (n < 2) return null;
  var vals = sh.getRange(2, 1, n - 1, 8).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === phone) return { row: i + 2, v: vals[i] };
  }
  return null;
}
function cleanPhone_(p) {
  var d = String(p || "").replace(/[^0-9]/g, "");
  return (d.length >= 7 && d.length <= 15) ? d : null;
}
function loadOwnerData_() {
  var user = ownerUser_();
  if (!user) return null;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = "d_" + user;
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(["key", "updated", "stored", "json"]);
    sh.setFrozenRows(1);
  }
  var n = sh.getLastRow();
  var vals = n > 1 ? sh.getRange(2, 1, n - 1, 4).getValues() : [];
  var idx = {};
  for (var i = 0; i < vals.length; i++) idx[String(vals[i][0])] = i;
  return { user: user, sh: sh, vals: vals, idx: idx, dirty: false };
}
function getRec_(d, key) {
  var i = d.idx[key];
  if (i === undefined) return null;
  try { return JSON.parse(d.vals[i][3]); } catch (e) { return null; }
}
function putRec_(d, key, rec, now) {
  rec._u = now;
  var row = [key, now, now, JSON.stringify(rec)];
  var i = d.idx[key];
  if (i === undefined) { d.vals.push(row); d.idx[key] = d.vals.length - 1; }
  else d.vals[i] = row;
  d.dirty = true;
}
function saveOwnerData_(d) {
  if (d.dirty && d.vals.length) d.sh.getRange(2, 1, d.vals.length, 4).setValues(d.vals);
}
function allSessions_(d) {
  var out = [];
  for (var i = 0; i < d.vals.length; i++) {
    var k = String(d.vals[i][0]);
    if (k.charAt(0) !== "s") continue;
    try { out.push(JSON.parse(d.vals[i][3])); } catch (e) {}
  }
  return out;
}
function obSettings_(cfg) {
  var ob = (cfg && cfg.ob) || {};
  return {
    on: ob.on !== false,
    days: Math.max(1, Math.min(30, Number(ob.days) || 7)),
    notice: Math.max(0, Number(ob.notice) || 15),
    max: Math.max(10, Math.min(480, Number(ob.max) || 180))
  };
}
function sessionBusy_(s, now) {
  if (s.status === "booked") return [Number(s.startAt), Number(s.endsAt)];
  if (s.status === "active") return [Number(s.checkIn || s.startAt), Math.max(Number(s.endsAt), now)];
  return null;
}
function cAuth_(req) {
  var phone = cleanPhone_(req.phone);
  if (!phone) return null;
  var sh = customers_();
  var f = findCustomer_(sh, phone);
  if (!f) return null;
  var tokens = parseTokens_(f.v[4]);
  var h = hash_(String(req.token || ""));
  return tokens.indexOf(h) >= 0 ? { phone: phone, sh: sh, f: f, tokens: tokens, name: String(f.v[1]) } : null;
}

function cInfo_(req) {
  var d = loadOwnerData_();
  if (!d) return { ok: false, error: "This lounge has not set up online booking yet." };
  var cfg = getRec_(d, "c:main");
  if (!cfg || !cfg.stations || !cfg.stations.length) {
    return { ok: false, error: "The lounge is still setting things up. Please try again later." };
  }
  var ob = obSettings_(cfg);
  return {
    ok: true, now: Date.now(), on: ob.on,
    lounge: cfg.lounge || "", currency: cfg.currency || "", perPlayer: cfg.perPlayer !== false,
    open: cfg.open, close: cfg.close, slotMin: cfg.slotMin,
    days: ob.days, notice: ob.notice, max: ob.max,
    stations: cfg.stations.map(function (s) { return { id: s.id, name: s.name, rate: s.rate }; })
  };
}

function cRegister_(req) {
  var name = String(req.name || "").replace(/\s+/g, " ").trim();
  if (name.length < 2 || name.length > 40) return { ok: false, error: "Please enter your name (2 to 40 letters)." };
  var phone = cleanPhone_(req.phone);
  if (!phone) return { ok: false, error: "Please enter a valid mobile number." };
  var pin = String(req.pin || "");
  if (!/^[0-9]{4,6}$/.test(pin)) return { ok: false, error: "PIN must be 4 to 6 digits." };
  var d = loadOwnerData_();
  if (!d) return { ok: false, error: "This lounge has not set up online booking yet." };
  var cfg = getRec_(d, "c:main");
  if (!obSettings_(cfg).on) return { ok: false, error: "Online booking is switched off right now." };
  var sh = customers_();
  if (findCustomer_(sh, phone)) return { ok: false, error: "This number is already registered. Please log in." };
  var salt = Utilities.getUuid(), token = newToken_(), now = Date.now();
  sh.appendRow([phone, name, salt, hash_(salt + ":" + pin), JSON.stringify([hash_(token)]), now, 0, 0]);
  var key = "u:" + phone;
  if (!getRec_(d, key)) { putRec_(d, key, { phone: phone, name: name, created: now }, now); saveOwnerData_(d); }
  return { ok: true, token: token, name: name, phone: phone };
}

function cLogin_(req) {
  var phone = cleanPhone_(req.phone);
  var sh = customers_();
  var f = phone ? findCustomer_(sh, phone) : null;
  if (!f) return { ok: false, error: "Wrong mobile number or PIN." };
  var v = f.v, now = Date.now();
  if (Number(v[7]) > now) return { ok: false, error: "Too many wrong attempts. Try again in a few minutes." };
  var good = hash_(String(v[2]) + ":" + String(req.pin || "")) === String(v[3]);
  if (!good) {
    var fails = Number(v[6]) + 1, lockUntil = 0;
    if (fails >= MAX_FAILS) { lockUntil = now + LOCK_MINUTES * 60000; fails = 0; }
    sh.getRange(f.row, 7, 1, 2).setValues([[fails, lockUntil]]);
    return { ok: false, error: "Wrong mobile number or PIN." };
  }
  var tokens = parseTokens_(v[4]), token = newToken_();
  tokens.push(hash_(token));
  while (tokens.length > MAX_TOKENS) tokens.shift();
  sh.getRange(f.row, 5).setValue(JSON.stringify(tokens));
  sh.getRange(f.row, 7, 1, 2).setValues([[0, 0]]);
  return { ok: true, token: token, name: String(v[1]), phone: phone };
}

function cLogout_(req) {
  var a = cAuth_(req);
  if (!a) return { ok: true };
  var h = hash_(String(req.token || ""));
  a.sh.getRange(a.f.row, 5).setValue(JSON.stringify(a.tokens.filter(function (t) { return t !== h; })));
  return { ok: true };
}

function cAvail_(req) {
  var d = loadOwnerData_();
  if (!d) return { ok: false, error: "This lounge has not set up online booking yet." };
  var from = Number(req.from), to = Number(req.to), now = Date.now();
  if (!(from > 0) || !(to > from) || to - from > 9 * 86400000) return { ok: false, error: "Bad date range." };
  var busy = [];
  allSessions_(d).forEach(function (s) {
    var iv = sessionBusy_(s, now);
    if (iv && iv[0] < to && iv[1] > from) busy.push({ s: s.stationId, a: iv[0], b: iv[1] });
  });
  return { ok: true, now: now, busy: busy };
}

function cBook_(req) {
  var a = cAuth_(req);
  if (!a) return { ok: false, auth: true, error: "Please log in again." };
  var d = loadOwnerData_();
  if (!d) return { ok: false, error: "This lounge has not set up online booking yet." };
  var cfg = getRec_(d, "c:main");
  if (!cfg || !cfg.stations) return { ok: false, error: "The lounge is still setting things up." };
  var ob = obSettings_(cfg);
  if (!ob.on) return { ok: false, error: "Online booking is switched off right now." };
  var urec = getRec_(d, "u:" + a.phone);
  if (urec && urec.blocked) return { ok: false, error: "Online booking is not available for this number. Please speak to the lounge." };

  var st = null;
  cfg.stations.forEach(function (s) { if (s.id === req.stationId) st = s; });
  if (!st) return { ok: false, error: "That station is not available." };
  var start = Number(req.start), end = Number(req.end), now = Date.now();
  var slot = Number(cfg.slotMin) || 60, open = Number(cfg.open), close = Number(cfg.close);
  if (!(start > 0) || !(end > start)) return { ok: false, error: "Please choose a time." };
  var mins = (end - start) / 60000;
  if (mins > ob.max) return { ok: false, error: "You can book up to " + ob.max + " minutes at a time." };
  if (mins % slot !== 0) return { ok: false, error: "Please choose whole slots." };
  if (start < now + ob.notice * 60000 - 60000) return { ok: false, error: "That time is too soon. Please choose a later slot." };
  if (start > now + ob.days * 86400000) return { ok: false, error: "You can book up to " + ob.days + " days ahead." };

  var tz = Session.getScriptTimeZone();
  var day1 = Utilities.formatDate(new Date(start), tz, "yyyyMMdd"), day2 = Utilities.formatDate(new Date(end - 1), tz, "yyyyMMdd");
  var sMin = Number(Utilities.formatDate(new Date(start), tz, "H")) * 60 + Number(Utilities.formatDate(new Date(start), tz, "m"));
  var eMin = sMin + mins;
  if (day1 !== day2 || sMin < open * 60 || eMin > close * 60 || (sMin - open * 60) % slot !== 0) {
    return { ok: false, error: "That time is outside the lounge opening hours or slot times." };
  }
  var players = Math.max(1, Math.min(8, parseInt(req.players, 10) || 1));
  var sessions = allSessions_(d), upcoming = 0;
  for (var i = 0; i < sessions.length; i++) {
    var s = sessions[i];
    if (s.phone === a.phone && s.status === "booked" && Number(s.endsAt) > now) upcoming++;
    if (s.stationId === st.id) {
      var iv = sessionBusy_(s, now);
      if (iv && iv[0] < end && iv[1] > start) return { ok: false, taken: true, error: "Sorry, that slot was just taken. Please pick another." };
    }
  }
  if (upcoming >= CUSTOMER_MAX_UPCOMING) return { ok: false, error: "You already have " + CUSTOMER_MAX_UPCOMING + " upcoming bookings. Cancel one or visit the lounge." };

  var id = "o" + now.toString(36) + Math.random().toString(36).slice(2, 6);
  var rec = {
    id: id, stationId: st.id, name: a.name, phone: a.phone, players: players, status: "booked",
    startAt: start, endsAt: end, checkIn: null, checkOut: null, rate: st.rate, amount: null, pay: null,
    walkIn: false, source: "online", createdAt: now
  };
  putRec_(d, "s:" + id, rec, now);
  saveOwnerData_(d);
  return { ok: true, booking: { id: id, stationId: st.id, startAt: start, endsAt: end, players: players, status: "booked" } };
}

function cMy_(req) {
  var a = cAuth_(req);
  if (!a) return { ok: false, auth: true, error: "Please log in again." };
  var d = loadOwnerData_();
  if (!d) return { ok: false, error: "This lounge has not set up online booking yet." };
  var now = Date.now(), out = [];
  allSessions_(d).forEach(function (s) {
    if (s.phone !== a.phone) return;
    if (s.status !== "booked" && s.status !== "active" && s.status !== "cancelled") return;
    if (Number(s.endsAt) < now - 2 * 86400000) return;
    out.push({ id: s.id, stationId: s.stationId, startAt: s.startAt, endsAt: s.endsAt, players: s.players, status: s.status });
  });
  out.sort(function (x, y) { return y.startAt - x.startAt; });
  return { ok: true, now: now, bookings: out.slice(0, 20) };
}

function cCancel_(req) {
  var a = cAuth_(req);
  if (!a) return { ok: false, auth: true, error: "Please log in again." };
  var d = loadOwnerData_();
  if (!d) return { ok: false, error: "This lounge has not set up online booking yet." };
  var key = "s:" + String(req.id || "").replace(/[^A-Za-z0-9_\-]/g, "");
  var s = getRec_(d, key), now = Date.now();
  if (!s || s.phone !== a.phone) return { ok: false, error: "Booking not found." };
  if (s.status !== "booked") return { ok: false, error: "This booking can no longer be cancelled here." };
  if (Number(s.startAt) < now) return { ok: false, error: "This booking has already started." };
  s.status = "cancelled"; s.cancelledBy = "customer";
  putRec_(d, key, s, now);
  saveOwnerData_(d);
  return { ok: true };
}
