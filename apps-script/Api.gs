/**
 * Cash Recon Desk — dashboard API (second file in the "Brinks Cash Recon Engine" project).
 * Uses SHEET_ID / TABS / helpers from Code.gs.
 *
 * Deploy: Deploy → New deployment → Web app · Execute as: Me · Who has access: Anyone
 * ("Anyone" only means the URL answers. Everything except sign-in needs a signed session
 *  for an email listed as Active on the Users tab.)
 *
 * Roles (Users tab, column C):
 *   Admin         everything + manage users
 *   Editor        all stores, can log manual adjustments / notes
 *   Super Viewer  all stores, read-only
 *   DM Viewer     only their stores (Store Map "DM Email" = their email, or DM names in Scope), read-only
 *   Store Viewer  only their stores (Store Map "SM Email" = their email, or WZ #s in Scope), read-only
 */

var DESK = { NAME: 'Cash Recon Desk', SESSION_DAYS: 30, CODE_MINUTES: 10, USERS: 'Users', ACT: 'Desk_Activity' };
var DESK_ROLES = ['Admin', 'Editor', 'Super Viewer', 'DM Viewer', 'Store Viewer'];
var OPEN_STATUSES = { 'MISSING DEPOSIT': 1, 'SHORT': 1, 'OVER': 1, 'BRINKS ONLY': 1 };

/* ---------- one-time setup (select deskSetup → Run) ---------- */
function deskSetup() {
  var ss = ss_();
  var u = ss.getSheetByName(DESK.USERS);
  if (!u) {
    u = ss.insertSheet(DESK.USERS);
    u.getRange(1, 1, 1, 6).setValues([['Email', 'Name', 'Role', 'Scope', 'Active', 'Added']]).setFontWeight('bold').setBackground('#f1f3f4');
    u.appendRow(['jon.sanchez@nyconnectionli.com', 'Jonathan Sanchez', 'Admin', '', true, new Date()]);
    u.setFrozenRows(1);
    u.getRange('C2:C500').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(DESK_ROLES, true).build());
    u.getRange('E2:E500').insertCheckboxes();
    u.getRange('D1').setNote('DM Viewer: DM name(s) as written in Store Map, comma-separated (blank = match by DM Email).\n' +
      'Store Viewer: WZ #s, comma-separated (blank = match by SM Email).');
  }
  var a = ss.getSheetByName(DESK.ACT);
  if (!a) {
    a = ss.insertSheet(DESK.ACT);
    a.getRange(1, 1, 1, 7).setValues([['Time', 'User', 'Date', 'WZ #', 'Store', 'Change', 'Detail']]).setFontWeight('bold').setBackground('#f1f3f4');
    a.setFrozenRows(1);
  }
  deskSecret_();
  desk2Setup();
  return 'Desk setup complete';
}

/* ---------- HTTP ---------- */
function doGet() { return deskJson_({ ok: true, app: DESK.NAME }); }

function doPost(e) {
  var b = {};
  try { b = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { return deskJson_({ ok: false, error: 'bad_request' }); }
  try {
    if (b.action === 'requestCode') return deskJson_(deskRequestCode_(b.email));
    if (b.action === 'verifyCode') return deskJson_(deskVerifyCode_(b.email, b.code));
    if (b.action === 'requestAccess') return deskJson_(desk2RequestAccess_(b));
    var me = deskAuth_(b.token);
    if (!me) return deskJson_({ ok: false, error: 'signed_out' });
    switch (b.action) {
      case 'bootstrap': return deskJson_(deskBootstrap_(me));
      case 'day':       return deskJson_(deskDay_(me, String(b.date || '')));
      case 'store':     return deskJson_(deskStore_(me, String(b.wz || ''), Number(b.days) || 30));
      case 'saveAdj':   return deskJson_(deskSaveAdj_(me, b));
      case 'decideAdj': return deskJson_(deskDecideAdj_(me, b));
      case 'users':     return deskJson_(deskUsers_(me));
      case 'saveUser':  return deskJson_(deskSaveUser_(me, b));
    }
    var more = desk2Route_(me, b);
    if (more) return deskJson_(more);
    return deskJson_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return deskJson_({ ok: false, error: 'server_error', message: String(err && err.message || err) });
  }
}
function deskJson_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- auth ---------- */
function deskEmail_(e) { return String(e || '').trim().toLowerCase(); }
function deskTrue_(v) { return v === true || /^(true|yes|y|1)$/i.test(String(v).trim()); }

function deskFindUser_(email) {
  var sh = ss_().getSheetByName(DESK.USERS);
  if (!sh) return null;
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (deskEmail_(v[i][0]) !== email) continue;
    if (!deskTrue_(v[i][4])) return null;
    var role = DESK_ROLES.indexOf(String(v[i][2]).trim()) >= 0 ? String(v[i][2]).trim() : 'Store Viewer';
    return { email: email, name: String(v[i][1] || email.split('@')[0]), role: role, scope: String(v[i][3] || '').trim(),
      canEdit: role === 'Admin' || role === 'Editor', isAdmin: role === 'Admin' };
  }
  return null;
}

function deskRequestCode_(raw) {
  var email = deskEmail_(raw);
  if (!/^[^\s@]+@[^\s@]+\.[a-z]{2,}$/i.test(email)) return { ok: false, error: 'bad_email' };
  var cache = CacheService.getScriptCache();
  if (cache.get('desk_sent:' + email)) return { ok: true };
  var user = deskFindUser_(email);
  if (user) {
    var code = String(Math.floor(100000 + Math.random() * 900000));
    cache.put('desk_code:' + email, JSON.stringify({ code: code, tries: 0 }), DESK.CODE_MINUTES * 60);
    cache.put('desk_sent:' + email, '1', 60);
    MailApp.sendEmail({ to: email, subject: DESK.NAME + ' sign-in code: ' + code,
      htmlBody: '<div style="font-family:-apple-system,Helvetica,Arial,sans-serif;font-size:15px;color:#1c1c1e">' +
        '<p>Hi ' + user.name.split(' ')[0] + ',</p><p>Your ' + DESK.NAME + ' sign-in code is</p>' +
        '<p style="font-size:30px;font-weight:700;letter-spacing:6px;margin:12px 0">' + code + '</p>' +
        '<p style="color:#6e6e73">It expires in ' + DESK.CODE_MINUTES + ' minutes. If you didn\'t try to sign in, ignore this email.</p></div>' });
  }
  return { ok: true };                      // same answer either way — never reveals who has access
}

function deskVerifyCode_(raw, rawCode) {
  var email = deskEmail_(raw), cache = CacheService.getScriptCache(), k = 'desk_code:' + email;
  var rec = cache.get(k);
  if (!rec) return { ok: false, error: 'code_expired' };
  var r = JSON.parse(rec);
  if (r.tries >= 5) { cache.remove(k); return { ok: false, error: 'too_many_tries' }; }
  if (String(rawCode || '').replace(/\D/g, '') !== r.code) {
    r.tries++; cache.put(k, JSON.stringify(r), DESK.CODE_MINUTES * 60);
    return { ok: false, error: 'wrong_code', triesLeft: 5 - r.tries };
  }
  cache.remove(k);
  var user = deskFindUser_(email);
  if (!user) return { ok: false, error: 'no_access' };
  var exp = Date.now() + DESK.SESSION_DAYS * 864e5;
  deskLog_(user, '', '', '', 'Signed in', '');
  return { ok: true, token: deskSign_(email + '|' + exp), expires: exp, me: user };
}

function deskAuth_(token) {
  if (!token || String(token).indexOf('.') < 0) return null;
  var p = String(token).split('.')[0], payload;
  try { payload = Utilities.newBlob(Utilities.base64DecodeWebSafe(p)).getDataAsString(); } catch (e) { return null; }
  if (deskSign_(payload) !== token) return null;
  var parts = payload.split('|');
  if (!parts[1] || Date.now() > Number(parts[1])) return null;
  return deskFindUser_(parts[0]);           // unticking Active locks them out immediately
}
function deskSecret_() {
  var props = PropertiesService.getScriptProperties(), s = props.getProperty('DESK_SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('DESK_SECRET', s); }
  return s;
}
function deskSign_(payload) {
  return Utilities.base64EncodeWebSafe(payload) + '.' +
    Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(payload, deskSecret_()));
}

/* ---------- scope ---------- */
function deskScope_(me) {
  var map = loadStoreMap_(), rows = ss_().getSheetByName(TABS.MAP).getDataRange().getValues();
  var h = rows[0].map(String), iW = h.indexOf('WZ Store ID'), iSM = h.indexOf('SM Email');
  var all = Object.keys(map.byWZ), allowed = {};
  if (me.role === 'Admin' || me.role === 'Editor' || me.role === 'Super Viewer') all.forEach(function (w) { allowed[w] = 1; });
  else if (me.role === 'DM Viewer') {
    var names = me.scope ? me.scope.split(',').map(function (s) { return s.trim().toLowerCase(); }).filter(String) : [];
    all.forEach(function (w) {
      var s = map.byWZ[w];
      if (names.length ? names.some(function (n) { return String(s.dm).toLowerCase().indexOf(n) >= 0; })
                       : String(s.dmEmail).toLowerCase().split(/[,;\s\/]+/).indexOf(me.email) >= 0) allowed[w] = 1;
    });
  } else {
    var wzs = me.scope ? me.scope.toUpperCase().split(/[,;\s]+/).filter(String) : [];
    if (wzs.length) wzs.forEach(function (w) { if (map.byWZ[w]) allowed[w] = 1; });
    else if (iSM >= 0) rows.slice(1).forEach(function (r) {
      if (String(r[iSM]).toLowerCase().split(/[,;\s\/]+/).indexOf(me.email) >= 0) allowed[String(r[iW]).trim()] = 1;
    });
  }
  return { map: map, allowed: allowed };
}

/* ---------- data ---------- */
function deskLogRows_() { return readTab_(TABS.LOG); }
function deskRow_(r) {
  return { date: normDate_(r[0]), wz: String(r[1]), store: String(r[2]), company: String(r[3]), dm: String(r[4]),
    rq: Number(r[5]) || 0, brinks: Number(r[6]) || 0, diff: Number(r[7]) || 0, status: String(r[8]), bags: Number(r[9]) || 0,
    notes: String(r[10] || ''), adj: r[11] === '' ? '' : Number(r[11]) || 0, adjReason: String(r[12] || ''), adjBy: String(r[13] || ''),
    note: String(r[14] || ''), finalDiff: r[15] === '' ? Number(r[7]) || 0 : Number(r[15]) || 0, finalStatus: String(r[16] || r[8]),
    updated: r[17] instanceof Date ? r[17].getTime() : null };
}
function deskIsOpen_(row) {
  var fs = row.finalStatus;
  if (/RESOLVED/.test(fs)) return false;
  return !!OPEN_STATUSES[fs.replace(/ \(adj\)$/, '')];
}

function deskBootstrap_(me) {
  var sc = deskScope_(me), rows = deskLogRows_().map(deskRow_).filter(function (r) { return sc.allowed[r.wz]; });
  var byDate = {};
  rows.forEach(function (r) {
    var d = byDate[r.date] = byDate[r.date] || { date: r.date, total: 0, matched: 0, open: 0, missing: 0, short: 0, over: 0, brinksOnly: 0, openAmt: 0 };
    if (r.status === 'NO ACTIVITY') return;
    d.total++;
    if (!deskIsOpen_(r)) { d.matched++; return; }
    d.open++; d.openAmt += Math.abs(r.finalDiff);
    var base = r.finalStatus.replace(/ \(adj\)$/, '');
    if (base === 'MISSING DEPOSIT') d.missing++; else if (base === 'SHORT') d.short++; else if (base === 'OVER') d.over++; else d.brinksOnly++;
  });
  var days = Object.keys(byDate).sort().reverse().slice(0, 90).map(function (k) { var d = byDate[k]; d.openAmt = round2_(d.openAmt); return d; });
  // Open items older than the latest day (still unresolved) — the "carry-over" list
  var latest = days.length ? days[0].date : '';
  var carry = rows.filter(function (r) { return r.date < latest && r.date >= addDays_(latest, -30) && deskIsOpen_(r); })
    .sort(function (a, b) { return a.date < b.date ? 1 : -1; });
  var stores = Object.keys(sc.allowed).filter(function (w) { return sc.map.byWZ[w]; }).map(function (w) {
    var s = sc.map.byWZ[w]; return { wz: w, name: String(s.name), company: String(s.company || ''), dm: s.dm, active: s.active };
  }).sort(function (a, b) { return a.name < b.name ? -1 : 1; });
  var cfg = getCfg_(), run = readTab_(TABS.RUN).slice(-1)[0];
  var pending = 0;
  if (me.isAdmin) readTab_(DESK2.REQ).forEach(function (r) { if (String(r[6]) === 'Pending') pending++; });
  return { ok: true, me: me, days: days, carry: carry, stores: stores, tolerance: Number(cfg.TOLERANCE) || 5,
    safeLimit: Number(cfg.SAFE_LIMIT) || 5000, locks: desk2Locks_(), pendingRequests: pending, pendingAdj: me.isAdmin ? deskPending_().length : 0, testMode: isTest_(cfg), lastRun: run ? { at: run[0] instanceof Date ? run[0].getTime() : null, msg: String(run[1]).slice(0, 160) } : null };
}

function deskDay_(me, date) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return { ok: false, error: 'bad_date' };
  var sc = deskScope_(me);
  var rows = deskLogRows_().map(deskRow_).filter(function (r) { return r.date === date && sc.allowed[r.wz]; });
  var bags = {}, cash = {}, allocs = desk2Allocs_().filter(function (a) { return sc.allowed[a.wz] && (a.date === date || a.depDate === date); });
  readTab_(TABS.BR).forEach(function (b) {
    var w = String(b[13] || ''); if (normDate_(b[12]) !== date || !sc.allowed[w]) return;
    (bags[w] = bags[w] || []).push({ barcode: String(b[0]), created: b[2] instanceof Date ? b[2].getTime() : String(b[2]), by: String(b[5]),
      status: String(b[6]), amount: Number(b[7]) || 0, storeId: String(b[3]), comments: b[10] instanceof Date ? normDate_(b[10]) : String(b[10] || ''), note: String(b[14] || '') });
  });
  readTab_(TABS.RQ).forEach(function (r) {
    var w = String(r[1] || ''); if (normDate_(r[0]) !== date || !sc.allowed[w]) return;
    (cash[w] = cash[w] || []).push({ tx: String(r[3]), by: String(r[5]), amount: Number(r[6]) || 0 });
  });
  var pend = me.canEdit ? deskPending_(date).filter(function (p) { return sc.allowed[p.wz]; }) : [];
  return { ok: true, date: date, rows: rows, bags: bags, cash: cash, allocs: allocs, pending: pend, locked: desk2Locks_()[date] || null };
}

function deskStore_(me, wz, days) {
  var sc = deskScope_(me);
  if (!sc.allowed[wz]) return { ok: false, error: 'not_allowed' };
  var rows = deskLogRows_().map(deskRow_).filter(function (r) { return r.wz === wz; })
    .sort(function (a, b) { return a.date < b.date ? 1 : -1; }).slice(0, Math.min(days, 120));
  return { ok: true, wz: wz, rows: rows, allocs: desk2Allocs_().filter(function (a) { return a.wz === wz; }) };
}

/* ---------- edits (Editor / Admin) ---------- */
/** Script lock for a dashboard write. Returns null (→ error "busy", the page retries) instead of throwing. */
function deskLock_() {
  var l = LockService.getScriptLock(); if (l.tryLock(25000)) return l;
  try { var p = PropertiesService.getScriptProperties(), a = JSON.parse(p.getProperty('busy_log') || '[]'); a.push(Date.now()); p.setProperty('busy_log', JSON.stringify(a.slice(-300))); } catch (e) {}
  return null;
}

var ADJ_RULES = { NOTE_OVER: 100, SPLIT_OVER: 500, APPROVAL_OVER: 1000, PENDING: 'Pending_Adjustments',
  PENDING_H: ['ID', 'Date', 'WZ #', 'Store', 'Adjust $', 'Reason', 'Note', 'Requested By', 'Requested At', 'Status', 'Decided By', 'Decided At'] };

/**
 * Save an adjustment / note. Guardrails (server-enforced so every client behaves the same):
 *  - conflict:      someone changed this store-day's adjustment/note since the page loaded it  → refresh first
 *  - note_required: |adj| > $100 needs a note (who confirmed it)
 *  - wrong_sign:    the adjustment moves the difference further from $0                         → confirm (ack "sign")
 *  - use_split:     |adj| > $500 or a "late / drop / combined" reason — money for other days    → confirm (ack "split")
 *  - pending:       non-admin |adj| > $1,000 goes to Pending_Adjustments for an admin to approve
 */
function deskSaveAdj_(me, b, approved) {
  if (!me.canEdit) return { ok: false, error: 'read_only' };
  var date = String(b.date || ''), wz = String(b.wz || ''), ack = b.ack || [];
  var adj = b.adj === '' || b.adj === null || b.adj === undefined ? '' : Number(b.adj);
  if (adj !== '' && (isNaN(adj) || Math.abs(adj) > 100000)) return { ok: false, error: 'bad_amount' };
  var reason = String(b.reason || '').trim().slice(0, 200), note = String(b.note || '').trim().slice(0, 1000);
  if (adj !== '' && adj !== 0 && !reason) return { ok: false, error: 'reason_required' };
  if (desk2Locks_()[date]) return { ok: false, error: 'day_locked' };
  var lock = deskLock_(); if (!lock) return { ok: false, error: 'busy' };
  try {
    var sh = sheet_(TABS.LOG, LOG_HEADERS), v = sh.getDataRange().getValues();
    for (var i = 1; i < v.length; i++) {
      if (normDate_(v[i][0]) !== date || String(v[i][1]) !== wz) continue;
      var cur = deskRow_(v[i]);
      if (b.prev && !approved) {
        var same = function (x, y) { return String(x === '' || x === null || x === undefined ? '' : x).trim() === String(y === '' || y === null || y === undefined ? '' : y).trim(); };
        var pa = b.prev.adj === '' || b.prev.adj === null || b.prev.adj === undefined ? '' : Number(b.prev.adj);
        if (!same(pa, cur.adj) || !same(b.prev.reason, cur.adjReason) || !same(b.prev.note, cur.note))
          return { ok: false, error: 'conflict', by: cur.adjBy, row: cur };
      }
      var a = adj === '' ? 0 : adj, changed = round2_(a) !== round2_(cur.adj === '' ? 0 : cur.adj);
      if (changed && !approved) {
        var tolC = Number(getCfg_().TOLERANCE) || 5, diff = Number(v[i][7]) || 0;
        if (Math.abs(a) > ADJ_RULES.NOTE_OVER && !note) return { ok: false, error: 'note_required' };
        if (a && Math.abs(diff + a) > Math.abs(diff) + tolC && ack.indexOf('sign') < 0)
          return { ok: false, error: 'wrong_sign', suggest: round2_(-diff) };
        if (a && (Math.abs(a) > ADJ_RULES.SPLIT_OVER || /late|drop|combined|several days|multiple days/i.test(reason)) && ack.indexOf('split') < 0)
          return { ok: false, error: 'use_split' };
        if (!me.isAdmin && Math.abs(a) > ADJ_RULES.APPROVAL_OVER) {
          var id = 'PA-' + Utilities.getUuid().slice(0, 8);
          desk2Tab_(ADJ_RULES.PENDING, ADJ_RULES.PENDING_H).appendRow([id, date, wz, cur.store, a, reason, note, me.name + ' <' + me.email + '>', new Date(), 'Pending', '', '']);
          deskLog_(me, date, wz, cur.store, 'Adjustment requested ' + money_(a), [reason, note].filter(String).join(' — '));
          try { MailApp.sendEmail(adminEmail_(), 'Cash Recon: approve ' + money_(a) + ' adjustment — ' + cur.store + ' ' + date,
            me.name + ' asked for a ' + money_(a) + ' adjustment on ' + cur.store + ' (' + wz + ') for ' + date + '.\n\nReason: ' + reason +
            '\nNote: ' + note + '\n\nApprove or reject it in the Cash Recon Desk (open the store for that day): ' + DESK2.SITE); } catch (e) {}
          return { ok: true, pending: true, id: id, row: cur };
        }
      }
      var tol = Number(getCfg_().TOLERANCE) || 5, base = String(v[i][8]), fd = round2_((Number(v[i][7]) || 0) + a);
      var fs = a ? (Math.abs(fd) <= tol ? 'RESOLVED (manual)' : base + ' (adj)')
                 : (/late drop/.test(String(v[i][16])) ? String(v[i][16]) : base);
      if (!a && /late drop/.test(String(v[i][16]))) fd = Number(v[i][15]) || 0;
      var by = approved ? approved.by : me.name;
      sh.getRange(i + 1, 12, 1, 7).setValues([[adj === 0 ? '' : adj, reason, by, note, fd, fs, new Date()]]);
      deskLog_(me, date, wz, String(v[i][2]), adj !== '' && adj !== 0 ? 'Adjustment ' + money_(adj) + (approved ? ' (approved)' : '') : 'Note',
        [reason, note].filter(String).join(' — '));
      return { ok: true, row: deskRow_(sh.getRange(i + 1, 1, 1, 18).getValues()[0]) };
    }
    return { ok: false, error: 'not_found' };
  } finally { lock.releaseLock(); }
}

/** Pending (>$1,000, non-admin) adjustments for a day — admins see and decide them on the store sheet. */
function deskPending_(date) {
  var sh = ss_().getSheetByName(ADJ_RULES.PENDING); if (!sh) return [];
  return sh.getDataRange().getValues().slice(1).filter(function (r) { return String(r[9]) === 'Pending' && (!date || normDate_(r[1]) === date); })
    .map(function (r) { return { id: String(r[0]), date: normDate_(r[1]), wz: String(r[2]), store: String(r[3]), adj: Number(r[4]) || 0,
      reason: String(r[5]), note: String(r[6]), by: String(r[7]).replace(/ <.*$/, ''), at: r[8] instanceof Date ? r[8].getTime() : null }; });
}
function deskDecideAdj_(me, b) {
  if (!me.isAdmin) return { ok: false, error: 'admin_only' };
  var sh = ss_().getSheetByName(ADJ_RULES.PENDING); if (!sh) return { ok: false, error: 'not_found' };
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]) !== String(b.id) || String(v[i][9]) !== 'Pending') continue;
    var res = { ok: true };
    if (b.approve) {
      res = deskSaveAdj_(me, { date: normDate_(v[i][1]), wz: String(v[i][2]), adj: Number(v[i][4]), reason: String(v[i][5]),
        note: String(v[i][6]) + ' [approved by ' + me.name + ']' }, { by: String(v[i][7]).replace(/ <.*$/, '') });
      if (!res.ok) return res;
    }
    sh.getRange(i + 1, 10, 1, 3).setValues([[b.approve ? 'Approved' : 'Rejected', me.name, new Date()]]);
    deskLog_(me, normDate_(v[i][1]), String(v[i][2]), String(v[i][3]), (b.approve ? 'Approved' : 'Rejected') + ' adjustment ' + money_(v[i][4]), String(v[i][0]));
    return res;
  }
  return { ok: false, error: 'not_found' };
}

function deskLog_(me, date, wz, store, change, detail) {
  try {
    var a = ss_().getSheetByName(DESK.ACT);
    if (a) a.appendRow([new Date(), me.name + ' <' + me.email + '>', date, wz, store, change, String(detail || '').slice(0, 500)]);
  } catch (e) {}
}

/* ---------- team access (Admin) ---------- */
function deskUsers_(me) {
  if (!me.isAdmin) return { ok: false, error: 'admin_only' };
  var v = ss_().getSheetByName(DESK.USERS).getDataRange().getValues().slice(1).filter(function (r) { return r[0]; });
  return { ok: true, roles: DESK_ROLES, users: v.map(function (r) {
    return { email: deskEmail_(r[0]), name: String(r[1] || ''), role: String(r[2] || ''), scope: String(r[3] || ''), active: deskTrue_(r[4]) };
  }) };
}

function deskSaveUser_(me, b) {
  if (!me.isAdmin) return { ok: false, error: 'admin_only' };
  var email = deskEmail_(b.email);
  if (!/^[^\s@]+@[^\s@]+\.[a-z]{2,}$/i.test(email)) return { ok: false, error: 'bad_email' };
  var role = DESK_ROLES.indexOf(b.role) >= 0 ? b.role : 'Store Viewer';
  if (email === me.email && (b.active === false || role !== 'Admin')) return { ok: false, error: 'cant_remove_self' };
  var sh = ss_().getSheetByName(DESK.USERS), v = sh.getDataRange().getValues();
  var row = [email, String(b.name || '').slice(0, 80), role, String(b.scope || '').slice(0, 300), b.active !== false];
  var firstEmpty = -1;
  for (var i = 1; i < v.length; i++) {
    if (deskEmail_(v[i][0]) === email) { sh.getRange(i + 1, 1, 1, 5).setValues([row]); deskLog_(me, '', '', '', 'User updated', email + ' · ' + role); return deskUsers_(me); }
    if (firstEmpty < 0 && !String(v[i][0]).trim()) firstEmpty = i + 1;
  }
  sh.getRange(firstEmpty > 0 ? firstEmpty : sh.getLastRow() + 1, 1, 1, 6).setValues([row.concat([new Date()])]);
  deskLog_(me, '', '', '', 'User added', email + ' · ' + role);
  return deskUsers_(me);
}

/** Editor-only check: what an Admin and a DM would see. Select deskSelfTest → Run. */
function deskSelfTest() {
  var admin = { email: 'test', name: 'Test', role: 'Admin', scope: '', canEdit: true, isAdmin: true };
  var b = deskBootstrap_(admin);
  console.log('days=' + b.days.length + ' latest=' + JSON.stringify(b.days[0]) + ' carry=' + b.carry.length + ' stores=' + b.stores.length);
  var d = deskDay_(admin, b.days[0].date);
  console.log('day rows=' + d.rows.length + ' bagStores=' + Object.keys(d.bags).length + ' cashStores=' + Object.keys(d.cash).length + ' WZ1448=' + JSON.stringify(d.bags.WZ1448 || []));
  var dm = deskBootstrap_({ email: 'alex.levitin@wireless-zone.com', name: 'Alex', role: 'DM Viewer', scope: '', canEdit: false, isAdmin: false });
  console.log('DM Alex stores=' + dm.stores.map(function (s) { return s.wz; }).join(','));
  console.log('store WZ1371 rows=' + deskStore_(admin, 'WZ1371', 30).rows.length);
}
