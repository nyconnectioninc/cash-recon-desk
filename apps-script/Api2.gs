/**
 * Cash Recon Desk — dashboard features (third file in the project; uses Code.gs + Api.gs helpers).
 *   Combined deposits / extra-$ matcher (Deposit_Allocations) · Close/reopen day (Day_Locks)
 *   Safe Tracker (Store Map "Pickup Days" / "Safe Limit") · Scorecard + Calendar + Export data (range)
 *   Settings (tolerance, safe limit) · Access requests (Access_Requests)
 */
var DESK2 = {
  ALLOC: 'Deposit_Allocations', LOCKS: 'Day_Locks', REQ: 'Access_Requests',
  ALLOC_H: ['ID', 'WZ #', 'Kind', 'Deposit Date', 'Covers Date', 'Amount', 'Reason', 'Logged By', 'Logged At', 'Status'],
  LOCK_H: ['Date', 'Closed By', 'Closed At'],
  REQ_H: ['Email', 'Name', 'Title', 'Wants', 'Reason', 'Requested', 'Status', 'Decided By', 'Decided At', 'Role', 'Scope'],
  EXCUSED: ['Store closed / holiday', 'Brinks missed pickup', 'Safe or bag malfunction'],
  REQ_DOMAINS: /@(wireless-zone\.com|nyconnectionli\.com)$/i,
  SITE: 'https://nyconnectioninc.github.io/cash-recon-desk/'
};

function desk2Tab_(name, headers) {
  var ss = ss_(), sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#f1f3f4');
    sh.setFrozenRows(1);
  }
  return sh;
}
function desk2Setup() {
  desk2Tab_(DESK2.ALLOC, DESK2.ALLOC_H); desk2Tab_(DESK2.LOCKS, DESK2.LOCK_H); desk2Tab_(DESK2.REQ, DESK2.REQ_H);
  desk2MapCol_('Pickup Days'); desk2MapCol_('Safe Limit');
  var cfg = sheet_(TABS.CFG, ['Key', 'Value', 'What it does']), have = cfg.getDataRange().getValues().map(function (r) { return r[0]; });
  if (have.indexOf('SAFE_LIMIT') < 0) cfg.appendRow(['SAFE_LIMIT', '5000', 'Safe Tracker warning level ($), unless the store has its own Safe Limit on the Store Map']);
  if (have.indexOf('PICKUP_HOUR') < 0) cfg.appendRow(['PICKUP_HOUR', '12', 'Bags dropped after this hour on a pickup day stay in the safe until the next pickup']);
  return 'Desk features ready';
}
function desk2MapCol_(name) {
  var sh = ss_().getSheetByName(TABS.MAP), h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  var i = h.indexOf(name);
  if (i >= 0) return i + 1;
  sh.getRange(1, h.length + 1).setValue(name).setFontWeight('bold');
  return h.length + 1;
}

/** Router for the extra actions (called from doPost in Api.gs). Returns null if the action isn't ours. */
function desk2Route_(me, b) {
  switch (b.action) {
    case 'range':       return desk2Range_(me, String(b.from || ''), String(b.to || ''));
    case 'safe':        return desk2Safe_(me);
    case 'allocate':    return desk2Allocate_(me, b);
    case 'undoAlloc':   return desk2UndoAlloc_(me, String(b.id || ''));
    case 'lockDay':     return desk2Lock_(me, String(b.date || ''), b.lock !== false);
    case 'settings':    return desk2Settings_(me, b);
    case 'saveStore':   return desk2SaveStore_(me, b);
    case 'requests':    return desk2Requests_(me);
    case 'decide':      return desk2Decide_(me, b);
    case 'exportSheet': return desk2ExportSheet_(me, b);
  }
  return null;
}

/* ---------- shared ---------- */
function desk2Locks_() {
  var out = {};
  readTab_(DESK2.LOCKS).forEach(function (r) { if (r[0]) out[normDate_(r[0])] = { by: String(r[1]), at: r[2] instanceof Date ? r[2].getTime() : null }; });
  return out;
}
function desk2Allocs_() {
  return readTab_(DESK2.ALLOC).filter(function (r) { return r[0] && String(r[9]) !== 'Undone'; }).map(function (r) {
    return { id: String(r[0]), wz: String(r[1]), kind: String(r[2]), depDate: normDate_(r[3]), date: normDate_(r[4]), amount: Number(r[5]) || 0,
      reason: String(r[6]), by: String(r[7]), excused: DESK2.EXCUSED.indexOf(String(r[6])) >= 0 };
  });
}
function desk2Final_(diff, adj, base, prevFs, prevFd, tol) {
  if (adj) { var fd = round2_(diff + adj); return { fd: fd, fs: Math.abs(fd) <= tol ? 'RESOLVED (manual)' : base + ' (adj)' }; }
  if (/late drop/.test(String(prevFs))) return { fd: Number(prevFd) || 0, fs: String(prevFs) };
  return { fd: diff, fs: base };
}

/* ---------- range data (scorecard, calendar, export) ---------- */
function desk2Range_(me, from, to) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(from) || !/^\d{4}-\d{2}-\d{2}$/.test(to) || from > to) return { ok: false, error: 'bad_date' };
  if (to > addDays_(from, 123)) from = addDays_(to, -123);
  var sc = deskScope_(me), allocs = desk2Allocs_(), byKey = {};
  allocs.forEach(function (a) { (byKey[a.date + '|' + a.wz] = byKey[a.date + '|' + a.wz] || []).push(a); });
  var rows = deskLogRows_().map(deskRow_).filter(function (r) { return r.date >= from && r.date <= to && sc.allowed[r.wz]; })
    .map(function (r) {
      var al = byKey[r.date + '|' + r.wz] || [];
      return [r.date, r.wz, r.store, r.dm, r.company, r.rq, r.brinks, r.diff, r.status, r.finalDiff, r.finalStatus,
        r.adj === '' ? '' : r.adj, r.adjReason, r.adjBy, r.note, al.map(function (a) { return a.kind + ':' + a.id + ':' + (a.excused ? 1 : 0); }).join(',')];
    });
  return { ok: true, from: from, to: to, cols: ['date', 'wz', 'store', 'dm', 'company', 'rq', 'brinks', 'diff', 'status', 'finalDiff', 'finalStatus', 'adj', 'adjReason', 'adjBy', 'note', 'alloc'],
    rows: rows, allocs: allocs.filter(function (a) { return sc.allowed[a.wz] && a.date >= addDays_(from, -14) && a.date <= to; }), locks: desk2Locks_() };
}

/* ---------- Safe Tracker ---------- */
function desk2Safe_(me) {
  var sc = deskScope_(me), cfg = getCfg_(), limitAll = Number(cfg.SAFE_LIMIT) || 5000, pickHour = Number(cfg.PICKUP_HOUR) || 12;
  var tz = Session.getScriptTimeZone(), today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var mapRows = ss_().getSheetByName(TABS.MAP).getDataRange().getValues(), h = mapRows[0].map(String);
  var iW = h.indexOf('WZ Store ID'), iP = h.indexOf('Pickup Days'), iL = h.indexOf('Safe Limit');
  var meta = {};
  mapRows.slice(1).forEach(function (r) { var w = String(r[iW]).trim(); if (w) meta[w] = { days: iP >= 0 ? String(r[iP] || '') : '', limit: iL >= 0 ? Number(r[iL]) || 0 : 0 }; });
  var DOW = { SUN: 0, MON: 1, TUE: 2, WED: 3, THU: 4, FRI: 5, SAT: 6 };
  var bags = readTab_(TABS.BR).filter(function (b) { return b[0] && String(b[6]).toUpperCase() !== 'VOIDED'; });
  var rq = readTab_(TABS.RQ);
  var out = Object.keys(sc.allowed).filter(function (w) { return sc.map.byWZ[w] && sc.map.byWZ[w].active; }).map(function (w) {
    var s = sc.map.byWZ[w], m = meta[w] || { days: '', limit: 0 };
    var dows = m.days.toUpperCase().split(/[^A-Z]+/).map(function (x) { return DOW[x.slice(0, 3)]; }).filter(function (x) { return x !== undefined; });
    var row = { wz: w, name: String(s.name), dm: s.dm, pickupDays: m.days, limit: m.limit || limitAll, last: null, drops: 0, rq: 0, bags: 0 };
    if (!dows.length) return row;
    var d = today;
    for (var k = 0; k < 8; k++) { if (dows.indexOf(new Date(d + 'T12:00:00').getDay()) >= 0) break; d = addDays_(d, -1); }
    row.last = d;
    var cut = new Date(d + 'T00:00:00'); cut.setHours(pickHour);
    bags.forEach(function (b) {
      var t = b[2] instanceof Date ? b[2] : new Date(b[2]);
      var sid = String(b[3]).trim(), owner = String(b[13] || '') || (/^WZ\d/i.test(sid) ? sid.toUpperCase() : ((sc.map.byBrinksId[sid] || [])[0] || ''));
      if (owner === w && t >= cut) { row.drops += Number(b[7]) || 0; row.bags++; }
    });
    rq.forEach(function (r) { if (String(r[1]) === w && normDate_(r[0]) >= d) row.rq += Number(r[6]) || 0; });
    row.drops = round2_(row.drops); row.rq = round2_(row.rq);
    return row;
  });
  return { ok: true, today: today, limit: limitAll, pickupHour: pickHour, stores: out };
}

function desk2SaveStore_(me, b) {
  if (!me.isAdmin) return { ok: false, error: 'admin_only' };
  var sh = ss_().getSheetByName(TABS.MAP), cP = desk2MapCol_('Pickup Days'), cL = desk2MapCol_('Safe Limit');
  var v = sh.getDataRange().getValues(), iW = v[0].map(String).indexOf('WZ Store ID');
  for (var i = 1; i < v.length; i++) if (String(v[i][iW]).trim() === String(b.wz)) {
    sh.getRange(i + 1, cP).setValue(String(b.pickupDays || '').slice(0, 60));
    sh.getRange(i + 1, cL).setValue(b.safeLimit === '' || b.safeLimit == null ? '' : Number(b.safeLimit) || '');
    deskLog_(me, '', b.wz, String(v[i][0]), 'Store settings', 'Pickup ' + (b.pickupDays || '—') + ' · limit ' + (b.safeLimit || 'default'));
    return desk2Safe_(me);
  }
  return { ok: false, error: 'not_found' };
}

/* ---------- Combined deposit / extra-$ matcher ---------- */
// Moves an overage on depDate back to earlier short / missing days (oldest first), as tracked manual adjustments.
function desk2Allocate_(me, b) {
  if (!me.canEdit) return { ok: false, error: 'read_only' };
  var wz = String(b.wz || ''), dep = String(b.depDate || ''), kind = b.kind === 'combo' ? 'combo' : 'match';
  var days = (b.days || []).map(String).filter(function (d) { return /^\d{4}-\d{2}-\d{2}$/.test(d) && d < dep; }).sort();
  var reason = String(b.reason || '').trim();
  if (!days.length) return { ok: false, error: 'pick_days' };
  if (kind === 'combo' && !reason) return { ok: false, error: 'reason_required' };
  if (!deskScope_(me).allowed[wz]) return { ok: false, error: 'not_allowed' };
  var locks = desk2Locks_();
  if (locks[dep] || days.some(function (d) { return locks[d]; })) return { ok: false, error: 'day_locked' };
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = sheet_(TABS.LOG, LOG_HEADERS), v = sh.getDataRange().getValues(), idx = {};
    for (var i = 1; i < v.length; i++) if (String(v[i][1]) === wz) idx[normDate_(v[i][0])] = i;
    if (idx[dep] === undefined) return { ok: false, error: 'not_found' };
    var tol = Number(getCfg_().TOLERANCE) || 5, depRow = deskRow_(v[idx[dep]]);
    var extra = round2_(depRow.finalDiff);
    if (extra <= tol) return { ok: false, error: 'no_extra' };
    var id = (kind === 'combo' ? 'CD-' : 'XM-') + dep.slice(5).replace('-', '') + '-' + wz + '-' + Utilities.getUuid().slice(0, 4);
    var moves = [], rem = extra;
    days.forEach(function (d) {
      if (idx[d] === undefined || rem <= 0.004) return;
      var need = round2_(-deskRow_(v[idx[d]]).finalDiff);
      if (need <= 0.004) return;
      var amt = round2_(Math.min(rem, need)); rem = round2_(rem - amt); moves.push({ date: d, amt: amt });
    });
    if (!moves.length) return { ok: false, error: 'nothing_to_move' };
    var total = round2_(moves.reduce(function (t, m) { return t + m.amt; }, 0));
    var label = (kind === 'combo' ? 'Combined deposit ' : 'Matched extra deposit ') + id + (reason ? ' · ' + reason : '');
    var now = new Date(), alloc = desk2Tab_(DESK2.ALLOC, DESK2.ALLOC_H);
    moves.concat([{ date: dep, amt: -total }]).forEach(function (m) {
      var r = idx[m.date], row = v[r], prevAdj = row[11] === '' ? 0 : Number(row[11]) || 0, adj = round2_(prevAdj + m.amt);
      var f = desk2Final_(Number(row[7]) || 0, adj, String(row[8]), row[16], row[15], tol);
      var rs = String(row[12] || ''); rs = rs ? rs + ' | ' + label : label;
      sh.getRange(r + 1, 12, 1, 3).setValues([[adj || '', rs.slice(0, 400), me.name]]);
      sh.getRange(r + 1, 16, 1, 3).setValues([[f.fd, f.fs, now]]);
      alloc.appendRow([id, wz, kind, dep, m.date, m.amt, reason || (kind === 'match' ? 'Matched extra deposit' : ''), me.name, now, 'Active']);
    });
    deskLog_(me, dep, wz, depRow.store, kind === 'combo' ? 'Combined deposit' : 'Matched extra $',
      money_(total) + ' → ' + moves.map(function (m) { return m.date.slice(5) + ' ' + money_(m.amt); }).join(', ') + (reason ? ' · ' + reason : ''));
    return { ok: true, id: id, moved: total, left: rem, moves: moves };
  } finally { lock.releaseLock(); }
}

function desk2UndoAlloc_(me, id) {
  if (!me.canEdit) return { ok: false, error: 'read_only' };
  var a = desk2Tab_(DESK2.ALLOC, DESK2.ALLOC_H), av = a.getDataRange().getValues(), items = [];
  for (var i = 1; i < av.length; i++) if (String(av[i][0]) === id && String(av[i][9]) !== 'Undone') items.push({ r: i + 1, wz: String(av[i][1]), date: normDate_(av[i][4]), amt: Number(av[i][5]) || 0 });
  if (!items.length) return { ok: false, error: 'not_found' };
  var locks = desk2Locks_();
  if (items.some(function (x) { return locks[x.date]; })) return { ok: false, error: 'day_locked' };
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = sheet_(TABS.LOG, LOG_HEADERS), v = sh.getDataRange().getValues(), tol = Number(getCfg_().TOLERANCE) || 5;
    items.forEach(function (x) {
      for (var j = 1; j < v.length; j++) if (String(v[j][1]) === x.wz && normDate_(v[j][0]) === x.date) {
        var adj = round2_((v[j][11] === '' ? 0 : Number(v[j][11]) || 0) - x.amt);
        var rs = String(v[j][12] || '').split(' | ').filter(function (p) { return p.indexOf(id) < 0; }).join(' | ');
        var f = desk2Final_(Number(v[j][7]) || 0, adj, String(v[j][8]), v[j][16], v[j][15], tol);
        sh.getRange(j + 1, 12, 1, 2).setValues([[adj || '', rs]]);
        sh.getRange(j + 1, 16, 1, 3).setValues([[f.fd, f.fs, new Date()]]);
        v[j][11] = adj || ''; v[j][12] = rs;
      }
      a.getRange(x.r, 10).setValue('Undone');
    });
    deskLog_(me, items[0].date, items[0].wz, '', 'Undo allocation', id);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* ---------- Close / reopen a day ---------- */
function desk2Lock_(me, date, lockIt) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return { ok: false, error: 'bad_date' };
  if (lockIt && !me.canEdit) return { ok: false, error: 'read_only' };
  if (!lockIt && !me.isAdmin) return { ok: false, error: 'admin_only' };
  var sh = desk2Tab_(DESK2.LOCKS, DESK2.LOCK_H), v = sh.getDataRange().getValues();
  for (var i = v.length - 1; i >= 1; i--) if (normDate_(v[i][0]) === date) { if (lockIt) return { ok: true, locks: desk2Locks_() }; sh.deleteRow(i + 1); }
  if (lockIt) sh.appendRow([date, me.name, new Date()]);
  deskLog_(me, date, '', '', lockIt ? 'Closed day' : 'Reopened day', '');
  return { ok: true, locks: desk2Locks_() };
}

/* ---------- Settings (Admin) ---------- */
function desk2Settings_(me, b) {
  if (!me.isAdmin) return { ok: false, error: 'admin_only' };
  var sh = sheet_(TABS.CFG, ['Key', 'Value', 'What it does']), v = sh.getDataRange().getValues();
  var set = function (k, val) {
    for (var i = 1; i < v.length; i++) if (String(v[i][0]).trim() === k) { sh.getRange(i + 1, 2).setValue(String(val)); return; }
    sh.appendRow([k, String(val), '']);
  };
  var t = Number(b.tolerance), l = Number(b.safeLimit);
  if (b.tolerance !== undefined && !isNaN(t) && t >= 0 && t <= 100) set('TOLERANCE', t);
  if (b.safeLimit !== undefined && !isNaN(l) && l >= 0) set('SAFE_LIMIT', l);
  deskLog_(me, '', '', '', 'Settings', 'Tolerance $' + (b.tolerance == null ? '—' : b.tolerance) + ' · safe limit $' + (b.safeLimit == null ? '—' : b.safeLimit));
  var c = getCfg_();
  return { ok: true, tolerance: Number(c.TOLERANCE) || 5, safeLimit: Number(c.SAFE_LIMIT) || 5000 };
}

/* ---------- Access requests ---------- */
function desk2RequestAccess_(b) {
  var email = deskEmail_(b.email);
  if (!DESK2.REQ_DOMAINS.test(email)) return { ok: false, error: 'domain' };
  var name = String(b.name || '').trim(), reason = String(b.reason || '').trim();
  if (!name || !reason) return { ok: false, error: 'missing_fields' };
  if (deskFindUser_(email)) return { ok: true, status: 'has_access' };
  var sh = desk2Tab_(DESK2.REQ, DESK2.REQ_H), v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (deskEmail_(v[i][0]) === email) {
    var st = String(v[i][6]);
    if (st === 'Denied') return { ok: true, status: 'denied' };
    if (st === 'Pending') return { ok: true, status: 'pending' };
  }
  var cache = CacheService.getScriptCache();
  if (cache.get('desk_req:' + email)) return { ok: true, status: 'pending' };
  cache.put('desk_req:' + email, '1', 300);
  sh.appendRow([email, name.slice(0, 80), String(b.title || '').slice(0, 80), String(b.wants || '').slice(0, 120), reason.slice(0, 500), new Date(), 'Pending', '', '', '', '']);
  var admins = ss_().getSheetByName(DESK.USERS).getDataRange().getValues().slice(1)
    .filter(function (r) { return String(r[2]) === 'Admin' && deskTrue_(r[4]); }).map(function (r) { return deskEmail_(r[0]); });
  if (admins.length) MailApp.sendEmail({ to: admins.join(','), subject: DESK.NAME + ': access request from ' + name,
    htmlBody: '<div style="font-family:-apple-system,Helvetica,Arial,sans-serif;font-size:15px">' +
      '<p><b>' + name + '</b> (' + email + ') asked for access.</p><p>Title: ' + (b.title || '—') + '<br>Wants to see: ' + (b.wants || '—') + '<br>Reason: ' + reason + '</p>' +
      '<p><a href="' + DESK2.SITE + '">Open Cash Recon Desk → People &amp; Access</a> to approve or deny.</p></div>' });
  return { ok: true, status: 'pending' };
}

function desk2Requests_(me) {
  if (!me.isAdmin) return { ok: false, error: 'admin_only' };
  var v = desk2Tab_(DESK2.REQ, DESK2.REQ_H).getDataRange().getValues().slice(1).filter(function (r) { return r[0]; });
  return { ok: true, requests: v.map(function (r) {
    return { email: deskEmail_(r[0]), name: String(r[1]), title: String(r[2]), wants: String(r[3]), reason: String(r[4]),
      at: r[5] instanceof Date ? r[5].getTime() : null, status: String(r[6]), by: String(r[7]) };
  }).reverse() };
}

function desk2Decide_(me, b) {
  if (!me.isAdmin) return { ok: false, error: 'admin_only' };
  var email = deskEmail_(b.email), sh = desk2Tab_(DESK2.REQ, DESK2.REQ_H), v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (deskEmail_(v[i][0]) === email) {
    var status = b.approve === true ? 'Approved' : b.approve === false ? 'Denied' : 'Pending';
    sh.getRange(i + 1, 7, 1, 5).setValues([[status, me.name, new Date(), b.role || '', b.scope || '']]);
    if (status === 'Approved') {
      deskSaveUser_(me, { email: email, name: v[i][1], role: b.role, scope: b.scope, active: true });
      MailApp.sendEmail({ to: email, subject: 'You have access to ' + DESK.NAME,
        htmlBody: '<div style="font-family:-apple-system,Helvetica,Arial,sans-serif;font-size:15px"><p>Hi ' + String(v[i][1]).split(' ')[0] + ',</p>' +
          '<p>Your access to ' + DESK.NAME + ' was approved (' + (b.role || 'Viewer') + ').</p><p><a href="' + DESK2.SITE + '">Sign in here</a> with this email. A code is emailed each time.</p></div>' });
    }
    deskLog_(me, '', '', '', 'Access ' + status.toLowerCase(), email + (b.role ? ' · ' + b.role : ''));
    return desk2Requests_(me);
  }
  return { ok: false, error: 'not_found' };
}

/* ---------- Export → Google Sheet ---------- */
function desk2ExportSheet_(me, b) {
  var data = desk2Range_(me, String(b.from || ''), String(b.to || ''));
  if (!data.ok) return data;
  var rows = data.rows.filter(function (r) { return (!b.dm || String(r[3]).trim() === b.dm) && (!b.wz || r[1] === b.wz); });
  var secs = b.sections || ['exceptions', 'adjustments', 'log'];
  var ss = SpreadsheetApp.create('Cash Recon export ' + data.from + ' to ' + data.to);
  var add = function (title, head, body) {
    var sh = ss.getSheets().length === 1 && ss.getSheets()[0].getLastRow() === 0 ? ss.getSheets()[0].setName(title) : ss.insertSheet(title);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#f1f3f4'); sh.setFrozenRows(1);
    if (body.length) sh.getRange(2, 1, body.length, head.length).setValues(body);
  };
  var H = ['Date', 'WZ #', 'Store', 'DM', 'RQ', 'Brinks', 'Difference', 'Status', 'Final Diff', 'Final Status', 'Adj $', 'Adj Reason', 'Adj By', 'Note'];
  var pick = function (r) { return [r[0], r[1], r[2], r[3], r[5], r[6], r[7], r[8], r[9], r[10], r[11], r[12], r[13], r[14]]; };
  var open = function (r) { return !/RESOLVED/.test(r[10]) && /MISSING|SHORT|OVER|BRINKS ONLY/.test(r[10]); };
  if (secs.indexOf('exceptions') >= 0) add('Exceptions', H, rows.filter(open).map(pick));
  if (secs.indexOf('adjustments') >= 0) add('Adjustments', H, rows.filter(function (r) { return r[11] !== '' || r[14]; }).map(pick));
  if (secs.indexOf('combined') >= 0) add('Combined & matched', ['ID', 'WZ #', 'Kind', 'Deposit Date', 'Covers Date', 'Amount', 'Reason', 'Logged By'],
    data.allocs.filter(function (a) { return a.date >= data.from && (!b.wz || a.wz === b.wz); }).map(function (a) { return [a.id, a.wz, a.kind, a.depDate, a.date, a.amount, a.reason, a.by]; }));
  if (secs.indexOf('score') >= 0) {
    var sc = {};
    rows.forEach(function (r) {
      var k = b.scoreBy === 'dm' ? String(r[3]).trim() : r[2] + ' (' + r[1] + ')', s = sc[k] = sc[k] || [k, 0, 0, 0, 0];
      if (/MISSING/.test(r[10]) && open(r)) s[1]++; else if (/SHORT|OVER/.test(r[10]) && open(r)) s[2]++;
      else if (/late drop/.test(r[10]) || /Combined deposit .*Manager error/.test(r[12])) s[3]++;
      if (/Combined deposit/.test(r[12]) && !/Manager error/.test(r[12])) s[4]++;
    });
    add('Scorecard', [b.scoreBy === 'dm' ? 'DM' : 'Store', 'Missing', 'Short/Over', 'Late', 'Excused'], Object.keys(sc).map(function (k) { return sc[k]; }));
  }
  if (secs.indexOf('log') >= 0) add('Daily log', H, rows.map(pick));
  var f = DriveApp.getFileById(ss.getId());
  var dest = folderByPath_('Brinks Recon/Exports', true); dest.addFile(f); DriveApp.getRootFolder().removeFile(f);
  try { f.addViewer(me.email); } catch (e) {}
  deskLog_(me, data.from + '→' + data.to, b.wz || '', '', 'Export Google Sheet', secs.join(','));
  return { ok: true, url: ss.getUrl() };
}
