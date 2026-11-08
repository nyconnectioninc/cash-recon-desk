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
  desk2Tab_(DESK2.PICK, DESK2.PICK_H); desk2MapCol_('Safe Limit');
  var cfg = sheet_(TABS.CFG, ['Key', 'Value', 'What it does']), have = cfg.getDataRange().getValues().map(function (r) { return r[0]; });
  if (have.indexOf('SAFE_LIMIT') < 0) cfg.appendRow(['SAFE_LIMIT', '5000', 'Safe Tracker warning level ($), unless the store has its own Safe Limit on the Store Map']);
  if (have.indexOf('PICKUP_QUERY') < 0) cfg.appendRow(['PICKUP_QUERY', 'from:no-reply@brinksinc.com subject:"Pick up and Delivery" has:attachment newer_than:30d', 'Gmail search for the Brinks Pick up and Delivery report']);
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

/* ---------- Safe Tracker (bag level: Deposit Report barcodes vs Pick up & Delivery barcodes) ---------- */
DESK2.PICK = 'Brinks_Pickups';
DESK2.PICK_H = ['Bag', 'Picked At', 'Store ID', 'Location', 'Type', 'Ingested'];

/** Weekly "Pick up and Delivery" email (and any xlsx dropped in Drive › Brinks Recon › Brinks Pickups) → Brinks_Pickups. */
function ingestPickups() {
  if (ss_().getSpreadsheetTimeZone() !== Session.getScriptTimeZone()) { desk2FixTimezone(); return; }   // one-time self-heal
  var P0 = PropertiesService.getScriptProperties();
  if (!P0.getProperty('brinks_backfill_done')) { brinksBackfill(); P0.setProperty('brinks_backfill_done', '1'); }   // one-time: deposit history since Sept
  var cfg = getCfg_(), props = PropertiesService.getScriptProperties();
  var q = cfg.PICKUP_QUERY || 'from:no-reply@brinksinc.com subject:"Pick up and Delivery" has:attachment newer_than:30d';
  var sh = desk2Tab_(DESK2.PICK, DESK2.PICK_H), have = {}, added = 0;
  sh.getDataRange().getValues().slice(1).forEach(function (r) { if (r[1] instanceof Date) have[r[0] + '|' + Math.floor(r[1].getTime() / 60000)] = 1; });
  var take = function (blob, name) {
    var add = desk2ParsePickups_(xlsxToValues_(blob, name), have);
    if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, add[0].length).setValues(add);
    added += add.length;
  };
  GmailApp.search(q).forEach(function (t) {
    t.getMessages().forEach(function (msg) {
      if (props.getProperty('pk_' + msg.getId())) return;
      msg.getAttachments().forEach(function (att) { if (/\.xlsx?$/i.test(att.getName())) take(att.copyBlob(), att.getName()); });
      props.setProperty('pk_' + msg.getId(), '1');
    });
  });
  var it = folderByPath_('Brinks Recon/Brinks Pickups', true).getFiles();
  while (it.hasNext()) {
    var f = it.next(), key = 'pkf_' + f.getId() + '_' + f.getLastUpdated().getTime();
    if (!/\.(xlsx?|txt)$/i.test(f.getName()) || props.getProperty(key)) continue;
    if (/\.txt$/i.test(f.getName())) {
      var add = desk2ParseCompact_(f.getBlob().getDataAsString(), have);
      if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, add[0].length).setValues(add);
      added += add.length;
    } else take(f.getBlob(), f.getName());
    props.setProperty(key, '1');
  }
  if (added) logRun('Pickups ingested: ' + added + ' bag rows');
  return added;
}
function desk2ParsePickups_(rows, have) {
  var hr = -1;
  for (var r = 0; r < Math.min(rows.length, 20); r++) { var hh = rows[r].map(function (x) { return String(x).trim(); }); if (hh.indexOf('Store ID') >= 0 && hh.some(function (x) { return /^Bag\//.test(x); })) { hr = r; break; } }
  if (hr < 0) throw new Error('Pickup report header not found');
  var h = rows[hr].map(function (x) { return String(x).trim(); }), c = function (n) { return h.indexOf(n); };
  var iD = c('Date'), iS = c('Store ID'), iL = c('Location Name'), iB = h.findIndex(function (x) { return /^Bag\//.test(x); }), iT = c('Transaction Type'), now = new Date(), out = [];
  for (var k = hr + 1; k < rows.length; k++) {
    var x = rows[k], bag = String(x[iB] || '').trim();
    if (!bag || !(x[iD] instanceof Date)) continue;
    var key = bag + '|' + Math.floor(x[iD].getTime() / 60000);
    if (have[key]) continue;
    have[key] = 1;
    out.push([bag, x[iD], String(x[iS]).trim(), String(x[iL] || ''), String(x[iT] || ''), now]);
  }
  return out;
}

/** Compact history file (PUv1: "yyyy-MM-dd HH:mm:ss<TAB>Store ID<TAB>P|D<TAB>bag prefix<TAB>suffixes / ranges"). */
function desk2ParseCompact_(txt, have) {
  var lines = String(txt).split(/\r?\n/), out = [], now = new Date();
  if (lines[0].trim() !== 'PUv1') throw new Error('Unknown pickup history format');
  lines.slice(1).forEach(function (ln) {
    if (!ln.trim()) return;
    var p = ln.split('\t'), m = p[0].match(/(\d+)-(\d+)-(\d+) (\d+):(\d+):(\d+)/);
    var t = new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]), type = p[2] === 'D' ? 'Delivery' : 'Pickup';
    p[4].split(',').forEach(function (part) {
      var r = part.match(/^(\d+)-(\d+)(.*)$/), bags = [];
      if (r) { for (var n = +r[1]; n <= +r[2]; n++) bags.push(p[3] + ('0000000000' + n).slice(-r[1].length) + r[3]); }
      else bags.push(p[3] + part);
      bags.forEach(function (bag) {
        var key = bag + '|' + Math.floor(t.getTime() / 60000);
        if (have[key]) return; have[key] = 1;
        out.push([bag, t, p[1], '', type, now]);
      });
    });
  });
  return out;
}

/** One-time fix: Control Sheet → Eastern time, then reload the pickup log. */
function desk2FixTimezone() {
  ss_().setSpreadsheetTimeZone(Session.getScriptTimeZone());
  var props = PropertiesService.getScriptProperties(), all = props.getProperties();
  Object.keys(all).forEach(function (k) { if (/^pkf?_/.test(k)) props.deleteProperty(k); });
  var sh = desk2Tab_(DESK2.PICK, DESK2.PICK_H); if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 6).clearContent();
  console.log('Sheet TZ now ' + ss_().getSpreadsheetTimeZone() + '; pickups re-ingested: ' + ingestPickups());
}

/** One-time: pull every Deposit Report back to a date so pickups have dollar amounts behind them. Safe to re-run. */
function brinksBackfill() {
  ingestBrinks('from:no-reply@brinksinc.com subject:"Deposit Report" has:attachment after:2026/08/25');
}

function desk2Safe_(me) {
  var sc = deskScope_(me), cfg = getCfg_(), limitAll = Number(cfg.SAFE_LIMIT) || 5000, now = Date.now();
  var mapRows = ss_().getSheetByName(TABS.MAP).getDataRange().getValues(), h = mapRows[0].map(String);
  var iW = h.indexOf('WZ Store ID'), iL = h.indexOf('Safe Limit'), lim = {};
  mapRows.slice(1).forEach(function (r) { var w = String(r[iW]).trim(); if (w && iL >= 0) lim[w] = Number(r[iL]) || 0; });
  // one safe per Brinks Store ID (stores sharing an ID are shown together)
  var safes = {};
  Object.keys(sc.allowed).forEach(function (w) {
    var s = sc.map.byWZ[w]; if (!s || !s.active) return;
    var id = s.brinksId || w, o = safes[id] = safes[id] || { id: id, stores: [], limit: 0 };
    o.stores.push({ wz: w, name: String(s.name), dm: s.dm });
    if (lim[w]) o.limit = Math.max(o.limit, lim[w]);
  });
  var bySid = {}, depositBags = {}, owner = {}, cover = Infinity;
  readTab_(TABS.BR).forEach(function (b) {
    if (!b[0]) return;
    var t = b[2] instanceof Date ? b[2].getTime() : new Date(b[2]).getTime();
    if (/^SD-/.test(String(b[0]))) return;   // smart-safe (bill feeder) deposits have no bag barcode to match to pickups
    cover = Math.min(cover, t); depositBags[String(b[0])] = 1; owner[String(b[0])] = String(b[3]).trim();
    if (/^(voided|test)$/i.test(String(b[6]).trim())) return;
    (bySid[String(b[3]).trim()] = bySid[String(b[3]).trim()] || []).push({ bag: String(b[0]), t: t, amt: Number(b[7]) || 0, by: String(b[5]), wz: String(b[13] || '') });
  });
  // Brinks sometimes files a pickup under another store's ID (Quality under "Hooksett") or an old ID ("NHVER", "MA 555").
  // Credit each picked-up bag to the safe it was dropped in; fall back to Config PICKUP_ALIASES, then the report's Store ID.
  var alias = {};
  String(cfg.PICKUP_ALIASES || 'NHVER=Hooksett, MA 555=WZ1456').split(/[,;]/).forEach(function (x) { var kv = x.split('='); if (kv.length === 2) alias[kv[0].trim()] = kv[1].trim(); });
  var picked = {}, events = {};
  readTab_(DESK2.PICK).forEach(function (r) {
    if (!r[0] || !/pickup/i.test(String(r[4]))) return;
    var t = r[1] instanceof Date ? r[1].getTime() : new Date(r[1]).getTime(), raw = String(r[2]).trim();
    var sid = owner[String(r[0])] || alias[raw] || raw;
    picked[String(r[0])] = t;
    var e = (events[sid] = events[sid] || {}); e[t] = (e[t] || 0) + 1;
  });
  var rq = readTab_(TABS.RQ), rqFirst = rq.reduce(function (m, r) { var d = normDate_(r[0]); return !m || d < m ? d : m; }, '');
  var out = Object.keys(safes).map(function (id) {
    var o = safes[id], ev = Object.keys(events[id] || {}).map(Number).sort(function (a, b) { return b - a; });
    var last = ev.length ? ev[0] : null, bags = bySid[id] || [];
    var inSafe = bags.filter(function (b) { return !picked[b.bag]; }).sort(function (a, b) { return a.t - b.t; });
    var shared = o.stores.length > 1;   // two safes on one Brinks ID: one store's pickup doesn't collect the other's bags
    var missed = last && !shared ? inSafe.filter(function (b) { return b.t < last - 2 * 3600e3; }) : [];
    var noRecord = [];
    ev.forEach(function (t, k) {
      var prev = ev[k + 1];                 // ev is newest first
      if (!prev || prev < cover) return;    // only pickups whose whole window is inside the deposit history
      Object.keys(picked).forEach(function (bag) { if (picked[bag] === t && !depositBags[bag] && bags.length) noRecord.push({ bag: bag, t: t }); });
    });
    var drops = round2_(inSafe.reduce(function (s, b) { return s + b.amt; }, 0)), rqSince = null;
    if (last) {
      var lastDay = Utilities.formatDate(new Date(last), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (rqFirst && lastDay >= rqFirst) {
        var ws = o.stores.map(function (s) { return s.wz; });
        rqSince = round2_(rq.reduce(function (s, r) { return ws.indexOf(String(r[1])) >= 0 && normDate_(r[0]) >= lastDay ? s + (Number(r[6]) || 0) : s; }, 0));
      }
    }
    return { id: id, shared: shared, stores: o.stores, limit: o.limit || limitAll, last: last, daysSince: last ? Math.floor((now - last) / 864e5) : null,
      history: ev.slice(0, 8).map(function (t) { return { at: t, bags: events[id][t] }; }),
      drops: drops, bags: inSafe.length, oldest: inSafe.length ? inSafe[0].t : null, rqSince: rqSince,
      inSafe: inSafe.slice(-60).map(function (b) { return [b.bag, b.t, b.amt, b.by]; }),
      missed: missed.map(function (b) { return [b.bag, b.t, b.amt, b.by]; }), noRecord: noRecord.slice(0, 30).map(function (x) { return [x.bag, x.t]; }) };
  });
  return { ok: true, limit: limitAll, coverage: isFinite(cover) ? cover : null, rqFirst: rqFirst, safes: out,
    lastReport: (function () { var r = readTab_(DESK2.PICK); return r.length ? r[r.length - 1][5] instanceof Date ? r[r.length - 1][5].getTime() : null : null; })() };
}

function desk2SaveStore_(me, b) {
  if (!me.isAdmin) return { ok: false, error: 'admin_only' };
  var sh = ss_().getSheetByName(TABS.MAP), cL = desk2MapCol_('Safe Limit');
  var v = sh.getDataRange().getValues(), h = v[0].map(String), iW = h.indexOf('WZ Store ID'), iB = h.indexOf('Brinks Store ID'), n = 0;
  for (var i = 1; i < v.length; i++) if (String(v[i][iB]).trim() === String(b.id) || String(v[i][iW]).trim() === String(b.id)) {
    sh.getRange(i + 1, cL).setValue(b.safeLimit === '' || b.safeLimit == null ? '' : Number(b.safeLimit) || ''); n++;
  }
  if (!n) return { ok: false, error: 'not_found' };
  deskLog_(me, '', String(b.id), '', 'Safe limit', b.safeLimit || 'default');
  return desk2Safe_(me);
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

/** Editor-only read check of the new endpoints. */
function desk2SelfTest() {
  var admin = { email: 'test', name: 'Test', role: 'Admin', scope: '', canEdit: true, isAdmin: true };
  var r = desk2Range_(admin, '2026-11-01', '2026-11-04');
  console.log('range rows=' + r.rows.length + ' allocs=' + r.allocs.length);
  var s = desk2Safe_(admin);
  console.log('safes=' + s.safes.length + ' withPickup=' + s.safes.filter(function (x) { return x.last; }).length + ' coverage=' + new Date(s.coverage) + ' rqFirst=' + s.rqFirst);
  s.safes.slice().sort(function (a, b) { return b.drops - a.drops; }).slice(0, 12).forEach(function (x) {
    console.log(x.id + ' last=' + (x.last ? new Date(x.last).toISOString().slice(0, 16) : '—') + ' bags=' + x.bags + ' $' + x.drops + ' missed=' + x.missed.length + ' noRecord=' + x.noRecord.length + ' rqSince=' + x.rqSince);
  });
}
