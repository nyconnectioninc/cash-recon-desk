/**
 * Brinks Cash Recon — engine (Google Apps Script, bound to "Brinks Deposit Recon – Control Sheet")
 *
 * Daily flow (trigger every 15 min, 6:00–11:00 AM ET):
 *   1. ingestRQ()      — new RQ Cash Audit files in Drive "Brinks Recon/RQ Cash Audit" -> RQ_Raw
 *   2. ingestBrinks()  — new Brinks "Deposit Report" emails in Gmail -> Brinks_Raw (deduped by bag barcode)
 *   3. When BOTH yesterday's RQ file and a Brinks report dated today exist:
 *        reconcile yesterday + 2 prior days (late bags re-assign) -> Daily_Log
 *        send the alert for yesterday once (TEST_MODE: admin only)
 *   4. 10:45 AM: if anything is still missing, email the admin.
 *
 * Manual columns in Daily_Log (Manual Adj $, Adj Reason, Adj By, Note) are NEVER overwritten.
 */

// ───────────────────────── CONFIG ─────────────────────────
var SHEET_ID = '1GqU-YrYVZax1sH5p7qaTZgDWr3Nmnx4uWBHHgNOk1Oo';   // Brinks Deposit Recon – Control Sheet
function ss_() { return SpreadsheetApp.openById(SHEET_ID); }
var TABS = { MAP: 'Store Map', CFG: 'Config', RQ: 'RQ_Raw', BR: 'Brinks_Raw', LOG: 'Daily_Log', RUN: 'Run_Log' };
var DEFAULT_CFG = {
  TEST_MODE: 'TRUE',                 // TRUE = alerts go to ADMIN_EMAIL only
  ADMIN_EMAIL: '',                   // blank = the account running the script
  EXTRA_ALERT_TO: '',                // comma-separated, added when TEST_MODE is FALSE
  RECON_TEAM: '',                    // comma-separated, ALWAYS get the daily alert (back-office deposit matchers)
  TOLERANCE: '5',                    // $ — differences within this are "MINOR"
  NEXT_MORNING_CUTOFF_HOUR: '11',    // bags created before this hour count for the PREVIOUS business day
  RQ_FOLDER_PATH: 'Brinks Recon/RQ Cash Audit',
  BRINKS_FOLDER_PATH: 'Brinks Recon/Brinks Deposits',
  BRINKS_QUERY: 'from:no-reply@brinksinc.com subject:"Deposit Report" has:attachment newer_than:10d',
  DEVICE_QUERY: 'from:no-reply@brinksinc.com subject:"Device Activity" has:attachment newer_than:10d',   // all safe types (bags + smart safes)
  BRINKS_IMPORT_PATH: 'Brinks Recon/Brinks Imports',   // drop a manual Device Activity / Deposit Report export here
  RECHECK_DAYS: '3',                 // re-reconcile this many recent days each run (late bags)
  MISSING_FILE_ALERT_HOUR: '11'      // after this hour, email admin if inputs still missing
};

var LOG_HEADERS = ['Date', 'WZ #', 'Store', 'Company', 'DM', 'RQ Deposit', 'Brinks Declared', 'Difference',
  'Status', 'Bags', 'Auto Notes', 'Manual Adj $', 'Adj Reason', 'Adj By', 'Note', 'Final Difference', 'Final Status', 'Updated'];
var MANUAL_COLS = [11, 12, 13, 14]; // 0-based: Manual Adj $, Adj Reason, Adj By, Note
var RQ_HEADERS = ['Business Date', 'WZ #', 'Location', 'Transaction #', 'Created On', 'Employee', 'Deposit', 'Entered', 'Expected', 'Float', 'Source File'];
var BR_HEADERS = ['Barcode', 'Business Day', 'Created At', 'Store ID', 'Location Name', 'Created By', 'Status', 'Declared',
  'Verified', 'Variance', 'Comments', 'Bag Received', 'Assigned Date', 'Assigned WZ', 'Assign Note', 'Report Date', 'Source Msg',
  'Location #', 'Deposit Type'];

// ───────────────────────── MENU / TRIGGERS ─────────────────────────
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Cash Recon')
    .addItem('Run now (ingest + recon + alert if ready)', 'runDaily')
    .addItem('Re-run recon for a date…', 'promptReconDate')
    .addItem('Ingest only (no alert)', 'ingestAll')
    .addSeparator()
    .addItem('Set up tabs', 'setupTabs')
    .addItem('Install daily triggers', 'installTriggers')
    .addToUi();
}

function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'runDaily') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runDaily').timeBased().everyMinutes(15).create();
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'onOpen') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('onOpen').forSpreadsheet(SHEET_ID).onOpen().create();   // adds the Cash Recon menu to the sheet
  logRun('Triggers installed (runDaily every 15 min; acts only 6:00–11:59 AM)');
}

function promptReconDate() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Re-run recon', 'Business date (YYYY-MM-DD):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var d = r.getResponseText().trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) { ui.alert('Use YYYY-MM-DD'); return; }
  ingestAll();
  reconcileAndWrite_(d);
  ui.alert('Recon updated for ' + d + ' — see Daily_Log.');
}

// ───────────────────────── MAIN ─────────────────────────
function runDaily() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    var tz = Session.getScriptTimeZone();
    var now = new Date();
    var hour = Number(Utilities.formatDate(now, tz, 'H'));
    if (hour < 6 || hour > 11) return;                       // quiet outside the morning window

    setupTabs();
    ingestAll();

    var cfg = getCfg_();
    var props = PropertiesService.getScriptProperties();
    var today = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    var target = addDays_(today, -1);
    var haveRQ = rqDates_().indexOf(target) >= 0;
    var haveBR = brinksReportDates_().indexOf(today) >= 0;

    if (haveRQ && haveBR) {
      var n = Number(cfg.RECHECK_DAYS) || 3;
      var win = [];
      for (var i = n - 1; i >= 0; i--) { win.push(addDays_(target, -i)); reconcileAndWrite_(addDays_(target, -i)); }
      netWindow_(win);
      if (props.getProperty('alert_sent_' + target) !== '1') {
        sendAlert_(target);
        props.setProperty('alert_sent_' + target, '1');
      }
    } else if (hour >= Number(cfg.MISSING_FILE_ALERT_HOUR) && props.getProperty('missing_sent_' + target) !== '1') {
      var miss = [];
      if (!haveRQ) miss.push('RQ Cash Audit file for ' + target + ' (check the Mac mini — ~/Library/Logs/rq-pull.log)');
      if (!haveBR) miss.push('Brinks Deposit Report email for ' + today);
      MailApp.sendEmail(adminEmail_(), '⚠️ Cash Recon blocked — missing input for ' + target,
        'The recon for ' + target + ' could not run because these inputs are missing:\n\n• ' + miss.join('\n• ') +
        '\n\nIt will run automatically as soon as they arrive (checks every 15 min until noon), or use Cash Recon → Run now.');
      props.setProperty('missing_sent_' + target, '1');
      logRun('Missing inputs for ' + target + ': ' + miss.join(' | '));
    }
  } catch (e) {
    logRun('ERROR ' + e.message + '\n' + e.stack);
    MailApp.sendEmail(adminEmail_(), '❌ Cash Recon error', e.message + '\n\n' + e.stack);
  } finally {
    lock.releaseLock();
  }
}

/** Manual catch-up: ingest everything, reconcile the last RECHECK_DAYS days ending yesterday, send yesterday's alert (ignores the morning window). */
function catchUp() {
  setupTabs();
  ingestAll();
  var tz = Session.getScriptTimeZone();
  var target = addDays_(Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'), -1);
  var n = Number(getCfg_().RECHECK_DAYS) || 3;
  var win = [];
      for (var i = n - 1; i >= 0; i--) { win.push(addDays_(target, -i)); reconcileAndWrite_(addDays_(target, -i)); }
      netWindow_(win);
  sendAlert_(target);
  PropertiesService.getScriptProperties().setProperty('alert_sent_' + target, '1');
}

/** Find a Gmail label case-insensitively; create it (or a safe fallback name) if missing. */
function gmailLabel_(name) {
  var all = GmailApp.getUserLabels();
  for (var i = 0; i < all.length; i++) if (all[i].getName().toLowerCase() === name.toLowerCase()) return all[i];
  try { return GmailApp.createLabel(name); }
  catch (e) {
    var alt = name.replace('/', ' - ');
    for (var j = 0; j < all.length; j++) if (all[j].getName().toLowerCase() === alt.toLowerCase()) return all[j];
    return GmailApp.createLabel(alt);
  }
}

/** Re-run recon + net matching for the window ending yesterday. Sends NO email. */
function reconOnly() {
  ingestAll();
  var tz = Session.getScriptTimeZone();
  var target = addDays_(Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'), -1);
  var n = Number(getCfg_().RECHECK_DAYS) || 3, win = [];
  for (var i = n - 1; i >= 0; i--) { win.push(addDays_(target, -i)); reconcileAndWrite_(addDays_(target, -i)); }
  netWindow_(win);
}

function ingestAll() {
  ingestRQ(); ingestBrinks();
  try { if (typeof ingestPickups === 'function') ingestPickups(); } catch (e) { logRun('Pickups ingest error: ' + e.message); }
}

// ───────────────────────── INGEST: RQ ─────────────────────────
function ingestRQ() {
  var cfg = getCfg_();
  var folder = folderByPath_(cfg.RQ_FOLDER_PATH, false);
  if (!folder) { logRun('RQ folder not found: ' + cfg.RQ_FOLDER_PATH); return; }
  var props = PropertiesService.getScriptProperties();
  var files = folder.getFiles();
  while (files.hasNext()) {
    var f = files.next();
    var m = f.getName().match(/^cash-audit-trail_(\d{4}-\d{2}-\d{2})\.xlsx$/);
    if (!m) continue;
    var key = 'rq_' + f.getId() + '_' + f.getLastUpdated().getTime();
    if (props.getProperty(key)) continue;
    var rows = xlsxToValues_(f.getBlob(), f.getName());
    var parsed = parseRQRows(rows, m[1], f.getName());
    replaceRowsByDate_(TABS.RQ, RQ_HEADERS, 0, m[1], parsed);
    props.setProperty(key, '1');
    logRun('RQ ingested ' + f.getName() + ': ' + parsed.length + ' cashouts');
  }
}

/** Pure: Cash Audit Trail rows -> RQ_Raw rows. */
function parseRQRows(rows, bizDate, fileName) {
  var h = rows[0].map(function (x) { return String(x).trim(); });
  var c = function (name) { var i = h.indexOf(name); if (i < 0) throw new Error('RQ column missing: ' + name); return i; };
  var iLoc = c('Location'), iCre = c('Created On'), iTx = c('Transaction Number'), iEmp = c('Signed In Employee'),
      iDep = c('Deposit'), iEnt = c('EnteredAmount'), iExp = c('ExpectedAmount'), iFl = c('Float Retained'), iType = c('Transaction Type');
  var out = [];
  for (var r = 1; r < rows.length; r++) {
    var row = rows[r];
    if (!row[iLoc]) continue;
    if (row[iType] && String(row[iType]).toUpperCase() !== 'CASHOUT') continue;
    var m = String(row[iLoc]).match(/(WZ\d+[A-Z]?)\s*$/);
    out.push([bizDate, m ? m[1] : '', row[iLoc], row[iTx], row[iCre], row[iEmp],
      num_(row[iDep]), num_(row[iEnt]), num_(row[iExp]), num_(row[iFl]), fileName]);
  }
  return out;
}

// ───────────────────────── INGEST: BRINKS ─────────────────────────
function ingestBrinks(query) {
  var cfg = getCfg_();
  if (!query) {
    ingestBrinks(cfg.BRINKS_QUERY);
    if (cfg.DEVICE_QUERY) ingestBrinks(cfg.DEVICE_QUERY);
    ingestBrinksImports_();
    return;
  }
  var props = PropertiesService.getScriptProperties();
  var outFolder = folderByPath_(cfg.BRINKS_FOLDER_PATH, true);
  var label = gmailLabel_('Brinks/Processed');
  var threads = GmailApp.search(query || cfg.BRINKS_QUERY);
  threads.forEach(function (t) {
    t.getMessages().forEach(function (msg) {
      if (props.getProperty('br_' + msg.getId())) return;
      var reportDate = Utilities.formatDate(msg.getDate(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      var count = 0;
      try {
        msg.getAttachments().forEach(function (att) {
          if (!/\.xlsx?$/i.test(att.getName())) return;
          var rows = xlsxToValues_(att.copyBlob(), att.getName());
          var parsed = parseBrinksRows(rows, reportDate, msg.getId());
          outFolder.createFile(att.copyBlob()).setName('Brinks_' + att.getName().replace(/\.xlsx?$/i, '').replace(/\W+/g, '_') + '_' + reportDate + '.xlsx');
          upsertBrinks_(parsed);
          activateNewInstalls_(parsed);
          count += parsed.length;
        });
      } catch (e) {
        if (/Deposit Barcode/.test(e.message)) { props.setProperty('br_' + msg.getId(), 'not-a-bag-report'); return; }   // other Brinks report (no barcodes)
        logRun('Brinks email ' + reportDate + ' skipped: ' + e.message); return;
      }
      props.setProperty('br_' + msg.getId(), reportDate);
      addBrinksReportDate_(reportDate);
      t.addLabel(label);
      logRun('Brinks ingested (email ' + reportDate + '): ' + count + ' bags');
    });
  });
}

/** Pure: Brinks Deposit Report rows -> Brinks_Raw rows (assignment columns blank). */
function parseBrinksRows(rows, reportDate, msgId) {
  var hr = -1;
  for (var r = 0; r < Math.min(rows.length, 30); r++) if (rows[r].map(function (v) { return String(v).trim(); }).indexOf('Business Day') >= 0) { hr = r; break; }
  if (hr < 0) throw new Error('Brinks header row not found');
  var h = rows[hr].map(function (x) { return String(x).trim(); });
  var c = function (name, re, optional) {
    var i = h.indexOf(name);
    if (i < 0 && re) for (var k = 0; k < h.length; k++) if (re.test(h[k])) { i = k; break; }
    if (i < 0 && !optional) throw new Error('Brinks column missing: ' + name + ' (headers: ' + h.filter(String).join(', ') + ')');
    return i;
  };
  var iBD = c('Business Day', /business/i), iCA = c('Created At', /created (at|on|date)|date created/i), iSID = c('Store ID', /store ?id|store ?#/i),
      iLN = c('Location Name', /location/i), iBC = c('Deposit Barcode', /barcode|^bag/i), iBy = c('Created By', /created by|user/i, true),
      iSt = c('Status', /status/i, true), iDec = c('Declared', /declared|transaction total|deposit amount/i), iVer = c('Verified', /verified/i, true), iVar = c('Variance', /variance/i, true),
      iCom = c('Comments', /comment/i, true), iRec = c('Bag Received Date', /received/i, true),
      iLoc = c('Location Number', /location (number|no|#)/i, true), iTyp = c('Deposit Type', /deposit type/i, true);
  var g = function (x, i) { return i < 0 ? '' : x[i]; };
  var out = [], lastBD = '';
  for (var r2 = hr + 1; r2 < rows.length; r2++) {
    var x = rows[r2];
    if (x[iBD] !== '' && x[iBD] != null) lastBD = x[iBD];      // Device Activity email lists Business Day once per group
    if (!x[iBC] || !x[iCA]) continue;
    x = x.slice(); x[iBD] = lastBD;
    var code = String(x[iBC]).trim(), type = String(g(x, iTyp) || '').trim(), loc = String(g(x, iLoc) || '').trim(), by = g(x, iBy),
        st = String(g(x, iSt) || 'ACH').trim(), amt = num_(x[iDec]);
    if (/safedeposit/i.test(type) || /\s/.test(code)) {          // smart safe (bill feeder): no bag barcode, "Created By" is the device #
      var t = x[iCA] instanceof Date ? x[iCA].getTime() : new Date(x[iCA]).getTime();
      code = 'SD-' + (loc || String(x[iSID]).trim()) + '-' + Math.floor(t / 1000) + '-' + amt;
      type = 'SafeDeposit (device ' + by + ')'; by = '';
      if (amt <= 1) st = 'Test';                                  // $1 install test deposits
    }
    out.push([code, x[iBD], x[iCA], String(x[iSID]).trim(), String(x[iLN]).trim(), by, st,
      amt, g(x, iVer) === '' ? '' : num_(x[iVer]), g(x, iVar) === '' ? '' : num_(x[iVar]), g(x, iCom) || '', g(x, iRec) || '',
      '', '', '', reportDate, msgId, loc, type]);
  }
  return out;
}

function upsertBrinks_(parsed) {
  var sh = sheet_(TABS.BR, BR_HEADERS);
  if (sh.getLastColumn() < BR_HEADERS.length) sh.getRange(1, 1, 1, BR_HEADERS.length).setValues([BR_HEADERS]).setFontWeight('bold');
  var data = sh.getDataRange().getValues();
  var idx = {};
  for (var i = 1; i < data.length; i++) idx[data[i][0]] = i + 1;
  var appends = [];
  parsed.forEach(function (row) {
    if (idx[row[0]]) {
      var rn = idx[row[0]];
      var old = sh.getRange(rn, 1, 1, BR_HEADERS.length).getValues()[0];
      row[12] = old[12]; row[13] = old[13]; row[14] = old[14];     // keep assignment columns
      [5, 8, 9, 10, 11, 17, 18].forEach(function (k) { if (row[k] === '' || row[k] == null) row[k] = old[k]; });   // one report fills what the other lacks
      sh.getRange(rn, 1, 1, row.length).setValues([row]);         // status/verified may change
    } else appends.push(row);
  });
  if (appends.length) sh.getRange(sh.getLastRow() + 1, 1, appends.length, appends[0].length).setValues(appends);
}

// ───────────────────────── RECON ─────────────────────────
function reconcileAndWrite_(bizDate) {
  var cfg = getCfg_();
  var map = loadStoreMap_();
  var rq = readTab_(TABS.RQ);
  var brSheet = sheet_(TABS.BR, BR_HEADERS);
  var br = brSheet.getDataRange().getValues().slice(1);
  var tz = Session.getScriptTimeZone();

  var result = reconcileCore(bizDate, rq, br, map, {
    tolerance: Number(cfg.TOLERANCE) || 5,
    cutoffHour: Number(cfg.NEXT_MORNING_CUTOFF_HOUR) || 11,
    fmt: function (d) { return Utilities.formatDate(d, tz, 'yyyy-MM-dd'); },
    hour: function (d) { return Number(Utilities.formatDate(d, tz, 'H')); }
  });

  // write assignments back to Brinks_Raw (only for bags assigned to this date)
  if (result.assignments.length) {
    var byCode = {};
    result.assignments.forEach(function (a) { byCode[a.barcode] = a; });
    var vals = brSheet.getRange(2, 1, br.length, 15).getValues();
    for (var i = 0; i < vals.length; i++) {
      var a = byCode[vals[i][0]];
      if (a) brSheet.getRange(i + 2, 13, 1, 3).setValues([[a.date, a.wz, a.note]]);
    }
  }
  upsertDailyLog_(bizDate, result.rows);
  logRun('Recon ' + bizDate + ': ' + JSON.stringify(result.summary));
  return result;
}

/**
 * Pure core. rqRows = RQ_Raw values (no header). brRows = Brinks_Raw values (no header).
 * map = { byWZ: {WZ: {name, company, dm, dmEmail, active, brinksId, location}}, byBrinksId: {id: [WZ...]} }
 */
function reconcileCore(bizDate, rqRows, brRows, map, opt) {
  // 1. RQ deposits per WZ for the date
  var rqBy = {};
  rqRows.forEach(function (r) {
    if (normDate_(r[0], opt) !== bizDate || !r[1]) return;
    rqBy[r[1]] = (rqBy[r[1]] || 0) + (Number(r[6]) || 0);
  });

  // 2. Brinks bags -> business date (+ candidate stores)
  var bags = [];
  // 1b. Strongest evidence: same employee + same amount as an RQ cashout (beats Brinks Store ID and created date)
  var empMatch = {}, usedTx = {};
  var nm = function (s) { return String(s || '').toLowerCase().replace(/[^a-z]/g, ''); };
  var outs = rqRows.filter(function (r) { return r[1] && r[5] && Number(r[6]) > 0 && map.byWZ[String(r[1])]; })
    .map(function (r) { return { date: normDate_(r[0], opt), wz: String(r[1]), emp: nm(r[5]), amt: Number(r[6]), tx: String(r[3]) }; });
  brRows.filter(function (b) { return b[0] && b[5] && !brSkip_(b); })
    .map(function (b) { return { b: b, t: b[2] instanceof Date ? b[2] : new Date(b[2]) }; })
    .sort(function (x, y) { return x.t - y.t; })
    .forEach(function (o) {
      var emp = nm(o.b[5]), amt = Number(o.b[7]) || 0, cd = opt.fmt(o.t), lo = addDays_(cd, -14);
      var hit = outs.filter(function (x) { return !usedTx[x.tx] && x.emp === emp && Math.abs(x.amt - amt) <= 1 && x.date <= cd && x.date >= lo; })
        .sort(function (x, y) { return x.date < y.date ? 1 : x.date > y.date ? -1 : 0; })[0];
      if (!hit) return;
      usedTx[hit.tx] = 1;
      var sid = String(o.b[3]).trim(), def = brCands_(o.b, map);
      var notes = [];
      if (def.length !== 1 || def[0] !== hit.wz) notes.push("Brinks ID '" + sid + "' \u2192 matched to " + o.b[5] + "'s cashout");
      if (hit.date !== bagDate_(o.b, opt).date) notes.push('late drop (cashout ' + hit.date + ')');
      empMatch[o.b[0]] = { date: hit.date, wz: hit.wz, note: notes.join('; ') };
    });
  // 1c. No employee match: exact amount (\u00b1$1) vs an unmatched cashout at the bag's store(s) in the prior 14 days (catches late drops)
  var rqMax = outs.reduce(function (m, x) { return x.date > m ? x.date : m; }, '');
  brRows.filter(function (b) {
    if (!b[0] || empMatch[b[0]] || brSkip_(b) || !(Number(b[7]) > 0)) return false;
    var n = brCands_(b, map).length;
    return n > 1 || (n === 1 && bagDate_(b, opt).date <= rqMax);   // single store: only once that day's RQ is in, so its own cashout gets first claim
  })
    .map(function (b) { return { b: b, t: b[2] instanceof Date ? b[2] : new Date(b[2]) }; })
    .sort(function (x, y) { return x.t - y.t; })
    .forEach(function (o) {
      var sid = String(o.b[3]).trim(), cands = brCands_(o.b, map), amt = Number(o.b[7]) || 0, cd = opt.fmt(o.t), lo = addDays_(cd, -14);
      var hit = outs.filter(function (x) { return !usedTx[x.tx] && cands.indexOf(x.wz) >= 0 && Math.abs(x.amt - amt) <= 1 && x.date <= cd && x.date >= lo; })
        .sort(function (x, y) { return x.date < y.date ? 1 : x.date > y.date ? -1 : 0; })[0];
      if (!hit) return;
      usedTx[hit.tx] = 1;
      var notes = cands.length > 1 ? ["Brinks ID '" + sid + "' shared \u2192 matched to " + hit.wz + ' cashout by amount'] : [];
      if (hit.date !== bagDate_(o.b, opt).date) notes.push('late drop (cashout ' + hit.date + ')');
      empMatch[o.b[0]] = { date: hit.date, wz: hit.wz, note: notes.join('; ') };
    });
  brRows.forEach(function (b) {
    if (!b[0] || brSkip_(b)) return;
    var sid = String(b[3]).trim(), em = empMatch[b[0]];
    var d = em ? { date: em.date, note: em.note } : bagDate_(b, opt);
    if (d.date !== bizDate) return;
    var cands = em ? [em.wz] : brCands_(b, map);
    if (cands.length > 1) {
      var byLoc = cands.filter(function (w) { return map.byWZ[w].location === String(b[4]).trim(); });
      if (byLoc.length === 1) cands = byLoc;
    }
    bags.push({ barcode: b[0], amount: Number(b[7]) || 0, status: String(b[6]), cands: cands, sid: sid, dateNote: d.note, comments: String(b[10] || '') });
  });

  var brBy = {}, bagsBy = {}, notesBy = {}, assignments = [], unmapped = [];
  var addNote = function (w, n) { (notesBy[w] = notesBy[w] || []).push(n); };
  var assign = function (bag, w, note) {
    brBy[w] = (brBy[w] || 0) + bag.amount;
    (bagsBy[w] = bagsBy[w] || []).push(bag);
    if (bag.status.toUpperCase() !== 'ACH') addNote(w, '$' + bag.amount + ' bag ' + bag.status);
    if (bag.dateNote) addNote(w, bag.dateNote);
    if (note) addNote(w, note);
    assignments.push({ barcode: bag.barcode, date: bizDate, wz: w, note: [bag.dateNote, note].filter(String).join('; ') });
  };
  // single-candidate bags first, then ambiguous ones by closest remaining amount
  bags.filter(function (b) { return b.cands.length === 1; }).forEach(function (b) { assign(b, b.cands[0], ''); });
  bags.filter(function (b) { return b.cands.length > 1; }).forEach(function (b) {
    var pick = b.cands.slice().sort(function (x, y) {
      return Math.abs((rqBy[x] || 0) - (brBy[x] || 0) - b.amount) - Math.abs((rqBy[y] || 0) - (brBy[y] || 0) - b.amount);
    })[0];
    assign(b, pick, "Brinks ID '" + b.sid + "' shared — matched by amount");
  });
  bags.filter(function (b) { return b.cands.length === 0; }).forEach(function (b) { unmapped.push(b); });

  // 3. Rows for every active store (+ any store with activity)
  var stores = Object.keys(map.byWZ).filter(function (w) { return map.byWZ[w].active || rqBy[w] !== undefined && brBy[w] !== undefined; });
  var rows = [], summary = {};
  stores.forEach(function (w) {
    var s = map.byWZ[w];
    if (!s.active && !brBy[w]) return;           // pending stores: skip until Brinks is live
    var hasRQ = rqBy[w] !== undefined, r = round2_(rqBy[w] || 0), b = round2_(brBy[w] || 0), d = round2_(b - r);
    var st;
    if (!hasRQ && !b) st = 'NO ACTIVITY';
    else if (!hasRQ) st = 'BRINKS ONLY';
    else if (r > 0 && b === 0) st = 'MISSING DEPOSIT';
    else if (d === 0) st = r === 0 ? 'MATCH ($0 day)' : 'MATCH';
    else if (Math.abs(d) <= opt.tolerance) st = 'MINOR';
    else st = d < 0 ? 'SHORT' : 'OVER';
    summary[st] = (summary[st] || 0) + 1;
    rows.push({ date: bizDate, wz: w, store: s.name, company: s.company, dm: s.dm, dmEmail: s.dmEmail, rq: r, brinks: b, diff: d,
      status: st, bags: (bagsBy[w] || []).length, notes: (notesBy[w] || []).filter(function (n, i, x) { return x.indexOf(n) === i; }).join(' | ') });
  });
  return { rows: rows, summary: summary, assignments: assignments, unmapped: unmapped };
}

/** A bag's business date: explicit date in comments wins; else created date, minus 1 day if before cutoff hour. */
/** Stores a Brinks row can belong to: Location # (unique per store) first, then a WZ-style Store ID, then the (possibly shared) Brinks Store ID. */
function brCands_(b, map) {
  var loc = String(b[17] || '').trim();
  if (loc && map.byLocNum && map.byLocNum[loc]) return [map.byLocNum[loc]];
  var sid = String(b[3]).trim();
  return /^WZ\d/i.test(sid) ? [sid.toUpperCase()] : (map.byBrinksId[sid] || []);
}
function brSkip_(b) { return /^(voided|test)$/i.test(String(b[6]).trim()); }

function bagDate_(b, opt) {
  var created = b[2] instanceof Date ? b[2] : new Date(b[2]);
  var cDate = opt.fmt(created);
  if (b[10] instanceof Date) {                                   // comment that is only a date gets stored as a Date by Sheets
    var cd0 = opt.fmt(b[10]), df = (new Date(cDate) - new Date(cd0)) / 864e5;
    if (df >= 0 && df <= 14) return { date: cd0, note: df > 0 ? 'late drop (comment says ' + cd0 + ')' : '' };
  }
  var com = String(b[10] || '');
  var m = com.match(/(?<![$\d.])\b(1[0-2]|0?[1-9])[.\/-](3[01]|[12]\d|0?[1-9])(?:[.\/-](\d{2,4}))?\b/);
  if (m) {
    var y = m[3] ? (m[3].length === 2 ? '20' + m[3] : m[3]) : cDate.slice(0, 4);
    var d = y + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2);
    var diff = (new Date(cDate) - new Date(d)) / 864e5;
    if (diff >= 0 && diff <= 14) return { date: d, note: diff > 0 ? 'late drop (comment says ' + d + ')' : '' };
  }
  if (opt.hour(created) < opt.cutoffHour) return { date: addDays_(cDate, -1), note: 'next-morning drop' };
  return { date: cDate, note: '' };
}

// ───────────────────────── DAILY LOG ─────────────────────────
function upsertDailyLog_(bizDate, rows) {
  var sh = sheet_(TABS.LOG, LOG_HEADERS);
  var data = sh.getDataRange().getValues();
  var idx = {};
  for (var i = 1; i < data.length; i++) idx[normDate_(data[i][0]) + '|' + data[i][1]] = i;
  var now = new Date(), tol = Number(getCfg_().TOLERANCE) || 5, appends = [];
  rows.forEach(function (r) {
    var key = bizDate + '|' + r.wz, existing = idx[key] !== undefined ? data[idx[key]] : null;
    var manual = existing ? MANUAL_COLS.map(function (c) { return existing[c]; }) : ['', '', '', ''];
    var adj = Number(manual[0]) || 0, fd = round2_(r.diff + adj);
    var fs = adj ? (Math.abs(fd) <= tol ? 'RESOLVED (manual)' : r.status + ' (adj)') : r.status;
    var row = [bizDate, r.wz, r.store, r.company, r.dm, r.rq, r.brinks, r.diff, r.status, r.bags, r.notes]
      .concat(manual).concat([fd, fs, now]);
    if (existing) sh.getRange(idx[key] + 1, 1, 1, row.length).setValues([row]);
    else appends.push(row);
  });
  if (appends.length) sh.getRange(sh.getLastRow() + 1, 1, appends.length, appends[0].length).setValues(appends);
}

// ───────────────────────── ALERT ─────────────────────────
function sendAlert_(bizDate) {
  var cfg = getCfg_();
  var log = readTab_(TABS.LOG).filter(function (r) { return normDate_(r[0]) === bizDate; });
  var groups = { 'MISSING DEPOSIT': [], 'SHORT': [], 'OVER': [], 'BRINKS ONLY': [], 'MINOR': [], 'NETTED': [] };
  var matched = 0, total = 0;
  log.forEach(function (r) {
    var st = String(r[16] || r[8]);
    if (/RESOLVED/.test(st)) { matched++; total++; if (/late drop/.test(st)) groups.NETTED.push(r); return; }
    var base = String(r[8]);
    if (base === 'NO ACTIVITY') return;
    total++;
    if (/^MATCH/.test(base)) matched++;
    else if (groups[base]) groups[base].push(r);
  });
  var money = function (n) { return (n < 0 ? '-$' : '$') + Math.abs(Number(n)).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); };
  var table = function (title, color, rows) {
    if (!rows.length) return '';
    var h = '<h3 style="color:' + color + ';margin:18px 0 6px">' + title + ' (' + rows.length + ')</h3>' +
      '<table cellpadding="6" style="border-collapse:collapse;font:13px Arial">' +
      '<tr style="background:#f1f3f4"><th align="left">Store</th><th align="left">DM</th><th align="right">RQ</th><th align="right">Brinks</th><th align="right">Diff</th><th align="left">Notes</th></tr>';
    rows.sort(function (a, b) { return Math.abs(b[7]) - Math.abs(a[7]); }).forEach(function (r) {
      h += '<tr style="border-top:1px solid #e0e0e0"><td>' + r[2] + ' <span style="color:#888">' + r[1] + '</span></td><td>' + r[4] +
        '</td><td align="right">' + money(r[5]) + '</td><td align="right">' + money(r[6]) + '</td><td align="right"><b>' + money(r[7]) +
        '</b></td><td style="color:#666">' + (r[10] || '') + '</td></tr>';
    });
    return h + '</table>';
  };
  var day = Utilities.formatDate(new Date(bizDate + 'T12:00:00'), Session.getScriptTimeZone(), 'EEE M/d');
  var subj = '💰 Cash Recon ' + day + ' — ' + groups['MISSING DEPOSIT'].length + ' missing · ' +
    (groups.SHORT.length + groups.OVER.length) + ' variance · ' + matched + '/' + total + ' matched';
  var html = '<div style="font:14px Arial;color:#202124">' +
    (isTest_(cfg) ? '<p style="background:#fff3cd;padding:8px">TEST MODE — this alert went to the admin only.</p>' : '') +
    '<p><b>' + matched + ' of ' + total + '</b> stores matched for <b>' + day + '</b>. Bags dropped the next morning are counted for the prior day; late bags update this log automatically.</p>' +
    table('❌ Missing deposit', '#c5221f', groups['MISSING DEPOSIT']) +
    table('⚠️ Short', '#e37400', groups.SHORT) +
    table('⚠️ Over', '#e37400', groups.OVER) +
    table('🔵 Brinks bag, no RQ cashout', '#1a73e8', groups['BRINKS ONLY']) +
    table('🟡 Within $' + cfg.TOLERANCE + ' (coins)', '#9aa0a6', groups.MINOR) +
    table('✅ Resolved — late drop nets out', '#188038', groups.NETTED) +
    '<p style="color:#888;font-size:12px;margin-top:20px">Log manual adjustments in Daily_Log (Manual Adj $ / Adj Reason). They are never overwritten.</p></div>';
  var to = [adminEmail_()].concat(String(cfg.RECON_TEAM || '').split(/[,;\s]+/));
  if (!isTest_(cfg)) {
    log.forEach(function (r) { var s = loadStoreMap_().byWZ[r[1]]; if (s && s.dmEmail) to.push(s.dmEmail); });
    if (cfg.EXTRA_ALERT_TO) to = to.concat(cfg.EXTRA_ALERT_TO.split(','));
  }
  to = to.map(function (e) { return String(e).trim().toLowerCase(); }).filter(function (e, i, a) { return e && a.indexOf(e) === i; });
  MailApp.sendEmail({ to: to.join(','), subject: subj, htmlBody: html });
  logRun('Alert sent for ' + bizDate + ' to ' + to.join(','));
}

// ───────────────────────── SETUP / HELPERS ─────────────────────────
function setupTabs() {
  var ss = ss_();
  if (!ss.getSheetByName(TABS.MAP) && ss.getSheetByName('Untitled')) ss.getSheetByName('Untitled').setName(TABS.MAP);
  var cfgSh = sheet_(TABS.CFG, ['Key', 'Value', 'What it does']);
  var have = cfgSh.getDataRange().getValues().map(function (r) { return r[0]; });
  var help = {
    TEST_MODE: 'TRUE = alerts only to ADMIN_EMAIL. Set FALSE to send to each store\'s DM.',
    ADMIN_EMAIL: 'Blank = the Google account running the script', EXTRA_ALERT_TO: 'Extra recipients (comma-separated), live mode only',
    TOLERANCE: '$ difference treated as coins/minor', NEXT_MORNING_CUTOFF_HOUR: 'Bags created before this hour count for the previous day',
    RQ_FOLDER_PATH: 'Drive folder the Mac mini saves RQ files to', BRINKS_FOLDER_PATH: 'Drive folder for saved Brinks attachments',
    RECON_TEAM: 'Back-office team \u2014 always gets the daily alert (comma-separated emails)',
    BRINKS_QUERY: 'Gmail search for the Brinks Deposit Report', DEVICE_QUERY: 'Gmail search for the Brinks Device Activity report (all safe types)',
    BRINKS_IMPORT_PATH: 'Drive folder for manual Brinks exports (loaded automatically)', RECHECK_DAYS: 'Recent days re-reconciled each run (late bags)',
    MISSING_FILE_ALERT_HOUR: 'Hour after which a missing-input email is sent'
  };
  Object.keys(DEFAULT_CFG).forEach(function (k) { if (have.indexOf(k) < 0) cfgSh.appendRow([k, DEFAULT_CFG[k], help[k] || '']); });
  sheet_(TABS.RQ, RQ_HEADERS); sheet_(TABS.BR, BR_HEADERS); sheet_(TABS.LOG, LOG_HEADERS); sheet_(TABS.RUN, ['Time', 'Message']);
}

function sheet_(name, headers) {
  var ss = ss_(), sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#f1f3f4');
    sh.setFrozenRows(1);
  }
  return sh;
}
function readTab_(name) { var sh = ss_().getSheetByName(name); return sh ? sh.getDataRange().getValues().slice(1) : []; }

function getCfg_() {
  var out = JSON.parse(JSON.stringify(DEFAULT_CFG));
  readTab_(TABS.CFG).forEach(function (r) { if (r[0]) out[String(r[0]).trim()] = String(r[1]).trim(); });
  return out;
}
/** TEST_MODE is ON unless explicitly FALSE/NO/OFF/0 — fails safe (admin-only) on any odd value. */
function isTest_(cfg) { return !/^(false|no|off|0)$/i.test(String(cfg.TEST_MODE).trim()); }
function adminEmail_() { return getCfg_().ADMIN_EMAIL || Session.getEffectiveUser().getEmail(); }

function loadStoreMap_() {
  var rows = ss_().getSheetByName(TABS.MAP).getDataRange().getValues();
  var h = rows[0].map(String), c = function (n) { return h.indexOf(n); };
  var iN = c('Store Name'), iCo = c('Brand'), iDM = c('DM'), iDE = c('DM Email'), iA = c('Active'), iB = c('Brinks Store ID'),
      iL = c('Brinks Location Name'), iW = c('WZ Store ID'), iLN = c('Brinks ID');
  var byWZ = {}, byBrinksId = {}, byLocNum = {};
  rows.slice(1).forEach(function (r) {
    var w = String(r[iW]).trim(); if (!w) return;
    byWZ[w] = { name: r[iN], company: r[iCo], dm: String(r[iDM]).trim(), dmEmail: String(r[iDE]).trim(), active: String(r[iA]).trim().toUpperCase() === 'Y',
      brinksId: String(r[iB]).trim(), location: String(r[iL]).trim() };
    if (byWZ[w].brinksId) (byBrinksId[byWZ[w].brinksId] = byBrinksId[byWZ[w].brinksId] || []).push(w);
    var ln = iLN >= 0 ? String(r[iLN]).trim() : '';
    if (/^\d+$/.test(ln)) byLocNum[ln] = w;
  });
  return { byWZ: byWZ, byBrinksId: byBrinksId, byLocNum: byLocNum };
}

function replaceRowsByDate_(tab, headers, dateCol, date, newRows) {
  var sh = sheet_(tab, headers);
  var data = sh.getDataRange().getValues();
  var keep = [data[0]].concat(data.slice(1).filter(function (r) { return normDate_(r[dateCol]) !== date; }));
  var all = keep.concat(newRows);
  sh.clearContents();
  sh.getRange(1, 1, all.length, headers.length).setValues(all.map(function (r) { return r.slice(0, headers.length); }));
}

function rqDates_() { return readTab_(TABS.RQ).map(function (r) { return normDate_(r[0]); }).filter(function (d, i, a) { return a.indexOf(d) === i; }); }
function brinksReportDates_() { return JSON.parse(PropertiesService.getScriptProperties().getProperty('brinks_report_dates') || '[]'); }
function addBrinksReportDate_(d) {
  var a = brinksReportDates_(); if (a.indexOf(d) < 0) a.push(d);
  PropertiesService.getScriptProperties().setProperty('brinks_report_dates', JSON.stringify(a.slice(-60)));
}

/** xlsx blob -> 2D values, via a temporary Google Sheets conversion (Advanced Drive service). */
function xlsxToValues_(blob, name) {
  var file = Drive.Files.create({ name: 'tmp_' + name, mimeType: MimeType.GOOGLE_SHEETS }, blob);
  try {
    var ss = SpreadsheetApp.openById(file.id);
    ss.setSpreadsheetTimeZone(Session.getScriptTimeZone());   // read Brinks/RQ wall-clock times as Eastern, not the Drive default
    return ss.getSheets()[0].getDataRange().getValues();
  } finally {
    Drive.Files.remove(file.id);
  }
}

function folderByPath_(path, create) {
  var f = DriveApp.getRootFolder();
  var parts = path.split('/');
  for (var i = 0; i < parts.length; i++) {
    var it = f.getFoldersByName(parts[i]);
    if (it.hasNext()) f = it.next();
    else if (create) f = f.createFolder(parts[i]);
    else return null;
  }
  return f;
}

function logRun(msg) { try { sheet_(TABS.RUN, ['Time', 'Message']).appendRow([new Date(), String(msg).slice(0, 5000)]); } catch (e) {} console.log(msg); }
function num_(v) { var n = Number(String(v).replace(/[$,]/g, '')); return isNaN(n) ? 0 : n; }
function round2_(n) { return Math.round(n * 100) / 100; }
function addDays_(ymd, n) { var d = new Date(ymd + 'T12:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); }
function normDate_(v, opt) {
  if (v instanceof Date) return opt ? opt.fmt(v) : Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v).slice(0, 10);
}

// ───────────────────────── NET MATCHING ─────────────────────────
/** Rolling net matching across the recheck window (pure core + sheet wrapper). */
// Pure: items = [{key, wz, date, diff, adj}] -> [{wz, date, status, note, finalDiff}] for rows that change
function money_(n) { return (n < 0 ? '-$' : '$') + Math.abs(n).toFixed(2).replace(/\.00$/, ''); }
function netCore(items, tol) {
  var out = [], groups = {};
  items.forEach(function (it) { (groups[it.key] = groups[it.key] || []).push(it); });
  Object.keys(groups).forEach(function (k) {
    var rows = groups[k];
    var dates = rows.map(function (r) { return r.date; }).filter(function (d, i, a) { return a.indexOf(d) === i; }).sort();
    var byDate = dates.map(function (d) {
      var rs = rows.filter(function (r) { return r.date === d; });
      return { date: d, rows: rs, diff: Math.round(rs.reduce(function (s, r) { return s + r.diff; }, 0) * 100) / 100,
               adj: rs.some(function (r) { return r.adj; }) };
    });
    var used = {};
    for (var a = 0; a < byDate.length; a++) {
      if (used[a] || byDate[a].adj || Math.abs(byDate[a].diff) <= tol) continue;
      var sum = byDate[a].diff, best = null;
      for (var b = a + 1; b < byDate.length; b++) {
        if (byDate[b].adj) break;
        sum = Math.round((sum + byDate[b].diff) * 100) / 100;
        if (Math.abs(byDate[b].diff) <= tol) continue;
        var opposite = (byDate[a].diff < 0) !== (byDate[b].diff < 0);
        if (!opposite) continue;
        if (Math.abs(sum) <= tol) { best = { b: b, sum: sum, full: true }; break; }
        // partial: the pair cancels most of the smaller side
        var small = Math.min(Math.abs(byDate[a].diff), Math.abs(byDate[b].diff));
        if (!best && Math.abs(sum) < 0.5 * Math.max(Math.abs(byDate[a].diff), Math.abs(byDate[b].diff)) && small > tol)
          best = { b: b, sum: sum, full: false };
      }
      if (!best) continue;
      var run = byDate.slice(a, best.b + 1);
      var label = run.map(function (x) { return x.date.slice(5).replace('-', '/'); }).join(' + ');
      run.forEach(function (x, i) {
        if (best.full) used[a + i] = 1;
        x.rows.forEach(function (r) {
          if (Math.abs(r.diff) <= tol) return;
          out.push(best.full
            ? { wz: r.wz, date: r.date, status: 'RESOLVED (late drop)', finalDiff: best.sum,
                note: 'nets to ' + money_(best.sum) + ' across ' + label + (k !== r.wz ? ' (shared Brinks ID)' : '') }
            : { wz: r.wz, date: r.date, status: null, finalDiff: null,
                note: 'likely late drop across ' + label + ' — residual ' + money_(best.sum) });
        });
      });
      if (best.full) a = best.b;
    }
  });
  return out;
}

function netWindow_(dates) {
  var tol = Number(getCfg_().TOLERANCE) || 5;
  var map = loadStoreMap_();
  var sh = sheet_(TABS.LOG, LOG_HEADERS);
  var data = sh.getDataRange().getValues();
  var items = [], rowOf = {};
  for (var i = 1; i < data.length; i++) {
    var d = normDate_(data[i][0]);
    if (dates.indexOf(d) < 0) continue;
    var wz = data[i][1], s = map.byWZ[wz] || {};
    var shared = s.brinksId && (map.byBrinksId[s.brinksId] || []).length > 1;
    items.push({ key: shared ? 'BID:' + s.brinksId : wz, wz: wz, date: d, diff: Number(data[i][7]) || 0, adj: Number(data[i][11]) || 0 });
    rowOf[d + '|' + wz] = i;
  }
  var changes = netCore(items, tol), n = 0;
  changes.forEach(function (c) {
    var r = rowOf[c.date + '|' + c.wz]; if (r === undefined) return;
    data[r][10] = (data[r][10] ? data[r][10] + ' | ' : '') + c.note;
    if (c.status) { data[r][15] = c.finalDiff; data[r][16] = c.status; n++; }
    sh.getRange(r + 1, 11, 1, 1).setValues([[data[r][10]]]);
    sh.getRange(r + 1, 16, 1, 2).setValues([[data[r][15], data[r][16]]]);
  });
  logRun('Net matching ' + dates.join(',') + ': ' + n + ' rows resolved, ' + (changes.length - n) + ' partial notes');
}


// ───────────────────────── BRINKS: manual imports, new installs, catch-up ─────────────────────────
/** Any .xlsx dropped in Brinks Recon/Brinks Imports is loaded once (Device Activity or Deposit Report export). */
function ingestBrinksImports_() {
  var cfg = getCfg_(), props = PropertiesService.getScriptProperties();
  var folder = folderByPath_(cfg.BRINKS_IMPORT_PATH || 'Brinks Recon/Brinks Imports', true), files = folder.getFiles();
  var tz = Session.getScriptTimeZone();
  while (files.hasNext()) {
    var f = files.next();
    var gs = f.getMimeType() === MimeType.GOOGLE_SHEETS;
    if (!gs && !/\.xlsx?$/i.test(f.getName())) continue;
    var key = 'bi_' + f.getId() + '_' + f.getLastUpdated().getTime();
    if (props.getProperty(key)) continue;
    try {
      var rd = Utilities.formatDate(f.getDateCreated(), tz, 'yyyy-MM-dd'), vals;
      if (gs) { var s2 = SpreadsheetApp.openById(f.getId()); if (s2.getSpreadsheetTimeZone() !== tz) s2.setSpreadsheetTimeZone(tz); vals = s2.getSheets()[0].getDataRange().getValues(); }
      else vals = xlsxToValues_(f.getBlob(), f.getName());
      var parsed = parseBrinksRows(vals, rd, 'import:' + f.getName());
      upsertBrinks_(parsed);
      activateNewInstalls_(parsed);
      props.setProperty(key, '1');
      logRun('Brinks import ' + f.getName() + ': ' + parsed.length + ' rows');
    } catch (e) { logRun('Brinks import ' + f.getName() + ' skipped: ' + e.message); }
  }
}

/** A deposit under a WZ#### Store ID turns on the matching pending Store Map row (Brinks ID = Location #, Active = Y). */
function activateNewInstalls_(parsed) {
  var seen = {};
  parsed.forEach(function (r) { if (/^WZ\d+/i.test(r[3]) && /^\d+$/.test(String(r[17]))) seen[r[3].toUpperCase()] = r; });
  var wzs = Object.keys(seen); if (!wzs.length) return;
  var sh = ss_().getSheetByName(TABS.MAP), data = sh.getDataRange().getValues(), h = data[0].map(String);
  var c = function (n) { return h.indexOf(n); }, iW = c('WZ Store ID'), iB = c('Brinks ID'), iS = c('Brinks Store ID'), iL = c('Brinks Location Name'),
      iA = c('Active'), iN = c('Notes'), iName = c('Store Name');
  var tz = Session.getScriptTimeZone(), today = Utilities.formatDate(new Date(), tz, 'M/d/yyyy'), done = [];
  for (var i = 1; i < data.length; i++) {
    var w = String(data[i][iW]).trim().toUpperCase(), r = seen[w];
    if (!r || String(data[i][iB]).trim()) continue;
    var row = i + 1;
    sh.getRange(row, iB + 1).setValue(String(r[17]));
    if (iS >= 0 && !String(data[i][iS]).trim()) sh.getRange(row, iS + 1).setValue(r[3]);
    if (iL >= 0 && !String(data[i][iL]).trim()) sh.getRange(row, iL + 1).setValue(r[4]);
    sh.getRange(row, iA + 1).setValue('Y');
    if (iN >= 0) sh.getRange(row, iN + 1).setValue('Auto-activated ' + today + ': first Brinks deposit (Location # ' + r[17] + ')');
    done.push(w + ' ' + data[i][iName] + ' → Location # ' + r[17]);
  }
  if (!done.length) return;
  logRun('New Brinks installs activated: ' + done.join(', '));
  try {
    MailApp.sendEmail(adminEmail_(), 'Cash Recon: ' + done.length + ' new Brinks install(s) activated',
      'These stores made their first Brinks deposit and are now in the daily recon:\n\n' + done.join('\n') +
      '\n\nFill in DM / Store Manager on the Store Map so they get the right alerts.');
  } catch (e) { logRun('Install email failed: ' + e.message); }
}

/** Manual: load everything new (incl. Brinks Imports), then re-run recon + netting for the last N days. */
function brinksCatchUp(days) {
  var n = Number(days) || 7, tz = Session.getScriptTimeZone();
  var hdr = ss_().getSheetByName(TABS.BR);
  if (hdr && hdr.getLastColumn() < BR_HEADERS.length) hdr.getRange(1, 1, 1, BR_HEADERS.length).setValues([BR_HEADERS]).setFontWeight('bold');
  ingestBrinks();
  var target = addDays_(Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'), -1), win = [];
  for (var i = n - 1; i >= 0; i--) { var d = addDays_(target, -i); win.push(d); reconcileAndWrite_(d); }
  netWindow_(win);
  logRun('brinksCatchUp: re-ran ' + win[0] + ' → ' + win[win.length - 1]);
}
