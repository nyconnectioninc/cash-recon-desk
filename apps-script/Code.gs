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
  SM_ALERTS: 'issues',               // store-manager emails: issues = only stores with a problem · all = every store daily · off
  TOLERANCE: '5',                    // $ — differences within this are reconciled as "MATCH (coins)"
  NEXT_MORNING_CUTOFF_HOUR: '11',    // bags created before this hour count for the PREVIOUS business day
  RQ_FOLDER_PATH: 'Brinks Recon/RQ Cash Audit',
  BRINKS_FOLDER_PATH: 'Brinks Recon/Brinks Deposits',
  BRINKS_QUERY: 'from:no-reply@brinksinc.com subject:"Deposit Report" has:attachment newer_than:10d',
  DEVICE_QUERY: 'from:no-reply@brinksinc.com subject:"Device Activity" has:attachment newer_than:10d',   // all safe types (bags + smart safes)
  BRINKS_IMPORT_PATH: 'Brinks Recon/Brinks Imports',   // drop a manual Device Activity / Deposit Report export here
  ACH_QUERY: 'from:brinksinc.com subject:ACH has:attachment newer_than:10d',   // Brinks ACH report (backup source for smart safes)
  ACH_FOLDER_PATH: 'Brinks Recon/ACH Reports',         // or drop an ACH report export here
  RECHECK_DAYS: '3',                 // re-reconcile this many recent days each run (late bags)
  MISSING_FILE_ALERT_HOUR: '11',     // after this hour, email admin if inputs still missing
  ALERT_SEND_HOUR: '12',             // daily recon email goes out at/after this hour (recon itself keeps refreshing from 6 AM)
  RUN_END_HOUR: '15',                // runDaily stops acting at this hour (checks every 15 min until then)
  MAIL_VIA: 'outlook'                // outlook = send recon emails from the work account via the Zapier relay (falls back to Gmail) · gmail = Gmail only
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
    .addItem('Load new files + re-run last 14 days', 'achCatchUp')
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
  logRun('Triggers installed (runDaily every 15 min; acts 6 AM until RUN_END_HOUR; email at ALERT_SEND_HOUR)');
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
    var cfg = getCfg_();
    var endHour = Number(cfg.RUN_END_HOUR) || 15, sendHour = Number(cfg.ALERT_SEND_HOUR) || 12;
    if (hour < 6 || hour >= endHour) return;                  // quiet outside the 6 AM – RUN_END_HOUR window

    setupTabs();
    var props = PropertiesService.getScriptProperties();
    var sig0 = dataSig_();
    ingestAll();
    var changed = dataSig_() !== sig0 || props.getProperty('recon_sig') !== dataSig_();
    var today = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    var target = addDays_(today, -1);
    var haveRQ = rqDates_().indexOf(target) >= 0;
    var haveBR = brinksReportDates_().indexOf(today) >= 0;

    // Once yesterday's email is out, only re-run the (slow, lock-holding) recon when new data actually arrived,
    // so the dashboard's saves aren't blocked by a 6-minute run every 15 minutes.
    if (haveRQ && haveBR && props.getProperty('alert_sent_' + target) === '1' && !changed) return;
    if (haveRQ && haveBR) {
      // Yesterday first, so the email never waits on the older re-checks (each day takes ~1-3 min with ~100 stores
      // and a run is capped at 6 min). Older days are re-checked with whatever time is left; the next run continues.
      var t0 = now.getTime(), n = Number(cfg.RECHECK_DAYS) || 3, win = [target];
      reconcileAndWrite_(target);
      var due = hour >= sendHour && props.getProperty('alert_sent_' + target) !== '1';
      if (due) netWindow_([target]);
      if (due) { sendAlert_(target); props.setProperty('alert_sent_' + target, '1'); }   // email waits for ALERT_SEND_HOUR
      for (var i = 1; i < n; i++) {
        if (Date.now() - t0 > 240000) { logRun('Re-check of older days deferred to the next run (time limit)'); break; }
        var d = addDays_(target, -i); reconcileAndWrite_(d); win.push(d);
      }
      if (win.length > 1) netWindow_(win);
      props.setProperty('recon_sig', dataSig_());   // this data is reconciled; next full run only when new rows land
    } else if (hour >= Number(cfg.MISSING_FILE_ALERT_HOUR) && props.getProperty('missing_sent_' + target) !== '1') {
      var miss = [];
      if (!haveRQ) miss.push('RQ Cash Audit file for ' + target + ' (check the Mac mini — ~/Library/Logs/rq-pull.log)');
      if (!haveBR) miss.push('Brinks Deposit Report email for ' + today);
      MailApp.sendEmail(adminEmail_(), '⚠️ Cash Recon blocked — missing input for ' + target,
        'The recon for ' + target + ' could not run because these inputs are missing:\n\n• ' + miss.join('\n• ') +
        '\n\nIt will run automatically as soon as they arrive (checks every 15 min until ' + endHour + ':00; the recon email goes out at ' + sendHour + ':00 or as soon as inputs arrive after that), or use Cash Recon → Run now.');
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
  try { ingestACH(); } catch (e) { logRun('ACH ingest error: ' + e.message); }
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
    var bad = rqFileProblem_(parsed, m[1]);
    if (bad) {                                   // wrong / partial file: keep yesterday's good data, tell the admin once
      logRun('RQ file REJECTED ' + f.getName() + ': ' + bad);
      if (!props.getProperty('rqbad_' + key)) {
        props.setProperty('rqbad_' + key, '1');
        try {
          MailApp.sendEmail({ to: adminEmail_(), subject: '⚠️ RQ file rejected — ' + f.getName(),
            htmlBody: '<p>The RQ Cash Audit file <b>' + f.getName() + '</b> was <b>not imported</b>: ' + bad + '.</p>' +
              '<p>Nothing was overwritten. Most likely a custom Cash Audit Trail run in Hub got picked up. Fix: in Hub re-run the ' +
              '"Yesterday · all stores · CASHOUT" report, export it, name it <code>' + f.getName() + '</code> and replace the file in ' +
              '<i>Brinks Recon › RQ Cash Audit</i>, then run catchUp (or wait for the next run).</p>' });
        } catch (e) { logRun('RQ reject email failed: ' + e.message); }
      }
      continue;
    }
    replaceRowsByDate_(TABS.RQ, RQ_HEADERS, 0, m[1], parsed);
    props.setProperty(key, '1');
    logRun('RQ ingested ' + f.getName() + ': ' + parsed.length + ' cashouts');
  }
}

/** Sanity check before an RQ file replaces a day: right dates inside, and (close to) all stores. '' = OK. */
function rqFileProblem_(parsed, bizDate) {
  if (!parsed.length) return 'it has no cashout rows';
  var tz = Session.getScriptTimeZone(), next = Utilities.formatDate(new Date(new Date(bizDate + 'T12:00:00').getTime() + 864e5), tz, 'yyyy-MM-dd');
  var dOf = function (v) {
    if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    var s = String(v).trim(), a = s.match(/^(\d{4})-(\d{2})-(\d{2})/), b = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    return a ? a[0] : b ? b[3] + '-' + ('0' + b[1]).slice(-2) + '-' + ('0' + b[2]).slice(-2) : '';
  };
  var seen = {}, ok = 0, dated = 0, other = {};
  parsed.forEach(function (r) { var d = dOf(r[4]); if (r[1]) seen[r[1]] = 1; if (!d) return; dated++;
    if (d === bizDate || d === next) ok++; else other[d] = (other[d] || 0) + 1; });
  if (dated && ok / dated < 0.8) return 'only ' + ok + ' of ' + dated + ' cashouts are dated ' + bizDate + ' (file also has ' + Object.keys(other).sort().join(', ') + ')';
  var smap = loadStoreMap_(), active = Object.keys(smap.byWZ).filter(function (w) { return smap.byWZ[w].active; }).length;
  var n = Object.keys(seen).length, min = Math.max(5, Math.floor(active * 0.5));
  if (n < min) return 'it only covers ' + n + ' stores (expected about ' + active + ' — looks filtered)';
  return '';
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

  // 1a. Range deposits: a bag whose comment covers a date range ("1556 Cash 10/30 - 11/05") is split across those days
  var rng = rangeDeposits_(bizDate, rqRows, brRows, map, opt);
  brRows = brRows.filter(function (b) { return !rng.skip[String(b[0])]; });

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
      var redir = map.redirect && map.redirect[String(cands[0]).toUpperCase()];
      var notes = cands.length > 1 ? [redir ? (hit.wz !== cands[0] ? 'Dropped in ' + cands[0] + ' safe \u2192 matched to ' + hit.wz + ' cashout by amount' : '')
        : "Brinks ID '" + sid + "' shared \u2192 matched to " + hit.wz + ' cashout by amount'].filter(String) : [];
      if (hit.date !== bagDate_(o.b, opt).date) notes.push('late drop (cashout ' + hit.date + ')');
      empMatch[o.b[0]] = { date: hit.date, wz: hit.wz, note: notes.join('; ') };
    });
  brRows.forEach(function (b) {
    if (!b[0] || brSkip_(b)) return;
    var sid = String(b[3]).trim(), em = empMatch[b[0]];
    var d = em ? { date: em.date, note: em.note } : bagDate_(b, opt);
    if (d.date !== bizDate) return;
    var cands = em ? [em.wz] : brCands_(b, map);
    if (cands.length > 1 && !(map.redirect && map.redirect[String(cands[0]).toUpperCase()])) {
      var byLoc = cands.filter(function (w) { return map.byWZ[w].location === String(b[4]).trim(); });
      if (byLoc.length === 1) cands = byLoc;
    }
    bags.push({ barcode: b[0], amount: Number(b[7]) || 0, status: String(b[6]), cands: cands, sid: sid, dateNote: d.note, comments: String(b[10] || '') });
  });

  rng.today.forEach(function (x) { bags.push(x); });
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
    else if (Math.abs(d) <= opt.tolerance) st = 'MATCH (coins)';   // within tolerance = reconciled, still reported
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
  var loc = String(b[17] || '').trim(), sid = String(b[3]).trim();
  var base = loc && map.byLocNum && map.byLocNum[loc] ? [map.byLocNum[loc]] : /^WZ\d/i.test(sid) ? [sid.toUpperCase()] : (map.byBrinksId[sid] || []);
  // Store Map "Deposits At": another store's cash goes into this safe (e.g. 1556Flat drops at 5th St), so its bags can belong to either
  if (base.length === 1 && map.redirect && map.redirect[base[0].toUpperCase()]) return base.concat(map.redirect[base[0].toUpperCase()]);
  return base;
}
function brSkip_(b) { return /^(voided|test|superseded)$/i.test(String(b[6]).trim()); }

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
    if (base === 'MINOR' || /coins/.test(base)) { matched++; groups.MINOR.push(r); }   // within tolerance: reconciled, listed for info
    else if (/^MATCH/.test(base)) matched++;
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
    table('🟡 Reconciled within $' + cfg.TOLERANCE + ' (coins, FYI — counted as matched)', '#9aa0a6', groups.MINOR) +
    table('✅ Resolved — late drop nets out', '#188038', groups.NETTED) +
    '<p style="color:#888;font-size:12px;margin-top:20px">Log manual adjustments in Daily_Log (Manual Adj $ / Adj Reason). They are never overwritten.</p></div>';
  var uniq = function (l) { return l.join(',').split(/[,;\s]+/).map(function (e) { return e.trim().toLowerCase(); }).filter(function (e, i, a) { return /@/.test(e) && a.indexOf(e) === i; }); };
  var smap = loadStoreMap_(), dms = [];
  log.forEach(function (r) { var s = smap.byWZ[r[1]]; if (s && s.dmEmail) dms.push(s.dmEmail); });
  var live = uniq([adminEmail_(), cfg.RECON_TEAM || ''].concat(dms, [cfg.EXTRA_ALERT_TO || '']));
  var to = isTest_(cfg) ? uniq([adminEmail_(), cfg.RECON_TEAM || '']) : live;
  if (isTest_(cfg)) html = html.replace('TEST MODE — this alert went to the admin only.</p>',
    'TEST MODE — this alert went to the admin only.<br><span style="font-size:12px">When live it goes to ' + live.length + ' people: ' + live.join(', ') + '</span></p>');
  var via = sendMail_({ to: to.join(','), subject: subj, htmlBody: html });
  logRun('Alert sent for ' + bizDate + ' via ' + via + ' to ' + to.join(','));
  try { sendStoreAlerts_(bizDate, log, cfg, day, money); } catch (e) { logRun('Store-manager alerts failed: ' + e); }
}

// ───────────────────────── STORE-MANAGER ALERTS ─────────────────────────
/** One email per store, to that store's SM (cc its DM). Only their own store.
 *  SM_ALERTS: issues (default) = stores with an open problem · all = every store with activity · off.
 *  TEST_MODE: nothing goes to SMs/DMs — the admin gets ONE preview email with every message that would have gone out. */
var SM_ACTION = {
  'MISSING DEPOSIT': 'RQ shows cash for this day but no Brinks bag was found. Drop the deposit today and reply with the bag # (or the reason it was not dropped).',
  'SHORT': 'Your Brinks bag is LESS than the RQ cashout. Recount the drawer/safe and reply with the reason (approved skim? who approved it?).',
  'OVER': 'Your Brinks bag is MORE than the RQ cashout. Check whether cash from another day was included and reply with the explanation.',
  'BRINKS ONLY': 'A Brinks bag was received but there is no RQ cashout for this day. Make sure the cash audit was completed in RQ and reply with the details.'
};
function sendStoreAlerts_(bizDate, log, cfg, day, money) {
  var mode = String(cfg.SM_ALERTS || 'issues').trim().toLowerCase();
  if (mode === 'off') return;
  var map = loadStoreMap_(), test = isTest_(cfg), out = [], noSM = [];
  log.forEach(function (r) {
    var fs = String(r[16] || r[8]), base = fs.replace(/ \(adj\)$/, '');
    if (/NO ACTIVITY/.test(base)) return;
    var open = !/RESOLVED/.test(fs) && !!SM_ACTION[base];
    if (!open && mode !== 'all') return;
    var s = map.byWZ[r[1]] || {};
    if (!s.smEmail) { if (open) noSM.push(r[2] + ' (' + r[1] + ')'); return; }
    var coins = base === 'MINOR' || /coins/.test(base);
    var label = open ? ({ 'MISSING DEPOSIT': 'Missing deposit', 'SHORT': 'Short', 'OVER': 'Over', 'BRINKS ONLY': 'Bag, no RQ cashout' })[base]
                     : coins ? 'Reconciled (within $' + cfg.TOLERANCE + ')' : 'Reconciled';
    var color = open ? (base === 'MISSING DEPOSIT' ? '#c5221f' : '#e37400') : '#188038';
    var diff = Number(r[15] !== '' && r[15] != null ? r[15] : r[7]) || 0;
    var subj = (open ? '⚠️ ' : '✅ ') + r[2] + ' — ' + label + ' for ' + day + (open ? ' (' + money(diff) + ')' : '');
    var html = '<div style="font:14px Arial;color:#202124;max-width:560px">' +
      '<p>Hi ' + (String(s.sm).split(/[ ,&]/)[0] || 'team') + ',</p>' +
      '<p>Here is the cash deposit check for <b>' + r[2] + '</b> (' + r[1] + ') for <b>' + day + '</b>:</p>' +
      '<table cellpadding="7" style="border-collapse:collapse;font:14px Arial;border:1px solid #e0e0e0">' +
      '<tr><td style="color:#666">Status</td><td><b style="color:' + color + '">' + label + '</b></td></tr>' +
      '<tr style="border-top:1px solid #eee"><td style="color:#666">RQ cash deposit</td><td>' + money(r[5]) + '</td></tr>' +
      '<tr style="border-top:1px solid #eee"><td style="color:#666">Brinks bag(s)</td><td>' + money(r[6]) + (r[9] ? ' · ' + r[9] + ' bag' + (r[9] == 1 ? '' : 's') : '') + '</td></tr>' +
      '<tr style="border-top:1px solid #eee"><td style="color:#666">Difference</td><td><b>' + money(diff) + '</b></td></tr>' +
      (r[10] ? '<tr style="border-top:1px solid #eee"><td style="color:#666">Notes</td><td style="color:#555">' + r[10] + '</td></tr>' : '') +
      '</table>' +
      (open ? '<p style="margin-top:14px"><b>What to do:</b> ' + SM_ACTION[base] + '</p><p>Reply to this email so your DM and the recon team can close it out.</p>'
            : '<p style="margin-top:14px">No action needed — thanks for keeping deposits on time.</p>') +
      '<p style="color:#888;font-size:12px;margin-top:18px">Automated from the Cash Recon system. Bags dropped the next morning are credited to the prior day automatically.</p></div>';
    out.push({ to: s.smEmail, cc: s.dmEmail, subj: subj, html: html });
  });
  if (!out.length && !noSM.length) return;
  if (test) {
    var prev = '<div style="font:14px Arial"><p style="background:#fff3cd;padding:8px">TEST MODE — preview of the ' + out.length +
      ' store-manager email(s) that WOULD go out for ' + day + '. Nothing was sent to stores.</p>' +
      (noSM.length ? '<p><b>Open issues with no SM email on the Store Map (' + noSM.length + '):</b> ' + noSM.join(', ') + '</p>' : '') +
      out.map(function (m) {
        return '<hr><p style="color:#666;font-size:12px">To: ' + m.to + (m.cc ? ' · Cc: ' + m.cc : '') + '<br>Subject: <b>' + m.subj + '</b></p>' + m.html;
      }).join('') + '</div>';
    sendMail_({ to: adminEmail_(), subject: '🧪 SM alert preview ' + day + ' — ' + out.length + ' store email(s)', htmlBody: prev });
    logRun('SM alert preview for ' + bizDate + ': ' + out.length + ' email(s), ' + noSM.length + ' open with no SM');
    return;
  }
  var sent = 0, skipped = 0;
  out.forEach(function (m) {
    var need = (m.to + (m.cc ? ',' + m.cc : '')).split(',').length;
    if (!mailRelay_() && MailApp.getRemainingDailyQuota() < need + 5) { skipped++; return; }   // Gmail only: keep headroom for the admin alert tomorrow
    sendMail_({ to: m.to, cc: m.cc || '', subject: m.subj, htmlBody: m.html, replyTo: [m.cc, adminEmail_()].filter(String).join(',') });
    sent++;
  });
  logRun('SM alerts for ' + bizDate + ': ' + sent + ' sent' + (skipped ? ', ' + skipped + ' skipped (daily email quota)' : '') +
    (noSM.length ? ', no SM email: ' + noSM.join(', ') : ''));
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
    SM_ALERTS: 'Store-manager emails: issues / all / off', TOLERANCE: '$ difference treated as coins/minor', NEXT_MORNING_CUTOFF_HOUR: 'Bags created before this hour count for the previous day',
    RQ_FOLDER_PATH: 'Drive folder the Mac mini saves RQ files to', BRINKS_FOLDER_PATH: 'Drive folder for saved Brinks attachments',
    RECON_TEAM: 'Back-office team \u2014 always gets the daily alert (comma-separated emails)',
    BRINKS_QUERY: 'Gmail search for the Brinks Deposit Report', DEVICE_QUERY: 'Gmail search for the Brinks Device Activity report (all safe types)',
    BRINKS_IMPORT_PATH: 'Drive folder for manual Brinks exports (loaded automatically)', RECHECK_DAYS: 'Recent days re-reconciled each run (late bags)',
    MISSING_FILE_ALERT_HOUR: 'Hour after which a missing-input email is sent',
    ALERT_SEND_HOUR: 'Hour the daily recon email goes out (24h clock; 12 = noon). Recon still refreshes from 6 AM.',
    RUN_END_HOUR: 'Hour the 15-min checks stop for the day (24h clock)',
    MAIL_VIA: 'outlook = recon emails go out from the work Outlook account (via Zapier, falls back to Gmail) · gmail = Gmail only',
    ACH_QUERY: 'Gmail search for the Brinks ACH report (backup for safe drops missing from Device Activity)',
    ACH_FOLDER_PATH: 'Drive folder for manual ACH report exports (loaded automatically)'
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
      iL = c('Brinks Location Name'), iW = c('WZ Store ID'), iLN = c('Brinks ID'), iDA = c('Deposits At'),
      iSM = c('Store Manager'), iSE = c('SM Email');
  var byWZ = {}, byBrinksId = {}, byLocNum = {}, redirect = {};
  rows.slice(1).forEach(function (r) {
    var w = String(r[iW]).trim(); if (!w) return;
    byWZ[w] = { name: r[iN], company: r[iCo], dm: String(r[iDM]).trim(), dmEmail: String(r[iDE]).trim(), active: String(r[iA]).trim().toUpperCase() === 'Y',
      brinksId: String(r[iB]).trim(), location: String(r[iL]).trim(),
      sm: iSM >= 0 ? String(r[iSM]).trim() : '', smEmail: iSE >= 0 ? String(r[iSE]).trim() : '' };
    if (byWZ[w].brinksId) (byBrinksId[byWZ[w].brinksId] = byBrinksId[byWZ[w].brinksId] || []).push(w);
    var ln = iLN >= 0 ? String(r[iLN]).trim() : '';
    if (/^\d+$/.test(ln)) byLocNum[ln] = w;
    var da = iDA >= 0 ? String(r[iDA]).trim().toUpperCase() : '';   // store whose cash is dropped in another store's safe
    if (da && da !== w.toUpperCase() && String(r[iA]).trim().toUpperCase() === 'Y') (redirect[da] = redirect[da] || []).push(w);
  });
  return { byWZ: byWZ, byBrinksId: byBrinksId, byLocNum: byLocNum, redirect: redirect };
}

function replaceRowsByDate_(tab, headers, dateCol, date, newRows) {
  var sh = sheet_(tab, headers);
  var data = sh.getDataRange().getValues();
  var keep = [data[0]].concat(data.slice(1).filter(function (r) { return normDate_(r[dateCol]) !== date; }));
  var all = keep.concat(newRows);
  sh.clearContents();
  sh.getRange(1, 1, all.length, headers.length).setValues(all.map(function (r) { return r.slice(0, headers.length); }));
}

/** Cheap fingerprint of the recon inputs (row counts); changes when new RQ / Brinks / ACH rows land. */
function dataSig_() {
  var ss = ss_();
  return ['RQ', 'BR'].map(function (k) { var sh = ss.getSheetByName(TABS[k]); return sh ? sh.getLastRow() : 0; }).join('|') +
    '|' + (PropertiesService.getScriptProperties().getProperty('brinks_report_dates') || '').length;
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

/** Manual: load everything new (incl. Brinks Imports), then re-run recon + netting for the last N days.
 *  Reconciling ~100 stores takes ~35 s/day, so long windows continue in a follow-up run (1 min later) instead of hitting the 6-min limit. */
function brinksCatchUp(days) {
  var t0 = Date.now(), n = Number(days) || 7, tz = Session.getScriptTimeZone();
  var hdr = ss_().getSheetByName(TABS.BR);
  if (hdr && hdr.getLastColumn() < BR_HEADERS.length) hdr.getRange(1, 1, 1, BR_HEADERS.length).setValues([BR_HEADERS]).setFontWeight('bold');
  ingestRQ();
  ingestBrinks();
  try { ingestACH(); } catch (e) { logRun('ACH ingest error: ' + e.message); }
  var target = addDays_(Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'), -1), win = [];
  for (var i = n - 1; i >= 0; i--) win.push(addDays_(target, -i));
  PropertiesService.getScriptProperties().setProperty('catchup_queue', JSON.stringify({ todo: win, all: win }));
  catchUpContinue(t0);
}

/** Works through the catch-up queue; re-schedules itself if it runs low on time, nets the whole window at the end. */
function catchUpContinue(t0) {
  t0 = typeof t0 === 'number' ? t0 : Date.now();
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'catchUpContinue') ScriptApp.deleteTrigger(t); });
  var props = PropertiesService.getScriptProperties(), q = JSON.parse(props.getProperty('catchup_queue') || 'null');
  if (!q) return;
  while (q.todo.length) {
    if (Date.now() - t0 > 270000) {
      props.setProperty('catchup_queue', JSON.stringify(q));
      ScriptApp.newTrigger('catchUpContinue').timeBased().after(60 * 1000).create();
      logRun('Catch-up paused (time limit) — continuing ' + q.todo[0] + ' → ' + q.todo[q.todo.length - 1] + ' in ~1 min');
      return;
    }
    reconcileAndWrite_(q.todo.shift());
    props.setProperty('catchup_queue', JSON.stringify(q));
  }
  netWindow_(q.all);
  props.deleteProperty('catchup_queue');
  logRun('Catch-up done: ' + q.all[0] + ' → ' + q.all[q.all.length - 1]);
}


/** Bags whose comment names a date range ("10/30 - 11/05") are pooled per store + range and spread over those days:
 *  each day gets up to its RQ cash (oldest first); what's left goes to the first day in the range with no RQ file (cash from
 *  before the RQ pull started), else to the last day. The store is the bag's own, or — when several stores can use the safe
 *  (shared ID / "Deposits At") — the one the comment names ("1556", "WZ1370", "1556Flat"). Returns today's share + bags to skip. */
function rangeDeposits_(bizDate, rqRows, brRows, map, opt) {
  var re = /(?<![$\d.])(1[0-2]|0?[1-9])[\/.](3[01]|[12]\d|0?[1-9])(?:[\/.]\d{2,4})?\s*(?:-|–|to|thru|through)\s*(1[0-2]|0?[1-9])[\/.](3[01]|[12]\d|0?[1-9])\b/i;
  var groups = {}, skip = {}, today = [];
  var tokens = function (w) {
    var s = map.byWZ[w] || {}, nm = String(s.name || '').toLowerCase(), t = [w.toLowerCase(), nm.replace(/\s+/g, '')];
    var d = w.replace(/\D/g, ''); if (d.length >= 3) t.push(d);
    var lead = (nm.match(/^\d{3,}/) || [])[0]; if (lead) t.push(lead);
    return t.filter(String);
  };
  var names = function (com, w) {
    return tokens(w).some(function (t) { return /^\d+$/.test(t) ? new RegExp('(^|\\D)' + t + '(\\D|$)').test(com) : com.indexOf(t) >= 0; });
  };
  brRows.forEach(function (b) {
    if (!b[0] || brSkip_(b)) return;
    var com = String(b[10] || ''), m = com.match(re); if (!m) return;
    var cands = brCands_(b, map), lc = com.toLowerCase(), target = null;
    if (cands.length === 1) target = cands[0];
    else { var hit = cands.filter(function (w) { return names(lc, w); }); if (hit.length === 1) target = hit[0]; }
    if (!target) return;
    var created = b[2] instanceof Date ? b[2] : new Date(b[2]), cd = opt.fmt(created), y = +cd.slice(0, 4);
    var mk = function (mo, da) { var yy = +mo > +cd.slice(5, 7) ? y - 1 : y; return yy + '-' + ('0' + mo).slice(-2) + '-' + ('0' + da).slice(-2); };
    var a = mk(m[1], m[2]), z = mk(m[3], m[4]);
    if (z < a || z > cd || (new Date(z) - new Date(a)) / 864e5 > 31) return;     // nonsense range: leave the bag alone
    var k = target + '|' + a + '|' + z, g = groups[k] = groups[k] || { wz: target, from: a, to: z, amt: 0, codes: [] };
    g.amt += Number(b[7]) || 0; g.codes.push(String(b[0])); skip[String(b[0])] = 1;
  });
  var rqDay = {}, rqHas = {};
  rqRows.forEach(function (r) { var d = normDate_(r[0], opt); rqHas[d] = 1; if (r[1]) rqDay[r[1] + '|' + d] = (rqDay[r[1] + '|' + d] || 0) + (Number(r[6]) || 0); });
  var used = {};   // overlapping ranges for the same store ("1 of 2" 10/30-11/05, "2 of 2" 10/30-11/04) share each day's cash
  Object.keys(groups).sort(function (x, y) { var gx = groups[x], gy = groups[y]; return gx.from < gy.from ? -1 : gx.from > gy.from ? 1 : gx.to < gy.to ? -1 : gx.to > gy.to ? 1 : 0; }).forEach(function (k) {
    var g = groups[k], days = [], d = g.from;
    while (d <= g.to) { days.push(d); d = addDays_(d, 1); }
    var share = {}, rem = round2_(g.amt);
    days.forEach(function (x) { var key = g.wz + '|' + x, need = round2_((rqDay[key] || 0) - (used[key] || 0)); var take = round2_(Math.min(rem, need));
      if (take > 0) { share[x] = take; used[key] = round2_((used[key] || 0) + take); rem = round2_(rem - take); } });
    if (rem > 0.004) { var spot = days.filter(function (x) { return !rqHas[x]; })[0] || g.to; share[spot] = round2_((share[spot] || 0) + rem); }
    if (share[bizDate]) today.push({ barcode: 'RANGE:' + g.codes.join('+'), amount: share[bizDate], status: 'ACH', cands: [g.wz], sid: '', comments: '',
      dateNote: 'Range deposit ' + g.from.slice(5).replace('-', '/') + '–' + g.to.slice(5).replace('-', '/') + ' (' + g.codes.length + ' bag' + (g.codes.length > 1 ? 's' : '') + ', $' + round2_(g.amt) + ') → $' + share[bizDate] + ' for this day' });
  });
  return { skip: skip, today: today };
}

// ───────────────────────── INGEST: ACH REPORT (backup source) ─────────────────────────
// The Brinks ACH report lists every credit Brinks sends to the bank, per bag (bag safes) or per drop (smart safes, "Bag Number" = device UUID).
// Every row is kept on ACH_Raw. Bag safes already come through the Deposit Report, so ACH only backs up SMART SAFES:
// per location + business day, if ACH credits exceed the Device Activity drops on Brinks_Raw, the gap is added as one
// "ACHB-<store>-<date>" row. It is recalculated every run, so when the missing drops arrive later the backup row drops to $0 ("Superseded").
var ACH_TAB = 'ACH_Raw';
var ACH_HEADERS = ['Key', 'Business Day', 'Customer', 'Location Name', 'Bag / Device', 'Created By', 'Activity', 'Process Date', 'Sent To ACH', 'Amount', 'Source', 'Loaded'];

function ingestACH() {
  var cfg = getCfg_(), props = PropertiesService.getScriptProperties(), tz = Session.getScriptTimeZone(), got = 0;
  var take = function (vals, src) { var p = parseACHRows_(vals, src); got += upsertACH_(p); logRun('ACH report ' + src + ': ' + p.length + ' rows'); };
  if (cfg.ACH_QUERY) GmailApp.search(cfg.ACH_QUERY).forEach(function (t) {
    t.getMessages().forEach(function (msg) {
      if (props.getProperty('ach_' + msg.getId())) return;
      try {
        msg.getAttachments().forEach(function (att) {
          if (/\.xlsx?$/i.test(att.getName())) take(xlsxToValues_(att.copyBlob(), att.getName()), 'email ' + Utilities.formatDate(msg.getDate(), tz, 'yyyy-MM-dd') + ' ' + att.getName());
        });
        props.setProperty('ach_' + msg.getId(), '1');
      } catch (e) {
        if (/ACH header/.test(e.message)) props.setProperty('ach_' + msg.getId(), 'not-ach');
        logRun('ACH email skipped: ' + e.message);
      }
    });
  });
  var folder = folderByPath_(cfg.ACH_FOLDER_PATH || 'Brinks Recon/ACH Reports', true), files = folder.getFiles();
  while (files.hasNext()) {
    var f = files.next(), gs = f.getMimeType() === MimeType.GOOGLE_SHEETS;
    if (!gs && !/\.xlsx?$/i.test(f.getName())) continue;
    var key = 'achf_' + f.getId() + '_' + f.getLastUpdated().getTime();
    if (props.getProperty(key)) continue;
    try {
      var vals;
      if (gs) { var s2 = SpreadsheetApp.openById(f.getId()); if (s2.getSpreadsheetTimeZone() !== tz) s2.setSpreadsheetTimeZone(tz); vals = s2.getSheets()[0].getDataRange().getValues(); }
      else vals = xlsxToValues_(f.getBlob(), f.getName());
      take(vals, f.getName());
      props.setProperty(key, '1');
    } catch (e) { logRun('ACH file ' + f.getName() + ' skipped: ' + e.message); }
  }
  achBackfill_();
  return got;
}

/** Pure: ACH report rows (either export layout) -> ACH_Raw rows. Identical rows in one file get #1, #2… so re-loads dedupe. */
function parseACHRows_(rows, src) {
  var hr = -1;
  for (var r = 0; r < Math.min(rows.length, 30); r++) {
    var hh = rows[r].map(function (x) { return String(x).trim(); });
    if (hh.indexOf('Bag Number') >= 0 && hh.indexOf('Business Day') >= 0) { hr = r; break; }
  }
  if (hr < 0) throw new Error('ACH header row not found');
  var h = rows[hr].map(function (x) { return String(x).trim(); }), c = function (n) { return h.indexOf(n); };
  var iBD = c('Business Day'), iCu = c('Customer Name'), iL = c('Location Name'), iB = c('Bag Number'), iBy = c('Created By'),
      iA = c('Activity'), iP = c('Process Date'), iS = c('Sent To ACH Date'), iAm = c('Amount');
  if ([iBD, iL, iB, iA, iAm].some(function (i) { return i < 0; })) throw new Error('ACH column missing (headers: ' + h.filter(String).join(', ') + ')');
  var g = function (x, i) { return i < 0 ? '' : x[i]; }, out = [], seen = {}, now = new Date();
  for (var r2 = hr + 1; r2 < rows.length; r2++) {
    var x = rows[r2], bag = String(x[iB] || '').trim(), bd = ymd_(x[iBD]);
    if (!bag || !bd) continue;
    var amt = round2_(num_(x[iAm])), act = String(x[iA]).trim(), base = bag + '|' + bd + '|' + act + '|' + amt.toFixed(2);
    seen[base] = (seen[base] || 0) + 1;
    out.push([base + '#' + seen[base], bd, String(g(x, iCu)).trim(), String(x[iL]).replace(/\s+/g, ' ').trim(), bag, String(g(x, iBy)).trim(),
      act, ymd_(g(x, iP)), ymd_(g(x, iS)), amt, src, now]);
  }
  return out;
}

function upsertACH_(parsed) {
  var sh = sheet_(ACH_TAB, ACH_HEADERS), have = {};
  sh.getDataRange().getValues().slice(1).forEach(function (r) { have[r[0]] = 1; });
  var add = parsed.filter(function (r) { return !have[r[0]] && (have[r[0]] = 1); });
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, ACH_HEADERS.length).setValues(add);
  return add.length;
}

/** Pure: ACH_Raw + Brinks_Raw -> backup rows for smart-safe credits missing from Device Activity (+ stores to activate). */
function achBackupRows_(ach, br, map, today, tz) {
  var nl = function (s) { return String(s || '').replace(/\s+/g, ' ').trim().toLowerCase(); };
  var isDev = function (s) { return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-/i.test(String(s).trim()); };
  var T = {}, locOf = {};
  ach.forEach(function (r) {
    var amt = Number(r[9]) || 0, act = String(r[6]);
    if (!isDev(r[4]) || /variance/i.test(act) || !/credit/i.test(act) || amt <= 1) return;   // smart-safe credits only; $1 = install test
    var k = nl(r[3]) + '|' + ymd_(r[1]);
    T[k] = round2_((T[k] || 0) + amt); locOf[k] = String(r[3]).replace(/\s+/g, ' ').trim();
  });
  var S = {}, have = {};
  br.forEach(function (b) {
    var code = String(b[0]);
    if (/^ACHB-/.test(code)) { have[code] = 1; return; }
    if (!/^SD-/.test(code) || brSkip_(b)) return;
    var k = nl(b[4]) + '|' + ymd_(b[1]);
    S[k] = round2_((S[k] || 0) + (Number(b[7]) || 0));
  });
  var byLoc = {};
  Object.keys(map.byWZ).forEach(function (w) { var l = nl(map.byWZ[w].location); if (l) byLoc[l] = w; });
  var rows = [], activate = {};
  Object.keys(T).sort().forEach(function (k) {
    var loc = locOf[k], d = k.split('|')[1], gap = round2_(T[k] - (S[k] || 0));
    var m = loc.match(/\bWZ\d{3,4}[A-Z]?\b/i), wz = m ? m[0].toUpperCase() : (byLoc[nl(loc)] || '');
    var sid = m ? wz : (wz && map.byWZ[wz] && map.byWZ[wz].brinksId) || loc.replace(/^Verizon Wireless\s*/i, '');
    var code = 'ACHB-' + (wz || sid) + '-' + d, live = gap > 0.01;
    if (!live && !have[code]) return;
    rows.push([code, d, Utilities.parseDate(d + ' 12:00', tz, 'yyyy-MM-dd HH:mm'), sid, loc, '', live ? 'ACH' : 'Superseded', live ? gap : 0,
      '', '', '', '', '', '', '', today, 'ach-backup', '',
      'ACH backup (smart safe): ACH $' + T[k] + ' vs Device Activity $' + (S[k] || 0)]);
    if (live && wz) activate[wz] = loc;
  });
  return { rows: rows, activate: activate };
}

function achBackfill_() {
  var ach = readTab_(ACH_TAB); if (!ach.length) return;
  var tz = Session.getScriptTimeZone(), today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var res = achBackupRows_(ach, readTab_(TABS.BR), loadStoreMap_(), today, tz);
  if (res.rows.length) upsertBrinks_(res.rows);
  var live = res.rows.filter(function (r) { return r[6] === 'ACH'; });
  logRun('ACH backup: ' + live.length + ' store-days, $' + round2_(live.reduce(function (s, r) { return s + r[7]; }, 0)) + ' not on Device Activity');
  activateFromACH_(res.activate);
}

/** Pending Store Map rows with real ACH safe credits go live (Store ID = WZ#; Location # is filled in later by activateNewInstalls_). */
function activateFromACH_(activate) {
  var wzs = Object.keys(activate); if (!wzs.length) return;
  var sh = ss_().getSheetByName(TABS.MAP), data = sh.getDataRange().getValues(), h = data[0].map(String);
  var c = function (n) { return h.indexOf(n); }, iW = c('WZ Store ID'), iS = c('Brinks Store ID'), iL = c('Brinks Location Name'), iA = c('Active'), iN = c('Notes'), iName = c('Store Name');
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d/yyyy'), done = [];
  for (var i = 1; i < data.length; i++) {
    var w = String(data[i][iW]).trim().toUpperCase();
    if (!activate[w] || String(data[i][iA]).trim().toUpperCase() === 'Y') continue;
    var row = i + 1;
    if (iS >= 0 && !String(data[i][iS]).trim()) sh.getRange(row, iS + 1).setValue(w);
    if (iL >= 0 && !String(data[i][iL]).trim()) sh.getRange(row, iL + 1).setValue(activate[w]);
    sh.getRange(row, iA + 1).setValue('Y');
    if (iN >= 0) sh.getRange(row, iN + 1).setValue('Auto-activated ' + today + ' from ACH report (not on scheduled Brinks report yet; Location # pending)');
    done.push(w + ' ' + data[i][iName]);
  }
  if (!done.length) return;
  logRun('Stores activated from ACH report: ' + done.join(', '));
  try {
    MailApp.sendEmail(adminEmail_(), 'Cash Recon: ' + done.length + ' store(s) activated from the ACH report',
      'These stores have smart-safe credits on the Brinks ACH report but are not on the scheduled Brinks report yet. They are now in the daily recon using ACH data:\n\n' +
      done.join('\n') + '\n\nAsk Brinks to add them to the scheduled Device Activity report.');
  } catch (e) { logRun('ACH activation email failed: ' + e.message); }
}

/** Manual: load new RQ / Brinks / ACH files, then re-run recon + netting for the last 14 days (continues automatically if it runs long). */
function achCatchUp() { setupTabs(); brinksCatchUp(14); }

function ymd_(v) {
  if (v === '' || v == null) return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var s = String(v).trim(), m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2);
  return /^\d{4}-\d{2}-\d{2}/.test(s) ? s.slice(0, 10) : '';
}

// ───────────────────────── EMAIL: Outlook relay (Zapier) with Gmail fallback ─────────────────────────
// Recon emails go out from the work Outlook account through a private Zapier workflow ("Cash Recon Email Relay (Outlook)").
// Script Properties: MAIL_RELAY_URL (the workflow's catch-hook URL) + MAIL_RELAY_SECRET. Config MAIL_VIA = outlook | gmail.
// If the relay isn't set up, is switched off, or the POST fails, the email goes out through Gmail (MailApp) as before.
function mailRelay_() {
  var p = PropertiesService.getScriptProperties(), url = p.getProperty('MAIL_RELAY_URL'), secret = p.getProperty('MAIL_RELAY_SECRET');
  if (!url || !secret || /^gmail$/i.test(String(getCfg_().MAIL_VIA || '').trim())) return null;
  return { url: url, secret: secret };
}
function sendMail_(o) {
  var relay = mailRelay_();
  if (relay) {
    try {
      var res = UrlFetchApp.fetch(relay.url, { method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        payload: JSON.stringify({ secret: relay.secret, to: o.to, cc: o.cc || '', bcc: o.bcc || '', reply_to: o.replyTo || '',
          subject: o.subject, html: o.htmlBody || null, text: o.htmlBody ? null : (o.body || '') }) });
      var code = res.getResponseCode();
      if (code >= 200 && code < 300) return 'outlook';
      logRun('Outlook relay HTTP ' + code + ' — sending via Gmail instead: ' + String(res.getContentText()).slice(0, 300));
    } catch (e) { logRun('Outlook relay error — sending via Gmail instead: ' + e.message); }
  }
  MailApp.sendEmail(o);
  return 'gmail';
}
/** One-time setup from the editor: paste the relay URL + secret, run once, then delete the values from this function. */
function setMailRelay(url, secret) {
  if (!url || !secret) throw new Error('Pass the Zapier catch-hook URL and the relay secret');
  PropertiesService.getScriptProperties().setProperties({ MAIL_RELAY_URL: url, MAIL_RELAY_SECRET: secret });
  logRun('Outlook mail relay configured');
}
/** Sends a test through the relay to the admin. */
function testMailRelay() {
  var via = sendMail_({ to: adminEmail_(), subject: 'Cash Recon — test email via ' + (mailRelay_() ? 'Outlook relay' : 'Gmail'),
    htmlBody: '<p>Test from the Cash Recon engine at ' + new Date() + '.</p>' });
  logRun('testMailRelay sent via ' + via);
}
