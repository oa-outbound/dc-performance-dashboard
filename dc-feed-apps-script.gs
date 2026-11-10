/**
 * DC Feed: pulls per-SDR, per-day DC outcomes from the Outbound Dashboard
 * into the "DC Feed" tab, which the DC Performance Dashboard reads.
 *
 * Setup (once):
 *   1. Project Settings > Script Properties: add DC_FEED_KEY (never put the key in code or the sheet).
 *   2. Run setupTriggers() once and approve the permissions.
 *   3. File > Share > Publish to web > "DC Feed" tab > CSV.
 *
 * How it writes: every pull REPLACES all rows whose date is inside the window the feed
 * returned, then writes the new rows. This is an upsert on (date, sdr) that also removes
 * rows that dropped out (e.g. an SDR's only DC that day was rescheduled to another date).
 */

var FEED_URL = 'https://oa-dashboard-web-vdd2e.ondigitalocean.app/api/dc-outcomes';
var TZ = 'Asia/Manila';
var DATA_TAB = 'DC Feed';
var STATUS_TAB = 'DC Feed Status';
var FIELDS = [
  'date', 'sdr', 'total_dcs',
  'both_attended', 'lead_attended', 'bpo_attended', 'both_attended_unqualified',
  'no_show', 'canceled_by_lead', 'canceled_by_partner', 'canceled_unattributed',
  'invalid', 'scheduled', 'rescheduled', 'no_outcome'
];

/**
 * Every 30 minutes: the last 7 days through the end of this month. Going to month end
 * brings in meetings already booked for later this month, so the dashboard's month total
 * matches the Outbound Dashboard (held + upcoming).
 */
function pullRecent() {
  var now = new Date();
  var from = fmt_(new Date(now.getTime() - 7 * 24 * 3600 * 1000));
  runPull_('recent', from, monthEnd_(now));
  pullActivitySafe_(monthStart_(now), fmt_(now));   // SDR Activity: this month to date
}

/** Last day of the current month, Manila, as YYYY-MM-DD. */
function monthEnd_(now) {
  var y = Number(Utilities.formatDate(now, TZ, 'yyyy'));
  var m = Number(Utilities.formatDate(now, TZ, 'M'));
  var last = new Date(Date.UTC(y, m, 0)).getUTCDate();
  return y + '-' + (m < 10 ? '0' + m : m) + '-' + (last < 10 ? '0' + last : last);
}

/**
 * Nightly: the current month. During the first 7 days of a month it also re-pulls the
 * previous month, so late outcomes on last month's final days still land before month-end
 * reporting. Range stays well under the 92-day limit.
 */
function pullMonth() {
  var now = new Date();
  var today = fmt_(now);
  var y = Number(Utilities.formatDate(now, TZ, 'yyyy'));
  var m = Number(Utilities.formatDate(now, TZ, 'M'));
  var d = Number(Utilities.formatDate(now, TZ, 'd'));
  if (d <= 7) { m -= 1; if (m === 0) { m = 12; y -= 1; } }
  var from = y + '-' + (m < 10 ? '0' + m : m) + '-01';
  runPull_('month', from, monthEnd_(now));
  // SDR Activity: settle last month's totals, then this month to date
  var lastEnd = new Date(Date.UTC(Number(Utilities.formatDate(now, TZ, 'yyyy')), Number(Utilities.formatDate(now, TZ, 'M')) - 1, 0));
  var lastEndIso = Utilities.formatDate(lastEnd, 'UTC', 'yyyy-MM-dd');
  pullActivitySafe_(lastEndIso.slice(0, 8) + '01', lastEndIso);
  pullActivitySafe_(monthStart_(now), fmt_(now));
}

/**
 * Run once by hand: loads history from BACKFILL_FROM to today in 90-day chunks (the feed
 * allows 92 days per call), so last-month comparisons, rep profiles and the YTD export have
 * data. After this, the two triggers keep everything current.
 */
var BACKFILL_FROM = '2025-01-01';
function backfill() {
  var start = new Date(BACKFILL_FROM + 'T00:00:00+08:00');
  var today = fmt_(new Date());
  while (fmt_(start) <= today) {
    var end = new Date(start.getTime() + 89 * 24 * 3600 * 1000);
    var to = fmt_(end) > today ? today : fmt_(end);
    runPull_('backfill', fmt_(start), to);
    start = new Date(end.getTime() + 24 * 3600 * 1000);
  }
}

/** Run once to create both triggers (removes any old copies first). */
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (fn === 'pullRecent' || fn === 'pullMonth') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('pullRecent').timeBased().everyMinutes(30).create();
  ScriptApp.newTrigger('pullMonth').timeBased().atHour(1).everyDays(1).inTimezone(TZ).create();
  pullMonth();   // fill the sheet now
}

// ---------------------------------------------------------------------------

function runPull_(mode, from, to) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(60 * 1000)) { console.warn('Another pull is running, skipped ' + mode); return; }
  try {
    var data = fetchFeed_(from, to);
    if (data.rows_not_reconciling > 0) {
      console.warn(data.rows_not_reconciling + ' row(s) did not reconcile (' + data.from + ' to ' + data.to + ')');
    }
    writeWindow_(data.from, data.to, data.rows);
    writeStatus_(mode, data);
  } finally {
    lock.releaseLock();
  }
}

function fetchFeed_(from, to) {
  var key = PropertiesService.getScriptProperties().getProperty('DC_FEED_KEY');
  if (!key) throw new Error('DC_FEED_KEY is missing from Script Properties');
  var url = FEED_URL;
  if (from && to) url += '?from=' + from + '&to=' + to;
  var res = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + key },
    muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  if (code !== 200) {
    throw new Error('DC feed returned ' + code + ': ' + res.getContentText().slice(0, 300));
  }
  var data = JSON.parse(res.getContentText());
  if (!data.from || !data.to || !Array.isArray(data.rows)) {
    throw new Error('DC feed response missing from/to/rows');
  }
  return data;
}

function writeWindow_(from, to, rows) {
  var sh = getTab_(DATA_TAB);
  var last = sh.getLastRow();
  var kept = [];
  if (last > 1) {
    var existing = sh.getRange(2, 1, last - 1, FIELDS.length).getValues();
    kept = existing.filter(function (r) {
      var d = String(r[0]);
      return d && (d < from || d > to);   // keep only rows outside the pulled window
    });
  }
  var fresh = rows.map(function (o) {
    return FIELDS.map(function (f) { return o[f] === undefined || o[f] === null ? (f === 'sdr' ? '' : 0) : o[f]; });
  });
  var all = kept.concat(fresh);
  all.sort(function (a, b) {
    if (a[0] !== b[0]) return a[0] < b[0] ? 1 : -1;   // newest date first
    return String(a[1]).localeCompare(String(b[1]));
  });

  if (last > 1) sh.getRange(2, 1, last - 1, FIELDS.length).clearContent();
  if (all.length) {
    var range = sh.getRange(2, 1, all.length, FIELDS.length);
    sh.getRange(2, 1, all.length, 1).setNumberFormat('@');   // keep dates as YYYY-MM-DD text
    range.setValues(all);
  }
}

function writeStatus_(mode, data) {
  var sh = getTab_(STATUS_TAB, ['field', 'value']);
  var rows = [
    ['last_sync_manila', Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm')],
    ['last_mode', mode],
    ['window_from', data.from],
    ['window_to', data.to],
    ['row_count', data.row_count],
    ['total_dcs', data.total_dcs],
    ['rows_not_reconciling', data.rows_not_reconciling]
  ];
  sh.getRange(2, 1, rows.length, 2).setNumberFormat('@').setValues(rows.map(function (r) { return [r[0], String(r[1])]; }));
}

function getTab_(name, header) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  var h = header || FIELDS;
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function fmt_(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }

// ---------------------------------------------------------------------------
// SDR Activity: per-SDR range totals from the Outbound Dashboard dials page (dials, leads
// dialed/handled, connected leads, leads booked, worked days). One row per SDR per month,
// keyed on the month's first day, replaced on every pull. Leads dialed and connected are
// distinct leads, so the endpoint returns range totals; never add daily rows together.
var ACT_URL = 'https://oa-dashboard-web-vdd2e.ondigitalocean.app/api/sdr-activity';
var ACT_TAB = 'SDR Activity';
var ACT_FIELDS = ['month', 'sdr', 'wave', 'dials', 'leads_dialed', 'leads_handled', 'connected_leads',
                  'leads_booked', 'dcs', 'worked_days', 'lead_connection_rate', 'as_of', 'stale'];

function monthStart_(now) { return Utilities.formatDate(now, TZ, 'yyyy-MM') + '-01'; }

// A failed activity pull must not stop the DC feed, and must not overwrite good rows with zeros:
// the endpoint answers 503 on failure, so the previous figures simply stay until the next pull.
function pullActivitySafe_(from, to) {
  try { pullActivity_(from, to); }
  catch (e) { console.warn('SDR Activity pull skipped (' + from + ' to ' + to + '): ' + e.message); }
}

function pullActivity_(from, to) {
  var key = PropertiesService.getScriptProperties().getProperty('DC_FEED_KEY');
  var res = UrlFetchApp.fetch(ACT_URL + '?from=' + from + '&to=' + to, {
    headers: { Authorization: 'Bearer ' + key }, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('activity feed returned ' + res.getResponseCode());
  var data = JSON.parse(res.getContentText());
  if (!Array.isArray(data.rows)) throw new Error('activity feed response missing rows');
  if (data.stale) console.warn('SDR Activity is stale by ' + data.stale_minutes + ' min; writing last-good figures');
  var month = from.slice(0, 8) + '01';
  var sh = getTab_(ACT_TAB, ACT_FIELDS);
  var last = sh.getLastRow();
  var kept = last > 1 ? sh.getRange(2, 1, last - 1, ACT_FIELDS.length).getValues()
                              .filter(function (r) { return String(r[0]) !== month; }) : [];
  var fresh = data.rows.map(function (o) {
    return [month, o.sdr, o.wave || '', o.dials || 0, o.leads_dialed || 0, o.leads_handled || 0,
            o.connected_leads || 0, o.leads_booked || 0, o.dcs || 0, o.worked_days || 0,
            o.lead_connection_rate == null ? '' : o.lead_connection_rate, to, data.stale ? 'yes' : ''];
  });
  var all = kept.concat(fresh).sort(function (a, b) {
    if (a[0] !== b[0]) return a[0] < b[0] ? 1 : -1;
    return String(a[1]).localeCompare(String(b[1]));
  });
  if (last > 1) sh.getRange(2, 1, last - 1, ACT_FIELDS.length).clearContent();
  if (all.length) {
    sh.getRange(2, 1, all.length, 1).setNumberFormat('@');
    sh.getRange(2, 12, all.length, 1).setNumberFormat('@');
    sh.getRange(2, 1, all.length, ACT_FIELDS.length).setValues(all);
  }
}

/** Run once by hand after pasting this version: fills the SDR Activity tab for last month and this month. */
function setupActivity() {
  var now = new Date();
  var lastEnd = new Date(Date.UTC(Number(Utilities.formatDate(now, TZ, 'yyyy')), Number(Utilities.formatDate(now, TZ, 'M')) - 1, 0));
  var lastEndIso = Utilities.formatDate(lastEnd, 'UTC', 'yyyy-MM-dd');
  pullActivity_(lastEndIso.slice(0, 8) + '01', lastEndIso);
  pullActivity_(monthStart_(now), fmt_(now));
}
