/* =====================================================================
   Sales Revenue page: Apps Script backend (complete Code.gs)

   Serves the Sales Revenue page on the Armodilo internal tools site:
   sign-in, invoice detail, and unbilled (open) orders.

   SET-UP (once)
   1. Project Settings (gear icon) > Script properties > Add script property
        Property: SR_USERS
        Value:    username:password
      For several people, separate them with commas:  vishnu:pass1,james:pass2
      Use the same username and password as the other tools so the single
      sign-in on the home page opens everything.
   2. Deploy > Manage deployments > pencil icon > Version: New version > Deploy.
      Keep the same deployment so the page's web address does not change.
      "Execute as" = Me, "Who has access" = Anyone.
   ===================================================================== */

var SHEET_ID = '1YUZNRsEFKvILXj9nld7kuPwvMu7DbB0BMBKq61ErTtE';   // the Sales Revenue spreadsheet
var DETAIL_TAB = 'Sales_Revenue by Invoice Detail';
var UNBILLED_ORDERS_TAB = 'OS_Orders';
var UNBILLED_DETAIL_TAB = 'OS_Orders Detail';

var SESSION_HOURS = 6;        // how long a sign-in lasts (6 is the most Apps Script allows)
var MAX_ATTEMPTS = 10;        // wrong passwords before a username is locked
var LOCK_MINUTES = 15;

/* ---------------- entry points ---------------- */

function doPost(e) {
  var body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json_({ ok: false, error: 'bad_request' }); }

  try {
    if (body.action === 'login') return json_(login_(body.user, body.pass));
    if (body.action === 'logout') { if (body.token) cache_().remove('tok_' + body.token); return json_({ ok: true }); }
    if (!validToken_(body.token)) return json_({ ok: false, error: 'auth' });
    if (body.action === 'all') return json_(readAll_());
    return json_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return json_({ ok: false, error: 'server', message: String(err && err.message || err) });
  }
}

function doGet() {
  return json_({ ok: true, service: 'Armodilo Sales Revenue' });
}

/* ---------------- sign-in ---------------- */

function login_(user, pass) {
  user = String(user || '').trim().toLowerCase();
  pass = String(pass || '');
  var c = cache_(), lockKey = 'lock_' + user, failKey = 'fail_' + user;
  if (c.get(lockKey)) return { ok: false, error: 'locked' };

  if (user && checkPassword_(user, pass)) {
    c.remove(failKey);
    var token = Utilities.getUuid();
    c.put('tok_' + token, user, SESSION_HOURS * 3600);
    return { ok: true, token: token };
  }

  var fails = Number(c.get(failKey) || 0) + 1;
  if (fails >= MAX_ATTEMPTS) {
    c.put(lockKey, '1', LOCK_MINUTES * 60);
    c.remove(failKey);
    return { ok: false, error: 'locked' };
  }
  c.put(failKey, String(fails), LOCK_MINUTES * 60);
  return { ok: false, error: 'bad_login', left: MAX_ATTEMPTS - fails };
}

function checkPassword_(user, pass) {
  var list = String(PropertiesService.getScriptProperties().getProperty('SR_USERS') || '');
  var ok = false;
  list.split(',').forEach(function (pair) {
    var i = pair.indexOf(':');
    if (i < 1) return;
    var u = pair.slice(0, i).trim().toLowerCase(), p = pair.slice(i + 1).trim();
    if (u === user && sameText_(p, pass)) ok = true;
  });
  return ok;
}

function sameText_(a, b) {          // compares without stopping at the first difference
  var diff = a.length ^ b.length;
  for (var i = 0; i < Math.max(a.length, b.length); i++) diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return diff === 0;
}

function validToken_(token) {
  return !!(token && cache_().get('tok_' + token));
}

/* ---------------- data ---------------- */

/* Reply for the 'all' action:
   parts:     [[part no, description], ...]
   customers: [[customer name, customer no], ...]
   rows:      [invoice no, 'yyyy-mm-dd', part index, customer index, quantity,
               unit price CAD, discount %, line total CAD, currency, rate]
   unbilled:  see readUnbilled_() */
function readAll_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var tz = ss.getSpreadsheetTimeZone();
  var detail = readTab_(ss, DETAIL_TAB);
  if (!detail.length) throw new Error('The "' + DETAIL_TAB + '" tab is missing or empty.');

  var parts = [], partIx = {}, customers = [], custIx = {}, rows = [];
  detail.forEach(function (r) {
    var inv = text_(r.invoice_no), date = dateText_(r.invoice_date, tz);
    var pn = partNo_(r.part_no);
    if (!inv || !date || !pn) return;

    if (!(pn in partIx)) { partIx[pn] = parts.length; parts.push([pn, text_(r.description)]); }
    var cname = text_(r.cust_name) || text_(r.cust_no), cno = text_(r.cust_no);
    var ck = cname.toLowerCase() + '|' + cno.toLowerCase();
    if (!(ck in custIx)) { custIx[ck] = customers.length; customers.push([cname, cno]); }

    rows.push([inv, date, partIx[pn], custIx[ck], number_(r.committed_qty), number_(r.unit_cdn),
               number_(r.line_disc), number_(r.sales_price_cdn), text_(r.currency) || 'CAD',
               number_(r.currency_rate) || 1]);
  });

  return {
    ok: true,
    generated: Utilities.formatDate(new Date(), tz, "yyyy-MM-dd'T'HH:mm"),
    labels: { order: 'Invoice Number', date: 'Invoice Date', part: 'Part Number',
              customer: 'Customer', custNo: 'Customer Number' },
    parts: parts, customers: customers, rows: rows,
    unbilled: readUnbilled_(ss, tz)
  };
}

/* Unbilled (open) orders.
   rows:  [order no, order date, customer name, customer no, reference, currency, rate,
           subtotal CAD, deposits CAD, on hold]
   lines: [order no, part no, description, quantity, unit price CAD, discount %, line total CAD] */
function readUnbilled_(ss, tz) {
  var orders = readTab_(ss, UNBILLED_ORDERS_TAB);
  var detail = readTab_(ss, UNBILLED_DETAIL_TAB);

  var extra = {};       // deposits and hold are only on the detail tab
  detail.forEach(function (r) {
    var o = text_(r.order_no);
    if (!o || extra[o]) return;
    extra[o] = { order_no: o, order_date: r.order_date, cust_name: r.cust_name, cust_no: r.cust_no,
                 ref_no: r.ref_no, currency: r.currency, currency_rate: r.currency_rate,
                 subtotal_cdn: r.subtotal_cdn, deposits: r.deposits, hold: r.hold };
  });

  var source = orders.length ? orders : Object.keys(extra).map(function (o) { return extra[o]; });
  var rows = [];
  source.forEach(function (r) {
    var o = text_(r.order_no);
    if (!o) return;                       // skips the totals row at the bottom of OS_Orders
    var e = extra[o] || {};
    rows.push([o, dateText_(r.order_date, tz), text_(r.cust_name), text_(r.cust_no), text_(r.ref_no),
               text_(r.currency) || 'CAD', number_(r.currency_rate) || 1, number_(r.subtotal_cdn),
               e.deposits === undefined || e.deposits === '' ? '' : number_(e.deposits),
               e.hold === true || String(e.hold).toUpperCase() === 'TRUE']);
  });

  var lines = [];
  detail.forEach(function (r) {
    var o = text_(r.order_no);
    if (!o) return;
    lines.push([o, partNo_(r.part_no), text_(r.description), number_(r.order_qty),
                number_(r.unit_cdn), number_(r.line_disc), number_(r.sales_price_cdn)]);
  });

  return { rows: rows, lines: lines };
}

/* ---------------- helpers ---------------- */

function readTab_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  var values = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  var head = values[0].map(function (h) { return String(h).trim().toLowerCase(); });
  return values.slice(1).map(function (row) {
    var o = {};
    head.forEach(function (h, i) { if (h) o[h] = row[i]; });
    return o;
  });
}

function text_(v) {
  v = String(v == null ? '' : v).trim();
  return v === 'NULL' ? '' : v;
}

function partNo_(v) {                     // "- FLOOR-O-R-B" -> "FLOOR-O-R-B"
  return text_(v).replace(/^-\s*/, '');
}

function number_(v) {
  if (typeof v === 'number') return v;
  var n = Number(String(v == null ? '' : v).replace(/[$,\s]/g, ''));
  return isFinite(n) ? n : 0;
}

function dateText_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  var s = String(v || '').trim();
  return /^\d{4}-\d{2}-\d{2}/.test(s) ? s.slice(0, 10) : '';
}

function cache_() { return CacheService.getScriptCache(); }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
