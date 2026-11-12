/* =====================================================================
   Unbilled orders for the Sales Revenue page.

   HOW TO ADD THIS TO THE SALES REVENUE APPS SCRIPT
   1. Open the Sales Revenue Apps Script. Click + next to "Files",
      choose Script, name it "unbilled", and paste this whole file in.
   2. In Code.gs, find where the 'all' action builds the reply that is
      sent back to the page (the object with ok: true, rows, parts,
      customers ...). Add this one line to it, just before it is returned:

          reply.unbilled = readUnbilled_();

      (use whatever name Code.gs gives that object instead of "reply").
   3. Deploy > Manage deployments > pencil icon > Version: New version > Deploy.
      Keep the same deployment so the web address does not change.

   The page shows the Unbilled orders section as soon as this data arrives.
   ===================================================================== */

var UNBILLED_SHEET_ID = '1YUZNRsEFKvILXj9nld7kuPwvMu7DbB0BMBKq61ErTtE';   // the Sales Revenue spreadsheet
var UNBILLED_ORDERS_TAB = 'OS_Orders';
var UNBILLED_DETAIL_TAB = 'OS_Orders Detail';

/* Returns { rows: [...], lines: [...] }
   rows:  [order no, order date, customer name, customer no, reference, currency, rate, subtotal CAD, deposits CAD, on hold]
   lines: [order no, part no, description, quantity, unit price CAD, discount %, line total CAD] */
function readUnbilled_() {
  var ss = SpreadsheetApp.openById(UNBILLED_SHEET_ID);
  var tz = ss.getSpreadsheetTimeZone();
  var orders = readTab_(ss, UNBILLED_ORDERS_TAB);
  var detail = readTab_(ss, UNBILLED_DETAIL_TAB);

  var dateText = function (v) {
    if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    return String(v || '').slice(0, 10);
  };
  var text = function (v) { v = String(v == null ? '' : v).trim(); return v === 'NULL' ? '' : v; };
  var number = function (v) {
    if (typeof v === 'number') return v;
    var n = Number(String(v || '').replace(/[$,\s]/g, ''));
    return isFinite(n) ? n : 0;
  };

  /* Order-level facts from the detail tab (deposits and hold are only there). */
  var extra = {};
  detail.forEach(function (r) {
    var o = text(r.order_no);
    if (!o || extra[o]) return;
    extra[o] = {
      date: r.order_date, name: r.cust_name, no: r.cust_no, ref: r.ref_no,
      cur: r.currency, rate: r.currency_rate, sub: r.subtotal_cdn,
      dep: r.deposits, hold: r.hold
    };
  });

  /* One row per order. OS_Orders is used when it exists; otherwise orders
     are built from the detail tab. The totals row at the bottom is skipped. */
  var source = orders.length ? orders : Object.keys(extra).map(function (o) {
    var e = extra[o];
    return { order_no: o, order_date: e.date, cust_name: e.name, cust_no: e.no, ref_no: e.ref,
             currency: e.cur, currency_rate: e.rate, subtotal_cdn: e.sub };
  });
  var rows = [];
  source.forEach(function (r) {
    var o = text(r.order_no);
    if (!o) return;
    var e = extra[o] || {};
    rows.push([
      o, dateText(r.order_date), text(r.cust_name), text(r.cust_no), text(r.ref_no),
      text(r.currency) || 'CAD', number(r.currency_rate) || 1, number(r.subtotal_cdn),
      e.dep === undefined || e.dep === '' ? '' : number(e.dep),
      e.hold === true || String(e.hold).toUpperCase() === 'TRUE'
    ]);
  });

  var lines = [];
  detail.forEach(function (r) {
    var o = text(r.order_no);
    if (!o) return;
    lines.push([o, text(r.part_no), text(r.description), number(r.order_qty),
                number(r.unit_cdn), number(r.line_disc), number(r.sales_price_cdn)]);
  });

  return { rows: rows, lines: lines };
}

/* Reads a tab into objects keyed by its header row. Missing tab = no rows. */
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
