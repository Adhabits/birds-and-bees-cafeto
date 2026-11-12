/**
 * Birds & Bees Cafeto — table orders
 * Logs every order from the website into the "Orders" sheet,
 * then forwards the guest to WhatsApp with the full order.
 *
 * Deploy: Deploy > New deployment > Web app
 *   Execute as: Me   |   Who has access: Anyone
 */
const SHEET_NAME = 'Orders';
const WHATSAPP = '919826033910';           // cafe WhatsApp (with country code)
const HEADERS = ['Order ID', 'Time', 'Table', 'Name', 'Phone', 'Items', 'Qty', 'Total (₹)', 'Note', 'Status'];

// Run this once from the editor (Run ▶) to grant permission and create the Orders tab
function setup() { sheet_(); }

function doGet(e)  { return handle_(e && e.parameter ? e.parameter : {}, false); }
function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) { p = (e && e.parameter) || {}; }
  return handle_(p, true);
}

function handle_(p, asJson) {
  // Health check: open the /exec URL with no order to confirm it's live
  if (!p.id) {
    return asJson ? json_({ ok: true, live: true })
                  : HtmlService.createHtmlOutput('<p style="font:16px system-ui;padding:24px">Birds &amp; Bees order endpoint is live ✅</p>');
  }

  const order = {
    id:    clean_(p.id, 40),
    table: clean_(p.table, 4),
    name:  clean_(p.name, 60),
    phone: clean_(p.phone, 16),
    items: clean_(p.items, 1800),
    qty:   Number(p.qty) || 0,
    total: Number(p.total) || 0,
    note:  clean_(p.note, 300),
  };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = sheet_();
    // Ignore repeats of the same order (page refresh, double tap)
    const last = sh.getLastRow();
    const from = Math.max(2, last - 200);
    const ids = last >= 2 ? sh.getRange(from, 1, last - from + 1, 1).getValues().flat() : [];
    if (ids.indexOf(order.id) === -1) {
      sh.appendRow([order.id, new Date(), order.table, order.name, "'" + order.phone,
                    order.items, order.qty, order.total, order.note, 'New']);
    }
  } finally {
    lock.releaseLock();
  }

  const msg =
    '🐝 *New table order — Birds & Bees Cafeto*\n' +
    'Order ID: ' + order.id + '\n' +
    'Table: ' + order.table + '\n' +
    'Name: ' + order.name + '\n' +
    'Phone: ' + order.phone + '\n\n' +
    order.items + '\n\n' +
    '*Total: ₹' + order.total + '*' +
    (order.note ? '\nNote: ' + order.note : '');
  const wa = 'https://wa.me/' + WHATSAPP + '?text=' + encodeURIComponent(msg);

  if (asJson) return json_({ ok: true, id: order.id, whatsapp: wa });
  return confirmPage_(order, wa);
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#0E1D15').setFontColor('#F3B53F');
    sh.setFrozenRows(1);
    sh.setColumnWidth(6, 380);
    sh.getRange('B:B').setNumberFormat('dd-mmm-yyyy hh:mm am/pm');
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(['New', 'Preparing', 'Served', 'Paid', 'Cancelled'], true).build();
    sh.getRange('J2:J').setDataValidation(rule);
  }
  return sh;
}

// Trim, cap length, and stop spreadsheet formula injection
function clean_(v, max) {
  let s = String(v == null ? '' : v).replace(/\r/g, '').slice(0, max).trim();
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function esc_(s) {
  return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

function confirmPage_(o, wa) {
  const html = `
  <style>
    body{margin:0;min-height:100vh;display:grid;place-items:center;background:#0E1D15;color:#F1F3EA;font:16px/1.5 system-ui,-apple-system,Segoe UI,sans-serif;padding:20px;box-sizing:border-box}
    .c{max-width:420px;width:100%;text-align:center}
    .tick{width:72px;height:72px;border-radius:0 46% 0 46%;background:#F3B53F;color:#0E1D15;display:grid;place-items:center;font-size:38px;margin:0 auto 16px}
    h1{font-size:26px;margin:0 0 6px} p{margin:0;color:#A9BCAD}
    .id{display:inline-block;margin:14px 0 22px;font-family:ui-monospace,Menlo,monospace;border:1px dashed #F3B53F;color:#F3B53F;padding:6px 12px;border-radius:999px}
    a.b{display:block;background:#25D366;color:#fff;text-decoration:none;font-weight:700;padding:16px;border-radius:999px;font-size:18px}
    small{display:block;margin-top:14px;color:#A9BCAD}
  </style>
  <div class="c">
    <div class="tick">✓</div>
    <h1>Order received, ${esc_(o.name)}!</h1>
    <p>Table ${esc_(o.table)} · ₹${esc_(o.total)}</p>
    <div class="id">${esc_(o.id)}</div>
    <a class="b" href="${esc_(wa)}" target="_top">Send order on WhatsApp</a>
    <small>Tap the button to send your order to the cafe. Opening WhatsApp…</small>
  </div>
  <script>setTimeout(function(){ try{ window.top.location.href = ${JSON.stringify(wa)}; }catch(e){} }, 900);</script>`;
  return HtmlService.createHtmlOutput(html)
    .setTitle('Order received · Birds & Bees Cafeto')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
