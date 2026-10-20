/**
 * Inventory check-out backend.
 * Paste into a Google Apps Script project and deploy as a Web App
 * (Execute as: Me, Who has access: Anyone). See README.md.
 */

const SPREADSHEET_ID = '18dOTIriWlQVI4ZQFUH_BV6XI2x2KFmoa9ijRz4nNkV4';
const LOG_SHEET = 'Checkout Log';
const LOG_HEADERS = ['ID', 'Checked Out', 'Office', 'Item', 'Box / Zone', 'Qty', 'Name', 'NetID', 'Event', 'Status', 'Checked In'];

// Office tab in the app -> sheet tab and first data row.
const OFFICES = {
  'OFSA Office': { sheet: 'OFSA Basement Inventory', firstRow: 3 },
  'DM Office': { sheet: 'DM Office', firstRow: 3 },
};

function doGet(e) {
  return handle_(() => getData_(e.parameter.office));
}

function doPost(e) {
  return handle_(() => {
    const body = JSON.parse(e.postData.contents);
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      if (body.action === 'checkout') return checkout_(body);
      if (body.action === 'checkin') return checkin_(body);
      throw new Error('Unknown action: ' + body.action);
    } finally {
      lock.releaseLock();
    }
  });
}

function handle_(fn) {
  let out;
  try {
    out = { ok: true, data: fn() };
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function ss_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function logSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(LOG_SHEET);
    sh.appendRow(LOG_HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, LOG_HEADERS.length).setFontWeight('bold');
  }
  return sh;
}

function logRows_() {
  const sh = logSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, LOG_HEADERS.length).getValues().map((r, i) => ({
    row: i + 2,
    id: String(r[0]),
    out: r[1] instanceof Date ? r[1].toISOString() : String(r[1]),
    office: r[2], item: r[3], box: r[4], qty: r[5],
    name: r[6], netid: r[7], event: r[8], status: r[9],
    in: r[10] instanceof Date ? r[10].toISOString() : String(r[10] || ''),
  }));
}

function getData_(office) {
  const cfg = OFFICES[office];
  if (!cfg) throw new Error('Unknown office: ' + office);
  const sh = ss_().getSheetByName(cfg.sheet);
  const last = sh.getLastRow();
  const values = last >= cfg.firstRow ? sh.getRange(cfg.firstRow, 1, last - cfg.firstRow + 1, 3).getDisplayValues() : [];

  const items = [];
  values.forEach((r, i) => {
    parseRow_(r[0], r[1], r[2]).forEach(it => items.push(Object.assign(it, { row: cfg.firstRow + i })));
  });

  const log = logRows_();
  const officeLog = log.filter(l => l.office === office);
  // People and events from every office, so autofill works across tabs.
  const people = {};
  log.forEach(l => { if (l.netid) people[String(l.netid).toLowerCase()] = l.name; });
  const events = Array.from(new Set(log.map(l => l.event).filter(String)));

  return {
    items,
    out: officeLog.filter(l => l.status === 'OUT').reverse(),
    recent: officeLog.slice(-25).reverse(),
    people,
    events,
  };
}

// Turns one messy inventory row into one or more {name, box, qty} items.
function parseRow_(box, desc, qty) {
  box = String(box || '').trim();
  desc = String(desc || '').trim();
  qty = String(qty || '').trim();
  if (!desc && !qty) return [];

  // "Yellow Wristband/Orange Wristband" with "80/105" -> two items.
  if (desc) {
    const names = desc.split('/').map(s => s.trim()).filter(String);
    const qtys = qty.split('/').map(s => s.trim());
    if (names.length > 1 && names.length === qtys.length) {
      return names.map((n, i) => ({ name: n, box, qty: qtys[i] }));
    }
    return [{ name: desc, box, qty }];
  }

  // No description: "10 unused FTK ribbons, 3 rolls of raffle tickets" -> one item per part.
  return qty.split(/[,;]/).map(s => s.trim()).filter(String).map(part => {
    const m = part.match(/^(~?\d[\d+]*)\s+(.*)$/);
    return m ? { name: m[2], box, qty: m[1] } : { name: part, box, qty: '' };
  });
}

function checkout_(b) {
  ['office', 'item', 'name', 'netid'].forEach(k => {
    if (!String(b[k] || '').trim()) throw new Error('Missing ' + k);
  });
  if (!OFFICES[b.office]) throw new Error('Unknown office: ' + b.office);
  const id = Utilities.getUuid().slice(0, 8);
  logSheet_().appendRow([
    id, new Date(), b.office, b.item, b.box || '', Number(b.qty) || 1,
    b.name.trim(), b.netid.trim().toLowerCase(), (b.event || '').trim(), 'OUT', '',
  ]);
  return { id };
}

function checkin_(b) {
  const entry = logRows_().find(l => l.id === String(b.id));
  if (!entry) throw new Error('Checkout not found');
  if (entry.status !== 'OUT') throw new Error('Already checked in');
  const sh = logSheet_();
  sh.getRange(entry.row, 10, 1, 2).setValues([['IN', new Date()]]);
  return { id: entry.id };
}
