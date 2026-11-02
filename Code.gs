// SAADO POS - Part 1 backend. Paste into a script bound to the NEW Google Sheet.
// 1) Change the PINs below. 2) Run setup() once. 3) Deploy as Web app (Execute as: Me, Access: Anyone).
const USERS = [ // Name, Role, Business (shop / fast / both), PIN (4-6 digits)
  ['Zakariye', 'owner', 'both', '1111'],
  ['Hussein', 'owner', 'both', '2222'],
  ['Raxma', 'employee', 'shop', '3333'],
  ['Sacad', 'employee', 'fast', '4444']];
const SETTINGS = [['rate', 11850], ['rounding', 500], ['shop_zaad', '436908'], ['shop_edahab', '93051'],
  ['fast_zaad', '473389'], ['fast_edahab', '98379'], ['fast_close', '']];
const HEAD = {
  Settings: ['Key', 'Value'],
  Users: ['Name', 'Role', 'Business', 'PinHash', 'Active'],
  Products: ['ID', 'Business', 'Name', 'Barcode', 'Price', 'Cost', 'Stock', 'LowStock', 'Expiry', 'Active'],
  Sales: ['SaleID', 'Date', 'Business', 'Cashier', 'Subtotal', 'Discount', 'Total', 'Payment', 'Tendered', 'Change'],
  SaleItems: ['SaleID', 'ProductID', 'Name', 'Qty', 'Price']};
const SAMPLE = [['P001', 'shop', 'Ahlan', '634240257094', 4000, 2500, 0, 5, '', true],
  ['P002', 'shop', 'Benson & Hedges LTD (Unit)', '6008165015093', 750, 500, 0, 20, '', true],
  ['P003', 'fast', 'Chicken', '112233445566', 15000, 9000, 0, 0, '', true],
  ['P004', 'fast', 'Chips', '662233445566', 5000, 3000, 0, 0, '', true]];

const sh = n => SpreadsheetApp.getActive().getSheetByName(n);
const hash = p => Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, p + ':saado'));
const rows = n => { const d = sh(n).getDataRange().getValues(), h = d.shift(); return d.map(r => Object.fromEntries(h.map((k, i) => [k, r[i]]))); };
const settings = () => Object.fromEntries(rows('Settings').map(r => [r.Key, r.Value]));

function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(HEAD).forEach(n => {
    const s = ss.getSheetByName(n) || ss.insertSheet(n);
    if (s.getLastRow() === 0) s.getRange(1, 1, 1, HEAD[n].length).setValues([HEAD[n]]).setFontWeight('bold').setBackground('#003366').setFontColor('#FFBF00');
    s.setFrozenRows(1);
  });
  sh('Products').getRange('D:D').setNumberFormat('@'); // keep barcodes as text
  if (sh('Settings').getLastRow() < 2) sh('Settings').getRange(2, 1, SETTINGS.length, 2).setValues(SETTINGS);
  if (sh('Users').getLastRow() < 2) sh('Users').getRange(2, 1, USERS.length, 5).setValues(USERS.map(u => [u[0], u[1], u[2], hash(u[3]), true]));
  if (sh('Products').getLastRow() < 2) sh('Products').getRange(2, 1, SAMPLE.length, 10).setValues(SAMPLE);
  const d = ss.getSheetByName('Sheet1'); if (d && ss.getSheets().length > 1) ss.deleteSheet(d);
}

function doGet() { return ContentService.createTextOutput('Saado POS API is running'); }
function doPost(e) {
  let r;
  try { r = route(JSON.parse(e.postData.contents)); } catch (x) { r = { ok: false, error: String(x.message || x) }; }
  return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
}

function route(b) {
  if (b.action === 'users') return { ok: true, users: rows('Users').filter(u => u.Active !== false).map(u => u.Name) };
  if (b.action === 'login') {
    const u = rows('Users').find(x => x.Name === b.name && x.Active !== false);
    if (!u || u.PinHash !== hash(String(b.pin))) return { ok: false, error: 'PIN' };
    const user = { name: u.Name, role: u.Role, business: u.Business }, token = Utilities.getUuid();
    CacheService.getScriptCache().put('t_' + token, JSON.stringify(user), 21600); // 6 hours
    return { ok: true, token, user };
  }
  const c = CacheService.getScriptCache().get('t_' + b.token);
  if (!c) return { ok: false, error: 'AUTH' };
  const u = JSON.parse(c), owner = u.role === 'owner';
  const mine = biz => u.business === 'both' || u.business === biz;
  if (b.action === 'data') {
    const ps = rows('Products').filter(p => p.Active !== false && mine(p.Business)).map(p => {
      if (p.Expiry instanceof Date) p.Expiry = Utilities.formatDate(p.Expiry, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (!owner) delete p.Cost;
      return p;
    });
    return { ok: true, settings: settings(), products: ps };
  }
  if (b.action === 'sale') {
    if (!mine(b.sale.business)) return { ok: false, error: 'Not allowed for this business' };
    return recordSale(b.sale, u);
  }
  if (!owner) return { ok: false, error: 'Owner only' };
  if (b.action === 'saveProduct') return saveProduct(b.product);
  if (b.action === 'saveSettings') {
    const S = sh('Settings'), d = S.getDataRange().getValues();
    Object.keys(b.values).forEach(k => {
      const i = d.findIndex(r => r[0] === k);
      if (i > 0) S.getRange(i + 1, 2).setValue(b.values[k]); else S.appendRow([k, b.values[k]]);
    });
    return { ok: true };
  }
  return { ok: false, error: 'Unknown action' };
}

function recordSale(s, u) {
  const L = LockService.getScriptLock(); L.waitLock(20000);
  try {
    const S = sh('Sales');
    if (S.createTextFinder(s.id).matchEntireCell(true).findNext()) return { ok: true, dup: true }; // already saved
    const P = sh('Products'), d = P.getDataRange().getValues(), step = Number(settings().rounding) || 1;
    let sub = 0; const lines = [];
    s.items.forEach(it => {
      const i = d.findIndex(r => r[0] === it.id);
      if (i < 1) throw new Error('Unknown product ' + it.id);
      sub += Number(d[i][4]) * it.qty;
      lines.push([s.id, d[i][0], d[i][2], it.qty, d[i][4]]);
      if (d[i][1] === 'shop') { d[i][6] = (Number(d[i][6]) || 0) - it.qty; P.getRange(i + 1, 7).setValue(d[i][6]); }
    });
    const disc = Math.min(Number(s.discount) || 0, sub), total = Math.round((sub - disc) / step) * step;
    S.appendRow([s.id, new Date(s.time || Date.now()), s.business, u.name, sub, disc, total, s.payment, s.tendered || '', s.change || '']);
    sh('SaleItems').getRange(sh('SaleItems').getLastRow() + 1, 1, lines.length, 5).setValues(lines);
    return { ok: true };
  } finally { L.releaseLock(); }
}

function saveProduct(p) {
  const P = sh('Products'), d = P.getDataRange().getValues(), i = p.ID ? d.findIndex(r => r[0] === p.ID) : -1;
  const row = [p.ID || 'P' + Date.now().toString(36), p.Business, p.Name, String(p.Barcode || ''), Number(p.Price), Number(p.Cost) || 0,
    Number(p.Stock) || 0, Number(p.LowStock) || 0, p.Expiry || '', true];
  if (i > 0) P.getRange(i + 1, 1, 1, 10).setValues([row]); else P.appendRow(row);
  return { ok: true };
}
