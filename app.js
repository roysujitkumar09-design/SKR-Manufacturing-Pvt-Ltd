'use strict';
/* SKR Manufacturing ERP - app.js
   Vanilla JS, data persisted in localStorage, Chart.js for dashboards.
   Demo login: admin@skrmfg.in / admin123 */

/* ---------- Helpers ---------- */
const KEY = 'skr_erp_v2', PG = 8;
const $ = (s, r = document) => r.querySelector(s), $$ = (s, r = document) => [...r.querySelectorAll(s)];
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const inr = n => '₹' + (+n || 0).toLocaleString('en-IN', { maximumFractionDigits: 0 });
const num = n => (+n || 0).toLocaleString('en-IN');
const sum = (a, k) => a.reduce((t, r) => t + (+r[k] || 0), 0);
const ymd = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
const ago = n => ymd(new Date(Date.now() - n * 864e5)), today = () => ago(0);
let DB, route = 'dashboard', charts = [], ST = {}, CONF = null, CU = {}, STREAM = null;

/* ---------- Roles & permissions ---------- */
const ALL = ['dashboard', 'materials', 'production', 'fg', 'costing', 'sales', 'customers', 'suppliers', 'purchases', 'scan', 'reports'];
const PERM = { Admin: ['*'], Manager: ALL, Operator: ['dashboard', 'materials', 'production', 'fg', 'scan'], Accountant: ['dashboard', 'costing', 'sales', 'customers', 'suppliers', 'purchases', 'reports'], Viewer: ALL };
const can = r => (PERM[CU.role] || []).some(x => x == '*' || x == r), RO = () => CU.role == 'Viewer';
const WR = new Set(['adj', 'pay', 'open', 'recv', 'costs', 'spay']); // row actions that change data

/* ---------- Shared helpers: money, words, QR, files, costing ---------- */
const inr2 = n => '₹' + (+n || 0).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const words = n => {
  const a = 'One Two Three Four Five Six Seven Eight Nine Ten Eleven Twelve Thirteen Fourteen Fifteen Sixteen Seventeen Eighteen Nineteen'.split(' '), b = ',,Twenty,Thirty,Forty,Fifty,Sixty,Seventy,Eighty,Ninety'.split(',');
  const w2 = x => x < 20 ? (a[x - 1] || '') : b[x / 10 | 0] + (x % 10 ? ' ' + a[x % 10 - 1] : '');
  const w3 = x => x > 99 ? a[(x / 100 | 0) - 1] + ' Hundred' + (x % 100 ? ' ' + w2(x % 100) : '') : w2(x);
  n = Math.round(n); let s = '';
  [[1e7, 'Crore'], [1e5, 'Lakh'], [1e3, 'Thousand']].forEach(([v, l]) => { if (n >= v) { s += w3(n / v | 0) + ' ' + l + ' '; n %= v; } });
  return (s + (n ? w3(n) : '')).trim() + ' Rupees Only';
};
const qr = (el, text, size = 110) => { el.innerHTML = ''; if (window.QRCode) new QRCode(el, { text, width: size, height: size }); else el.textContent = text; };
async function saveFile(name, data) { // uses the viewer download prompt when hosted, a normal download otherwise
  try { const d = window.claude && await claude.use('downloads'); if (d) { await d.save({ filename: name, data }); return toast('Saved ' + name); } } catch (e) { if (e && e.code == 'declined') return; }
  const a = document.createElement('a'); a.href = URL.createObjectURL(data instanceof Blob ? data : new Blob([data])); a.download = name; a.click(); toast('Downloaded ' + name);
}
const xl = (wb, name) => saveFile(name, XLSX.write(wb, { type: 'array', bookType: 'xlsx' }));
const noXL = () => !window.XLSX && (toast('The Excel library did not load. Check your connection.', 'warn'), true);
const mp = n => (DB.purchases.find(p => p.material == n && p.status == 'Received') || {}).price || 0; // latest material price
const unitCost = n => { const p = P(n); return p ? +(p.ratio * mp(p.mat) + p.labour + p.oh).toFixed(2) : 0; };
const recv = r => { const m = DB.materials.find(x => x.name == r.material), s = DB.suppliers.find(x => x.name == r.supplier); if (m) m.stock += r.qty; if (s) s.payable += r.total; };

/* ---------- Sample data ---------- */
function seed() {
  const P = [['Industrial Fasteners', 3000, 12, 'Steel Wire', .05, 1.5, .8, '7318'], ['Packaging Boxes', 2000, 35, 'Corrugated Board', .4, 4, 1.5, '4819'], ['Plastic Components', 2500, 48, 'PVC Granules', .08, 6, 2.5, '3926'], ['Electrical Parts', 1500, 120, 'Copper Wire', .02, 18, 6, '8536']];
  const S = ['ABC Metals', 'Bengal Packaging', 'Eastern Chemicals'], C = ['ABC Traders', 'XYZ Enterprises', 'Kolkata Industrial Supply'], O = ['Rakesh Das', 'Anita Roy', 'Suman Ghosh'];
  const id = (p, n) => p + String(n).padStart(3, '0');
  const mats = [['Steel Wire', 'kg', 1200, 500, 0], ['Aluminium Sheet', 'sheet', 300, 400, 0], ['Corrugated Board', 'sheet', 2000, 800, 1], ['Kraft Paper', 'kg', 150, 300, 1], ['PVC Granules', 'kg', 900, 400, 2], ['ABS Resin', 'kg', 600, 500, 2], ['Copper Wire', 'kg', 80, 100, 0]];
  const t = n => new Date(Date.now() - n * 36e5).toISOString();
  return {
    products: P.map(p => ({ name: p[0], opening: p[1], rate: p[2], mat: p[3], ratio: p[4], labour: p[5], oh: p[6], hsn: p[7] })),
    suppliers: S.map((n, i) => ({ id: id('SP-', i + 1), name: n, phone: '+91 98310 0' + (2000 + i * 111), email: 'sales@' + n.split(' ')[0].toLowerCase() + '.in', address: ['Burrabazar', 'Posta', 'Dankuni'][i] + ', West Bengal', gst: '19AAAB' + (4321 + i) + 'C1Z' + (7 + i), payable: [25000, 0, 12000][i] })),
    materials: mats.map((m, i) => ({ id: id('RM-', i + 1), name: m[0], unit: m[1], stock: m[2], reorder: m[3], supplier: S[m[4]] })),
    production: [...Array(16)].map((_, i) => ({ id: id('BN-', 16 - i), date: ago(i >> 1), product: P[i % 4][0], qty: 150 + i * 83 % 250, operator: O[i % 3], notes: i % 5 ? '' : 'Quality check passed' })),
    sales: [...Array(18)].map((_, i) => { const p = P[i % 4], q = 100 + i * 137 % 400, a = q * p[2]; return { id: id('INV-', 18 - i), date: ago(i * 10), customer: C[i % 3], product: p[0], qty: q, rate: p[2], gst: 18, amount: a, total: +(a * 1.18).toFixed(2), status: ['Paid', 'Pending', 'Paid', 'Partial'][i % 4] }; }),
    customers: C.map((n, i) => ({ id: id('CU-', i + 1), name: n, phone: '+91 98300 0' + (1000 + i * 111), email: n.split(' ')[0].toLowerCase() + '@example.com', address: ['Burrabazar', 'Salt Lake Sector V', 'Dum Dum'][i] + ', Kolkata', gst: ['19', '27', '19'][i] + 'AAAC' + (1234 + i) + 'A1Z' + (5 + i), outstanding: [18500, 0, 42000][i] })),
    purchases: [...Array(8)].map((_, i) => { const m = mats[i % 7], q = 40 + i * 37 % 120, pr = [62, 48, 38, 70, 80, 95, 720, 55][i]; return { id: id('PO-', 8 - i), date: ago(i * 20), supplier: S[m[4]], material: m[0], qty: q, price: pr, total: q * pr, status: 'Received' }; }),
    activity: [{ t: t(1), u: 'Admin', m: 'Added production batch BN-016' }, { t: t(5), u: 'Admin', m: 'Generated invoice INV-018' }, { t: t(26), u: 'Admin', m: 'Purchase PO-008 received from ABC Metals' }, { t: t(50), u: 'Admin', m: 'Stock adjusted for Kraft Paper' }],
    users: [['Admin User', 'admin@skrmfg.in', 'admin123', 'Admin'], ['Meera Sen', 'manager@skrmfg.in', 'manager123', 'Manager'], ['Rakesh Das', 'operator@skrmfg.in', 'operator123', 'Operator'], ['Anil Gupta', 'accounts@skrmfg.in', 'accounts123', 'Accountant'], ['Guest Viewer', 'viewer@skrmfg.in', 'viewer123', 'Viewer']].map((u, i) => ({ id: id('US-', i + 1), name: u[0], email: u[1], password: u[2], role: u[3] })),
    scans: [],
    settings: { name: 'SKR Manufacturing Pvt Ltd', gst: '19AAACS1234K1Z5', phone: '+91 33 4000 1234', email: 'accounts@skrmfg.in', address: 'Plot 21, Industrial Estate, Howrah, West Bengal 711101', logo: '', theme: 'light' }
  };
}
const save = () => localStorage.setItem(KEY, JSON.stringify(DB));
function load() { try { DB = JSON.parse(localStorage.getItem(KEY)); } catch (e) { } if (!DB) { DB = seed(); save(); } }
const log = m => { DB.activity.unshift({ t: new Date().toISOString(), u: CU.name || 'Admin', m }); DB.activity.length = Math.min(DB.activity.length, 40); };

/* ---------- Derived data ---------- */
const P = n => DB.products.find(p => p.name == n);
const fg = () => DB.products.map(p => {
  const pr = sum(DB.production.filter(r => r.product == p.name), 'qty'), so = sum(DB.sales.filter(r => r.product == p.name), 'qty');
  return { id: p.name, name: p.name, opening: p.opening, produced: pr, sold: so, closing: p.opening + pr - so };
});
const closing = n => (fg().find(r => r.name == n) || {}).closing || 0;
const nextId = m => m.px + String(Math.max(0, ...DB[m.k].map(r => parseInt(r.id.split('-')[1]) || 0)) + 1).padStart(3, '0');

/* ---------- Module definitions (drive tables, forms, validation) ----------
   f: form fields [key,label,type,options,optional]   c: table columns [key,label,formatter] */
const sup = () => DB.suppliers.map(s => s.name), pn = () => DB.products.map(p => p.name), mn = () => DB.materials.map(m => m.name), cn = () => DB.customers.map(c => c.name);
const bd = (t, c) => `<span class="b ${c}">${t}</span>`;
const pay = s => bd(s, { Paid: 'ok', Pending: 'bad', Partial: 'warn' }[s]);
const M = {
  materials: {
    t: 'Raw Materials', k: 'materials', px: 'RM-', add: 'Add material',
    f: [['id', 'Material ID', 'auto'], ['name', 'Material name'], ['unit', 'Unit', 'select', ['kg', 'pcs', 'm', 'ltr', 'sheet']], ['stock', 'Current stock', 'number'], ['reorder', 'Reorder level', 'number'], ['supplier', 'Supplier', 'select', sup]],
    c: [['id', 'ID'], ['name', 'Material'], ['unit', 'Unit'], ['stock', 'Stock', r => num(r.stock)], ['reorder', 'Reorder level', r => num(r.reorder)], ['supplier', 'Supplier'], ['s', 'Status', r => r.stock <= r.reorder ? bd('Low stock', 'bad') : bd('In stock', 'ok')]],
    flt: ['Status', () => ['Low stock', 'In stock'], (r, v) => (r.stock <= r.reorder) == (v == 'Low stock')],
    x: [['adj', 'Adjust stock'], ['qr', 'QR label']]
  },
  production: {
    t: 'Production', k: 'production', px: 'BN-', add: 'Production entry',
    tb: '<button class="btn o" data-a="go" data-id="reports">Daily report</button>',
    x: [['qr', 'QR label']],
    f: [['id', 'Batch number', 'auto'], ['date', 'Date', 'date'], ['product', 'Product name', 'select', pn], ['qty', 'Quantity produced', 'number'], ['operator', 'Operator name'], ['notes', 'Notes', 'textarea', 0, 1]],
    c: [['id', 'Batch'], ['date', 'Date'], ['product', 'Product'], ['qty', 'Qty', r => num(r.qty)], ['cost', 'Batch cost', r => inr(r.qty * unitCost(r.product))], ['operator', 'Operator'], ['notes', 'Notes']],
    def: () => ({ date: today() }),
    flt: ['Product', pn, (r, v) => r.product == v],
    // Consume raw material according to each product's usage ratio
    hook: (r, o) => { const p = P(r.product), m = DB.materials.find(x => x.name == p.mat); if (m) m.stock = Math.max(0, +(m.stock - (r.qty - (o ? o.qty : 0)) * p.ratio).toFixed(2)); }
  },
  fg: {
    t: 'Finished Goods', rows: fg,
    c: [['name', 'Product'], ['opening', 'Opening stock', r => num(r.opening)], ['produced', 'Produced', r => '+' + num(r.produced)], ['sold', 'Sold', r => '−' + num(r.sold)], ['closing', 'Closing stock', r => `<b>${num(r.closing)}</b>`]],
    x: [['open', 'Set opening']]
  },
  sales: {
    t: 'Sales', k: 'sales', px: 'INV-', add: 'New sale',
    def: () => ({ date: today(), rate: DB.products[0].rate, gst: 18, status: 'Pending' }),
    f: [['id', 'Invoice number', 'auto'], ['date', 'Date', 'date'], ['customer', 'Customer name', 'select', cn], ['product', 'Product', 'select', pn], ['qty', 'Quantity', 'number'], ['rate', 'Unit rate (₹)', 'number'], ['gst', 'GST %', 'select', [0, 5, 12, 18, 28]], ['status', 'Payment status', 'select', ['Paid', 'Pending', 'Partial']]],
    c: [['id', 'Invoice'], ['date', 'Date'], ['customer', 'Customer'], ['product', 'Product'], ['qty', 'Qty'], ['amount', 'Amount', r => inr(r.amount)], ['gst', 'GST', r => inr(r.total - r.amount)], ['total', 'Total', r => `<b>${inr(r.total)}</b>`], ['status', 'Payment', r => pay(r.status)]],
    flt: ['Payment', () => ['Paid', 'Pending', 'Partial'], (r, v) => r.status == v],
    x: [['inv', 'GST invoice'], ['wa', 'WhatsApp']],
    calc: r => { r.amount = r.qty * r.rate; r.total = +(r.amount * (1 + r.gst / 100)).toFixed(2); },
    chk: (r, o) => r.qty > closing(r.product) + (o && o.product == r.product ? o.qty : 0) ? `Only ${num(closing(r.product))} units of ${r.product} are in stock.` : '',
    hook: (r, o) => { const c = DB.customers.find(x => x.name == r.customer); if (!o && c && r.status != 'Paid') c.outstanding += r.total; }
  },
  customers: {
    t: 'Customers', k: 'customers', px: 'CU-', add: 'Add customer',
    f: [['id', 'Customer ID', 'auto'], ['name', 'Customer name'], ['phone', 'Phone', 'tel'], ['email', 'Email', 'email', 0, 1], ['address', 'Address', 'textarea', 0, 1], ['gst', 'GST number', 'text', 0, 1], ['outstanding', 'Outstanding balance (₹)', 'number']],
    c: [['name', 'Customer', r => `<b>${esc(r.name)}</b><br><small class="mut">${r.id}</small>`], ['phone', 'Phone'], ['email', 'Email'], ['gst', 'GST no.'], ['outstanding', 'Outstanding', r => r.outstanding > 0 ? bd(inr(r.outstanding), 'bad') : bd('Clear', 'ok')]],
    flt: ['Dues', () => ['Outstanding', 'Clear'], (r, v) => (r.outstanding > 0) == (v == 'Outstanding')],
    x: [['pay', 'Receive payment']]
  },
  purchases: {
    t: 'Purchase Orders', k: 'purchases', px: 'PO-', add: 'New purchase order',
    def: () => ({ date: today(), status: 'Ordered' }),
    f: [['id', 'PO number', 'auto'], ['date', 'Date', 'date'], ['supplier', 'Supplier', 'select', sup], ['material', 'Material', 'select', mn], ['qty', 'Quantity', 'number'], ['price', 'Price per unit (₹)', 'number'], ['status', 'Status', 'select', ['Ordered', 'Received', 'Cancelled']]],
    c: [['id', 'PO'], ['date', 'Date'], ['supplier', 'Supplier'], ['material', 'Material'], ['qty', 'Qty'], ['price', 'Price', r => inr(r.price)], ['total', 'Total', r => `<b>${inr(r.total)}</b>`], ['status', 'Status', r => bd(r.status, { Ordered: 'warn', Received: 'ok', Cancelled: 'bad' }[r.status])]],
    flt: ['Status', () => ['Ordered', 'Received', 'Cancelled'], (r, v) => r.status == v],
    x: [['recv', 'Receive', r => r.status == 'Ordered'], ['po', 'Print PO']],
    calc: r => r.total = r.qty * r.price,
    // Receiving a PO adds stock and raises the supplier payable
    hook: (r, o) => { if (r.status == 'Received' && (!o || o.status != 'Received')) recv(r); }
  },
  suppliers: {
    t: 'Suppliers', k: 'suppliers', px: 'SP-', add: 'Add supplier',
    f: [['id', 'Supplier ID', 'auto'], ['name', 'Supplier name'], ['phone', 'Phone', 'tel'], ['email', 'Email', 'email', 0, 1], ['address', 'Address', 'textarea', 0, 1], ['gst', 'GST number', 'text', 0, 1], ['payable', 'Payable balance (₹)', 'number']],
    c: [['name', 'Supplier', r => `<b>${esc(r.name)}</b><br><small class="mut">${r.id}</small>`], ['phone', 'Phone'], ['gst', 'GST no.'], ['n', 'Orders', r => DB.purchases.filter(p => p.supplier == r.name).length], ['payable', 'Payable', r => r.payable > 0 ? bd(inr(r.payable), 'warn') : bd('Clear', 'ok')]],
    flt: ['Dues', () => ['Payable', 'Clear'], (r, v) => (r.payable > 0) == (v == 'Payable')],
    x: [['spay', 'Pay supplier']]
  },
  costing: {
    t: 'Production Costing',
    rows: () => DB.products.map(p => { const mc = +(p.ratio * mp(p.mat)).toFixed(2), c = unitCost(p.name); return { id: p.name, name: p.name, mc, labour: p.labour, oh: p.oh, cost: c, rate: p.rate, margin: Math.round((p.rate - c) / p.rate * 100) }; }),
    c: [['name', 'Product'], ['mc', 'Material', r => inr2(r.mc)], ['labour', 'Labour', r => inr2(r.labour)], ['oh', 'Overhead', r => inr2(r.oh)], ['cost', 'Cost per unit', r => `<b>${inr2(r.cost)}</b>`], ['rate', 'Selling rate', r => inr2(r.rate)], ['margin', 'Margin', r => bd(r.margin + '%', r.margin >= 30 ? 'ok' : r.margin >= 15 ? 'warn' : 'bad')]],
    x: [['costs', 'Set labour and overhead']]
  },
  users: {
    t: 'Users and Roles', k: 'users', px: 'US-', add: 'Add user',
    f: [['id', 'User ID', 'auto'], ['name', 'Full name'], ['email', 'Email', 'email'], ['password', 'Password'], ['role', 'Role', 'select', Object.keys(PERM)]],
    c: [['name', 'Name'], ['email', 'Email'], ['role', 'Role', r => bd(r.role, r.role == 'Admin' ? 'bad' : 'ok')]],
    chk: (r, o) => DB.users.some(u => u.email == r.email && u !== o) ? 'That email is already in use.' : ''
  }
};
const NAV = [
  ['dashboard', 'Dashboard', 'M3 3h7v7H3zM14 3h7v7h-7zM14 14h7v7h-7zM3 14h7v7H3z'],
  ['materials', 'Raw Materials', 'M21 8l-9-5-9 5v8l9 5 9-5zM3 8l9 5 9-5M12 13v8'],
  ['production', 'Production', 'M2 20V8l6 4V8l6 4V4h4v16z'],
  ['fg', 'Finished Goods', 'M16.5 9.4l-9-5.2M21 16V8l-9-5-9 5v8l9 5zM3.3 7L12 12l8.7-5M12 22V12'],
  ['costing', 'Costing', 'M6 3h12M6 8h12M6 13l8.5 8M6 13h3a4 4 0 0 0 0-8'],
  ['sales', 'Sales', 'M23 6l-9.5 9.5-5-5L1 18M17 6h6v6'],
  ['customers', 'Customers', 'M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2M9 11a4 4 0 1 0 0-8 4 4 0 0 0 0 8M23 21v-2a4 4 0 0 0-3-3.9M16 3.1a4 4 0 0 1 0 7.8'],
  ['suppliers', 'Suppliers', 'M1 3h15v13H1zM16 8h4l3 3v5h-7zM5.5 19a1.5 1.5 0 1 0 0 .01M18.5 19a1.5 1.5 0 1 0 0 .01'],
  ['purchases', 'Purchase Orders', 'M6 6h15l-1.5 9h-12zM6 6L5 2H2M9 20a1 1 0 1 0 0 .01M18 20a1 1 0 1 0 0 .01'],
  ['scan', 'Scanner and QR', 'M3 7V5a2 2 0 0 1 2-2h2M17 3h2a2 2 0 0 1 2 2v2M21 17v2a2 2 0 0 1-2 2h-2M7 21H5a2 2 0 0 1-2-2v-2M7 12h10'],
  ['reports', 'Reports', 'M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8zM14 2v6h6M8 13h8M8 17h8'],
  ['users', 'Users and Roles', 'M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2M9 11a4 4 0 1 0 0-8 4 4 0 0 0 0 8M23 21v-2a4 4 0 0 0-3-3.9M16 3.1a4 4 0 0 1 0 7.8'],
  ['settings', 'Settings', 'M4 21v-7M4 10V3M12 21v-9M12 8V3M20 21v-5M20 12V3M1 14h6M9 8h6M17 16h6']
];

/* ---------- UI primitives: toast, modal, forms ---------- */
function toast(m, t = 'ok') { const e = document.createElement('div'); e.className = 't ' + t; e.textContent = m; $('#toasts').append(e); setTimeout(() => e.remove(), 3200); }
const modal = h => { $('#modal').innerHTML = `<div class="mo">${h}</div>`; }, closeM = () => { $('#modal').innerHTML = ''; };
const busy = (ms, cb) => { const l = $('#loader'); l.classList.remove('off'); setTimeout(() => { cb(); l.classList.add('off'); }, ms); };
function ask(msg, cb) { CONF = cb; modal(`<div class="md sm"><h3>Please confirm</h3><p>${esc(msg)}</p><div class="ft"><button class="btn o" data-x>Cancel</button><button class="btn d" data-a="yes">Delete</button></div></div>`); }

function openForm(title, F, v, ok) {
  const L = o => typeof o == 'function' ? o() : o;
  modal(`<div class="md"><h3>${esc(title)}</h3><form>${F.map(([k, l, t = 'text', o, opt]) => `<label>${l}${t == 'select'
    ? `<select name="${k}">${L(o).map(x => `<option ${x == v[k] ? 'selected' : ''}>${esc(x)}</option>`).join('')}</select>`
    : t == 'textarea' ? `<textarea name="${k}" rows="2">${esc(v[k])}</textarea>`
      : `<input name="${k}" type="${t == 'auto' ? 'text' : t}" value="${esc(v[k])}" ${t == 'auto' ? 'readonly' : opt ? '' : 'required'} ${t == 'number' ? 'step="any" min="0"' : ''}>`}</label>`).join('')}
    <div class="ft"><button type="button" class="btn o" data-x>Cancel</button><button class="btn p">Save</button></div></form></div>`);
  const f = $('#modal form'), E = f.elements;
  f.onchange = e => { if (e.target.name == 'product' && E.rate) E.rate.value = P(e.target.value).rate; }; // auto-fill unit rate
  f.onsubmit = e => { e.preventDefault(); const d = {}; F.forEach(([k, , t, o]) => { let x = E[k].value; if (t == 'number' || (Array.isArray(o) && typeof o[0] == 'number')) x = +x; d[k] = x; }); ok(d); };
}

/* ---------- Generic CRUD ---------- */
const openAdd = k => { const m = M[k]; openForm(m.add, m.f, { ...(m.def ? m.def() : {}), id: nextId(m) }, d => saveRec(k, d)); };
function saveRec(k, d, id) {
  const m = M[k], a = DB[k], o = id && a.find(r => r.id == id), r = { ...o, ...d };
  m.calc && m.calc(r);
  const e = m.chk && m.chk(r, o); if (e) return toast(e, 'bad');
  m.hook && m.hook(r, o);
  o ? Object.assign(o, r) : a.unshift(r);
  log((o ? 'Updated ' : 'Added ') + m.t.toLowerCase() + ' ' + r.id); save(); closeM(); draw(k); toast((o ? 'Updated ' : 'Saved ') + r.id);
}

/* ---------- Module page: search, filter, table, pagination ---------- */
function page(k) {
  const m = M[k], s = ST[k] = ST[k] || { q: '', f: '', p: 1 };
  $('#view').innerHTML = `<div class="head"><div><h2>${m.t}</h2><p class="mut" id="cnt"></p></div><div class="tools">${m.tb || ''}${k != 'users' ? '<button class="btn o" data-a="exp">Export Excel</button>' : ''}${m.f && !RO() ? '<button class="btn o" data-a="imp">Import Excel</button>' : ''}${m.add && !RO() ? `<button class="btn p" data-a="add">+ ${m.add}</button>` : ''}</div></div>
  <div class="card"><div class="bar"><input id="q" placeholder="Search ${m.t.toLowerCase()}…" value="${esc(s.q)}" aria-label="Search">${m.flt ? `<select id="fl" aria-label="Filter"><option value="">All ${m.flt[0].toLowerCase()}</option>${m.flt[1]().map(x => `<option ${x == s.f ? 'selected' : ''}>${esc(x)}</option>`).join('')}</select>` : ''}</div><div id="tbl"></div></div>`;
  $('#q').oninput = e => { s.q = e.target.value; s.p = 1; draw(k); };
  if (m.flt) $('#fl').onchange = e => { s.f = e.target.value; s.p = 1; draw(k); };
  draw(k);
}
function draw(k) {
  const m = M[k], s = ST[k], q = s.q.toLowerCase();
  const R = (m.rows ? m.rows() : DB[k]).filter(r => (!q || Object.values(r).join(' ').toLowerCase().includes(q)) && (!s.f || m.flt[2](r, s.f)));
  const n = Math.max(1, Math.ceil(R.length / PG)); s.p = Math.min(s.p, n);
  $('#cnt').textContent = R.length + ' records';
  const btn = (a, l, c, id) => `<button class="btn ${c} s" data-a="${a}" data-id="${esc(id)}">${l}</button>`;
  $('#tbl').innerHTML = `<div class="tw"><table><thead><tr>${m.c.map(c => `<th>${c[1]}</th>`).join('')}<th></th></tr></thead><tbody>${R.slice((s.p - 1) * PG, s.p * PG).map(r => `<tr>${m.c.map(c => `<td>${c[2] ? c[2](r) : esc(r[c[0]])}</td>`).join('')}<td class="ac">${(m.x || []).filter(x => (!x[2] || x[2](r)) && (!RO() || !WR.has(x[0]))).map(x => btn(x[0], x[1], 'o', r.id)).join('')}${m.f && !RO() ? btn('edit', 'Edit', 'o', r.id) + btn('del', 'Delete', 'd', r.id) : ''}</td></tr>`).join('') || '<tr><td colspan="9" class="mut">No records match. Try a different search or filter.</td></tr>'}</tbody></table></div>
  <div class="pg"><span class="mut">Page ${s.p} of ${n}</span><div><button class="btn o s" data-a="pv" ${s.p < 2 ? 'disabled' : ''}>Previous</button><button class="btn o s" data-a="nx" ${s.p >= n ? 'disabled' : ''}>Next</button></div></div>`;
}

/* ---------- Dashboard ---------- */
function count(el) {
  const to = +el.dataset.v, cur = el.dataset.c == '1', t0 = performance.now();
  (function f(t) { const p = Math.min(1, (t - t0) / 900); el.textContent = (cur ? '₹' : '') + Math.round(to * (1 - Math.pow(1 - p, 3))).toLocaleString('en-IN'); if (p < 1) requestAnimationFrame(f); })(t0);
}
function dash() {
  const td = today(), mo = td.slice(0, 7), sa = DB.sales, low = DB.materials.filter(m => m.stock <= m.reorder);
  const K = [['Total raw materials', sum(DB.materials, 'stock'), 0, 'units across ' + DB.materials.length + ' items', 'pri'], ['Total finished goods', sum(fg(), 'closing'), 0, 'units ready to ship', 'ok'], ["Today's production", sum(DB.production.filter(r => r.date == td), 'qty'), 0, 'units today', 'ok'], ['Monthly sales', sum(sa.filter(r => r.date.startsWith(mo)), 'total'), 1, 'this month, incl. GST', 'warn'], ['Total orders', sa.length, 0, 'all invoices', 'pri'], ['Pending orders', sa.filter(r => r.status != 'Paid').length, 0, 'awaiting full payment', 'bad'], ['Profit', sum(sa, 'amount') - sum(DB.purchases.filter(r => r.status == 'Received'), 'total'), 1, 'sales less purchases', 'ok']];
  const QA = [['New production', 'production:add'], ['New sale', 'sales:add'], ['Add material', 'materials:add'], ['New purchase', 'purchases:add'], ['Add customer', 'customers:add'], ['Reports', 'reports']];
  $('#view').innerHTML = `<div class="head"><div><h2>Plant overview</h2><p class="mut">How production, stock and sales stand right now.</p></div><div class="tools">${QA.map(q => `<button class="btn o s" data-a="go" data-id="${q[1]}">${q[0]}</button>`).join('')}</div></div>
  <div class="kpis">${K.map(k => `<div class="card kpi ${k[4]}"><span class="mut">${k[0]}</span><b data-v="${k[1]}" data-c="${k[2]}">0</b><small class="mut">${k[3]}</small></div>`).join('')}</div>
  <div class="cg"><div class="card"><h3>Daily production</h3><div class="cv"><canvas id="c1"></canvas></div></div><div class="card"><h3>Monthly sales</h3><div class="cv"><canvas id="c2"></canvas></div></div><div class="card"><h3>Inventory consumption</h3><div class="cv"><canvas id="c3"></canvas></div></div></div>
  <div class="cg2"><div class="card"><h3>Recent activities</h3><div class="tw"><table><thead><tr><th>When</th><th>User</th><th>Activity</th></tr></thead><tbody>${DB.activity.slice(0, 6).map(a => `<tr><td>${new Date(a.t).toLocaleString('en-IN', { day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit' })}</td><td>${a.u}</td><td>${esc(a.m)}</td></tr>`).join('')}</tbody></table></div></div>
  <div class="card"><h3>Low stock alerts</h3>${low.map(m => `<div class="ls"><div><b>${esc(m.name)}</b><br><small class="mut">${num(m.stock)} ${m.unit} left, reorder at ${num(m.reorder)}</small></div>${bd('Low', 'bad')}<div class="pr"><i style="width:${Math.min(100, m.stock / m.reorder * 100)}%"></i></div></div>`).join('') || '<p class="mut">All materials are above their reorder level.</p>'}</div></div>`;
  $$('.kpi b').forEach(count);
  if (!window.Chart) return toast('Charts need an internet connection to load Chart.js', 'warn');
  const acc = getComputedStyle(document.documentElement).getPropertyValue('--acc').trim(), opt = { responsive: true, maintainAspectRatio: false, animation: { duration: 1200, easing: 'easeOutQuart' }, plugins: { legend: { display: false } } };
  const days = [...Array(7)].map((_, i) => ago(6 - i)), now = new Date();
  const mons = [...Array(6)].map((_, i) => ymd(new Date(now.getFullYear(), now.getMonth() - 5 + i, 1)).slice(0, 7));
  const use = {}; DB.production.forEach(r => { const p = P(r.product); if (p) use[p.mat] = (use[p.mat] || 0) + r.qty * p.ratio; });
  const mk = (id, cfg) => charts.push(new Chart($(id), cfg));
  mk('#c1', { type: 'line', data: { labels: days.map(d => d.slice(5)), datasets: [{ data: days.map(d => sum(DB.production.filter(r => r.date == d), 'qty')), borderColor: acc, backgroundColor: acc + '26', fill: true, tension: .4 }] }, options: opt });
  mk('#c2', { type: 'bar', data: { labels: mons, datasets: [{ data: mons.map(m => Math.round(sum(sa.filter(r => r.date.startsWith(m)), 'total'))), backgroundColor: '#10B981', borderRadius: 8 }] }, options: opt });
  mk('#c3', { type: 'doughnut', data: { labels: Object.keys(use), datasets: [{ data: Object.values(use).map(Math.round), backgroundColor: ['#1E3A8A', '#10B981', '#F59E0B', '#EF4444', '#6366F1'], borderWidth: 0 }] }, options: { ...opt, plugins: { legend: { display: true, position: 'bottom' } } } });
}

/* ---------- Reports ---------- */
function repData(t, d) {
  const mo = d.slice(0, 7), pr = DB.production, sl = DB.sales.filter(r => r.date.startsWith(mo));
  if (t == 'dp') return [`Daily Production Report, ${d}`, ['Batch', 'Product', 'Quantity', 'Operator', 'Notes'], pr.filter(r => r.date == d).map(r => [r.id, r.product, r.qty, r.operator, r.notes])];
  if (t == 'mp') return [`Monthly Production Report, ${mo}`, ['Product', 'Batches', 'Quantity'], DB.products.map(p => { const x = pr.filter(r => r.product == p.name && r.date.startsWith(mo)); return [p.name, x.length, sum(x, 'qty')]; })];
  if (t == 'inv') return ['Inventory Report, ' + d, ['Item', 'Type', 'Stock', 'Status'], [...DB.materials.map(m => [m.name, 'Raw material', m.stock + ' ' + m.unit, m.stock <= m.reorder ? 'Low stock' : 'OK']), ...fg().map(f => [f.name, 'Finished good', f.closing, 'OK'])]];
  if (t == 'sl') return [`Sales Report, ${mo}`, ['Invoice', 'Date', 'Customer', 'Product', 'Qty', 'Total', 'Payment'], sl.map(r => [r.id, r.date, r.customer, r.product, r.qty, inr(r.total), r.status])];
  const rev = sum(sl, 'amount'), gst = sum(sl, 'total') - rev, cost = sum(DB.purchases.filter(r => r.status == 'Received' && r.date.startsWith(mo)), 'total');
  return [`Profit & Loss Report, ${mo}`, ['Particular', 'Amount'], [['Sales revenue', inr(rev)], ['GST collected', inr(gst)], ['Purchases (cost)', inr(cost)], ['Gross profit', inr(rev - cost)]]];
}
function reps() {
  const s = ST.rep = ST.rep || { t: 'dp', d: today() }, T = { dp: 'Daily Production', mp: 'Monthly Production', inv: 'Inventory', sl: 'Sales', pl: 'Profit & Loss' };
  $('#view').innerHTML = `<div class="head"><div><h2>Reports</h2><p class="mut">Pick a report and date, then export or print.</p></div><div class="tools"><button class="btn o" data-a="pdf">PDF</button><button class="btn o" data-a="xls">Excel</button><button class="btn p" data-a="prn">Print</button></div></div>
  <div class="card bar"><select id="rt">${Object.entries(T).map(([k, v]) => `<option value="${k}" ${k == s.t ? 'selected' : ''}>${v} report</option>`).join('')}</select><input type="date" id="rd" value="${s.d}"></div>
  <div class="card print-area" id="rep"></div>`;
  const fill = () => { const [t, h, r] = repData(s.t, s.d), S = DB.settings; $('#rep').innerHTML = `<h3>${esc(S.name)}</h3><p class="mut">${esc(S.address)}, GSTIN ${esc(S.gst)}</p><h2 style="margin:1rem 0">${esc(t)}</h2><div class="tw"><table><thead><tr>${h.map(x => `<th>${x}</th>`).join('')}</tr></thead><tbody>${r.map(x => `<tr>${x.map(c => `<td>${esc(c)}</td>`).join('')}</tr>`).join('') || '<tr><td colspan="9" class="mut">No records for this period.</td></tr>'}</tbody></table></div>`; };
  $('#rt').onchange = e => { s.t = e.target.value; fill(); }; $('#rd').onchange = e => { s.d = e.target.value || today(); fill(); }; fill();
}

/* ---------- Invoice layout ---------- */
function invoice(r) { // GST tax invoice: CGST+SGST within the same state, IGST otherwise
  const s = DB.settings, c = DB.customers.find(x => x.name == r.customer) || {}, p = P(r.product) || {}, g = r.total - r.amount;
  const intra = (c.gst || s.gst).slice(0, 2) == s.gst.slice(0, 2), h = g / 2;
  const tax = intra ? `<div><span>CGST ${r.gst / 2}%</span><span>${inr2(h)}</span></div><div><span>SGST ${r.gst / 2}%</span><span>${inr2(h)}</span></div>` : `<div><span>IGST ${r.gst}%</span><span>${inr2(g)}</span></div>`;
  modal(`<div class="md wide"><div class="print-area inv"><div class="ih"><div>${s.logo ? `<img class="lg" src="${s.logo}" alt="">` : ''}<h3>${esc(s.name)}</h3><small>${esc(s.address)}<br>GSTIN ${esc(s.gst)}, ${esc(s.phone)}</small></div><div><h2>TAX INVOICE</h2><small>${r.id}<br>Date ${r.date}<br>Place of supply: state code ${esc((c.gst || s.gst).slice(0, 2))}</small><br>${pay(r.status)}</div></div>
  <p><b>Bill to</b><br>${esc(r.customer)}<br><small>${esc(c.address || '')}<br>GSTIN ${esc(c.gst || 'Not provided')}</small></p><br>
  <table><thead><tr><th>Product</th><th>HSN</th><th>Qty</th><th>Rate</th><th>Amount</th></tr></thead><tbody><tr><td>${esc(r.product)}</td><td>${esc(p.hsn || '')}</td><td>${r.qty}</td><td>${inr2(r.rate)}</td><td>${inr2(r.amount)}</td></tr></tbody></table>
  <div class="tot"><div><span>Taxable value</span><span>${inr2(r.amount)}</span></div>${tax}<div><b>Total</b><b>${inr2(r.total)}</b></div></div>
  <p><small><b>Amount in words:</b> ${words(r.total)}</small></p><div class="ih"><small>Payment terms: 15 days from invoice date.<br>This is a computer generated invoice.</small><div id="iq"></div></div></div>
  <div class="ft"><button class="btn o" data-x>Close</button><button class="btn o" data-a="wa" data-id="${r.id}">Share on WhatsApp</button><button class="btn p" data-a="prn">Print or save as PDF</button></div></div>`);
  qr($('#iq'), `SKR|${r.id}|${r.customer}|${r.total}`, 80);
}

/* ---------- Settings ---------- */
function sets() {
  const s = DB.settings;
  $('#view').innerHTML = `<div class="head"><div><h2>Settings</h2><p class="mut">Company profile and appearance</p></div></div><form class="card fg2" id="sf">
  ${[['name', 'Company name'], ['gst', 'GST number'], ['phone', 'Phone'], ['email', 'Email']].map(([k, l]) => `<label>${l}<input name="${k}" value="${esc(s[k])}"></label>`).join('')}
  <label class="full">Address<textarea name="address" rows="2">${esc(s.address)}</textarea></label>
  <label>Logo<input type="file" name="logo" accept="image/*"></label>
  <label>Theme<select name="theme"><option value="light">Light</option><option value="dark">Dark</option></select></label>
  <div class="full"><button class="btn p">Save settings</button></div></form>`;
  const f = $('#sf'), E = f.elements; E.theme.value = s.theme;
  f.onsubmit = e => {
    e.preventDefault();
    const done = () => { ['name', 'gst', 'phone', 'email', 'address', 'theme'].forEach(k => s[k] = E[k].value); log('Updated company settings'); save(); theme(); brand(); toast('Settings saved'); };
    const fl = E.logo.files[0]; if (fl) { const r = new FileReader(); r.onload = () => { s.logo = r.result; done(); }; r.readAsDataURL(fl); } else done();
  };
}

/* ---------- Theme, branding, routing ---------- */
function theme() {
  const d = DB.settings.theme == 'dark'; document.documentElement.dataset.theme = DB.settings.theme; $('#tg').textContent = d ? '☀' : '☾';
  if (window.Chart) { Chart.defaults.color = d ? '#94A3B8' : '#64748B'; Chart.defaults.borderColor = d ? 'rgba(148,163,184,.15)' : 'rgba(30,58,138,.08)'; Chart.defaults.font.family = 'Poppins'; }
}
const brand = () => { $('#bn').textContent = DB.settings.name; $('#bl').innerHTML = DB.settings.logo ? `<img src="${DB.settings.logo}" alt="">` : 'SKR'; };
function render() { stopCam(); charts.forEach(c => c.destroy()); charts = []; ({ dashboard: dash, reports: reps, settings: sets, scan }[route] || (() => page(route)))(); }
function nav(r) { if (!can(r)) return toast('Your role does not have access to that page.', 'warn'); route = r; $$('.nav a').forEach(a => a.classList.toggle('on', a.dataset.r == r)); $('#ttl').textContent = NAV.find(n => n[0] == r)[1]; $('.side').classList.remove('open'); render(); scrollTo(0, 0); }

/* ---------- Actions (wired through data-a attributes) ---------- */
const A = {
  nav, add: () => openAdd(route), go: v => { const [r, a] = v.split(':'); nav(r); a && can(r) && !RO() && openAdd(r); },
  edit: id => { const m = M[route]; openForm('Edit ' + id, m.f, DB[route].find(r => r.id == id), d => saveRec(route, d, id)); },
  del: id => id == CU.id && route == 'users' ? toast('You cannot delete your own account.', 'bad') : ask(`Delete ${id}? This cannot be undone.`, () => { DB[route] = DB[route].filter(r => r.id != id); log('Deleted ' + id); save(); draw(route); toast('Deleted ' + id, 'warn'); }),
  yes: () => { const c = CONF; closeM(); c && c(); },
  pv: () => { ST[route].p--; draw(route); }, nx: () => { ST[route].p++; draw(route); },
  menu: () => $('.side').classList.toggle('open'),
  dark: () => { DB.settings.theme = DB.settings.theme == 'dark' ? 'light' : 'dark'; save(); theme(); render(); },
  fp: () => { $('#lf').hidden = !$('#lf').hidden; $('#ff').hidden = !$('#ff').hidden; },
  out: () => { localStorage.removeItem('skr_auth'); sessionStorage.removeItem('skr_auth'); location.reload(); },
  inv: id => invoice(DB.sales.find(r => r.id == id)),
  adj: id => {
    const m = DB.materials.find(r => r.id == id);
    openForm('Adjust stock: ' + m.name, [['t', 'Adjustment', 'select', ['Add stock', 'Remove stock']], ['q', `Quantity (${m.unit})`, 'number'], ['why', 'Reason']], {}, d => {
      const n = m.stock + (d.t == 'Add stock' ? d.q : -d.q); if (n < 0) return toast(`You can remove at most ${m.stock} ${m.unit}.`, 'bad');
      m.stock = n; log(`${d.t} for ${m.name}: ${d.q} ${m.unit} (${d.why})`); save(); closeM(); draw('materials'); toast('Stock updated');
    });
  },
  pay: id => {
    const c = DB.customers.find(r => r.id == id);
    openForm('Receive payment: ' + c.name, [['q', 'Amount received (₹)', 'number'], ['mode', 'Mode', 'select', ['UPI', 'Bank transfer', 'Cash', 'Cheque']]], {}, d => {
      if (d.q > c.outstanding) return toast(`Amount is more than the ${inr(c.outstanding)} outstanding.`, 'bad');
      c.outstanding -= d.q; log(`Received ${inr(d.q)} from ${c.name} via ${d.mode}`); save(); closeM(); draw('customers'); toast('Payment recorded');
    });
  },
  open: id => { const p = P(id); openForm('Opening stock: ' + id, [['o', 'Opening stock', 'number']], { o: p.opening }, d => { p.opening = d.o; save(); closeM(); draw('fg'); toast('Opening stock updated'); }); },
  pdf: () => { toast('Choose "Save as PDF" as the destination in the print dialog'); setTimeout(print, 700); },
  prn: () => print(),
  xls: () => {
    if (noXL()) return; const [t, h, r] = repData(ST.rep.t, ST.rep.d), wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([[DB.settings.name], [t], [], h, ...r]), 'Report'); xl(wb, t.replace(/\W+/g, '_') + '.xlsx');
  }
};
document.addEventListener('click', e => {
  if (!e.target.closest('.side,[data-a=menu]')) $('.side').classList.remove('open');
  if (e.target.matches('.mo') || e.target.closest('[data-x]')) return closeM();
  const b = e.target.closest('[data-a]'); if (!b) return;
  if (b.tagName == 'A') e.preventDefault();
  A[b.dataset.a] && A[b.dataset.a](b.dataset.id);
});
document.addEventListener('keydown', e => { if (e.key == 'Escape') closeM(); });
$('#gs').onkeydown = e => { // global search: jump to first module with a match
  if (e.key != 'Enter') return; const q = e.target.value.toLowerCase();
  const k = ['materials', 'customers', 'sales', 'production', 'purchases'].find(k => DB[k].some(r => Object.values(r).join(' ').toLowerCase().includes(q)));
  k ? (ST[k] = { q: e.target.value, f: '', p: 1 }, nav(k)) : toast(`No match found for "${e.target.value}"`, 'warn');
};

/* ---------- Extra actions: Excel, QR, scanner, WhatsApp, PO, costing ---------- */
Object.assign(A, {
  exp: () => { // export the current module to a real .xlsx workbook
    if (noXL()) return; const m = M[route], cols = m.f ? m.f.map(f => [f[0], f[1]]) : m.c.map(c => [c[0], c[1]]), rows = m.rows ? m.rows() : DB[route];
    const ws = XLSX.utils.json_to_sheet(rows.map(r => Object.fromEntries(cols.map(([k, l]) => [l, r[k]]))), { header: cols.map(c => c[1]) }), wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, m.t.slice(0, 30)); xl(wb, `SKR_${m.t.replace(/\W+/g, '_')}_${today()}.xlsx`);
  },
  imp: () => $('#imp').click(),
  qr: id => {
    const m = DB.materials.find(r => r.id == id), b = DB.production.find(r => r.id == id), t = b ? [b.id, b.product, `${b.date}, ${num(b.qty)} units, ${b.operator}`] : [m.id, m.name, `Stock ${num(m.stock)} ${m.unit}`];
    modal(`<div class="md sm"><div class="print-area qrl"><b>${esc(DB.settings.name)}</b><div id="qb"></div><b>${t[0]}</b><span>${esc(t[1])}</span><small>${esc(t[2])}</small></div><div class="ft"><button class="btn o" data-x>Close</button><button class="btn p" data-a="prn">Print label</button></div></div>`);
    qr($('#qb'), `SKR|${t[0]}|${t[1]}`, 150);
  },
  wa: id => {
    const r = DB.sales.find(x => x.id == id), c = DB.customers.find(x => x.name == r.customer) || {}; let p = (c.phone || '').replace(/\D/g, ''); if (p.length == 10) p = '91' + p;
    const msg = `Hello ${r.customer}, your invoice ${r.id} dated ${r.date} from ${DB.settings.name} is ${inr2(r.total)} including GST. Payment status: ${r.status}. Thank you for your business.`;
    log('Shared invoice ' + r.id + ' on WhatsApp'); save(); window.open(`https://wa.me/${p}?text=${encodeURIComponent(msg)}`, '_blank');
  },
  recv: id => { const r = DB.purchases.find(x => x.id == id); recv(r); r.status = 'Received'; log(`Received ${r.id}: ${r.qty} ${r.material} from ${r.supplier}`); save(); draw('purchases'); toast('Stock updated from ' + r.id); },
  po: id => {
    const r = DB.purchases.find(x => x.id == id), s = DB.settings, v = DB.suppliers.find(x => x.name == r.supplier) || {};
    modal(`<div class="md wide"><div class="print-area inv"><div class="ih"><div><h3>${esc(s.name)}</h3><small>${esc(s.address)}<br>GSTIN ${esc(s.gst)}</small></div><div><h2>PURCHASE ORDER</h2><small>${r.id}<br>Date ${r.date}</small></div></div><p><b>Supplier</b><br>${esc(r.supplier)}<br><small>${esc(v.address || '')}<br>GSTIN ${esc(v.gst || 'Not provided')}</small></p><br><table><thead><tr><th>Material</th><th>Qty</th><th>Rate</th><th>Amount</th></tr></thead><tbody><tr><td>${esc(r.material)}</td><td>${r.qty}</td><td>${inr2(r.price)}</td><td>${inr2(r.total)}</td></tr></tbody></table><div class="tot"><div><b>Total</b><b>${inr2(r.total)}</b></div></div><p><small>Please quote this PO number on your delivery challan and invoice.</small></p></div><div class="ft"><button class="btn o" data-x>Close</button><button class="btn p" data-a="prn">Print PO</button></div></div>`);
  },
  spay: id => {
    const s = DB.suppliers.find(r => r.id == id);
    openForm('Pay supplier: ' + s.name, [['q', 'Amount paid (₹)', 'number'], ['mode', 'Mode', 'select', ['Bank transfer', 'UPI', 'Cash', 'Cheque']]], {}, d => {
      if (d.q > s.payable) return toast(`Amount is more than the ${inr(s.payable)} payable.`, 'bad');
      s.payable -= d.q; log(`Paid ${inr(d.q)} to ${s.name} via ${d.mode}`); save(); closeM(); draw('suppliers'); toast('Payment recorded');
    });
  },
  costs: id => { const p = P(id); openForm('Costs per unit: ' + id, [['l', 'Labour cost per unit (₹)', 'number'], ['o', 'Overhead per unit (₹)', 'number']], { l: p.labour, o: p.oh }, d => { p.labour = d.l; p.oh = d.o; save(); closeM(); draw('costing'); toast('Costing updated'); }); },
  look: () => lookup($('#sc').value),
  cam: async () => {
    if (!('BarcodeDetector' in window)) return toast('This browser has no camera scanner. Use a USB scanner or type the code.', 'warn');
    try {
      STREAM = await navigator.mediaDevices.getUserMedia({ video: { facingMode: 'environment' } }); const v = $('#vd'), d = new BarcodeDetector(); v.hidden = false; v.srcObject = STREAM; await v.play();
      const tick = async () => { if (!STREAM) return; try { const r = await d.detect(v); if (r[0]) { stopCam(); $('#sc').value = r[0].rawValue; return lookup(r[0].rawValue); } } catch (e) { } requestAnimationFrame(tick); }; tick();
    } catch (e) { stopCam(); toast('Camera access was blocked. Allow the camera or type the code.', 'bad'); }
  }
});
function stopCam() { STREAM && STREAM.getTracks().forEach(t => t.stop()); STREAM = null; const v = $('#vd'); v && (v.hidden = true); }

/* Scanner page: camera, USB scanner (types + Enter) or manual entry */
function scan() {
  $('#view').innerHTML = `<div class="head"><div><h2>Scanner and QR tracking</h2><p class="mut">Scan a batch, material, invoice or PO label to see where it stands.</p></div></div>
  <div class="cg2"><div class="card"><video id="vd" playsinline muted hidden></video><div class="bar"><input id="sc" placeholder="Scan with a USB scanner or type a code, then press Enter" aria-label="Code"><button class="btn o" data-a="cam">Use camera</button><button class="btn p" data-a="look">Look up</button></div><div id="res"><p class="mut">Results appear here. Print QR labels from the Production and Raw Materials pages.</p></div></div>
  <div class="card"><h3>Recent scans</h3><div id="hist"></div></div></div>`;
  $('#sc').onkeydown = e => e.key == 'Enter' && lookup($('#sc').value); $('#sc').focus(); hist();
}
const hist = () => { const h = $('#hist'); if (h) h.innerHTML = DB.scans.slice(0, 8).map(s => `<div class="ls"><div><b>${esc(s.code.slice(0, 40))}</b><br><small class="mut">${new Date(s.t).toLocaleString('en-IN', { day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit' })}</small></div>${bd(s.type, s.type == 'No match' ? 'bad' : 'ok')}</div>`).join('') || '<p class="mut">No scans yet.</p>'; };
function lookup(raw) {
  const code = raw.trim(); if (!code) return; const id = (code.split('|')[1] || code).trim().toLowerCase(), eq = r => String(r.id).toLowerCase() == id;
  const b = DB.production.find(eq), m = DB.materials.find(eq), s = DB.sales.find(eq), po = DB.purchases.find(eq), pr = DB.products.find(x => x.name.toLowerCase() == id);
  const row = (k, v) => `<tr><td class="mut">${k}</td><td>${v}</td></tr>`; let type = 'No match', h = '';
  if (b) { type = 'Production batch'; const p = P(b.product), sold = DB.sales.filter(x => x.product == b.product && x.date >= b.date);
    h = row('Batch', b.id) + row('Product', esc(b.product)) + row('Made on', b.date) + row('Quantity', num(b.qty)) + row('Operator', esc(b.operator)) + row('Material used', `${+(b.qty * p.ratio).toFixed(1)} of ${esc(p.mat)}`) + row('Batch cost', inr(b.qty * unitCost(b.product))) + row('Stock now', num(closing(b.product)) + ' units') + row('Sold in', sold.map(x => x.id + ' (' + esc(x.customer) + ')').join(', ') || 'Not sold yet'); }
  else if (m) { type = 'Raw material'; h = row('Material', esc(m.name)) + row('ID', m.id) + row('Stock', `${num(m.stock)} ${m.unit}`) + row('Reorder level', num(m.reorder)) + row('Supplier', esc(m.supplier)) + row('Status', m.stock <= m.reorder ? bd('Low stock', 'bad') : bd('In stock', 'ok')); }
  else if (s) { type = 'Invoice'; h = row('Invoice', s.id) + row('Customer', esc(s.customer)) + row('Product', `${esc(s.product)} x ${s.qty}`) + row('Total', inr2(s.total)) + row('Payment', pay(s.status)); }
  else if (po) { type = 'Purchase order'; h = row('PO', po.id) + row('Supplier', esc(po.supplier)) + row('Material', `${esc(po.material)} x ${po.qty}`) + row('Total', inr2(po.total)) + row('Status', po.status); }
  else if (pr) { type = 'Product'; const f = fg().find(x => x.name == pr.name); h = row('Product', esc(pr.name)) + row('HSN', pr.hsn) + row('Closing stock', num(f.closing)) + row('Cost per unit', inr2(unitCost(pr.name))) + row('Selling rate', inr2(pr.rate)); }
  $('#res').innerHTML = h ? `<h3>${type}</h3><div class="tw"><table><tbody>${h}</tbody></table></div>` : `<p>No record found for "${esc(code)}". Check the label and try again.</p>`;
  DB.scans.unshift({ t: new Date().toISOString(), code, type }); DB.scans.length = Math.min(DB.scans.length, 20); save(); hist();
}

/* Excel import: matches columns by label or key, then recalculates totals */
$('#imp').onchange = async e => {
  const f = e.target.files[0]; e.target.value = ''; if (!f || noXL()) return;
  const m = M[route], wb = XLSX.read(await f.arrayBuffer()), rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]); let n = 0;
  rows.forEach(x => {
    const r = { ...(m.def ? m.def() : {}) }; let got = 0;
    m.f.forEach(([k, l, t, o]) => { let v = x[l] ?? x[k]; if (v === undefined) return; if (t == 'number' || (Array.isArray(o) && typeof o[0] == 'number')) v = +v || 0; r[k] = v; got++; });
    if (!got) return; if (!r.id || DB[route].some(z => z.id == r.id)) r.id = nextId(m);
    m.calc && m.calc(r); DB[route].unshift(r); n++;
  });
  if (!n) return toast('No usable rows found. Use the same column names as an export.', 'bad');
  log(`Imported ${n} ${m.t.toLowerCase()} rows from Excel`); save(); draw(route); toast(`Imported ${n} rows`);
};

/* ---------- Auth + boot ---------- */
function enter() {
  $('#login').hidden = true; $('.app').hidden = false; brand(); $('#who').textContent = `${CU.name}, ${CU.role}`;
  $('#nav').innerHTML = NAV.filter(n => can(n[0])).map(n => `<a href="#" data-a="nav" data-id="${n[0]}" data-r="${n[0]}"><svg viewBox="0 0 24 24"><path d="${n[2]}"/></svg>${n[1]}</a>`).join(''); nav('dashboard');
}
$('#lf').onsubmit = e => {
  e.preventDefault(); const u = DB.users.find(x => x.email.toLowerCase() == $('#em').value.trim().toLowerCase() && x.password == $('#pw').value), rm = $('#rm').checked;
  if (!u) return toast('Email or password is incorrect. Use Demo login to explore.', 'bad');
  CU = u; (rm ? localStorage : sessionStorage).setItem('skr_auth', u.id); rm ? localStorage.setItem('skr_em', u.email) : localStorage.removeItem('skr_em'); busy(700, enter);
};
$('#demo').onclick = () => { const u = DB.users.find(x => x.role == $('#dr').value); $('#em').value = u.email; $('#pw').value = u.password; $('#lf').requestSubmit(); };
$('#ff').onsubmit = e => { e.preventDefault(); toast('Reset link sent. Check your inbox.'); A.fp(); };
load(); theme(); $('#em').value = localStorage.getItem('skr_em') || '';
busy(900, () => { const u = DB.users.find(x => x.id == (localStorage.getItem('skr_auth') || sessionStorage.getItem('skr_auth'))); if (u) { CU = u; enter(); } else $('#login').hidden = false; });
