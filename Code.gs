/**
 * Calm Fits Backend — Google Apps Script
 * Password admin HANYA di Script Properties (key: ADMIN_PASSWORD = Bismillaah)
 * Jangan simpan password di frontend / GitHub Pages.
 *
 * Deploy: Web app → Execute as Me → Anyone
 * Setelah deploy, jalankan setupSheets() sekali.
 */

const SHEET_OUTFITS = 'Outfits';
const SHEET_SHOP = 'ShopProducts';
const SHEET_ORDERS = 'Orders';
const PROP_PASSWORD = 'ADMIN_PASSWORD';
const PROP_SESSIONS = 'SESSIONS';

function doGet(e) {
  return handle_(e, 'GET');
}

function doPost(e) {
  return handle_(e, 'POST');
}

function handle_(e, method) {
  const out = ContentService.createTextOutput().setMimeType(ContentService.MimeType.JSON);
  try {
    let body = {};
    if (method === 'POST' && e.postData && e.postData.contents) {
      try { body = JSON.parse(e.postData.contents); } catch (_) {}
    }
    const params = e.parameter || {};
    Object.keys(params).forEach(k => { if (body[k] === undefined) body[k] = params[k]; });
    const action = String(body.action || '').toLowerCase();

    let result;
    switch (action) {
      case 'setup':
        result = setupSheets();
        break;
      case 'get_all':
        result = { success: true, outfits: getOutfits_().data, shop: getShop_().data };
        break;
      case 'get_outfits':
        result = getOutfits_();
        break;
      case 'get_shop':
        result = getShop_();
        break;
      case 'login':
        result = login_(body.password);
        break;
      case 'create_order':
        result = createOrder_(body);
        break;
      // Admin (token required)
      case 'create_outfit':
        result = requireAuth_(body.token, () => createOutfit_(body));
        break;
      case 'update_outfit':
        result = requireAuth_(body.token, () => updateOutfit_(body));
        break;
      case 'delete_outfit':
        result = requireAuth_(body.token, () => deleteOutfit_(body.id));
        break;
      case 'create_shop':
        result = requireAuth_(body.token, () => createShop_(body));
        break;
      case 'update_shop':
        result = requireAuth_(body.token, () => updateShop_(body));
        break;
      case 'delete_shop':
        result = requireAuth_(body.token, () => deleteShop_(body.id));
        break;
      case 'get_orders':
        result = requireAuth_(body.token, () => getOrders_());
        break;
      case 'update_order':
        result = requireAuth_(body.token, () => updateOrder_(body));
        break;
      case 'delete_order':
        result = requireAuth_(body.token, () => deleteOrder_(body.id));
        break;
      default:
        result = { success: false, error: 'Unknown action: ' + action };
    }
    out.setContent(JSON.stringify(result));
  } catch (err) {
    out.setContent(JSON.stringify({ success: false, error: String(err.message || err) }));
  }
  return out;
}

/* ===================== SETUP (auto sheets + headers + seed) ===================== */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Outfits
  let sh = ss.getSheetByName(SHEET_OUTFITS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_OUTFITS);
    sh.appendRow(['id', 'name', 'category', 'img', 'color', 'active']);
    sh.getRange(1, 1, 1, 6).setFontWeight('bold');
    const seed = [
      [1, 'Kawaii', 'cute', 'images/kawaii.jpg', 'pink', 'yes'],
      [2, 'Korean Casual', 'cute', 'images/korean-casual.jpg', 'beige', 'yes'],
      [3, 'Y2K Cute', 'cute', 'images/y2k-cute.jpg', 'pink', 'yes'],
      [4, 'College Girl', 'cute', 'images/college-girl.jpg', 'beige', 'yes'],
      [5, 'Comfy Cute', 'cute', 'images/comfy-cute.jpg', 'beige', 'yes'],
      [6, 'Graphic Tee Outfit', 'cute', 'images/graphic-tee.jpg', 'black', 'yes'],
      [7, 'Preppy', 'cute', 'images/preppy.jpg', 'beige', 'yes'],
      [8, 'Soft Girl', 'cute', 'images/soft-girl.jpg', 'pink', 'yes'],
      [25, 'Cottagecore', 'pretty', 'images/cottagecore.jpg', 'green', 'yes'],
      [26, 'Elegant', 'pretty', 'images/elegant.jpg', 'beige', 'yes'],
      [27, 'Fairycore', 'pretty', 'images/fairycore.jpg', 'green', 'yes'],
      [28, 'Feminine', 'pretty', 'images/feminine.jpg', 'pink', 'yes'],
      [29, 'Balletcore', 'pretty', 'images/balletcore.jpg', 'pink', 'yes'],
      [30, 'Coquette', 'pretty', 'images/coquette.jpg', 'pink', 'yes'],
      [31, 'Princesscore', 'pretty', 'images/princesscore.jpg', 'pink', 'yes'],
      [32, 'Romantic', 'pretty', 'images/romantic.jpg', 'pink', 'yes'],
      [40, 'T-shirt & Baggy Jeans', 'comfy', 'images/tshirt-baggy.jpg', 'blue', 'yes'],
      [41, 'Oversized Hoodie', 'comfy', 'images/oversized-hoodie.jpg', 'beige', 'yes'],
      [42, 'Jogger Outfit', 'comfy', 'images/jogger-outfit.jpg', 'black', 'yes'],
      [43, 'Cozy Knitwear', 'comfy', 'images/cozy-knitwear.jpg', 'beige', 'yes'],
      [44, 'Streetwear', 'comfy', 'images/streetwear.jpg', 'blue', 'yes'],
      [45, 'Soft Girl Casual', 'comfy', 'images/soft-girl-casual.jpg', 'beige', 'yes'],
      [46, 'Fluffy Sweater', 'comfy', 'images/fluffy-sweater.jpg', 'green', 'yes'],
      [47, 'Minimalist Casual', 'comfy', 'images/minimalist-casual.jpg', 'blue', 'yes'],
      [50, 'Preppy Feminine', 'soft-girly', 'images/preppy-feminine.jpg', 'black', 'yes'],
      [51, 'Korean Feminine', 'soft-girly', 'images/korean-feminine.jpg', 'blue', 'yes'],
      [52, 'Angelcore', 'soft-girly', 'images/angelcore.jpg', 'white', 'yes'],
      [53, 'Princesscore Soft', 'soft-girly', 'images/princesscore-soft.jpg', 'beige', 'yes'],
      [54, 'Soft Romantic', 'soft-girly', 'images/soft-romantic.jpg', 'pink', 'yes'],
      [55, 'Vintage Feminine', 'soft-girly', 'images/vintage-feminine.jpg', 'brown', 'yes'],
      [56, 'Girly Casual', 'soft-girly', 'images/girly-casual.jpg', 'pink', 'yes'],
      [57, 'Dollcore', 'soft-girly', 'images/dollcore.jpg', 'red', 'yes'],
      [23, 'Soft Hijab Look', 'hijabi', 'https://images.unsplash.com/photo-1583391733956-3750e0ff4e8b?w=400&h=550&fit=crop', 'beige', 'yes'],
      [24, 'Elegant Hijabi', 'hijabi', 'https://images.unsplash.com/photo-1610030469983-98e550d6193c?w=400&h=550&fit=crop', 'beige', 'yes']
    ];
    sh.getRange(2, 1, seed.length, 6).setValues(seed);
  }

  // ShopProducts
  let sh2 = ss.getSheetByName(SHEET_SHOP);
  if (!sh2) {
    sh2 = ss.insertSheet(SHEET_SHOP);
    sh2.appendRow(['id', 'outfitId', 'name', 'price', 'img', 'active']);
    sh2.getRange(1, 1, 1, 6).setFontWeight('bold');
    const seedShop = [
      ['s1', 1, 'Kawaii', 189000, 'images/kawaii.jpg', 'yes'],
      ['s2', 2, 'Korean Casual', 165000, 'images/korean-casual.jpg', 'yes'],
      ['s3', 3, 'Y2K Cute', 210000, 'images/y2k-cute.jpg', 'yes'],
      ['s4', 5, 'Comfy Cute', 155000, 'images/comfy-cute.jpg', 'yes'],
      ['s5', 8, 'Soft Girl', 175000, 'images/soft-girl.jpg', 'yes'],
      ['s6', 25, 'Cottagecore', 245000, 'images/cottagecore.jpg', 'yes'],
      ['s7', 26, 'Elegant', 289000, 'images/elegant.jpg', 'yes'],
      ['s8', 27, 'Fairycore', 268000, 'images/fairycore.jpg', 'yes'],
      ['s9', 29, 'Balletcore', 198000, 'images/balletcore.jpg', 'yes'],
      ['s10', 30, 'Coquette', 225000, 'images/coquette.jpg', 'yes'],
      ['s11', 40, 'T-shirt & Baggy Jeans', 149000, 'images/tshirt-baggy.jpg', 'yes'],
      ['s12', 41, 'Oversized Hoodie', 179000, 'images/oversized-hoodie.jpg', 'yes'],
      ['s13', 42, 'Jogger Outfit', 199000, 'images/jogger-outfit.jpg', 'yes'],
      ['s14', 43, 'Cozy Knitwear', 215000, 'images/cozy-knitwear.jpg', 'yes'],
      ['s15', 50, 'Preppy Feminine', 259000, 'images/preppy-feminine.jpg', 'yes'],
      ['s16', 51, 'Korean Feminine', 235000, 'images/korean-feminine.jpg', 'yes'],
      ['s17', 52, 'Angelcore', 275000, 'images/angelcore.jpg', 'yes'],
      ['s18', 57, 'Dollcore', 249000, 'images/dollcore.jpg', 'yes'],
      ['s19', 46, 'Fluffy Sweater', 169000, 'images/fluffy-sweater.jpg', 'yes'],
      ['s20', 54, 'Soft Romantic', 229000, 'images/soft-romantic.jpg', 'yes']
    ];
    sh2.getRange(2, 1, seedShop.length, 6).setValues(seedShop);
  }

  // Orders
  let sh3 = ss.getSheetByName(SHEET_ORDERS);
  if (!sh3) {
    sh3 = ss.insertSheet(SHEET_ORDERS);
    sh3.appendRow([
      'id', 'date', 'customerName', 'wa', 'items', 'total', 'payment', 'status', 'notes'
    ]);
    sh3.getRange(1, 1, 1, 9).setFontWeight('bold');
    // status: pending | processed | done
  }

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty(PROP_PASSWORD)) {
    props.setProperty(PROP_PASSWORD, 'Bismillaah');
  }

  return { success: true, message: 'Sheets ready + seeded. Password in Script Properties.' };
}

/* ===================== PUBLIC ===================== */
function getOutfits_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_OUTFITS);
  if (!sh) return { success: false, error: 'Outfits missing — run setup' };
  const data = sh.getDataRange().getValues();
  const list = data.slice(1)
    .filter(r => String(r[5]).toLowerCase() !== 'no')
    .map(r => ({
      id: Number(r[0]),
      name: String(r[1]),
      category: String(r[2]),
      img: String(r[3]),
      color: String(r[4] || '')
    }));
  return { success: true, data: list };
}

function getShop_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SHOP);
  if (!sh) return { success: false, error: 'Shop missing — run setup' };
  const data = sh.getDataRange().getValues();
  const list = data.slice(1)
    .filter(r => String(r[5]).toLowerCase() !== 'no')
    .map(r => ({
      id: String(r[0]),
      outfitId: Number(r[1]),
      name: String(r[2]),
      price: Number(r[3]),
      img: String(r[4])
    }));
  return { success: true, data: list };
}

function createOrder_(body) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ORDERS);
  if (!sh) return { success: false, error: 'Orders sheet missing' };
  const name = String(body.customerName || '').trim();
  const wa = String(body.wa || '').trim();
  if (!name || !wa) return { success: false, error: 'Nama dan nomor WA wajib diisi' };

  const data = sh.getDataRange().getValues();
  const ids = data.slice(1).map(r => Number(r[0])).filter(n => !isNaN(n));
  const newId = ids.length ? Math.max(...ids) + 1 : 1;
  const date = new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' });
  const items = typeof body.items === 'string' ? body.items : JSON.stringify(body.items || []);
  const total = Number(body.total) || 0;
  const payment = String(body.payment || '');
  const notes = String(body.notes || '');

  sh.appendRow([newId, date, name, wa, items, total, payment, 'pending', notes]);
  return { success: true, id: newId, message: 'Pesanan berhasil dikirim' };
}

/* ===================== AUTH ===================== */
function login_(password) {
  const props = PropertiesService.getScriptProperties();
  const stored = props.getProperty(PROP_PASSWORD);
  if (!stored || String(password) !== String(stored)) {
    return { success: false, error: 'Password salah' };
  }
  const token = Utilities.getUuid();
  const sessions = JSON.parse(props.getProperty(PROP_SESSIONS) || '{}');
  const now = Date.now();
  Object.keys(sessions).forEach(k => { if (sessions[k] < now) delete sessions[k]; });
  sessions[token] = now + 2 * 60 * 60 * 1000; // 2 jam
  props.setProperty(PROP_SESSIONS, JSON.stringify(sessions));
  return { success: true, token: token };
}

function requireAuth_(token, fn) {
  if (!token) return { success: false, error: 'Unauthorized' };
  const props = PropertiesService.getScriptProperties();
  const sessions = JSON.parse(props.getProperty(PROP_SESSIONS) || '{}');
  if (!sessions[token] || sessions[token] < Date.now()) {
    return { success: false, error: 'Sesi habis. Login lagi.' };
  }
  return fn();
}

/* ===================== CRUD OUTFITS ===================== */
function createOutfit_(body) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_OUTFITS);
  const data = sh.getDataRange().getValues();
  const ids = data.slice(1).map(r => Number(r[0]));
  const newId = ids.length ? Math.max(...ids) + 1 : 1;
  sh.appendRow([newId, body.name || '', body.category || 'cute', body.img || '', body.color || '', 'yes']);
  return { success: true, id: newId };
}

function updateOutfit_(body) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_OUTFITS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][0]) === Number(body.id)) {
      if (body.name !== undefined) sh.getRange(i + 1, 2).setValue(body.name);
      if (body.category !== undefined) sh.getRange(i + 1, 3).setValue(body.category);
      if (body.img !== undefined) sh.getRange(i + 1, 4).setValue(body.img);
      if (body.color !== undefined) sh.getRange(i + 1, 5).setValue(body.color);
      if (body.active !== undefined) sh.getRange(i + 1, 6).setValue(body.active);
      return { success: true };
    }
  }
  return { success: false, error: 'Outfit tidak ditemukan' };
}

function deleteOutfit_(id) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_OUTFITS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][0]) === Number(id)) {
      sh.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Outfit tidak ditemukan' };
}

/* ===================== CRUD SHOP ===================== */
function createShop_(body) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SHOP);
  const data = sh.getDataRange().getValues();
  const ids = data.slice(1).map(r => String(r[0]));
  let newId = 's' + (ids.length + 1);
  let n = 1;
  while (ids.includes(newId)) { n++; newId = 's' + (ids.length + n); }
  sh.appendRow([newId, body.outfitId || 0, body.name || '', Number(body.price) || 0, body.img || '', 'yes']);
  return { success: true, id: newId };
}

function updateShop_(body) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SHOP);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(body.id)) {
      if (body.outfitId !== undefined) sh.getRange(i + 1, 2).setValue(body.outfitId);
      if (body.name !== undefined) sh.getRange(i + 1, 3).setValue(body.name);
      if (body.price !== undefined) sh.getRange(i + 1, 4).setValue(Number(body.price));
      if (body.img !== undefined) sh.getRange(i + 1, 5).setValue(body.img);
      if (body.active !== undefined) sh.getRange(i + 1, 6).setValue(body.active);
      return { success: true };
    }
  }
  return { success: false, error: 'Produk tidak ditemukan' };
}

function deleteShop_(id) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SHOP);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sh.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Produk tidak ditemukan' };
}

/* ===================== ORDERS ADMIN ===================== */
function getOrders_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ORDERS);
  if (!sh) return { success: false, error: 'Orders sheet missing' };
  const data = sh.getDataRange().getValues();
  const list = data.slice(1).map(r => ({
    id: Number(r[0]),
    date: String(r[1]),
    customerName: String(r[2]),
    wa: String(r[3]),
    items: String(r[4]),
    total: Number(r[5]),
    payment: String(r[6]),
    status: String(r[7] || 'pending'),
    notes: String(r[8] || '')
  })).reverse(); // newest first
  return { success: true, data: list };
}

function updateOrder_(body) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ORDERS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][0]) === Number(body.id)) {
      if (body.status !== undefined) sh.getRange(i + 1, 8).setValue(body.status);
      if (body.notes !== undefined) sh.getRange(i + 1, 9).setValue(body.notes);
      return { success: true };
    }
  }
  return { success: false, error: 'Order tidak ditemukan' };
}

function deleteOrder_(id) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ORDERS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][0]) === Number(id)) {
      sh.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Order tidak ditemukan' };
}
