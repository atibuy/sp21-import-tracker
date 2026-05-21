/**
 * ====================================================================
 *  Shipment Tracker — Google Apps Script backend
 * ====================================================================
 *
 *  HOW TO INSTALL
 *  --------------
 *  1) เปิด Google Apps Script: https://script.google.com/  →  New project
 *  2) ลบไฟล์ Code.gs ที่ขึ้นมา แล้ว paste ไฟล์นี้ทั้งหมดลงไป
 *  3) เมนู Run → เลือกฟังก์ชัน  initSheets  → กด Run
 *     (ครั้งแรกจะให้ Authorize เข้าถึง Google Sheets + Google Drive — กดอนุญาตทั้งคู่)
 *     ระบบจะสร้าง header row บน sheet "Shipments", "Items", "Costs",
 *     "Photos", "PhotoData"
 *  4) เมนู Deploy → New deployment
 *       Type: Web app
 *       Description: shipment-tracker
 *       Execute as: Me (ตัวเองที่เป็นเจ้าของ sheet)
 *       Who has access: Anyone   ← จำเป็น เพื่อให้ index.html เรียกได้
 *     กด Deploy → คัดลอก "Web app URL" (ลงท้ายด้วย /exec)
 *  5) เปิดไฟล์ index.html แล้วใส่ URL นี้ในตัวแปร CLOUD_CONFIG.webAppUrl
 *
 *  หากแก้ไขโค้ดในไฟล์นี้  ต้อง  Deploy → Manage deployments → ปุ่ม ✎ →
 *  เปลี่ยน "Version" เป็น New version → Deploy   ไม่งั้น URL เดิมจะรันโค้ดเก่า
 *
 *  อัปเกรดจากเวอร์ชันก่อน (base64-in-Sheets → Drive)
 *  -------------------------------------------------
 *  เวอร์ชันนี้เก็บรูปภาพ/PDF จริงใน Google Drive (โฟลเดอร์ SP21/<containerNo>/)
 *  แทนที่จะเก็บ base64 chunk ใน sheet "PhotoData".  หลัง deploy แล้ว
 *  - รูปที่อัปใหม่จะไปอยู่ใน Drive ทันที
 *  - รูปเก่าที่ยังเก็บเป็น chunk ใน "PhotoData" ยังอ่านได้ตามปกติ
 *    (ระบบ fallback ไป reconstruct dataUrl ให้)
 *  - ครั้งแรกที่ถูกเรียก uploadPhoto, Apps Script จะขอ scope ใหม่ของ Drive
 *    ไปที่ Apps Script editor → Run → เลือก authorizeDriveScope → กด Run
 *    เพื่อ accept scope ล่วงหน้า (ไม่งั้น web-app call แรกจะ error)
 *
 *  หมายเหตุ
 *  --------
 *  • รูปภาพเก็บใน Google Drive: My Drive / SP21 / <containerNo> /
 *    ไฟล์ตั้ง sharing เป็น ANYONE_WITH_LINK + VIEW เพื่อให้ <img> ใน browser
 *    โหลดผ่าน lh3.googleusercontent.com ได้โดยไม่ต้องล็อกอิน
 *  • sheet "Photos" เก็บ metadata + Drive URLs;  "PhotoData" คงไว้สำหรับ
 *    legacy chunks ที่อัปไว้สมัยก่อน — ไม่ใช้เขียนใหม่อีก
 *  • ตาราง Costs เป็น snapshot ของต้นทุนล่าสุดที่คำนวณจาก client
 *    หากแก้ Step ใด ๆ ที่กระทบต้นทุน  row นี้จะถูกเขียนทับ
 */

const SPREADSHEET_ID = '11yVFC71onKOKSsQbDiTbQDCvyY12s067ZMVmS7VFBlE';

const SHEET_SHIPMENTS  = 'Shipments';
const SHEET_ITEMS      = 'Items';
const SHEET_COSTS      = 'Costs';
const SHEET_PHOTOS     = 'Photos';
const SHEET_PHOTO_DATA = 'PhotoData';
const SHEET_CARDS      = 'Cards';

// โฟลเดอร์ราก ใน My Drive ของบัญชีที่ deploy script
// โครงสร้าง:
//   My Drive / SP21 / <containerNo> / <step>_<category>_<ts>_<rand>.<ext>
//   My Drive / SP21 / Cards / <cardId>_<category>_<ts>_<rand>.<ext>
const DRIVE_ROOT_FOLDER = 'SP21';
const DRIVE_CARDS_SUBFOLDER = 'Cards';

// IMPORTANT: append-only — เพิ่ม column ใหม่ได้ที่ "ท้าย array" เท่านั้น
// เพื่อให้ index ตรงกับ sheet ที่มีอยู่ (data เก่าจะอยู่ตำแหน่งเดิม)
// raw_json ปล่อยให้อยู่ตรงนั้น  ไม่ใช่ตัวสุดท้ายแล้วก็ยังใช้ได้ (อ้างถึงด้วยชื่อ)
const SHIPMENT_COLS = [
  'id', 'createdAt', 'updatedAt', 'allocMethod',
  'confirmed', 'confirmedAt',
  'stepA_completed', 'stepA_orderDate', 'stepA_exchangeRate',
  'stepB_completed', 'stepB_shipDate', 'stepB_photoCount',
  'stepC_completed', 'stepC_boxes', 'stepC_volume', 'stepC_weight', 'stepC_carrier',
  'stepD_completed', 'stepD_arrivalDate', 'stepD_shippingCost',
  'stepD_customs', 'stepD_vat', 'stepD_tpiToNim',
  'stepE_completed', 'stepE_nimExpressCost', 'stepE_photoCount',
  'raw_json',
  // ===== fields ใหม่ — ห้ามแทรกกลาง  ต่อท้ายเท่านั้น =====
  'containerNo',
  'stepC_transportMode',
  'stepD_arrivalGroupSize',
  'stepE_smallBoxes', 'stepE_largeBoxes',
  'stepE_smallRate', 'stepE_largeRate'
];

const ITEM_COLS = [
  'shipmentId', 'idx', 'name', 'sku',
  'boxes', 'perBox', 'yuanAmount', 'weight', 'volume',
  // ===== fields ใหม่ — ต่อท้ายเท่านั้น =====
  'factory'
];

const COST_COLS = [
  'shipmentId', 'computedAt',
  'totalYuan', 'rate', 'totalProduct',
  'shippingCost', 'customs', 'vat', 'tpiToNim', 'nim',
  'totalExtra', 'grandTotal',
  'totalBoxes', 'totalWeight', 'totalVolume', 'totalPieces',
  'allocMethod', 'effectiveAlloc'
];

// แต่ละ row = รูป 1 ใบ ที่ผูกกับ shipment+step+category (metadata เท่านั้น)
// url/viewUrl/thumbUrl ปล่อยว่าง — base64 จริงอยู่ใน PhotoData
const PHOTO_COLS = [
  'shipmentId', 'step', 'category', 'idx',
  'fileId', 'url', 'viewUrl', 'thumbUrl', 'name', 'uploadedAt',
  // ===== fields ใหม่ — ต่อท้ายเท่านั้น =====
  'mimeType'
];

// แต่ละ row = chunk หนึ่งก้อนของ base64 — รูปใหญ่จะมีหลาย row ต่อ fileId
const PHOTO_DATA_COLS = [
  'fileId', 'chunkIdx', 'totalChunks', 'dataChunk'
];

// แต่ละ row = Card-purchase 1 รายการ — 1 การซื้อมีได้ทั้ง รายปี + ตลอดชีพ พร้อมกัน
// photos_json เก็บ array ของ {fileId,url,viewUrl,thumbUrl,name,uploadedAt,mimeType,category}
// IMPORTANT: append-only — คอลัมน์เก่า cardType/cardCount/pricePerCard ยังอยู่
// เพื่อ backward-compat กับข้อมูลเก่า  ไม่ใช้แล้วในการเขียน
const CARD_COLS = [
  'id', 'createdAt', 'updatedAt',
  'payDate', 'cardType', 'cardCount', 'pricePerCard',
  'photos_json',
  // ===== fields ใหม่ — ต่อท้ายเท่านั้น =====
  'annual_count', 'annual_price', 'annual_free',
  'lifetime_count', 'lifetime_price', 'lifetime_free',
  // annual_price/lifetime_price ตอนนี้คือ "ราคาต่อใบ (¥)" — ต้องคูณ exchange_rate ถึงได้บาท
  // exchange_rate = 0/null สำหรับการ์ดเก่า (ราคาต่อใบเป็นบาทอยู่แล้ว) — frontend treat 0 → 1
  'exchange_rate'
];

/* ================ Web app entry points ================ */

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'list';
  return handle({ action: action });
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json({ ok: false, error: 'Invalid JSON: ' + err.message });
  }
  return handle(req);
}

function handle(req) {
  try {
    if (req.action === 'list')            return json({ ok: true, data: listShipments() });
    if (req.action === 'save')            return json({ ok: true, data: upsertShipment(req.payload) });
    if (req.action === 'delete')          return json({ ok: true, data: deleteShipmentById(req.id) });
    if (req.action === 'init')            return json({ ok: true, data: initSheets() });
    if (req.action === 'ping')            return json({ ok: true, data: 'pong' });
    if (req.action === 'uploadPhoto')     return json({ ok: true, data: uploadPhoto(req.payload) });
    if (req.action === 'deletePhoto')     return json({ ok: true, data: deletePhoto(req.payload) });
    if (req.action === 'listCards')       return json({ ok: true, data: listCards() });
    if (req.action === 'saveCard')        return json({ ok: true, data: upsertCard(req.payload) });
    if (req.action === 'deleteCard')      return json({ ok: true, data: deleteCardById(req.id) });
    if (req.action === 'uploadCardPhoto') return json({ ok: true, data: uploadCardPhoto(req.payload) });
    return json({ ok: false, error: 'Unknown action: ' + req.action });
  } catch (err) {
    return json({ ok: false, error: String(err && err.message || err), stack: String(err && err.stack || '') });
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ================ Sheet helpers ================ */

function ss() { return SpreadsheetApp.openById(SPREADSHEET_ID); }

function getOrCreateSheet(name, headers) {
  let sheet = ss().getSheetByName(name);
  if (!sheet) {
    sheet = ss().insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  } else {
    // auto-migrate: เติม header ที่ขาดต่อท้าย (เรียงตาม headers[])
    const lastCol = Math.max(1, sheet.getLastColumn());
    const current = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (v) { return String(v); });
    const missing = headers.filter(function (h) { return current.indexOf(h) === -1; });
    if (missing.length > 0) {
      sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
    }
  }
  return sheet;
}

function initSheets() {
  getOrCreateSheet(SHEET_SHIPMENTS, SHIPMENT_COLS);
  getOrCreateSheet(SHEET_ITEMS, ITEM_COLS);
  getOrCreateSheet(SHEET_COSTS, COST_COLS);
  getOrCreateSheet(SHEET_PHOTOS, PHOTO_COLS);
  getOrCreateSheet(SHEET_PHOTO_DATA, PHOTO_DATA_COLS);
  getOrCreateSheet(SHEET_CARDS, CARD_COLS);
  return 'Sheets initialized: ' + [SHEET_SHIPMENTS, SHEET_ITEMS, SHEET_COSTS, SHEET_PHOTOS, SHEET_PHOTO_DATA, SHEET_CARDS].join(', ');
}

/* ================ Photo storage (Google Drive) ================ */

// เรียกฟังก์ชันนี้ครั้งแรกหลัง paste โค้ดใหม่ (Run → authorizeDriveScope)
// จะ trigger consent screen ให้ขอ scope ของ Drive ก่อน — ไม่งั้น web-app call แรก
// จะ fail ด้วย "Authorization is required to perform that action"
function authorizeDriveScope() {
  const folder = getOrCreateChildFolder(DriveApp.getRootFolder(), DRIVE_ROOT_FOLDER);
  return 'Drive scope OK. Root folder id: ' + folder.getId();
}

function getOrCreateChildFolder(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function getShipmentFolder(containerNo) {
  const root = getOrCreateChildFolder(DriveApp.getRootFolder(), DRIVE_ROOT_FOLDER);
  return getOrCreateChildFolder(root, sanitizeFolderName(containerNo));
}

// containerNo มาจาก user input — กันอักษร reserved ของ filesystem ไว้ก่อน
function sanitizeFolderName(name) {
  return String(name || '').replace(/[\\/:*?"<>|]/g, '_').trim() || 'unknown';
}

// สร้าง URL ที่ <img src> ใน browser โหลดได้ — ต้องการให้ไฟล์ถูก share แบบ ANYONE_WITH_LINK
// `lh3.googleusercontent.com/d/<id>=w<width>` ใช้งานได้ดี + รองรับ referrerpolicy=no-referrer
function buildDriveImageUrl(fileId, width) {
  return 'https://lh3.googleusercontent.com/d/' + fileId + '=w' + width;
}
function buildDriveViewUrl(fileId) {
  return 'https://drive.google.com/file/d/' + fileId + '/view';
}

function uploadPhoto(payload) {
  if (!payload || !payload.shipmentId) throw new Error('shipmentId required');
  if (!payload.dataUrl) throw new Error('dataUrl required');
  const m = String(payload.dataUrl).match(/^data:([^;]+);base64,(.+)$/);
  if (!m) throw new Error('Invalid dataUrl');
  const mime = m[1];
  const b64 = m[2];

  const containerNo = String(payload.containerNo || payload.shipmentId || '').trim();
  if (!containerNo) throw new Error('containerNo required');
  const step = String(payload.step || 'X').toUpperCase().replace(/[^A-Z0-9]/g, '') || 'X';
  const category = String(payload.category || 'general').replace(/[^A-Za-z0-9_-]/g, '_') || 'general';

  const ext = (mime.split('/')[1] || 'jpg').replace('jpeg', 'jpg');
  const ts = Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMdd_HHmmss');
  const rand = Math.random().toString(36).slice(2, 6);
  const safeName = step + '_' + category + '_' + ts + '_' + rand + '.' + ext;

  const folder = getShipmentFolder(containerNo);
  const bytes = Utilities.base64Decode(b64);
  const blob = Utilities.newBlob(bytes, mime, safeName);
  const file = folder.createFile(blob);

  // ANYONE_WITH_LINK + VIEW — ถ้า org policy ห้ามจะ throw และลบไฟล์ทิ้งให้
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    try { file.setTrashed(true); } catch (_) {}
    throw new Error('ตั้งสิทธิ์แชร์ไฟล์ไม่ได้ (org/domain policy ห้าม share สาธารณะ?): ' + (e && e.message || e));
  }

  const fileId = file.getId();
  return {
    fileId: fileId,
    url:      buildDriveImageUrl(fileId, 2000),
    viewUrl:  buildDriveViewUrl(fileId),
    thumbUrl: buildDriveImageUrl(fileId, 400),
    name: safeName,
    uploadedAt: new Date().toISOString(),
    mimeType: mime
  };
}

// legacy: เคยตัด base64 เป็น chunk เก็บใน "PhotoData" — เก็บฟังก์ชันลบไว้ใช้ตอน delete
function deletePhotoChunks(fileId) {
  const sheet = ss().getSheetByName(SHEET_PHOTO_DATA);
  if (!sheet || sheet.getLastRow() < 2) return;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (let i = rows.length - 1; i >= 0; i--) {
    if (String(rows[i][0]) === String(fileId)) sheet.deleteRow(i + 2);
  }
}

function deletePhoto(payload) {
  if (!payload || !payload.fileId) throw new Error('fileId required');
  const fileId = String(payload.fileId);
  // 1) Drive file (รูปใหม่)
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
  } catch (e) {
    // รูป legacy (ไม่มีจริงใน Drive) — ข้าม
  }
  // 2) Legacy chunks ใน PhotoData (ถ้ามี)
  deletePhotoChunks(fileId);
  return fileId;
}

/* ================ List ================ */

function listShipments() {
  const sheet = getOrCreateSheet(SHEET_SHIPMENTS, SHIPMENT_COLS);
  if (sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, SHIPMENT_COLS.length).getValues();
  const rawIdx = SHIPMENT_COLS.indexOf('raw_json');
  const idIdx = SHIPMENT_COLS.indexOf('id');
  const photosBySid = loadAllPhotosByShipment();
  const out = [];
  for (const row of rows) {
    const sid = String(row[idIdx] || '');
    const raw = row[rawIdx];
    let s = null;
    if (raw) {
      try { s = JSON.parse(raw); } catch (e) { s = null; }
    }
    if (!s) s = rowToShipment(row);
    mergePhotosIntoShipment(s, photosBySid[sid] || {});
    out.push(s);
  }
  return out;
}

function loadAllPhotosByShipment() {
  const photoSheet = getOrCreateSheet(SHEET_PHOTOS, PHOTO_COLS);
  const dataSheet  = ss().getSheetByName(SHEET_PHOTO_DATA);

  // 1) อ่าน legacy chunks (ถ้ามี sheet PhotoData) — ใช้ fallback สำหรับรูปเก่า
  const chunksByFileId = {};
  if (dataSheet && dataSheet.getLastRow() >= 2) {
    const rows = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, PHOTO_DATA_COLS.length).getValues();
    rows.forEach(function (r) {
      const fid = String(r[0] || '');
      if (!fid) return;
      (chunksByFileId[fid] = chunksByFileId[fid] || []).push({ idx: Number(r[1]) || 0, data: String(r[3] || '') });
    });
    Object.keys(chunksByFileId).forEach(function (fid) {
      chunksByFileId[fid].sort(function (a, b) { return a.idx - b.idx; });
    });
  }

  // 2) อ่าน metadata Photos
  const out = {};
  if (photoSheet.getLastRow() < 2) return out;
  const rows = photoSheet.getRange(2, 1, photoSheet.getLastRow() - 1, PHOTO_COLS.length).getValues();
  const ix = function (name) { return PHOTO_COLS.indexOf(name); };
  rows
    .slice()
    .sort(function (a, b) { return (Number(a[ix('idx')]) || 0) - (Number(b[ix('idx')]) || 0); })
    .forEach(function (r) {
      const sid = String(r[ix('shipmentId')] || '');
      const step = String(r[ix('step')] || '').toUpperCase();
      if (!sid || !step) return;
      const fileId = String(r[ix('fileId')] || '');
      const mime = String(r[ix('mimeType')] || 'image/jpeg');
      const savedUrl   = String(r[ix('url')] || '');
      const savedView  = String(r[ix('viewUrl')] || '');
      const savedThumb = String(r[ix('thumbUrl')] || '');

      let url = savedUrl, viewUrl = savedView || savedUrl, thumbUrl = savedThumb || savedUrl;
      if (!url) {
        // legacy: ไม่มี URL ใน sheet — ลอง reconstruct จาก chunks
        const chunks = chunksByFileId[fileId];
        if (chunks && chunks.length) {
          const dataUrl = 'data:' + mime + ';base64,' + chunks.map(function (c) { return c.data; }).join('');
          url = viewUrl = thumbUrl = dataUrl;
        } else if (fileId) {
          // ไม่มีทั้ง URL ทั้ง chunks — เดาว่าเป็นไฟล์ Drive ที่ยังไม่ได้บันทึก URL กลับลง sheet
          url      = buildDriveImageUrl(fileId, 2000);
          viewUrl  = buildDriveViewUrl(fileId);
          thumbUrl = buildDriveImageUrl(fileId, 400);
        }
      }

      const photo = {
        category: r[ix('category')] || 'general',
        fileId: fileId,
        url:      url,
        viewUrl:  viewUrl,
        thumbUrl: thumbUrl,
        name: r[ix('name')] || '',
        uploadedAt: r[ix('uploadedAt')] || '',
        mimeType: mime
      };
      if (!out[sid]) out[sid] = {};
      (out[sid][step] = out[sid][step] || []).push(photo);
    });
  return out;
}

function mergePhotosIntoShipment(s, photoByStep) {
  ['A', 'B', 'C', 'D', 'E'].forEach(function (step) {
    const list = photoByStep[step] || [];
    if (list.length === 0) return;
    const k = 'step' + step;
    if (!s[k]) s[k] = { photos: list, completed: false };
    else s[k].photos = list;
  });
}

function rowToShipment(row) {
  const get = name => row[SHIPMENT_COLS.indexOf(name)];
  const toBool = v => v === true || v === 'TRUE' || v === 'true';
  const toIso = v => v instanceof Date ? v.toISOString() : (v || null);
  const s = {
    id: get('id'),
    containerNo: get('containerNo') || get('id') || null,
    createdAt: toIso(get('createdAt')),
    updatedAt: toIso(get('updatedAt')),
    allocMethod: get('allocMethod') || null,
    confirmed: toBool(get('confirmed')),
    confirmedAt: toIso(get('confirmedAt')),
    stepA: null, stepB: null, stepC: null, stepD: null, stepE: null
  };
  if (toBool(get('stepA_completed'))) {
    s.stepA = {
      orderDate: toIso(get('stepA_orderDate')),
      exchangeRate: Number(get('stepA_exchangeRate')) || 0,
      items: getItemsFor(s.id),
      photos: [],
      completed: true
    };
  }
  if (toBool(get('stepB_completed'))) {
    s.stepB = {
      shipDate: toIso(get('stepB_shipDate')),
      photos: [],
      completed: true
    };
  }
  if (toBool(get('stepC_completed'))) {
    s.stepC = {
      boxes: Number(get('stepC_boxes')) || 0,
      volume: Number(get('stepC_volume')) || 0,
      weight: Number(get('stepC_weight')) || 0,
      carrier: get('stepC_carrier') || '',
      transportMode: get('stepC_transportMode') || null,
      photos: [],
      completed: true
    };
  }
  if (toBool(get('stepD_completed'))) {
    s.stepD = {
      arrivalDate: toIso(get('stepD_arrivalDate')),
      shippingCost: Number(get('stepD_shippingCost')) || 0,
      customs: Number(get('stepD_customs')) || 0,
      vat: Number(get('stepD_vat')) || 0,
      tpiToNim: Number(get('stepD_tpiToNim')) || 0,
      arrivalGroupSize: Number(get('stepD_arrivalGroupSize')) || 1,
      photos: [],
      completed: true
    };
  }
  if (toBool(get('stepE_completed'))) {
    s.stepE = {
      nimExpressCost: Number(get('stepE_nimExpressCost')) || 0,
      smallBoxes: Number(get('stepE_smallBoxes')) || 0,
      largeBoxes: Number(get('stepE_largeBoxes')) || 0,
      smallRate: Number(get('stepE_smallRate')) || 63,
      largeRate: Number(get('stepE_largeRate')) || 100,
      photos: [],
      completed: true
    };
  }
  return s;
}

function getItemsFor(shipmentId) {
  const sheet = getOrCreateSheet(SHEET_ITEMS, ITEM_COLS);
  if (sheet.getLastRow() < 2) return [];
  const factoryIdx = ITEM_COLS.indexOf('factory');
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, ITEM_COLS.length).getValues();
  return rows
    .filter(r => String(r[0]) === String(shipmentId))
    .sort((a, b) => (Number(a[1]) || 0) - (Number(b[1]) || 0))
    .map(r => ({
      name: r[2] || '', sku: r[3] || '',
      boxes: Number(r[4]) || 0,
      perBox: Number(r[5]) || 0,
      yuanAmount: Number(r[6]) || 0,
      weight: Number(r[7]) || 0,
      volume: Number(r[8]) || 0,
      factory: r[factoryIdx] || ''
    }));
}

/* ================ Upsert ================ */

function upsertShipment(payload) {
  if (!payload || !payload.shipment) throw new Error('payload.shipment is required');
  const s = payload.shipment;
  if (!s.id) throw new Error('shipment.id is required');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getOrCreateSheet(SHEET_SHIPMENTS, SHIPMENT_COLS);
    const row = shipmentToRow(s);
    const lastRow = sheet.getLastRow();
    let foundIdx = -1;
    if (lastRow >= 2) {
      const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(s.id)) { foundIdx = i; break; }
      }
    }
    if (foundIdx === -1) {
      sheet.getRange(sheet.getLastRow() + 1, 1, 1, SHIPMENT_COLS.length).setValues([row]);
    } else {
      sheet.getRange(foundIdx + 2, 1, 1, SHIPMENT_COLS.length).setValues([row]);
    }
    replaceItemsFor(s.id, (s.stepA && s.stepA.items) || []);
    replacePhotosFor(s.id, collectShipmentPhotos(s));
    if (payload.costSnapshot) replaceCostFor(s.id, payload.costSnapshot);
    return s.id;
  } finally {
    lock.releaseLock();
  }
}

function shipmentToRow(s) {
  const row = new Array(SHIPMENT_COLS.length).fill('');
  const set = (name, val) => {
    const idx = SHIPMENT_COLS.indexOf(name);
    if (idx >= 0) row[idx] = val == null ? '' : val;
  };
  set('id', s.id);
  set('containerNo', s.containerNo || s.id || '');
  set('createdAt', s.createdAt || '');
  set('updatedAt', s.updatedAt || '');
  set('allocMethod', s.allocMethod || '');
  set('confirmed', !!s.confirmed);
  set('confirmedAt', s.confirmedAt || '');

  if (s.stepA) {
    set('stepA_completed', true);
    set('stepA_orderDate', s.stepA.orderDate || '');
    set('stepA_exchangeRate', Number(s.stepA.exchangeRate) || 0);
  }
  if (s.stepB) {
    set('stepB_completed', true);
    set('stepB_shipDate', s.stepB.shipDate || '');
    set('stepB_photoCount', (s.stepB.photos || []).length);
  }
  if (s.stepC) {
    set('stepC_completed', true);
    set('stepC_boxes', Number(s.stepC.boxes) || 0);
    set('stepC_volume', Number(s.stepC.volume) || 0);
    set('stepC_weight', Number(s.stepC.weight) || 0);
    set('stepC_carrier', s.stepC.carrier || '');
    set('stepC_transportMode', s.stepC.transportMode || '');
  }
  if (s.stepD) {
    set('stepD_completed', true);
    set('stepD_arrivalDate', s.stepD.arrivalDate || '');
    set('stepD_shippingCost', Number(s.stepD.shippingCost) || 0);
    set('stepD_customs', Number(s.stepD.customs) || 0);
    set('stepD_vat', Number(s.stepD.vat) || 0);
    set('stepD_tpiToNim', Number(s.stepD.tpiToNim) || 0);
    set('stepD_arrivalGroupSize', Number(s.stepD.arrivalGroupSize) || 1);
  }
  if (s.stepE) {
    set('stepE_completed', true);
    set('stepE_nimExpressCost', Number(s.stepE.nimExpressCost) || 0);
    set('stepE_photoCount', (s.stepE.photos || []).length);
    set('stepE_smallBoxes', Number(s.stepE.smallBoxes) || 0);
    set('stepE_largeBoxes', Number(s.stepE.largeBoxes) || 0);
    set('stepE_smallRate', Number(s.stepE.smallRate) || 0);
    set('stepE_largeRate', Number(s.stepE.largeRate) || 0);
  }

  // raw_json — backup ของทุกฟิลด์  ตัด photos ทุก step ออก เพราะ base64 จะทำให้ cell บวมเกิน limit
  const lean = deepClone(s);
  ['A', 'B', 'C', 'D', 'E'].forEach(function (st) {
    const k = 'step' + st;
    if (lean[k] && lean[k].photos) lean[k].photos = [];
  });
  const rawStr = JSON.stringify(lean);
  // safety: ถ้ายังยาวเกิน 49,000 chars ก็ตัดทิ้ง  (ไม่น่าเกิดถ้าตัด photos แล้ว)
  set('raw_json', rawStr.length > 49000 ? '' : rawStr);
  return row;
}

function deepClone(obj) {
  return JSON.parse(JSON.stringify(obj));
}

function replaceItemsFor(shipmentId, items) {
  const sheet = getOrCreateSheet(SHEET_ITEMS, ITEM_COLS);
  if (sheet.getLastRow() >= 2) {
    const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (let i = ids.length - 1; i >= 0; i--) {
      if (String(ids[i][0]) === String(shipmentId)) sheet.deleteRow(i + 2);
    }
  }
  if (!items || items.length === 0) return;
  const rows = items.map((it, i) => {
    const row = new Array(ITEM_COLS.length).fill('');
    const set = (name, val) => {
      const idx = ITEM_COLS.indexOf(name);
      if (idx >= 0) row[idx] = val == null ? '' : val;
    };
    set('shipmentId', shipmentId);
    set('idx', i);
    set('name', it.name || '');
    set('sku', it.sku || '');
    set('boxes', Number(it.boxes) || 0);
    set('perBox', Number(it.perBox) || 0);
    set('yuanAmount', Number(it.yuanAmount) || 0);
    set('weight', Number(it.weight) || 0);
    set('volume', Number(it.volume) || 0);
    set('factory', it.factory || '');
    return row;
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, ITEM_COLS.length).setValues(rows);
}

function replaceCostFor(shipmentId, c) {
  const sheet = getOrCreateSheet(SHEET_COSTS, COST_COLS);
  if (sheet.getLastRow() >= 2) {
    const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (let i = ids.length - 1; i >= 0; i--) {
      if (String(ids[i][0]) === String(shipmentId)) sheet.deleteRow(i + 2);
    }
  }
  const row = [
    shipmentId, new Date().toISOString(),
    Number(c.totalYuan) || 0, Number(c.rate) || 0, Number(c.totalProduct) || 0,
    Number(c.shippingCost) || 0, Number(c.customs) || 0, Number(c.vat) || 0,
    Number(c.tpiToNim) || 0, Number(c.nim) || 0,
    Number(c.totalExtra) || 0, Number(c.grandTotal) || 0,
    Number(c.totalBoxes) || 0, Number(c.totalWeight) || 0,
    Number(c.totalVolume) || 0, Number(c.totalPieces) || 0,
    c.allocMethod || '', c.effectiveAlloc || ''
  ];
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, COST_COLS.length).setValues([row]);
}

/* ================ Photos sheet (metadata) ================ */

function collectShipmentPhotos(s) {
  const all = [];
  ['A', 'B', 'C', 'D', 'E'].forEach(function (step) {
    const stepObj = s['step' + step];
    const photos = stepObj && stepObj.photos;
    if (!Array.isArray(photos)) return;
    photos.forEach(function (p) {
      if (!p || !p.fileId) return;  // เก็บเฉพาะรูปที่อัปแล้ว (มี fileId)
      all.push({
        step: step,
        category: p.category || 'general',
        fileId: p.fileId,
        url:      p.url || '',
        viewUrl:  p.viewUrl || '',
        thumbUrl: p.thumbUrl || '',
        name: p.name || '',
        uploadedAt: p.uploadedAt || '',
        mimeType: p.mimeType || ''
      });
    });
  });
  return all;
}

function replacePhotosFor(shipmentId, photos) {
  const sheet = getOrCreateSheet(SHEET_PHOTOS, PHOTO_COLS);
  if (sheet.getLastRow() >= 2) {
    const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (let i = ids.length - 1; i >= 0; i--) {
      if (String(ids[i][0]) === String(shipmentId)) sheet.deleteRow(i + 2);
    }
  }
  if (!photos || photos.length === 0) return;
  const rows = photos.map(function (p, i) {
    const row = new Array(PHOTO_COLS.length).fill('');
    const set = function (name, val) {
      const idx = PHOTO_COLS.indexOf(name);
      if (idx >= 0) row[idx] = val == null ? '' : val;
    };
    set('shipmentId', shipmentId);
    set('step', p.step || '');
    set('category', p.category || '');
    set('idx', i);
    set('fileId', p.fileId || '');
    set('url',      p.url || '');
    set('viewUrl',  p.viewUrl || '');
    set('thumbUrl', p.thumbUrl || '');
    set('name', p.name || '');
    set('uploadedAt', p.uploadedAt || '');
    set('mimeType', p.mimeType || '');
    return row;
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, PHOTO_COLS.length).setValues(rows);
}

/* ================ Cards (Card-service purchases) ================ */

function getCardsFolder() {
  const root = getOrCreateChildFolder(DriveApp.getRootFolder(), DRIVE_ROOT_FOLDER);
  return getOrCreateChildFolder(root, DRIVE_CARDS_SUBFOLDER);
}

function uploadCardPhoto(payload) {
  if (!payload || !payload.cardId) throw new Error('cardId required');
  if (!payload.dataUrl) throw new Error('dataUrl required');
  const m = String(payload.dataUrl).match(/^data:([^;]+);base64,(.+)$/);
  if (!m) throw new Error('Invalid dataUrl');
  const mime = m[1];
  const b64 = m[2];

  const cardId = sanitizeFolderName(payload.cardId);
  const category = String(payload.category || 'general').replace(/[^A-Za-z0-9_-]/g, '_') || 'general';
  const ext = (mime.split('/')[1] || 'jpg').replace('jpeg', 'jpg');
  const ts = Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMdd_HHmmss');
  const rand = Math.random().toString(36).slice(2, 6);
  const safeName = cardId + '_' + category + '_' + ts + '_' + rand + '.' + ext;

  const folder = getCardsFolder();
  const bytes = Utilities.base64Decode(b64);
  const blob = Utilities.newBlob(bytes, mime, safeName);
  const file = folder.createFile(blob);

  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    try { file.setTrashed(true); } catch (_) {}
    throw new Error('ตั้งสิทธิ์แชร์ไฟล์ไม่ได้ (org/domain policy ห้าม share สาธารณะ?): ' + (e && e.message || e));
  }

  const fileId = file.getId();
  return {
    fileId: fileId,
    url:      buildDriveImageUrl(fileId, 2000),
    viewUrl:  buildDriveViewUrl(fileId),
    thumbUrl: buildDriveImageUrl(fileId, 400),
    name: safeName,
    uploadedAt: new Date().toISOString(),
    mimeType: mime
  };
}

function listCards() {
  const sheet = getOrCreateSheet(SHEET_CARDS, CARD_COLS);
  if (sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, CARD_COLS.length).getValues();
  const ix = function (name) { return CARD_COLS.indexOf(name); };
  const toIso = function (v) { return v instanceof Date ? v.toISOString() : (v || null); };
  const num = function (v) { return Number(v) || 0; };
  return rows.map(function (r) {
    let photos = [];
    const raw = r[ix('photos_json')];
    if (raw) { try { photos = JSON.parse(raw) || []; } catch (e) { photos = []; } }

    // ค่าใหม่ (annual_/lifetime_) — ถ้ามีคอลัมน์ใดมีค่า > 0 ถือเป็น new shape
    let annual   = { count: num(r[ix('annual_count')]),   pricePerCard: num(r[ix('annual_price')]),   freeCount: num(r[ix('annual_free')])   };
    let lifetime = { count: num(r[ix('lifetime_count')]), pricePerCard: num(r[ix('lifetime_price')]), freeCount: num(r[ix('lifetime_free')]) };

    // legacy fallback — ถ้า new cols ว่างหมดแต่มี cardType/cardCount เก่า → กระจายไปยังประเภทที่ตรง
    if (annual.count === 0 && lifetime.count === 0 && annual.pricePerCard === 0 && lifetime.pricePerCard === 0) {
      const legacyType = String(r[ix('cardType')] || '');
      const legacyCount = num(r[ix('cardCount')]);
      const legacyPrice = num(r[ix('pricePerCard')]);
      if (legacyCount > 0 || legacyPrice > 0) {
        const seed = { count: legacyCount, pricePerCard: legacyPrice, freeCount: 0 };
        if (legacyType === 'lifetime') lifetime = seed; else annual = seed;
      }
    }

    return {
      id: String(r[ix('id')] || ''),
      createdAt: toIso(r[ix('createdAt')]),
      updatedAt: toIso(r[ix('updatedAt')]),
      payDate: toIso(r[ix('payDate')]),
      exchangeRate: num(r[ix('exchange_rate')]),
      annual: annual,
      lifetime: lifetime,
      photos: Array.isArray(photos) ? photos : []
    };
  });
}

function upsertCard(payload) {
  if (!payload || !payload.card) throw new Error('payload.card is required');
  const c = payload.card;
  if (!c.id) throw new Error('card.id is required');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getOrCreateSheet(SHEET_CARDS, CARD_COLS);
    const row = new Array(CARD_COLS.length).fill('');
    const set = function (name, val) {
      const idx = CARD_COLS.indexOf(name);
      if (idx >= 0) row[idx] = val == null ? '' : val;
    };
    set('id', c.id);
    set('createdAt', c.createdAt || '');
    set('updatedAt', c.updatedAt || new Date().toISOString());
    set('payDate', c.payDate || '');
    // legacy fields — เคลียร์ทิ้ง  ใช้ annual_/lifetime_ แทน
    set('cardType', '');
    set('cardCount', 0);
    set('pricePerCard', 0);
    const photos = Array.isArray(c.photos) ? c.photos.filter(function (p) { return p && p.fileId; }) : [];
    set('photos_json', JSON.stringify(photos));
    const a = c.annual || {};
    const l = c.lifetime || {};
    set('annual_count',   Number(a.count) || 0);
    set('annual_price',   Number(a.pricePerCard) || 0);
    set('annual_free',    Number(a.freeCount) || 0);
    set('lifetime_count', Number(l.count) || 0);
    set('lifetime_price', Number(l.pricePerCard) || 0);
    set('lifetime_free',  Number(l.freeCount) || 0);
    set('exchange_rate',  Number(c.exchangeRate) || 0);

    const lastRow = sheet.getLastRow();
    let foundIdx = -1;
    if (lastRow >= 2) {
      const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(c.id)) { foundIdx = i; break; }
      }
    }
    if (foundIdx === -1) {
      sheet.getRange(sheet.getLastRow() + 1, 1, 1, CARD_COLS.length).setValues([row]);
    } else {
      sheet.getRange(foundIdx + 2, 1, 1, CARD_COLS.length).setValues([row]);
    }
    return c.id;
  } finally {
    lock.releaseLock();
  }
}

function deleteCardById(id) {
  if (!id) throw new Error('id is required');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = ss().getSheetByName(SHEET_CARDS);
    if (!sheet || sheet.getLastRow() < 2) return id;
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, CARD_COLS.length).getValues();
    const ix = function (name) { return CARD_COLS.indexOf(name); };
    const photoFileIds = [];
    for (let i = rows.length - 1; i >= 0; i--) {
      if (String(rows[i][ix('id')]) === String(id)) {
        const raw = rows[i][ix('photos_json')];
        if (raw) {
          try {
            const photos = JSON.parse(raw) || [];
            photos.forEach(function (p) { if (p && p.fileId) photoFileIds.push(String(p.fileId)); });
          } catch (e) { /* ignore */ }
        }
        sheet.deleteRow(i + 2);
      }
    }
    photoFileIds.forEach(function (fid) {
      try { DriveApp.getFileById(fid).setTrashed(true); } catch (e) { /* ignore */ }
    });
    return id;
  } finally {
    lock.releaseLock();
  }
}

/* ================ Delete ================ */

function deleteShipmentById(id) {
  if (!id) throw new Error('id is required');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    // 0) เก็บ containerNo ของ shipment ก่อน — ใช้หา Drive folder ตอนล้างไฟล์
    let containerNo = '';
    const shipSheet = ss().getSheetByName(SHEET_SHIPMENTS);
    if (shipSheet && shipSheet.getLastRow() >= 2) {
      const idIdx = SHIPMENT_COLS.indexOf('id');
      const cIdx  = SHIPMENT_COLS.indexOf('containerNo');
      const rows = shipSheet.getRange(2, 1, shipSheet.getLastRow() - 1, SHIPMENT_COLS.length).getValues();
      for (let i = 0; i < rows.length; i++) {
        if (String(rows[i][idIdx]) === String(id)) {
          containerNo = String(rows[i][cIdx] || rows[i][idIdx] || '');
          break;
        }
      }
    }

    // 1) เก็บ fileId ของรูปทุกใบใน shipment ก่อนลบ metadata
    const photoSheet = ss().getSheetByName(SHEET_PHOTOS);
    const fileIdsToDelete = [];
    if (photoSheet && photoSheet.getLastRow() >= 2) {
      const rows = photoSheet.getRange(2, 1, photoSheet.getLastRow() - 1, PHOTO_COLS.length).getValues();
      const sidIdx = PHOTO_COLS.indexOf('shipmentId');
      const fidIdx = PHOTO_COLS.indexOf('fileId');
      rows.forEach(function (r) {
        if (String(r[sidIdx]) === String(id) && r[fidIdx]) fileIdsToDelete.push(String(r[fidIdx]));
      });
    }
    // 2) ลบ row ใน Shipments / Items / Costs / Photos
    [SHEET_SHIPMENTS, SHEET_ITEMS, SHEET_COSTS, SHEET_PHOTOS].forEach(name => {
      const sheet = ss().getSheetByName(name);
      if (!sheet || sheet.getLastRow() < 2) return;
      const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
      for (let i = ids.length - 1; i >= 0; i--) {
        if (String(ids[i][0]) === String(id)) sheet.deleteRow(i + 2);
      }
    });
    // 3) ลบ chunks legacy ใน PhotoData ที่เกี่ยวข้อง
    if (fileIdsToDelete.length) {
      const dataSheet = ss().getSheetByName(SHEET_PHOTO_DATA);
      if (dataSheet && dataSheet.getLastRow() >= 2) {
        const fidSet = {};
        fileIdsToDelete.forEach(function (f) { fidSet[f] = true; });
        const fids = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, 1).getValues();
        for (let i = fids.length - 1; i >= 0; i--) {
          if (fidSet[String(fids[i][0])]) dataSheet.deleteRow(i + 2);
        }
      }
    }
    // 4) ลบไฟล์จริงใน Drive (best-effort — รูป legacy chunked จะ throw แต่ข้าม)
    fileIdsToDelete.forEach(function (fid) {
      try { DriveApp.getFileById(fid).setTrashed(true); } catch (e) { /* ignore */ }
    });
    // 5) ลบโฟลเดอร์ SP21/<containerNo>/ ทิ้ง (ถ้าโฟลเดอร์ว่างหรือมีแค่ orphan files ของ shipment นี้)
    if (containerNo) {
      try {
        const rootIt = DriveApp.getRootFolder().getFoldersByName(DRIVE_ROOT_FOLDER);
        if (rootIt.hasNext()) {
          const sp21 = rootIt.next();
          const subIt = sp21.getFoldersByName(sanitizeFolderName(containerNo));
          if (subIt.hasNext()) subIt.next().setTrashed(true);
        }
      } catch (e) { /* best-effort */ }
    }
    return id;
  } finally {
    lock.releaseLock();
  }
}
