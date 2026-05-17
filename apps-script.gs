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
 *     (ครั้งแรกจะให้ Authorize เข้าถึง Google Sheets — กดอนุญาต)
 *     ระบบจะสร้าง header row บน sheet "Shipments", "Items", "Costs"
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
 *  หมายเหตุ
 *  --------
 *  • รูปภาพ (base64) ถูกเก็บไว้ใน localStorage ของ browser เท่านั้น
 *    ไม่ส่งขึ้น Sheets เพราะ cell มี limit ~50,000 chars  รูปก้อนเดียวก็เกินได้
 *    บน sheet จะเก็บแค่ photoCount ไว้ดูว่ามีรูปกี่ใบ
 *  • ตาราง Costs เป็น snapshot ของต้นทุนล่าสุดที่คำนวณจาก client
 *    หากแก้ Step ใด ๆ ที่กระทบต้นทุน  row นี้จะถูกเขียนทับ
 */

const SPREADSHEET_ID = '11yVFC71onKOKSsQbDiTbQDCvyY12s067ZMVmS7VFBlE';

const SHEET_SHIPMENTS = 'Shipments';
const SHEET_ITEMS = 'Items';
const SHEET_COSTS = 'Costs';

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
    if (req.action === 'list')   return json({ ok: true, data: listShipments() });
    if (req.action === 'save')   return json({ ok: true, data: upsertShipment(req.payload) });
    if (req.action === 'delete') return json({ ok: true, data: deleteShipmentById(req.id) });
    if (req.action === 'init')   return json({ ok: true, data: initSheets() });
    if (req.action === 'ping')   return json({ ok: true, data: 'pong' });
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
  return 'Sheets initialized: ' + [SHEET_SHIPMENTS, SHEET_ITEMS, SHEET_COSTS].join(', ');
}

/* ================ List ================ */

function listShipments() {
  const sheet = getOrCreateSheet(SHEET_SHIPMENTS, SHIPMENT_COLS);
  if (sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, SHIPMENT_COLS.length).getValues();
  const rawIdx = SHIPMENT_COLS.indexOf('raw_json');
  const out = [];
  for (const row of rows) {
    const raw = row[rawIdx];
    if (raw) {
      try { out.push(JSON.parse(raw)); continue; } catch (e) {}
    }
    out.push(rowToShipment(row));
  }
  return out;
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

  // raw_json — backup ของทุกฟิลด์  ตัด photos ออกเพื่อให้ไม่เกิน cell limit
  const lean = deepClone(s);
  if (lean.stepB && lean.stepB.photos) lean.stepB.photos = [];
  if (lean.stepE && lean.stepE.photos) lean.stepE.photos = [];
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

/* ================ Delete ================ */

function deleteShipmentById(id) {
  if (!id) throw new Error('id is required');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    [SHEET_SHIPMENTS, SHEET_ITEMS, SHEET_COSTS].forEach(name => {
      const sheet = ss().getSheetByName(name);
      if (!sheet || sheet.getLastRow() < 2) return;
      const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
      for (let i = ids.length - 1; i >= 0; i--) {
        if (String(ids[i][0]) === String(id)) sheet.deleteRow(i + 2);
      }
    });
    return id;
  } finally {
    lock.releaseLock();
  }
}
