/*******************************************************
 * POS KASIR KEDAI KOPI - UTILS
 * Helper backend untuk:
 * - open spreadsheet by ID
 * - baca sheet jadi object
 * - normalisasi string/date/number
 * - config
 * - rupiah
 * - stock status
 *******************************************************/

const POS_SHEET = {
  MENU: 'MENU',
  SALES: 'SALES',
  SALE_ITEMS: 'SALE_ITEMS',
  EXPENSES: 'EXPENSES',
  THERAPISTS: 'THERAPISTS',
  THERAPIST_POINTS: 'THERAPIST_POINTS',
  CONFIG: 'CONFIG'
};

const POS_STATUS = {
  AVAILABLE: 'Available',
  LOW_STOCK: 'Low Stock',
  OUT_OF_STOCK: 'Out of Stock'
};

const POS_PAYMENT = {
  CASH: 'Cash',
  QRIS: 'QRIS',
  TRANSFER: 'Transfer'
};

/* =====================================================
   SPREADSHEET ACCESS
===================================================== */

function POS_openSpreadsheet_() {
  if (!SPREADSHEET_ID || SPREADSHEET_ID === 'ISI_DENGAN_SPREADSHEET_ID_KAMU') {
    throw new Error('SPREADSHEET_ID belum diisi. Isi dulu di file Setup.gs.');
  }

  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function POS_getSheet_(sheetName) {
  const ss = POS_openSpreadsheet_();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('Sheet tidak ditemukan: ' + sheetName);
  }

  return sheet;
}

function POS_getHeaders_(sheet) {
  const lastCol = sheet.getLastColumn();

  if (lastCol < 1) {
    return [];
  }

  return sheet
    .getRange(1, 1, 1, lastCol)
    .getValues()[0]
    .map(value => POS_toString_(value).trim());
}

function POS_getHeaderMap_(headers) {
  const map = {};

  headers.forEach((header, index) => {
    map[header] = index;
  });

  return map;
}

function POS_readObjects_(sheetName) {
  const sheet = POS_getSheet_(sheetName);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2 || lastCol < 1) {
    return [];
  }

  const headers = POS_getHeaders_(sheet);
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  return values.map(row => {
    const obj = {};

    headers.forEach((header, index) => {
      obj[header] = POS_normalizeCell_(row[index]);
    });

    return obj;
  });
}

function POS_appendObjects_(sheetName, objects) {
  if (!objects || objects.length === 0) {
    return;
  }

  const sheet = POS_getSheet_(sheetName);
  const headers = POS_getHeaders_(sheet);

  const rows = objects.map(obj => {
    return headers.map(header => {
      const value = obj[header];

      if (value === null || value === undefined) {
        return '';
      }

      return value;
    });
  });

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, headers.length).setValues(rows);
}

/* =====================================================
   NORMALIZATION
===================================================== */

function POS_normalizeCell_(value) {
  if (value === null || value === undefined) {
    return '';
  }

  if (Object.prototype.toString.call(value) === '[object Date]') {
    return POS_formatDateTime_(value);
  }

  if (typeof value === 'string') {
    return value.trim();
  }

  return value;
}

function POS_toString_(value) {
  if (value === null || value === undefined) {
    return '';
  }

  if (Object.prototype.toString.call(value) === '[object Date]') {
    return POS_formatDateTime_(value);
  }

  return String(value);
}

function POS_toNumber_(value) {
  if (value === null || value === undefined || value === '') {
    return 0;
  }

  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : 0;
  }

  const cleaned = String(value)
    .replace(/Rp/gi, '')
    .replace(/\./g, '')
    .replace(/,/g, '.')
    .replace(/[^\d.-]/g, '')
    .trim();

  const number = Number(cleaned);

  return Number.isFinite(number) ? number : 0;
}

function POS_toBoolean_(value) {
  const text = POS_toString_(value).toLowerCase();

  return text === 'true' || text === 'yes' || text === '1' || text === 'active';
}

function POS_safeJson_(data) {
  return JSON.parse(JSON.stringify(data));
}

/* =====================================================
   DATE STRING HELPERS
===================================================== */

function POS_getTimezone_() {
  return Session.getScriptTimeZone() || 'Asia/Jakarta';
}

function POS_now_() {
  return new Date();
}

function POS_formatDate_(dateObj) {
  return Utilities.formatDate(dateObj, POS_getTimezone_(), 'yyyy-MM-dd');
}

function POS_formatTime_(dateObj) {
  return Utilities.formatDate(dateObj, POS_getTimezone_(), 'HH:mm:ss');
}

function POS_formatDateTime_(dateObj) {
  return Utilities.formatDate(dateObj, POS_getTimezone_(), 'yyyy-MM-dd HH:mm:ss');
}

function POS_todayString_() {
  return POS_formatDate_(POS_now_());
}

function POS_timeString_() {
  return POS_formatTime_(POS_now_());
}

function POS_nowString_() {
  return POS_formatDateTime_(POS_now_());
}

function POS_addDays_(dateObj, days) {
  const copy = new Date(dateObj.getTime());
  copy.setDate(copy.getDate() + days);
  return copy;
}

/* =====================================================
   CONFIG HELPERS
===================================================== */

function POS_getConfigMap_() {
  const rows = POS_readObjects_(POS_SHEET.CONFIG);
  const map = {};

  rows.forEach(row => {
    const key = POS_toString_(row.Key).trim();
    const value = POS_toString_(row.Value);

    if (key) {
      map[key] = value;
    }
  });

  return map;
}

function POS_getConfigValue_(key, defaultValue) {
  const config = POS_getConfigMap_();

  if (config[key] === undefined || config[key] === '') {
    return defaultValue;
  }

  return config[key];
}

function POS_updateConfigValues_(updates) {
  const sheet = POS_getSheet_(POS_SHEET.CONFIG);
  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    throw new Error('CONFIG belum siap.');
  }

  const headers = values[0].map(h => POS_toString_(h));
  const keyIndex = headers.indexOf('Key');
  const valueIndex = headers.indexOf('Value');
  const updatedAtIndex = headers.indexOf('Updated_At');

  if (keyIndex === -1 || valueIndex === -1) {
    throw new Error('Header CONFIG tidak valid.');
  }

  const now = POS_nowString_();
  const updateKeys = Object.keys(updates || {});

  if (updateKeys.length === 0) {
    return;
  }

  for (let r = 1; r < values.length; r++) {
    const key = POS_toString_(values[r][keyIndex]);

    if (updates[key] !== undefined) {
      sheet.getRange(r + 1, valueIndex + 1).setValue(POS_toString_(updates[key]));

      if (updatedAtIndex !== -1) {
        sheet.getRange(r + 1, updatedAtIndex + 1).setValue(now);
      }
    }
  }
}

/* =====================================================
   MENU / STOCK HELPERS
===================================================== */

function POS_getLowStockLimit_() {
  return POS_toNumber_(POS_getConfigValue_('LOW_STOCK_LIMIT', '10')) || 10;
}

function POS_computeStockStatus_(stock) {
  const limit = POS_getLowStockLimit_();
  const numericStock = POS_toNumber_(stock);

  if (numericStock <= 0) {
    return POS_STATUS.OUT_OF_STOCK;
  }

  if (numericStock <= limit) {
    return POS_STATUS.LOW_STOCK;
  }

  return POS_STATUS.AVAILABLE;
}

function POS_buildMenuObject_(row) {
  const stock = POS_toNumber_(row.Stock);
  const price = POS_toNumber_(row.Price);
  const cost = POS_toNumber_(row.Cost || row.HPP || row.Modal);
  const active = POS_toBoolean_(row.Active);
  const stockStatus = POS_computeStockStatus_(stock);

  return {
    menuId: POS_toString_(row.Menu_ID),
    category: POS_toString_(row.Category),
    menuName: POS_toString_(row.Menu_Name),
    price: price,
    cost: cost,
    stock: stock,
    imageUrl: POS_toString_(row.Image_URL),
    active: active,
    stockStatus: stockStatus,
    createdAt: POS_toString_(row.Created_At),
    updatedAt: POS_toString_(row.Updated_At)
  };
}

function POS_getMenuRowMap_() {
  const sheet = POS_getSheet_(POS_SHEET.MENU);
  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    return {};
  }

  const headers = values[0].map(h => POS_toString_(h));
  const map = POS_getHeaderMap_(headers);
  const result = {};

  for (let r = 1; r < values.length; r++) {
    const menuId = POS_toString_(values[r][map.Menu_ID]);

    if (!menuId) {
      continue;
    }

    result[menuId] = {
      rowNumber: r + 1,
      row: values[r],
      headerMap: map
    };
  }

  return result;
}

/* =====================================================
   MONEY HELPERS
===================================================== */

function POS_roundToNearest100_(amount) {
  const number = POS_toNumber_(amount);
  return Math.round(number / 100) * 100;
}

function POS_formatRupiah_(amount) {
  const number = POS_toNumber_(amount);

  return 'Rp ' + number.toLocaleString('id-ID');
}

function POS_calculateTotals_(items, taxRate, serviceRate) {
  const subtotal = items.reduce((sum, item) => {
    return sum + POS_toNumber_(item.amount);
  }, 0);

  const taxAmount = Math.round(subtotal * POS_toNumber_(taxRate) / 100);
  const serviceAmount = Math.round(subtotal * POS_toNumber_(serviceRate) / 100);
  const grandTotal = subtotal + taxAmount + serviceAmount;
  const roundedTotal = POS_roundToNearest100_(grandTotal);

  return {
    subtotal: subtotal,
    taxRate: POS_toNumber_(taxRate),
    taxAmount: taxAmount,
    serviceRate: POS_toNumber_(serviceRate),
    serviceAmount: serviceAmount,
    grandTotal: grandTotal,
    roundedTotal: roundedTotal
  };
}

/* =====================================================
   TRANSACTION HELPERS
===================================================== */

function POS_generateTransactionId_() {
  const now = POS_now_();
  const datePart = Utilities.formatDate(now, POS_getTimezone_(), 'yyyyMMdd');
  const timePart = Utilities.formatDate(now, POS_getTimezone_(), 'HHmmss');
  const randomPart = Math.floor(Math.random() * 900 + 100);

  return 'TRX-' + datePart + '-' + timePart + '-' + randomPart;
}


function POS_generateExpenseId_() {
  const now = POS_now_();
  const datePart = Utilities.formatDate(now, POS_getTimezone_(), 'yyyyMMdd');
  const timePart = Utilities.formatDate(now, POS_getTimezone_(), 'HHmmss');
  const randomPart = Math.floor(Math.random() * 900 + 100);
  return 'EXP-' + datePart + '-' + timePart + '-' + randomPart;
}

function POS_validatePaymentMethod_(paymentMethod) {
  const method = POS_toString_(paymentMethod);

  if (
    method !== POS_PAYMENT.CASH &&
    method !== POS_PAYMENT.QRIS &&
    method !== POS_PAYMENT.TRANSFER
  ) {
    throw new Error('Metode pembayaran tidak valid.');
  }

  return method;
}

function POS_validateCartItems_(items) {
  if (!Array.isArray(items) || items.length === 0) {
    throw new Error('Cart masih kosong.');
  }

  return items.map(item => {
    const menuId = POS_toString_(item.menuId);
    const qty = POS_toNumber_(item.qty);

    if (!menuId) {
      throw new Error('Menu ID tidak valid.');
    }

    if (qty <= 0) {
      throw new Error('Qty harus lebih dari 0.');
    }

    return {
      menuId: menuId,
      qty: qty
    };
  });
}

/* =====================================================
   DASHBOARD HELPERS
===================================================== */

function POS_getLast7DateStrings_() {
  const today = POS_now_();
  const dates = [];

  for (let i = 6; i >= 0; i--) {
    dates.push(POS_formatDate_(POS_addDays_(today, -i)));
  }

  return dates;
}

function POS_isSameDateString_(value, targetDate) {
  return POS_toString_(value) === POS_toString_(targetDate);
}