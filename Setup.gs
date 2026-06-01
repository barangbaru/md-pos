/*******************************************************
 * POS KASIR KEDAI KOPI - BATCH 1
 * SETUP DATABASE SHEETS
 *
 * IMPORTANT:
 * - Tidak pakai getActiveSpreadsheet()
 * - Semua akses spreadsheet pakai SPREADSHEET_ID
 * - Date dan Time disimpan sebagai string
 * - Sheet dibuat ringkas: MENU, SALES, SALE_ITEMS, CONFIG
 *******************************************************/

const SPREADSHEET_ID = '1nGbkJ_pEK4RgFOHB-yfsO7aDDKWsN8dG4z6COjC7g9Q';

/**
 * Jalankan function ini sekali untuk setup database POS.
 */
function setupPOSDatabase() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  setupMenuSheet_(ss);
  setupSalesSheet_(ss);
  setupSaleItemsSheet_(ss);
  setupExpensesSheet_(ss);
  setupTherapistsSheet_(ss);
  setupTherapistPointsSheet_(ss);
  setupConfigSheet_(ss);

  SpreadsheetApp.flush();

  Logger.log('✅ Setup POS Database selesai.');
}

/**
 * OPTIONAL:
 * Jalankan ini kalau mau reset total database.
 * HATI-HATI: ini akan menghapus isi sheet POS.
 */
function resetPOSDatabase() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const sheetNames = ['MENU', 'SALES', 'SALE_ITEMS', 'EXPENSES', 'THERAPISTS', 'THERAPIST_POINTS', 'CONFIG'];

  sheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      ss.deleteSheet(sheet);
    }
  });

  setupPOSDatabase();

  Logger.log('✅ Reset POS Database selesai.');
}

/* =====================================================
   SETUP SHEET: MENU
===================================================== */

function setupMenuSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, 'MENU');
  sheet.clear();

  const headers = [
    'Menu_ID',
    'Category',
    'Menu_Name',
    'Price',
    'Cost',
    'Stock',
    'Image_URL',
    'Active',
    'Stock_Status',
    'Created_At',
    'Updated_At'
  ];

  const now = getNowString_();

  const dummyMenus = [
    [
      'M001',
      'Coffee',
      'Kopi Susu Gula Aren',
      22000,
      12000,
      25,
      'https://images.unsplash.com/photo-1509042239860-f550ce710b93',
      'TRUE',
      'Available',
      now,
      now
    ],
    [
      'M002',
      'Coffee',
      'Es Americano',
      18000,
      9000,
      18,
      'https://images.unsplash.com/photo-1517701604599-bb29b565090c',
      'TRUE',
      'Available',
      now,
      now
    ],
    [
      'M003',
      'Coffee',
      'Cafe Latte',
      24000,
      13000,
      12,
      'https://images.unsplash.com/photo-1541167760496-1628856ab772',
      'TRUE',
      'Available',
      now,
      now
    ],
    [
      'M004',
      'Coffee',
      'Cappuccino',
      24000,
      13000,
      8,
      'https://images.unsplash.com/photo-1572442388796-11668a67e53d',
      'TRUE',
      'Low Stock',
      now,
      now
    ],
    [
      'M005',
      'Non Coffee',
      'Matcha Latte',
      26000,
      14000,
      15,
      'https://images.unsplash.com/photo-1515823064-d6e0c04616a7',
      'TRUE',
      'Available',
      now,
      now
    ],
    [
      'M006',
      'Non Coffee',
      'Chocolate Ice',
      23000,
      12000,
      10,
      'https://images.unsplash.com/photo-1542990253-0d0f5be5f0ed',
      'TRUE',
      'Low Stock',
      now,
      now
    ],
    [
      'M007',
      'Tea',
      'Lemon Tea',
      17000,
      7000,
      30,
      'https://images.unsplash.com/photo-1556679343-c7306c1976bc',
      'TRUE',
      'Available',
      now,
      now
    ],
    [
      'M008',
      'Tea',
      'Lychee Tea',
      19000,
      8000,
      0,
      'https://images.unsplash.com/photo-1556679343-c7306c1976bc',
      'TRUE',
      'Out of Stock',
      now,
      now
    ],
    [
      'M009',
      'Snack',
      'Croissant Butter',
      25000,
      15000,
      20,
      'https://images.unsplash.com/photo-1555507036-ab1f4038808a',
      'TRUE',
      'Available',
      now,
      now
    ],
    [
      'M010',
      'Snack',
      'French Fries',
      22000,
      11000,
      9,
      'https://images.unsplash.com/photo-1576107232684-1279f390859f',
      'TRUE',
      'Low Stock',
      now,
      now
    ],
    [
      'M011',
      'Snack',
      'Chicken Sandwich',
      32000,
      19000,
      14,
      'https://images.unsplash.com/photo-1528735602780-2552fd46c7af',
      'TRUE',
      'Available',
      now,
      now
    ],
    [
      'M012',
      'Dessert',
      'Brownies',
      21000,
      10000,
      7,
      'https://images.unsplash.com/photo-1606313564200-e75d5e30476c',
      'TRUE',
      'Low Stock',
      now,
      now
    ]
  ];

  writeTable_(sheet, headers, dummyMenus);
  applyMenuFormatting_(sheet);
}

/* =====================================================
   SETUP SHEET: SALES
===================================================== */

function setupSalesSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, 'SALES');
  sheet.clear();

  const headers = [
    'Transaction_ID',
    'Date',
    'Time',
    'Cashier_Name',
    'Subtotal',
    'Tax_Rate',
    'Tax_Amount',
    'Service_Rate',
    'Service_Amount',
    'Grand_Total',
    'Payment_Method',
    'Paid_Amount',
    'Change_Amount',
    'Rounded_Total',
    'Status',
    'Created_At'
  ];

  writeTable_(sheet, headers, []);
  applySalesFormatting_(sheet);
}

/* =====================================================
   SETUP SHEET: SALE_ITEMS
===================================================== */

function setupSaleItemsSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, 'SALE_ITEMS');
  sheet.clear();

  const headers = [
    'Transaction_ID',
    'Menu_ID',
    'Menu_Name',
    'Category',
    'Qty',
    'Price',
    'Cost',
    'Amount',
    'Gross_Profit',
    'Created_At'
  ];

  writeTable_(sheet, headers, []);
  applySaleItemsFormatting_(sheet);
}


/* =====================================================
   SETUP SHEET: EXPENSES
===================================================== */

function setupExpensesSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, 'EXPENSES');
  sheet.clear();

  const headers = [
    'Expense_ID',
    'Date',
    'Time',
    'Cashier_Name',
    'Expense_Type',
    'Personal_Cashier',
    'Category',
    'Description',
    'Amount',
    'Created_At'
  ];

  writeTable_(sheet, headers, []);
  applyExpensesFormatting_(sheet);
}

/* =====================================================
   SETUP SHEET: THERAPISTS (master data)
===================================================== */

function setupTherapistsSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, 'THERAPISTS');
  sheet.clear();

  const headers = [
    'Therapist_ID',
    'Therapist_Name',
    'Phone',
    'Default_Point_Value', // nilai 1 point dalam Rupiah (untuk hitung tarif jika dibutuhkan)
    'Active',
    'Notes',
    'Created_At',
    'Updated_At'
  ];

  const now = getNowString_();

  // contoh dummy supaya UI tidak kosong saat first run
  const dummy = [
    ['T001', 'Terapis A', '0812-0000-0001', 10000, 'TRUE', '', now, now],
    ['T002', 'Terapis B', '0812-0000-0002', 10000, 'TRUE', '', now, now]
  ];

  writeTable_(sheet, headers, dummy);
  applyTherapistsFormatting_(sheet);
}

/* =====================================================
   SETUP SHEET: THERAPIST_POINTS (log point per transaksi)
===================================================== */

function setupTherapistPointsSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, 'THERAPIST_POINTS');
  sheet.clear();

  const headers = [
    'Point_ID',
    'Date',
    'Time',
    'Therapist_ID',
    'Therapist_Name',
    'Service_Name',
    'Qty',
    'Point_Per_Unit',
    'Total_Point',
    'Amount',          // nominal Rupiah dari layanan ini (opsional, untuk audit)
    'Cashier_Name',    // kasir yang input
    'Notes',
    'Created_At'
  ];

  writeTable_(sheet, headers, []);
  applyTherapistPointsFormatting_(sheet);
}

/* =====================================================
   SETUP SHEET: CONFIG
===================================================== */

function setupConfigSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, 'CONFIG');
  sheet.clear();

  const headers = ['Key', 'Value', 'Description', 'Updated_At'];

  const now = getNowString_();

  const configRows = [
    [
      'STORE_NAME',
      'Kedai Kopi AIM',
      'Nama toko yang tampil di struk dan dashboard',
      now
    ],
    [
      'STORE_ADDRESS',
      'BSD Serpong, Tangerang Selatan',
      'Alamat toko yang tampil di struk',
      now
    ],
    [
      'STORE_PHONE',
      '085719254547',
      'Nomor telepon toko',
      now
    ],
    [
      'TAX_RATE',
      '10',
      'Pajak dalam persen. Contoh: 10 berarti 10%',
      now
    ],
    [
      'SERVICE_RATE',
      '5',
      'Service charge dalam persen. Contoh: 5 berarti 5%',
      now
    ],
    [
      'CASHIER_PASSCODE',
      '123456',
      'Passcode login awal kasir',
      now
    ],
    [
      'DASHBOARD_PASSCODE',
      '456789',
      'Passcode untuk membuka dashboard penjualan',
      now
    ],
    [
      'CONFIG_PASSCODE',
      '999999',
      'Passcode untuk membuka tab config',
      now
    ],
    [
      'LOW_STOCK_LIMIT',
      '10',
      'Batas stock rendah. Di bawah angka ini akan tampil Sisa Sedikit',
      now
    ],
    [
      'CURRENCY',
      'IDR',
      'Mata uang aplikasi',
      now
    ],
    [
      'RECEIPT_FOOTER',
      'Terima kasih sudah berkunjung.',
      'Teks penutup struk pembayaran',
      now
    ],
    [
      'SUPERADMIN_PASSCODE',
      '777777',
      'Passcode untuk membuka tab Konsolidasi (superadmin)',
      now
    ],
    [
      'THERAPIST_PASSCODE',
      '111111',
      'Passcode untuk membuka master terapis & input point',
      now
    ],
    [
      'THERAPIST_SHARE_PERCENT',
      '50',
      'Persentase bagi hasil terapis dari total nilai point (default 50%)',
      now
    ]
  ];

  writeTable_(sheet, headers, configRows);
  applyConfigFormatting_(sheet);
}


/**
 * Jalankan function ini 1x untuk upgrade file lama tanpa menghapus data.
 */
function upgradePOSDatabaseSchema() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  addMissingColumns_(ss, 'MENU', ['Cost']);
  addMissingColumns_(ss, 'SALE_ITEMS', ['Cost', 'Gross_Profit']);
  if (!ss.getSheetByName('EXPENSES')) {
    setupExpensesSheet_(ss);
  } else {
    addMissingColumns_(ss, 'EXPENSES', ['Expense_ID','Date','Time','Cashier_Name','Expense_Type','Personal_Cashier','Category','Description','Amount','Created_At']);
  }

  // ===== Therapist sheets =====
  if (!ss.getSheetByName('THERAPISTS')) {
    setupTherapistsSheet_(ss);
  } else {
    addMissingColumns_(ss, 'THERAPISTS', ['Therapist_ID','Therapist_Name','Phone','Default_Point_Value','Active','Notes','Created_At','Updated_At']);
  }
  if (!ss.getSheetByName('THERAPIST_POINTS')) {
    setupTherapistPointsSheet_(ss);
  } else {
    addMissingColumns_(ss, 'THERAPIST_POINTS', ['Point_ID','Date','Time','Therapist_ID','Therapist_Name','Service_Name','Qty','Point_Per_Unit','Total_Point','Amount','Cashier_Name','Notes','Created_At']);
  }

  // ===== Item tambahan (pembukuan terpisah) =====
  if (!ss.getSheetByName('ADDON_MENU')) {
    const am = getOrCreateSheet_(ss, 'ADDON_MENU');
    writeTable_(am, ['Addon_ID', 'Name', 'Price', 'Cost', 'Stock', 'Stock_Status', 'Active', 'Created_At', 'Updated_At'], []);
  }
  if (!ss.getSheetByName('ADDON_SALES')) {
    const as = getOrCreateSheet_(ss, 'ADDON_SALES');
    writeTable_(as, ['Addon_Sale_ID', 'Transaction_ID', 'Date', 'Time', 'Addon_ID', 'Name', 'Qty', 'Price', 'Cost', 'Amount', 'Gross_Profit', 'Cashier_Name', 'Created_At'], []);
  }

  // ===== Tambah config baru kalau belum ada =====
  ensureConfigRow_(ss, 'SUPERADMIN_PASSCODE', '777777', 'Passcode untuk membuka tab Konsolidasi (superadmin)');
  ensureConfigRow_(ss, 'THERAPIST_PASSCODE', '111111', 'Passcode untuk membuka master terapis & input point');
  ensureConfigRow_(ss, 'THERAPIST_SHARE_PERCENT', '50', 'Persentase bagi hasil terapis dari total nilai point (default 50%)');

  SpreadsheetApp.flush();
  Logger.log('✅ Upgrade POS Database Schema selesai tanpa reset data.');
}

function ensureConfigRow_(ss, key, defaultValue, description) {
  const sheet = ss.getSheetByName('CONFIG');
  if (!sheet) return;
  const values = sheet.getDataRange().getValues();
  if (values.length < 1) return;
  const headers = values[0].map(String);
  const keyIdx = headers.indexOf('Key');
  if (keyIdx === -1) return;
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][keyIdx]) === key) return; // sudah ada
  }
  const now = getNowString_();
  sheet.appendRow([key, defaultValue, description, now]);
}

function addMissingColumns_(ss, sheetName, requiredHeaders) {
  const sheet = getOrCreateSheet_(ss, sheetName);
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  requiredHeaders.forEach(header => {
    if (!headers.includes(header)) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
      headers.push(header);
    }
  });
}

/* =====================================================
   SHARED HELPERS
===================================================== */

function getOrCreateSheet_(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  return sheet;
}

function writeTable_(sheet, headers, rows) {
  const totalColumns = headers.length;

  sheet.getRange(1, 1, 1, totalColumns).setValues([headers]);

  if (rows && rows.length > 0) {
    sheet.getRange(2, 1, rows.length, totalColumns).setValues(rows);
  }

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, totalColumns);
}

function applyBaseHeaderStyle_(sheet, totalColumns) {
  const headerRange = sheet.getRange(1, 1, 1, totalColumns);

  headerRange
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#3b2415')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.setRowHeight(1, 36);
}

function applyMenuFormatting_(sheet) {
  const totalColumns = 11;
  applyBaseHeaderStyle_(sheet, totalColumns);

  sheet.getRange('A:A').setNumberFormat('@'); // Menu_ID as text
  sheet.getRange('B:C').setNumberFormat('@');
  sheet.getRange('D:E').setNumberFormat('#,##0');
  sheet.getRange('F:F').setNumberFormat('#,##0');
  sheet.getRange('G:K').setNumberFormat('@');

  sheet.setColumnWidth(1, 110);
  sheet.setColumnWidth(2, 130);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 80);
  sheet.setColumnWidth(7, 280);
  sheet.setColumnWidth(8, 80);
  sheet.setColumnWidth(9, 130);
  sheet.setColumnWidth(10, 160);
  sheet.setColumnWidth(11, 160);

  applyFilterIfNeeded_(sheet, totalColumns);
}

function applySalesFormatting_(sheet) {
  const totalColumns = 16;
  applyBaseHeaderStyle_(sheet, totalColumns);

  sheet.getRange('A:D').setNumberFormat('@');
  sheet.getRange('E:J').setNumberFormat('#,##0');
  sheet.getRange('K:K').setNumberFormat('@');
  sheet.getRange('L:N').setNumberFormat('#,##0');
  sheet.getRange('O:P').setNumberFormat('@');

  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 140);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(6, 90);
  sheet.setColumnWidth(7, 110);
  sheet.setColumnWidth(8, 110);
  sheet.setColumnWidth(9, 130);
  sheet.setColumnWidth(10, 130);
  sheet.setColumnWidth(11, 140);
  sheet.setColumnWidth(12, 120);
  sheet.setColumnWidth(13, 130);
  sheet.setColumnWidth(14, 130);
  sheet.setColumnWidth(15, 110);
  sheet.setColumnWidth(16, 170);

  applyFilterIfNeeded_(sheet, totalColumns);
}

function applySaleItemsFormatting_(sheet) {
  const totalColumns = 10;
  applyBaseHeaderStyle_(sheet, totalColumns);

  sheet.getRange('A:D').setNumberFormat('@');
  sheet.getRange('E:E').setNumberFormat('#,##0');
  sheet.getRange('F:I').setNumberFormat('#,##0');
  sheet.getRange('J:J').setNumberFormat('@');

  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 80);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 110);
  sheet.setColumnWidth(8, 110);
  sheet.setColumnWidth(9, 130);
  sheet.setColumnWidth(10, 170);

  applyFilterIfNeeded_(sheet, totalColumns);
}


function applyExpensesFormatting_(sheet) {
  const totalColumns = 10;
  applyBaseHeaderStyle_(sheet, totalColumns);

  sheet.getRange('A:H').setNumberFormat('@');
  sheet.getRange('I:I').setNumberFormat('#,##0');
  sheet.getRange('J:J').setNumberFormat('@');

  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 140);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 160);
  sheet.setColumnWidth(7, 170);
  sheet.setColumnWidth(8, 320);
  sheet.setColumnWidth(9, 130);
  sheet.setColumnWidth(10, 170);

  applyFilterIfNeeded_(sheet, totalColumns);
}

function applyConfigFormatting_(sheet) {
  const totalColumns = 4;
  applyBaseHeaderStyle_(sheet, totalColumns);

  sheet.getRange('A:D').setNumberFormat('@');

  sheet.setColumnWidth(1, 190);
  sheet.setColumnWidth(2, 230);
  sheet.setColumnWidth(3, 420);
  sheet.setColumnWidth(4, 170);

  applyFilterIfNeeded_(sheet, totalColumns);
}

function applyTherapistsFormatting_(sheet) {
  const totalColumns = 8;
  applyBaseHeaderStyle_(sheet, totalColumns);

  sheet.getRange('A:C').setNumberFormat('@');
  sheet.getRange('D:D').setNumberFormat('#,##0');
  sheet.getRange('E:H').setNumberFormat('@');

  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 80);
  sheet.setColumnWidth(6, 280);
  sheet.setColumnWidth(7, 170);
  sheet.setColumnWidth(8, 170);

  applyFilterIfNeeded_(sheet, totalColumns);
}

function applyTherapistPointsFormatting_(sheet) {
  const totalColumns = 13;
  applyBaseHeaderStyle_(sheet, totalColumns);

  sheet.getRange('A:F').setNumberFormat('@');
  sheet.getRange('G:J').setNumberFormat('#,##0');
  sheet.getRange('K:M').setNumberFormat('@');

  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 110);
  sheet.setColumnWidth(5, 180);
  sheet.setColumnWidth(6, 200);
  sheet.setColumnWidth(7, 70);
  sheet.setColumnWidth(8, 110);
  sheet.setColumnWidth(9, 110);
  sheet.setColumnWidth(10, 110);
  sheet.setColumnWidth(11, 140);
  sheet.setColumnWidth(12, 250);
  sheet.setColumnWidth(13, 170);

  applyFilterIfNeeded_(sheet, totalColumns);
}

function applyFilterIfNeeded_(sheet, totalColumns) {
  const existingFilter = sheet.getFilter();

  if (existingFilter) {
    existingFilter.remove();
  }

  const lastRow = Math.max(sheet.getLastRow(), 2);
  sheet.getRange(1, 1, lastRow, totalColumns).createFilter();
}

/**
 * Semua tanggal kita simpan sebagai string.
 * Format: yyyy-MM-dd HH:mm:ss
 */
function getNowString_() {
  const timezone = Session.getScriptTimeZone() || 'Asia/Jakarta';
  return Utilities.formatDate(new Date(), timezone, 'yyyy-MM-dd HH:mm:ss');
}

/**
 * Untuk tanggal transaksi.
 * Format: yyyy-MM-dd
 */
function getTodayString_() {
  const timezone = Session.getScriptTimeZone() || 'Asia/Jakarta';
  return Utilities.formatDate(new Date(), timezone, 'yyyy-MM-dd');
}

/**
 * Untuk jam transaksi.
 * Format: HH:mm:ss
 */
function getTimeString_() {
  const timezone = Session.getScriptTimeZone() || 'Asia/Jakarta';
  return Utilities.formatDate(new Date(), timezone, 'HH:mm:ss');
}