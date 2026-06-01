/*******************************************************
 * POS Warung Kolak Lestari - BACKEND API
 * Batch 2
 *
 * Backend untuk:
 * - doGet web app
 * - login kasir
 * - validasi passcode tab protected
 * - ambil menu/config
 * - update stock
 * - checkout order
 * - dashboard penjualan
 * - update config
 *******************************************************/

/* =====================================================
   WEB APP ROUTING
===================================================== */

function doGet() {
  const template = HtmlService.createTemplateFromFile('Index');

  return template
    .evaluate()
    .setTitle('POS system')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* =====================================================
   INITIAL DATA
===================================================== */

function getInitialData() {
  try {
    const config = getConfig();
    const menus = getMenus();
    const therapists = getTherapists({ activeOnly: true });
    const addons = getAddons({ activeOnly: true });

    return {
      success: true,
      message: 'Initial data berhasil dimuat.',
      data: {
        config: config.data,
        menus: menus.data,
        therapists: therapists.data,  // hanya terapis aktif, tanpa passcode
        addons: addons.data           // item tambahan aktif (pembukuan terpisah)
      }
    };
  } catch (error) {
    return POS_errorResponse_(error);
  }
}

/* =====================================================
   AUTH
===================================================== */

function validateCashierLogin(payload) {
  try {
    const cashierName = POS_toString_(payload && payload.cashierName).trim();
    const passcode = POS_toString_(payload && payload.passcode).trim();

    if (!cashierName) {
      throw new Error('Nama kasir wajib diisi.');
    }

    if (!passcode) {
      throw new Error('Passcode wajib diisi.');
    }

    const cashierPasscode = POS_getConfigValue_('CASHIER_PASSCODE', '');

    if (passcode !== cashierPasscode) {
      throw new Error('Passcode kasir salah.');
    }

    return {
      success: true,
      message: 'Login berhasil.',
      data: {
        cashierName: cashierName,
        loggedInAt: POS_nowString_()
      }
    };
  } catch (error) {
    return POS_errorResponse_(error);
  }
}

function validateProtectedTabPasscode(payload) {
  try {
    const area = POS_toString_(payload && payload.area).trim().toLowerCase();
    const passcode = POS_toString_(payload && payload.passcode).trim();

    if (!area) {
      throw new Error('Area akses belum dipilih.');
    }

    if (!passcode) {
      throw new Error('Passcode wajib diisi.');
    }

    let configKey = '';

    if (area === 'dashboard') {
      configKey = 'DASHBOARD_PASSCODE';
    } else if (area === 'config') {
      configKey = 'CONFIG_PASSCODE';
    } else if (area === 'therapist') {
      configKey = 'THERAPIST_PASSCODE';
    } else if (area === 'superadmin') {
      configKey = 'SUPERADMIN_PASSCODE';
    } else {
      throw new Error('Area akses tidak valid.');
    }

    const correctPasscode = POS_getConfigValue_(configKey, '');

    if (passcode !== correctPasscode) {
      throw new Error('Passcode salah.');
    }

    return {
      success: true,
      message: 'Akses berhasil dibuka.',
      data: {
        area: area,
        unlockedAt: POS_nowString_()
      }
    };
  } catch (error) {
    return POS_errorResponse_(error);
  }
}

/* =====================================================
   MENU
===================================================== */

function getMenus() {
  try {
    const rows = POS_readObjects_(POS_SHEET.MENU);

    const menus = rows
      .map(POS_buildMenuObject_)
      .filter(menu => menu.menuId && menu.active);

    const categories = [...new Set(menus.map(menu => menu.category).filter(Boolean))];

    return {
      success: true,
      message: 'Menu berhasil dimuat.',
      data: {
        menus: menus,
        categories: categories
      }
    };
  } catch (error) {
    return POS_errorResponse_(error);
  }
}

function updateMenuStock(payload) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const menuId = POS_toString_(payload && payload.menuId).trim();
    const mode = POS_toString_(payload && payload.mode).trim().toUpperCase();
    const stockValue = POS_toNumber_(payload && payload.stockValue);

    if (!menuId) {
      throw new Error('Menu ID wajib diisi.');
    }

    if (!['SET', 'ADD', 'REDUCE'].includes(mode)) {
      throw new Error('Mode update stock tidak valid. Gunakan SET, ADD, atau REDUCE.');
    }

    if (stockValue < 0) {
      throw new Error('Nilai stock tidak boleh negatif.');
    }

    const sheet = POS_getSheet_(POS_SHEET.MENU);
    const rowMap = POS_getMenuRowMap_();
    const target = rowMap[menuId];

    if (!target) {
      throw new Error('Menu tidak ditemukan: ' + menuId);
    }

    const stockCol = target.headerMap.Stock + 1;
    const statusCol = target.headerMap.Stock_Status + 1;
    const updatedAtCol = target.headerMap.Updated_At + 1;

    const currentStock = POS_toNumber_(target.row[target.headerMap.Stock]);

    let newStock = currentStock;

    if (mode === 'SET') {
      newStock = stockValue;
    }

    if (mode === 'ADD') {
      newStock = currentStock + stockValue;
    }

    if (mode === 'REDUCE') {
      newStock = Math.max(0, currentStock - stockValue);
    }

    const newStatus = POS_computeStockStatus_(newStock);

    sheet.getRange(target.rowNumber, stockCol).setValue(newStock);
    sheet.getRange(target.rowNumber, statusCol).setValue(newStatus);
    sheet.getRange(target.rowNumber, updatedAtCol).setValue(POS_nowString_());

    SpreadsheetApp.flush();

    return {
      success: true,
      message: 'Stock berhasil diperbarui.',
      data: {
        menuId: menuId,
        oldStock: currentStock,
        newStock: newStock,
        stockStatus: newStatus
      }
    };
  } catch (error) {
    return POS_errorResponse_(error);
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}


function updateMenuCost(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const menuId = POS_toString_(payload && payload.menuId).trim();
    const costValue = POS_toNumber_(payload && payload.costValue);
    if (!menuId) throw new Error('Menu ID wajib diisi.');
    if (costValue < 0) throw new Error('Modal/HPP tidak boleh negatif.');
    const sheet = POS_getSheet_(POS_SHEET.MENU);
    const rowMap = POS_getMenuRowMap_();
    const target = rowMap[menuId];
    if (!target) throw new Error('Menu tidak ditemukan: ' + menuId);
    if (target.headerMap.Cost === undefined) throw new Error('Kolom Cost belum ada. Jalankan upgradePOSDatabaseSchema().');
    sheet.getRange(target.rowNumber, target.headerMap.Cost + 1).setValue(costValue);
    sheet.getRange(target.rowNumber, target.headerMap.Updated_At + 1).setValue(POS_nowString_());
    SpreadsheetApp.flush();
    return { success: true, message: 'Modal/HPP menu berhasil diperbarui.', data: { menuId: menuId, cost: costValue } };
  } catch (error) { return POS_errorResponse_(error); }
  finally { try { lock.releaseLock(); } catch (e) {} }
}

/* =====================================================
   CONFIG
===================================================== */

function getConfig() {
  try {
    const config = POS_getConfigMap_();

    const safeConfig = {
      storeName: POS_toString_(config.STORE_NAME || 'Kedai Kopi'),
      storeAddress: POS_toString_(config.STORE_ADDRESS || ''),
      storePhone: POS_toString_(config.STORE_PHONE || ''),
      taxRate: POS_toNumber_(config.TAX_RATE || 0),
      serviceRate: POS_toNumber_(config.SERVICE_RATE || 0),
      lowStockLimit: POS_toNumber_(config.LOW_STOCK_LIMIT || 10),
      currency: POS_toString_(config.CURRENCY || 'IDR'),
      receiptFooter: POS_toString_(config.RECEIPT_FOOTER || 'Terima kasih.'),

      qrisImageUrl: POS_toString_(config.QRIS_IMAGE_URL || ''),
      bankName: POS_toString_(config.BANK_NAME || ''),
      bankAccountNumber: POS_toString_(config.BANK_ACCOUNT_NUMBER || ''),
      bankAccountName: POS_toString_(config.BANK_ACCOUNT_NAME || ''),
      transferNote: POS_toString_(config.TRANSFER_NOTE || ''),

      therapistSharePercent: POS_toNumber_(config.THERAPIST_SHARE_PERCENT || 50),

      // Notifikasi
      reportEmail: POS_toString_(config.REPORT_EMAIL || ''),
      telegramBotToken: POS_toString_(config.TELEGRAM_BOT_TOKEN || ''),
      telegramChatId: POS_toString_(config.TELEGRAM_CHAT_ID || ''),

      // Maintenance & Log
      logRetentionDays: POS_toNumber_(config.LOG_RETENTION_DAYS || 60)
    };

    return {
      success: true,
      message: 'Config berhasil dimuat.',
      data: safeConfig
    };
  } catch (error) {
    return POS_errorResponse_(error);
  }
}

function updateConfig(payload) {
  try {
    const configPasscode = POS_toString_(payload && payload.configPasscode).trim();

    if (!configPasscode) {
      throw new Error('Passcode config wajib diisi.');
    }

    const correctPasscode = POS_getConfigValue_('CONFIG_PASSCODE', '');

    if (configPasscode !== correctPasscode) {
      throw new Error('Passcode config salah.');
    }

    const updates = {};

    if (payload.storeName !== undefined) {
      updates.STORE_NAME = POS_toString_(payload.storeName);
    }

    if (payload.storeAddress !== undefined) {
      updates.STORE_ADDRESS = POS_toString_(payload.storeAddress);
    }

    if (payload.storePhone !== undefined) {
      updates.STORE_PHONE = POS_toString_(payload.storePhone);
    }

    if (payload.taxRate !== undefined) {
      const taxRate = POS_toNumber_(payload.taxRate);

      if (taxRate < 0 || taxRate > 100) {
        throw new Error('Tax rate harus di antara 0 sampai 100.');
      }

      updates.TAX_RATE = taxRate;
    }

    if (payload.serviceRate !== undefined) {
      const serviceRate = POS_toNumber_(payload.serviceRate);

      if (serviceRate < 0 || serviceRate > 100) {
        throw new Error('Service rate harus di antara 0 sampai 100.');
      }

      updates.SERVICE_RATE = serviceRate;
    }

    if (payload.lowStockLimit !== undefined) {
      const lowStockLimit = POS_toNumber_(payload.lowStockLimit);

      if (lowStockLimit < 1) {
        throw new Error('Low stock limit minimal 1.');
      }

      updates.LOW_STOCK_LIMIT = lowStockLimit;
    }

    if (payload.receiptFooter !== undefined) {
      updates.RECEIPT_FOOTER = POS_toString_(payload.receiptFooter);
    }

    if (payload.cashierPasscode !== undefined && POS_toString_(payload.cashierPasscode)) {
      updates.CASHIER_PASSCODE = POS_toString_(payload.cashierPasscode);
    }

    if (payload.dashboardPasscode !== undefined && POS_toString_(payload.dashboardPasscode)) {
      updates.DASHBOARD_PASSCODE = POS_toString_(payload.dashboardPasscode);
    }

    if (payload.newConfigPasscode !== undefined && POS_toString_(payload.newConfigPasscode)) {
      updates.CONFIG_PASSCODE = POS_toString_(payload.newConfigPasscode);
    }

    if (payload.therapistPasscode !== undefined && POS_toString_(payload.therapistPasscode)) {
      updates.THERAPIST_PASSCODE = POS_toString_(payload.therapistPasscode);
    }

    if (payload.superadminPasscode !== undefined && POS_toString_(payload.superadminPasscode)) {
      updates.SUPERADMIN_PASSCODE = POS_toString_(payload.superadminPasscode);
    }

    if (payload.therapistSharePercent !== undefined && payload.therapistSharePercent !== '') {
      const sharePct = POS_toNumber_(payload.therapistSharePercent);
      if (sharePct < 0 || sharePct > 100) {
        throw new Error('Persentase bagi hasil terapis harus 0 - 100.');
      }
      updates.THERAPIST_SHARE_PERCENT = sharePct;
    }

    // Notifikasi
    if (payload.reportEmail !== undefined) updates.REPORT_EMAIL = POS_toString_(payload.reportEmail).trim();
    if (payload.telegramBotToken !== undefined && POS_toString_(payload.telegramBotToken).trim()) {
      updates.TELEGRAM_BOT_TOKEN = POS_toString_(payload.telegramBotToken).trim();
    }
    if (payload.telegramChatId !== undefined && POS_toString_(payload.telegramChatId).trim()) {
      updates.TELEGRAM_CHAT_ID = POS_toString_(payload.telegramChatId).trim();
    }
    if (payload.logRetentionDays !== undefined && payload.logRetentionDays !== '') {
      const r = POS_toNumber_(payload.logRetentionDays);
      if (r < 1 || r > 90) throw new Error('Retensi log harus antara 1–90 hari.');
      updates.LOG_RETENTION_DAYS = r;
    }

    POS_updateConfigValues_(updates);

    return {
      success: true,
      message: 'Config berhasil diperbarui.',
      data: getConfig().data
    };
  } catch (error) {
    return POS_errorResponse_(error);
  }
}

/* =====================================================
   CHECKOUT
===================================================== */

function checkoutOrder(payload) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(15000);

    const cashierName = POS_toString_(payload && payload.cashierName).trim();
    const paymentMethod = POS_validatePaymentMethod_(payload && payload.paymentMethod);
    const paidAmountInput = POS_toNumber_(payload && payload.paidAmount);

    // Item layanan (service) & item tambahan (addon) — boleh salah satu kosong,
    // tapi tidak boleh dua-duanya kosong.
    const rawItems = (payload && Array.isArray(payload.items)) ? payload.items : [];
    const rawAddons = (payload && Array.isArray(payload.addonItems)) ? payload.addonItems : [];
    if (rawItems.length === 0 && rawAddons.length === 0) throw new Error('Cart masih kosong.');
    const cartItems = rawItems.length > 0 ? POS_validateCartItems_(rawItems) : [];

    if (!cashierName) {
      throw new Error('Nama kasir tidak ditemukan. Silakan login ulang.');
    }

    const config = POS_getConfigMap_();
    const taxRate = POS_toNumber_(config.TAX_RATE || 0);
    const serviceRate = POS_toNumber_(config.SERVICE_RATE || 0);

    const menuSheet = POS_getSheet_(POS_SHEET.MENU);
    const rowMap = POS_getMenuRowMap_();

    const normalizedItems = [];

    cartItems.forEach(cartItem => {
      const target = rowMap[cartItem.menuId];

      if (!target) {
        throw new Error('Menu tidak ditemukan: ' + cartItem.menuId);
      }

      const menuId = POS_toString_(target.row[target.headerMap.Menu_ID]);
      const category = POS_toString_(target.row[target.headerMap.Category]);
      const menuName = POS_toString_(target.row[target.headerMap.Menu_Name]);
      const price = POS_toNumber_(target.row[target.headerMap.Price]);
      const cost = target.headerMap.Cost !== undefined ? POS_toNumber_(target.row[target.headerMap.Cost]) : 0;
      const stock = POS_toNumber_(target.row[target.headerMap.Stock]);
      const active = POS_toBoolean_(target.row[target.headerMap.Active]);

      if (!active) {
        throw new Error(menuName + ' sedang tidak aktif.');
      }

      if (stock <= 0) {
        throw new Error(menuName + ' sedang habis.');
      }

      if (cartItem.qty > stock) {
        throw new Error(menuName + ' stock tidak cukup. Sisa stock: ' + stock);
      }

      normalizedItems.push({
        menuId: menuId,
        category: category,
        menuName: menuName,
        qty: cartItem.qty,
        price: price,
        cost: cost * cartItem.qty,
        amount: price * cartItem.qty,
        grossProfit: (price * cartItem.qty) - (cost * cartItem.qty),
        currentStock: stock,
        rowNumber: target.rowNumber,
        headerMap: target.headerMap
      });
    });

    const totals = POS_calculateTotals_(normalizedItems, taxRate, serviceRate);

    // ── Proses item tambahan (addon) — pembukuan terpisah ─────────────────
    // Stok addon dikelola di sheet ADDON_MENU; penjualan dicatat di ADDON_SALES.
    const normalizedAddons = [];
    let addonSheet = null;
    let addonRowMap = {};
    if (rawAddons.length > 0) {
      addonSheet = POS_getOrCreateSheet_(POS_SHEET.ADDON_MENU, ADDON_MENU_HEADERS_);
      const aHeaders = POS_getHeaders_(addonSheet);
      const aHeaderMap = POS_getHeaderMap_(aHeaders);
      const aLastRow = addonSheet.getLastRow();
      if (aLastRow >= 2) {
        const aValues = addonSheet.getRange(2, 1, aLastRow - 1, aHeaders.length).getValues();
        for (let i = 0; i < aValues.length; i++) {
          const id = POS_toString_(aValues[i][aHeaderMap.Addon_ID]);
          if (id) addonRowMap[id] = { rowNumber: i + 2, row: aValues[i], headerMap: aHeaderMap };
        }
      }
      rawAddons.forEach(a => {
        const addonId = POS_toString_(a.addonId).trim();
        const qty = POS_toNumber_(a.qty);
        if (!addonId) throw new Error('Addon ID tidak valid.');
        if (qty <= 0) throw new Error('Qty item tambahan harus > 0.');
        const t = addonRowMap[addonId];
        if (!t) throw new Error('Item tambahan tidak ditemukan: ' + addonId);
        const hm = t.headerMap;
        const name = POS_toString_(t.row[hm.Name]);
        const price = POS_toNumber_(t.row[hm.Price]);
        const cost = POS_toNumber_(t.row[hm.Cost]);
        const stock = POS_toNumber_(t.row[hm.Stock]);
        const active = POS_toBoolean_(t.row[hm.Active]);
        if (!active) throw new Error(name + ' (tambahan) tidak aktif.');
        if (stock <= 0) throw new Error(name + ' (tambahan) sedang habis.');
        if (qty > stock) throw new Error(name + ' (tambahan) stock tidak cukup. Sisa: ' + stock);
        normalizedAddons.push({
          addonId: addonId, name: name, qty: qty, price: price,
          cost: cost * qty, amount: price * qty, grossProfit: (price - cost) * qty,
          currentStock: stock, rowNumber: t.rowNumber, headerMap: hm
        });
      });
    }
    const addonTotal = normalizedAddons.reduce((s, a) => s + a.amount, 0);

    // ── Total gabungan untuk pembayaran (layanan + tambahan) ──────────────
    const combinedGrand = totals.roundedTotal + addonTotal;
    const combinedRounded = POS_roundToNearest100_(combinedGrand);

    let paidAmount = paidAmountInput;
    let changeAmount = 0;
    if (paymentMethod === POS_PAYMENT.CASH) {
      if (paidAmount < combinedRounded) {
        throw new Error('Uang bayar kurang dari total pembayaran.');
      }
      changeAmount = paidAmount - combinedRounded;
    } else {
      paidAmount = combinedRounded;
      changeAmount = 0;
    }

    const transactionId = POS_generateTransactionId_();
    const today = POS_todayString_();
    const time = POS_timeString_();
    const now = POS_nowString_();

    // SALES (buku layanan) — hanya ditulis jika ada item layanan.
    // Paid/Change di buku layanan = nilai layanan saja (pembukuan bersih),
    // pembayaran gabungan & kembalian riil ada di receipt.
    if (normalizedItems.length > 0) {
      const salesRow = {
        Transaction_ID: transactionId,
        Date: today,
        Time: time,
        Cashier_Name: cashierName,
        Subtotal: totals.subtotal,
        Tax_Rate: totals.taxRate,
        Tax_Amount: totals.taxAmount,
        Service_Rate: totals.serviceRate,
        Service_Amount: totals.serviceAmount,
        Grand_Total: totals.grandTotal,
        Payment_Method: paymentMethod,
        Paid_Amount: totals.roundedTotal,
        Change_Amount: 0,
        Rounded_Total: totals.roundedTotal,
        Status: 'PAID',
        Created_At: now
      };
      const saleItemRows = normalizedItems.map(item => ({
        Transaction_ID: transactionId,
        Menu_ID: item.menuId,
        Menu_Name: item.menuName,
        Category: item.category,
        Qty: item.qty,
        Price: item.price,
        Cost: item.cost,
        Amount: item.amount,
        Gross_Profit: item.grossProfit,
        Created_At: now
      }));
      POS_appendObjects_(POS_SHEET.SALES, [salesRow]);
      POS_appendObjects_(POS_SHEET.SALE_ITEMS, saleItemRows);
    }

    // ADDON_SALES (buku tambahan, terpisah)
    if (normalizedAddons.length > 0) {
      const addonSaleRows = normalizedAddons.map((a, idx) => ({
        Addon_Sale_ID: POS_generateAddonSaleId_(idx),
        Transaction_ID: transactionId,
        Date: today,
        Time: time,
        Addon_ID: a.addonId,
        Name: a.name,
        Qty: a.qty,
        Price: a.price,
        Cost: a.cost,
        Amount: a.amount,
        Gross_Profit: a.grossProfit,
        Cashier_Name: cashierName,
        Created_At: now
      }));
      POS_appendObjects_(POS_SHEET.ADDON_SALES, addonSaleRows);

      // Kurangi stok addon
      normalizedAddons.forEach(a => {
        const hm = a.headerMap;
        const newStock = a.currentStock - a.qty;
        const newStatus = POS_computeStockStatus_(newStock);
        addonSheet.getRange(a.rowNumber, hm.Stock + 1).setValue(newStock);
        addonSheet.getRange(a.rowNumber, hm.Stock_Status + 1).setValue(newStatus);
        addonSheet.getRange(a.rowNumber, hm.Updated_At + 1).setValue(now);
      });
    }
    POS_invalidateDashboardCache_();

    // Batch stock update: 1 setValues per item (3 kolom sekaligus) vs 3 setValue terpisah
    normalizedItems.forEach(item => {
      const newStock = item.currentStock - item.qty;
      const newStatus = POS_computeStockStatus_(newStock);
      const hm = item.headerMap;
      // Temukan kolom paling kiri dan paling kanan dari ketiga kolom target
      const cols = [hm.Stock + 1, hm.Stock_Status + 1, hm.Updated_At + 1];
      const minCol = Math.min.apply(null, cols);
      const maxCol = Math.max.apply(null, cols);
      const width = maxCol - minCol + 1;
      // Baca nilai row saat ini lalu patch ketiga kolom sekaligus
      const rowValues = menuSheet.getRange(item.rowNumber, minCol, 1, width).getValues()[0];
      rowValues[hm.Stock - (minCol - 1)] = newStock;
      rowValues[hm.Stock_Status - (minCol - 1)] = newStatus;
      rowValues[hm.Updated_At - (minCol - 1)] = now;
      menuSheet.getRange(item.rowNumber, minCol, 1, width).setValues([rowValues]);
    });

    SpreadsheetApp.flush();

    const receiptData = {
      store: {
        name: POS_toString_(config.STORE_NAME || 'Kedai Kopi'),
        address: POS_toString_(config.STORE_ADDRESS || ''),
        phone: POS_toString_(config.STORE_PHONE || ''),
        footer: POS_toString_(config.RECEIPT_FOOTER || 'Terima kasih.')
      },
      transaction: {
        transactionId: transactionId,
        date: today,
        time: time,
        cashierName: cashierName,
        paymentMethod: paymentMethod
      },
      items: normalizedItems.map(item => {
        return {
          menuId: item.menuId,
          category: item.category,
          menuName: item.menuName,
          qty: item.qty,
          price: item.price,
          amount: item.amount
        };
      }),
      addonItems: normalizedAddons.map(a => ({
        addonId: a.addonId,
        name: a.name,
        qty: a.qty,
        price: a.price,
        amount: a.amount
      })),
      totals: {
        subtotal: totals.subtotal,
        taxRate: totals.taxRate,
        taxAmount: totals.taxAmount,
        serviceRate: totals.serviceRate,
        serviceAmount: totals.serviceAmount,
        grandTotal: totals.grandTotal,
        roundedTotal: totals.roundedTotal,        // total layanan (untuk point terapis)
        addonTotal: addonTotal,                    // total item tambahan
        combinedTotal: combinedRounded,            // total bayar gabungan
        paidAmount: paidAmount,
        changeAmount: changeAmount
      }
    };

    return {
      success: true,
      message: 'Checkout berhasil.',
      data: {
        transactionId: transactionId,
        receipt: receiptData,
        menus: getMenus().data,
        addons: getAddons({ activeOnly: true }).data
      }
    };
  } catch (error) {
    return POS_errorResponse_(error);
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}


/* =====================================================
   EXPENSES / PENGELUARAN
===================================================== */

function POS_isPersonalExpenseCategory_(category) {
  const text = POS_toString_(category).toLowerCase().trim();
  return text.indexOf('kasbon') !== -1 || text.indexOf('fee') !== -1;
}

function POS_normalizeExpenseRow_(row) {
  const savedType = POS_toString_(row && row.Expense_Type).trim();
  const category = POS_toString_(row && row.Category);
  let type;
  if (savedType === 'Therapist') {
    // Hormati flag terapis yang sudah tersimpan
    type = 'Therapist';
  } else {
    type = POS_isPersonalExpenseCategory_(category) ? 'Personal' : 'Sharing';
  }
  const copy = Object.assign({}, row);
  copy.Expense_Type = type;
  if (type === 'Personal' && !POS_toString_(copy.Personal_Cashier)) copy.Personal_Cashier = POS_toString_(copy.Cashier_Name);
  if (type === 'Sharing') copy.Personal_Cashier = '';
  // Therapist: keep Personal_Cashier as-is (nama terapis)
  return copy;
}

function addExpense(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const cashierName = POS_toString_(payload && payload.cashierName).trim();
    const category = POS_toString_(payload && payload.category).trim();
    const description = POS_toString_(payload && payload.description).trim();
    const amount = POS_toNumber_(payload && payload.amount);

    // Explicit target dari frontend: 'therapist' = kasbon terapis (terhubung ke master THERAPISTS)
    const explicitTarget = POS_toString_(payload && payload.targetType).trim().toLowerCase();

    let expenseType;
    let personalCashier;

    if (explicitTarget === 'therapist') {
      expenseType = 'Therapist';
      // Untuk kasbon terapis, Personal_Cashier = nama terapis (bukan kasir)
      personalCashier = POS_toString_(payload && payload.therapistName || payload && payload.personalCashier).trim();
      if (!personalCashier) throw new Error('Nama terapis penerima kasbon wajib diisi.');
    } else if (POS_isPersonalExpenseCategory_(category)) {
      expenseType = 'Personal';
      personalCashier = POS_toString_(payload && payload.personalCashier || cashierName).trim();
      if (!personalCashier) throw new Error('Nama kasir penerima kasbon/fee wajib diisi.');
    } else {
      expenseType = 'Sharing';
      personalCashier = '';
    }

    if (!cashierName) throw new Error('Nama kasir tidak ditemukan. Silakan login ulang.');
    if (!category) throw new Error('Kategori pengeluaran wajib diisi.');
    if (!description) throw new Error('Keterangan pengeluaran wajib diisi.');
    if (amount <= 0) throw new Error('Nilai pengeluaran harus lebih dari 0.');

    const row = {
      Expense_ID: POS_generateExpenseId_(), Date: POS_todayString_(), Time: POS_timeString_(), Cashier_Name: cashierName,
      Expense_Type: expenseType, Personal_Cashier: personalCashier, Category: category, Description: description, Amount: amount,
      Created_At: POS_nowString_()
    };
    POS_appendObjects_(POS_SHEET.EXPENSES, [row]);
    SpreadsheetApp.flush();
    POS_invalidateDashboardCache_();
    return { success: true, message: 'Pengeluaran berhasil disimpan.', data: row };
  } catch (error) { return POS_errorResponse_(error); }
  finally { try { lock.releaseLock(); } catch (e) {} }
}

function getExpenses() {
  try {
    const expenses = POS_readObjects_(POS_SHEET.EXPENSES)
      .filter(row => POS_toString_(row.Expense_ID))
      .map(POS_normalizeExpenseRow_)
      .sort((a, b) => (POS_toString_(b.Date) + ' ' + POS_toString_(b.Time)).localeCompare(POS_toString_(a.Date) + ' ' + POS_toString_(a.Time)));
    return { success: true, message: 'Data pengeluaran berhasil dimuat.', data: { expenses: expenses } };
  } catch (error) { return POS_errorResponse_(error); }
}

/* =====================================================
   DASHBOARD
===================================================== */

/**
 * Mengambil data dashboard dengan dukungan filter range tanggal.
 *
 * @param {Object} [filter] Opsional. { startDate: 'yyyy-MM-dd', endDate: 'yyyy-MM-dd' }
 *   - Jika tidak dikirim atau invalid, default range = hari ini (single day).
 *   - Range dibatasi maksimal 366 hari untuk menjaga performa.
 *
 * Semua kartu utama, ringkasan pembayaran, top kasir, breakdown pengeluaran,
 * dan tabel rekap penjualan mengikuti range yang dipilih. Chart "Last 7 Days"
 * tetap menampilkan 7 hari terakhir relatif terhadap endDate range.
 */
function getDashboardData(filter) {
  try {
    // ── SERVER-SIDE CACHE (60 detik, invalid otomatis saat ada checkout/expense) ──
    const cache = CacheService.getScriptCache();
    const ver = cache.get('dash_ver') || '0';
    const cacheKey = 'dash_v' + ver + '_' + JSON.stringify(filter || {});
    const cached = cache.get(cacheKey);
    if (cached) {
      try { return JSON.parse(cached); } catch (e) { /* cache corrupt, lanjut */ }
    }

    const sales = POS_readObjects_(POS_SHEET.SALES);
    const saleItems = POS_readObjects_(POS_SHEET.SALE_ITEMS);
    const expenses = POS_readObjects_(POS_SHEET.EXPENSES).map(POS_normalizeExpenseRow_);
    const allTherapistPoints = POS_readObjects_(POS_SHEET.THERAPIST_POINTS).filter(r => POS_toString_(r.Point_ID));
    const menuCostMap = POS_getMenuCostMap_();
    const today = POS_todayString_();

    // ----- Resolve & validate date range -----
    const range = POS_resolveDashboardRange_(filter, today);
    const rangeDates = POS_getDateStringsBetween_(range.startDate, range.endDate);

    const paidSales = sales.filter(row => POS_toString_(row.Status) === 'PAID');

    // ── PRE-INDEX untuk O(1) lookup (hindari O(n²) di loop harian) ────────
    // Map: date → paidSales[]
    const salesByDate = {};
    paidSales.forEach(row => {
      const d = POS_toString_(row.Date).slice(0, 10);
      if (!salesByDate[d]) salesByDate[d] = [];
      salesByDate[d].push(row);
    });
    // Map: transactionId → saleItems[]
    const itemsByTxId = {};
    saleItems.forEach(item => {
      const txId = POS_toString_(item.Transaction_ID);
      if (!txId) return;
      if (!itemsByTxId[txId]) itemsByTxId[txId] = [];
      itemsByTxId[txId].push(item);
    });
    // Map: date → expenses[]
    const expensesByDate = {};
    expenses.forEach(row => {
      const d = POS_toString_(row.Date).slice(0, 10);
      if (!expensesByDate[d]) expensesByDate[d] = [];
      expensesByDate[d].push(row);
    });

    // Helper: ambil items dari Set txIds (O(1) per txId)
    function getItemsForSales(salesArr) {
      const result = [];
      salesArr.forEach(row => {
        const txId = POS_toString_(row.Transaction_ID);
        if (itemsByTxId[txId]) result.push.apply(result, itemsByTxId[txId]);
      });
      return result;
    }
    function getDailyExpenses(date) { return expensesByDate[date] || []; }
    function getDailySales(date) { return salesByDate[date] || []; }

    // Sales / items / expenses dalam range
    const rangeSales = paidSales.filter(row => POS_isDateInRange_(row.Date, range.startDate, range.endDate));
    const rangeExpenses = expenses.filter(row => POS_isDateInRange_(row.Date, range.startDate, range.endDate));
    const rangeItems = getItemsForSales(rangeSales);

    // Agregat utama (kartu "hari ini" di UI sekarang berisi total range)
    const rangeRevenue = POS_sumSalesAmount_(rangeSales);
    const rangeSubtotal = rangeSales.reduce((sum, row) => sum + POS_toNumber_(row.Subtotal), 0);
    const rangeCost = POS_sumItemCost_(rangeItems, menuCostMap);
    const rangeGrossProfit = POS_sumItemGrossProfit_(rangeItems, menuCostMap);
    const rangeExpenseTotal = POS_sumExpenses_(rangeExpenses);
    // Sudah dinormalisasi — langsung filter tanpa re-normalize
    const rangeSharingExpenseTotal = rangeExpenses.filter(r => POS_toString_(r.Expense_Type) === 'Sharing').reduce((s, r) => s + POS_toNumber_(r.Amount), 0);
    const rangePersonalExpenseTotal = rangeExpenses.filter(r => { const t = POS_toString_(r.Expense_Type); return t === 'Personal' || t === 'Therapist'; }).reduce((s, r) => s + POS_toNumber_(r.Amount), 0);
    const rangeNetProfit = rangeGrossProfit - rangeExpenseTotal;
    const rangeTransactions = rangeSales.length;
    const rangeTherapistPoints = allTherapistPoints.filter(r => POS_isDateInRange_(r.Date, range.startDate, range.endDate));
    const topTherapistRange = POS_calculateTopTherapists_(rangeTherapistPoints, 1)[0] || { therapistId: '-', therapistName: '-', totalPoint: 0, totalAmount: 0, count: 0 };

    // Breakdown harian dalam range — O(1) lookup via pre-index
    const rangeDailyBreakdown = rangeDates.map(date => {
      const dailySales = getDailySales(date);
      const dailyItems = getItemsForSales(dailySales);
      const dailyExpenses = getDailyExpenses(date);
      const revenue = POS_sumSalesAmount_(dailySales);
      const grossProfit = POS_sumItemGrossProfit_(dailyItems, menuCostMap);
      const expenseTotal = POS_sumExpenses_(dailyExpenses);

      return {
        date: date,
        label: date.slice(8),
        revenue: revenue,
        grossProfit: grossProfit,
        expenses: expenseTotal,
        netProfit: grossProfit - expenseTotal,
        transactions: dailySales.length
      };
    });

    const rangeSummary = {
      revenue: rangeRevenue,
      grossProfit: rangeGrossProfit,
      expenses: rangeExpenseTotal,
      netProfit: rangeNetProfit,
      transactions: rangeTransactions
    };

    // Chart 7 hari terakhir — O(1) lookup via pre-index
    const last7Dates = POS_getLast7DateStringsFrom_(range.endDate);
    const last7Revenue = last7Dates.map(date => {
      const dailySales = getDailySales(date);
      const dailyItems = getItemsForSales(dailySales);
      const dailyExpenses = getDailyExpenses(date);
      const revenue = POS_sumSalesAmount_(dailySales);
      const grossProfit = POS_sumItemGrossProfit_(dailyItems, menuCostMap);
      const expenseTotal = POS_sumExpenses_(dailyExpenses);
      return { date: date, label: date.slice(5), revenue: revenue, grossProfit: grossProfit, expenses: expenseTotal, netProfit: grossProfit - expenseTotal, transactions: dailySales.length };
    });

    const paymentSummary = POS_calculatePaymentSummary_(rangeSales);
    const topTherapists = POS_calculateTopTherapists_(rangeTherapistPoints, 5);
    const sharingExpenses = POS_groupExpensesByCategory_(rangeExpenses.filter(row => {
      const t = POS_toString_(row.Expense_Type);
      return t !== 'Personal' && t !== 'Therapist';
    }));
    const personalExpenses = POS_groupPersonalExpenses_(rangeExpenses.filter(row => {
      const t = POS_toString_(row.Expense_Type);
      return t === 'Personal' || t === 'Therapist';
    }));

    // ── PEMBUKUAN TAMBAHAN (ADDON) — terpisah dari layanan ────────────────
    let addonSales = [];
    try {
      addonSales = POS_readObjects_(POS_SHEET.ADDON_SALES)
        .filter(r => POS_toString_(r.Addon_Sale_ID))
        .filter(r => POS_isDateInRange_(r.Date, range.startDate, range.endDate));
    } catch (e) { addonSales = []; } // sheet belum ada = belum ada penjualan tambahan
    const addonRevenue = addonSales.reduce((s, r) => s + POS_toNumber_(r.Amount), 0);
    const addonCost = addonSales.reduce((s, r) => s + POS_toNumber_(r.Cost), 0);
    const addonProfit = addonSales.reduce((s, r) => s + POS_toNumber_(r.Gross_Profit), 0);
    const addonQty = addonSales.reduce((s, r) => s + POS_toNumber_(r.Qty), 0);
    // Top item tambahan (by qty)
    const addonMap = {};
    addonSales.forEach(r => {
      const name = POS_toString_(r.Name) || '-';
      if (!addonMap[name]) addonMap[name] = { name: name, qty: 0, amount: 0 };
      addonMap[name].qty += POS_toNumber_(r.Qty);
      addonMap[name].amount += POS_toNumber_(r.Amount);
    });
    const topAddons = Object.values(addonMap)
      .sort((a, b) => b.qty !== a.qty ? b.qty - a.qty : b.amount - a.amount)
      .slice(0, 5);
    // Breakdown harian addon
    const addonByDate = {};
    addonSales.forEach(r => {
      const d = POS_toString_(r.Date).slice(0, 10);
      if (!addonByDate[d]) addonByDate[d] = { date: d, revenue: 0, cost: 0, profit: 0, qty: 0, count: 0 };
      addonByDate[d].revenue += POS_toNumber_(r.Amount);
      addonByDate[d].cost += POS_toNumber_(r.Cost);
      addonByDate[d].profit += POS_toNumber_(r.Gross_Profit);
      addonByDate[d].qty += POS_toNumber_(r.Qty);
      addonByDate[d].count += 1;
    });
    const addonDaily = rangeDates.map(d => addonByDate[d] || { date: d, revenue: 0, cost: 0, profit: 0, qty: 0, count: 0 });
    const addonSummary = {
      revenue: addonRevenue,
      cost: addonCost,
      profit: addonProfit,
      qty: addonQty,
      count: addonSales.length,
      topAddons: topAddons,
      daily: addonDaily
    };

    const result = {
      success: true,
      message: 'Dashboard berhasil dimuat.',
      data: {
        today: today,
        range: { startDate: range.startDate, endDate: range.endDate, days: rangeDates.length },
        cards: {
          // Catatan: nama field "today" dipertahankan agar kompatibel dengan UI lama.
          // Isinya sekarang adalah total range yang dipilih.
          revenueToday: rangeRevenue,
          subtotalToday: rangeSubtotal,
          costToday: rangeCost,
          grossProfitToday: rangeGrossProfit,
          expensesToday: rangeExpenseTotal,
          sharingExpensesToday: rangeSharingExpenseTotal,
          personalExpensesToday: rangePersonalExpenseTotal,
          netProfitToday: rangeNetProfit,
          transactionsToday: rangeTransactions,
          topTherapistToday: topTherapistRange,
          // Field "monthly..." sekarang berisi total range (sama dengan range...).
          // Dipertahankan supaya tidak break UI lama bila ada referensi.
          monthlyRevenue: rangeSummary.revenue,
          monthlyGrossProfit: rangeSummary.grossProfit,
          monthlyExpenses: rangeSummary.expenses,
          monthlyNetProfit: rangeSummary.netProfit,
          monthlyTransactions: rangeSummary.transactions
        },
        last7Revenue: last7Revenue,
        paymentSummary: paymentSummary,
        topTherapists: topTherapists,
        // monthlySales / monthlySummary sekarang berisi breakdown harian dari range
        monthlySales: rangeDailyBreakdown,
        monthlySummary: rangeSummary,
        sharingExpenses: sharingExpenses,
        personalExpenses: personalExpenses,
        addonSummary: addonSummary
      }
    };
    // Simpan ke cache (60 detik)
    try { cache.put(cacheKey, JSON.stringify(result), 60); } catch (e) { /* abaikan jika payload terlalu besar */ }
    return result;
  } catch (error) { return POS_errorResponse_(error); }
}

/**
 * Validasi & normalisasi range tanggal dari frontend.
 * Default = single day (today). Swap otomatis jika start > end.
 * Maksimal range 366 hari untuk menjaga performa.
 */
function POS_resolveDashboardRange_(filter, todayStr) {
  const fallback = { startDate: todayStr, endDate: todayStr };
  if (!filter || typeof filter !== 'object') return fallback;

  let start = POS_toString_(filter.startDate).trim();
  let end = POS_toString_(filter.endDate).trim();
  const dateRegex = /^\d{4}-\d{2}-\d{2}$/;

  if (!dateRegex.test(start) && !dateRegex.test(end)) return fallback;
  if (!dateRegex.test(start)) start = end;
  if (!dateRegex.test(end)) end = start;

  // swap jika terbalik
  if (start > end) { const tmp = start; start = end; end = tmp; }

  // batasi maksimal 366 hari
  const startD = POS_parseDateString_(start);
  const endD = POS_parseDateString_(end);
  if (!startD || !endD) return fallback;
  const diffDays = Math.round((endD.getTime() - startD.getTime()) / (24 * 60 * 60 * 1000));
  if (diffDays > 365) {
    const cappedStart = POS_addDays_(endD, -365);
    return { startDate: POS_formatDate_(cappedStart), endDate: end };
  }

  return { startDate: start, endDate: end };
}

/** Parse 'yyyy-MM-dd' menjadi Date (lokal). Return null jika invalid. */
function POS_parseDateString_(str) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(str || '');
  if (!m) return null;
  const y = Number(m[1]), mo = Number(m[2]) - 1, d = Number(m[3]);
  const dt = new Date(y, mo, d);
  if (dt.getFullYear() !== y || dt.getMonth() !== mo || dt.getDate() !== d) return null;
  return dt;
}

/** Generate semua tanggal (string yyyy-MM-dd) inklusif antara start..end. */
function POS_getDateStringsBetween_(startStr, endStr) {
  const startD = POS_parseDateString_(startStr);
  const endD = POS_parseDateString_(endStr);
  if (!startD || !endD) return [];
  const dates = [];
  let cursor = startD;
  while (cursor.getTime() <= endD.getTime()) {
    dates.push(POS_formatDate_(cursor));
    cursor = POS_addDays_(cursor, 1);
  }
  return dates;
}

/** Cek apakah value tanggal berada di range [startStr, endStr] inklusif. */
function POS_isDateInRange_(value, startStr, endStr) {
  if (!value) return false;
  let dateStr;
  if (value instanceof Date) {
    dateStr = POS_formatDate_(value);
  } else {
    dateStr = POS_toString_(value).trim().slice(0, 10);
  }
  return dateStr >= startStr && dateStr <= endStr;
}

/** 7 hari terakhir berakhir di endDateStr (inklusif). */
function POS_getLast7DateStringsFrom_(endDateStr) {
  const endD = POS_parseDateString_(endDateStr) || POS_now_();
  const dates = [];
  for (let i = 6; i >= 0; i--) {
    dates.push(POS_formatDate_(POS_addDays_(endD, -i)));
  }
  return dates;
}


function POS_getCurrentMonthDateStrings_() {
  const today = POS_now_();
  const year = today.getFullYear();
  const month = today.getMonth();
  const lastDay = new Date(year, month + 1, 0).getDate();
  const dates = [];

  for (let day = 1; day <= lastDay; day++) {
    dates.push(POS_formatDate_(new Date(year, month, day)));
  }

  return dates;
}

function POS_calculateTopProduct_(items) {
  const topProducts = POS_calculateTopProducts_(items, 1);

  if (topProducts.length === 0) {
    return {
      menuName: '-',
      qty: 0,
      amount: 0
    };
  }

  return topProducts[0];
}

function POS_calculateTopProducts_(items, limit) {
  const map = {};

  items.forEach(item => {
    const menuName = POS_toString_(item.Menu_Name);
    const qty = POS_toNumber_(item.Qty);
    const amount = POS_toNumber_(item.Amount);

    if (!menuName) {
      return;
    }

    if (!map[menuName]) {
      map[menuName] = {
        menuName: menuName,
        qty: 0,
        amount: 0
      };
    }

    map[menuName].qty += qty;
    map[menuName].amount += amount;
  });

  return Object.values(map)
    .sort((a, b) => {
      if (b.qty !== a.qty) {
        return b.qty - a.qty;
      }

      return b.amount - a.amount;
    })
    .slice(0, limit || 5);
}

function POS_calculatePaymentSummary_(sales) {
  const summary = {
    Cash: 0,
    QRIS: 0,
    Transfer: 0
  };

  sales.forEach(row => {
    const method = POS_toString_(row.Payment_Method);
    const amount = POS_toNumber_(row.Rounded_Total || row.Grand_Total);

    if (summary[method] === undefined) {
      summary[method] = 0;
    }

    summary[method] += amount;
  });

  return summary;
}


function POS_sumSalesAmount_(sales) { return sales.reduce((sum, row) => sum + POS_toNumber_(row.Rounded_Total || row.Grand_Total), 0); }
function POS_getMenuCostMap_() { const map={}; POS_readObjects_(POS_SHEET.MENU).forEach(row => { const id=POS_toString_(row.Menu_ID); if(id) map[id]=POS_toNumber_(row.Cost || row.HPP || row.Modal); }); return map; }
function POS_getItemCost_(item, menuCostMap) { const existing=POS_toNumber_(item.Cost); if (existing > 0 || POS_toString_(item.Cost) !== '') return existing; return POS_toNumber_((menuCostMap || {})[POS_toString_(item.Menu_ID)]) * POS_toNumber_(item.Qty); }
function POS_sumItemCost_(items, menuCostMap) { return items.reduce((sum, item) => sum + POS_getItemCost_(item, menuCostMap), 0); }
function POS_sumItemGrossProfit_(items, menuCostMap) { return items.reduce((sum, item) => { const existing=POS_toNumber_(item.Gross_Profit); if (existing !== 0 || POS_toString_(item.Gross_Profit) !== '') return sum + existing; return sum + POS_toNumber_(item.Amount) - POS_getItemCost_(item, menuCostMap); }, 0); }
function POS_sumExpenses_(expenses) { return expenses.reduce((sum, row) => sum + POS_toNumber_(row.Amount), 0); }
// Catatan: fungsi ini mengasumsikan expenses sudah dinormalisasi (Expense_Type sudah diisi).
// Jangan panggil dengan data mentah — gunakan setelah .map(POS_normalizeExpenseRow_).
function POS_sumExpensesByType_(expenses, type) { return expenses.filter(row => POS_toString_(row.Expense_Type || 'Sharing') === type).reduce((sum, row) => sum + POS_toNumber_(row.Amount), 0); }
// ── Cache invalidation ─────────────────────────────────────────────────────
function POS_invalidateDashboardCache_() {
  try {
    const cache = CacheService.getScriptCache();
    // Naikkan versi — getDashboardData akan membaca versi ini sebagai bagian dari cache key
    const v = (parseInt(cache.get('dash_ver') || '0', 10) + 1) % 9999;
    cache.put('dash_ver', String(v), 3600);
  } catch (e) {}
}
function POS_calculateTopCashiers_(sales, limit) { const map={}; sales.forEach(row => { const cashierName=POS_toString_(row.Cashier_Name) || '-'; const amount=POS_toNumber_(row.Rounded_Total || row.Grand_Total); if(!map[cashierName]) map[cashierName]={cashierName:cashierName,transactions:0,revenue:0}; map[cashierName].transactions += 1; map[cashierName].revenue += amount; }); return Object.values(map).sort((a,b)=> b.revenue !== a.revenue ? b.revenue-a.revenue : b.transactions-a.transactions).slice(0, limit || 5); }
function POS_calculateTopTherapists_(pointRows, limit) { const map={}; pointRows.forEach(row => { const id=POS_toString_(row.Therapist_ID) || '-'; const name=POS_toString_(row.Therapist_Name) || '-'; if(!map[id]) map[id]={therapistId:id,therapistName:name,totalPoint:0,totalAmount:0,count:0}; map[id].totalPoint += POS_toNumber_(row.Total_Point); map[id].totalAmount += POS_toNumber_(row.Amount); map[id].count += 1; }); return Object.values(map).sort((a,b)=> b.totalPoint !== a.totalPoint ? b.totalPoint-a.totalPoint : b.totalAmount-a.totalAmount).slice(0, limit || 5); }
function POS_groupExpensesByCategory_(expenses) { const map={}; expenses.forEach(row => { const category=POS_toString_(row.Category) || 'Lainnya'; if(!map[category]) map[category]={category:category,amount:0,count:0}; map[category].amount += POS_toNumber_(row.Amount); map[category].count += 1; }); return Object.values(map).sort((a,b)=>b.amount-a.amount); }
function POS_groupPersonalExpenses_(expenses) { const map={}; expenses.forEach(row => { const cashier=POS_toString_(row.Personal_Cashier) || POS_toString_(row.Cashier_Name) || '-'; const category=POS_toString_(row.Category) || 'Kasbon/Fee'; const key=cashier+'|'+category; if(!map[key]) map[key]={cashierName:cashier,category:category,amount:0,count:0}; map[key].amount += POS_toNumber_(row.Amount); map[key].count += 1; }); return Object.values(map).sort((a,b)=>b.amount-a.amount); }

/* =====================================================
   RECEIPT TEXT GENERATOR
   Ini belum dipakai di frontend Batch 3,
   tapi sudah disiapkan untuk Batch 5 / RAWBT.
===================================================== */

function generatePaymentReceiptText(receipt) {
  try {
    const lines = [];

    const store = receipt.store || {};
    const trx = receipt.transaction || {};
    const items = receipt.items || [];
    const totals = receipt.totals || {};

    lines.push(POS_centerText_(POS_toString_(store.name), 32));

    if (store.address) {
      lines.push(POS_centerText_(POS_toString_(store.address), 32));
    }

    if (store.phone) {
      lines.push(POS_centerText_(POS_toString_(store.phone), 32));
    }

    lines.push('--------------------------------');
    lines.push('No: ' + POS_toString_(trx.transactionId));
    lines.push('Tgl: ' + POS_toString_(trx.date) + ' ' + POS_toString_(trx.time));
    lines.push('Kasir: ' + POS_toString_(trx.cashierName));
    lines.push('--------------------------------');

    items.forEach(item => {
      lines.push(POS_toString_(item.menuName));
      const left = POS_toString_(item.qty) + ' x ' + POS_formatRupiah_(item.price);
      const right = POS_formatRupiah_(item.amount);
      lines.push(POS_leftRightText_(left, right, 32));
    });

    lines.push('--------------------------------');
    lines.push(POS_leftRightText_('Subtotal', POS_formatRupiah_(totals.subtotal), 32));
    lines.push(POS_leftRightText_('Tax ' + totals.taxRate + '%', POS_formatRupiah_(totals.taxAmount), 32));
    lines.push(POS_leftRightText_('Service ' + totals.serviceRate + '%', POS_formatRupiah_(totals.serviceAmount), 32));
    lines.push(POS_leftRightText_('Total', POS_formatRupiah_(totals.roundedTotal), 32));
    lines.push(POS_leftRightText_('Bayar', POS_formatRupiah_(totals.paidAmount), 32));
    lines.push(POS_leftRightText_('Kembali', POS_formatRupiah_(totals.changeAmount), 32));
    lines.push('--------------------------------');

    if (store.footer) {
      lines.push(POS_centerText_(POS_toString_(store.footer), 32));
    }

    return {
      success: true,
      message: 'Receipt text berhasil dibuat.',
      data: lines.join('\n')
    };
  } catch (error) {
    return POS_errorResponse_(error);
  }
}

function generateKitchenReceiptText(receipt) {
  try {
    const lines = [];

    const trx = receipt.transaction || {};
    const items = receipt.items || [];

    lines.push(POS_centerText_('KITCHEN ORDER', 32));
    lines.push('--------------------------------');
    lines.push('No: ' + POS_toString_(trx.transactionId));
    lines.push('Jam: ' + POS_toString_(trx.time));
    lines.push('Kasir: ' + POS_toString_(trx.cashierName));
    lines.push('--------------------------------');

    items.forEach(item => {
      lines.push(POS_toString_(item.qty) + 'x ' + POS_toString_(item.menuName));
    });

    lines.push('--------------------------------');

    return {
      success: true,
      message: 'Kitchen receipt text berhasil dibuat.',
      data: lines.join('\n')
    };
  } catch (error) {
    return POS_errorResponse_(error);
  }
}

function POS_centerText_(text, width) {
  const cleanText = POS_toString_(text);

  if (cleanText.length >= width) {
    return cleanText.slice(0, width);
  }

  const leftPadding = Math.floor((width - cleanText.length) / 2);
  return ' '.repeat(leftPadding) + cleanText;
}

function POS_leftRightText_(left, right, width) {
  const cleanLeft = POS_toString_(left);
  const cleanRight = POS_toString_(right);

  const space = width - cleanLeft.length - cleanRight.length;

  if (space <= 1) {
    return cleanLeft + ' ' + cleanRight;
  }

  return cleanLeft + ' '.repeat(space) + cleanRight;
}

/* =====================================================
   ERROR RESPONSE
===================================================== */

function POS_errorResponse_(error) {
  const message = error && error.message ? error.message : String(error);
  Logger.log('❌ POS Error: ' + message);
  // Non-blocking: jika log gagal, response utama tetap terkirim
  try { POS_logError_(POS_extractFnName_(error), message); } catch (e) {}
  return { success: false, message: message, data: null };
}

/** Ekstrak nama fungsi pemanggil dari stack trace. */
function POS_extractFnName_(error) {
  if (!error || !error.stack) return 'unknown';
  const lines = String(error.stack).split('\n');
  for (let i = 1; i < lines.length; i++) {
    const m = /at\s+([A-Za-z_]\w*)/.exec(lines[i]);
    if (m && m[1] !== 'POS_errorResponse_' && m[1] !== 'POS_extractFnName_') return m[1];
  }
  return 'unknown';
}

/* =====================================================
   SEND REPORT (Email & Telegram)
===================================================== */

/**
 * Kirim laporan harian ke Email dan/atau Telegram.
 * payload: { filter: { startDate, endDate }, channels: ['email','telegram'] }
 */
function sendReport(payload) {
  try {
    const filter = (payload && payload.filter) || {};
    const channels = (payload && payload.channels) || ['email', 'telegram'];

    const config = POS_getConfigMap_();
    const storeName = POS_toString_(config.STORE_NAME || 'POS System');
    const reportEmail = POS_toString_(config.REPORT_EMAIL || '').trim();
    const telegramToken = POS_toString_(config.TELEGRAM_BOT_TOKEN || '').trim();
    const telegramChatId = POS_toString_(config.TELEGRAM_CHAT_ID || '').trim();

    // Ambil data dashboard sesuai filter
    const dashResult = getDashboardData(filter);
    if (!dashResult.success) throw new Error('Gagal mengambil data: ' + dashResult.message);

    const data = dashResult.data;
    const cards = data.cards || {};
    const range = data.range || {};
    const summary = data.monthlySummary || {};

    // ── Hitung nilai bagi hasil ───────────────────────────────────────────
    const omzet = Number(cards.revenueToday || 0);
    const operasional = Number(cards.sharingExpensesToday || 0);
    const kasbon = Number(cards.personalExpensesToday || 0);
    const labaTerapis = omzet / 2;
    const sisaInvestor = (omzet / 2) - operasional;
    const labaInvestor1 = sisaInvestor * 0.60;
    const labaInvestor2 = sisaInvestor * 0.40;
    const transaksi = Number(cards.transactionsToday || 0);

    // Label range
    const rangeLabel = range.startDate === range.endDate
      ? range.startDate
      : range.startDate + ' s/d ' + range.endDate;
    const days = range.days || 1;

    // ── Format pesan teks (untuk Telegram) ───────────────────────────────
    const sep = '━━━━━━━━━━━━━━━━━━━━';
    const lines = [
      '📊 <b>LAPORAN ' + storeName.toUpperCase() + '</b>',
      '📅 ' + rangeLabel + (days > 1 ? ' (' + days + ' hari)' : ''),
      sep,
      '💰 <b>Total Omzet</b>: ' + POS_formatRupiah_(omzet),
      '🧾 <b>Transaksi</b>: ' + transaksi + ' transaksi',
      sep,
      '🏢 <b>Pengeluaran Operasional</b>: ' + POS_formatRupiah_(operasional),
      '💳 <b>Total Kasbon</b>: ' + POS_formatRupiah_(kasbon),
      sep,
      '👩‍⚕️ <b>Laba Terapis (50%)</b>: ' + POS_formatRupiah_(labaTerapis),
      sep,
      '👤 <b>Laba Investor 1 (60%)</b>: ' + POS_formatRupiah_(labaInvestor1),
      '   <i>60% × (omzet/2 - operasional)</i>',
      '👤 <b>Laba Investor 2 (40%)</b>: ' + POS_formatRupiah_(labaInvestor2),
      '   <i>40% × (omzet/2 - operasional)</i>',
      sep,
    ];

    // Breakdown pembayaran
    const payment = data.paymentSummary || {};
    lines.push('💵 <b>Metode Pembayaran</b>:');
    ['Cash', 'QRIS', 'Transfer'].forEach(function(m) {
      if (payment[m]) lines.push('   ' + m + ': ' + POS_formatRupiah_(payment[m]));
    });

    // Top terapis
    const topTherapists = data.topTherapists || [];
    if (topTherapists.length > 0) {
      lines.push(sep);
      lines.push('🏆 <b>Top Peringkat Terapis</b>:');
      topTherapists.forEach(function(t, i) {
        lines.push('   ' + (i + 1) + '. ' + t.therapistName + ' — ' + t.totalPoint + ' pt (' + POS_formatRupiah_(t.totalAmount) + ')');
      });
    }

    // Pengeluaran operasional
    const sharingExp = data.sharingExpenses || [];
    if (sharingExp.length > 0) {
      lines.push(sep);
      lines.push('🏢 <b>Detail Operasional</b>:');
      sharingExp.forEach(function(e) {
        lines.push('   • ' + e.category + ': ' + POS_formatRupiah_(e.amount));
      });
    }

    // Kasbon
    const personalExp = data.personalExpenses || [];
    if (personalExp.length > 0) {
      lines.push(sep);
      lines.push('💳 <b>Detail Kasbon</b>:');
      personalExp.forEach(function(e) {
        lines.push('   • ' + e.cashierName + ': ' + POS_formatRupiah_(e.amount));
      });
    }

    // Pembukuan tambahan (terpisah)
    const addon = data.addonSummary || {};
    if ((addon.count || 0) > 0) {
      lines.push(sep);
      lines.push('🥤 <b>Pembukuan Tambahan (Terpisah)</b>:');
      lines.push('   Omzet: ' + POS_formatRupiah_(addon.revenue || 0));
      lines.push('   Modal: ' + POS_formatRupiah_(addon.cost || 0));
      lines.push('   Laba: ' + POS_formatRupiah_(addon.profit || 0));
      lines.push('   Item terjual: ' + Number(addon.qty || 0) + ' pcs');
      (addon.topAddons || []).forEach(function(t, i) {
        lines.push('   ' + (i + 1) + '. ' + t.name + ' — ' + t.qty + ' pcs (' + POS_formatRupiah_(t.amount) + ')');
      });
    }

    lines.push(sep);
    lines.push('🕐 Dikirim: ' + POS_nowString_());

    const telegramMessage = lines.join('\n');

    // ── Format HTML untuk Email ───────────────────────────────────────────
    const emailHtml = POS_buildReportEmailHtml_(storeName, rangeLabel, days, {
      omzet, operasional, kasbon, labaTerapis, labaInvestor1, labaInvestor2,
      transaksi, sisaInvestor, payment, topTherapists, sharingExp, personalExp,
      addon: addon
    });

    const emailSubject = '[' + storeName + '] Laporan ' + rangeLabel;

    // ── Kirim ─────────────────────────────────────────────────────────────
    const results = {};

    if (channels.indexOf('email') !== -1) {
      if (!reportEmail) {
        const msg = 'Email penerima belum dikonfigurasi di Config.';
        try { POS_logError_('sendReport[email]', msg); } catch(e) {}
        results.email = { success: false, message: msg };
      } else {
        try {
          MailApp.sendEmail({ to: reportEmail, subject: emailSubject, htmlBody: emailHtml });
          results.email = { success: true, message: 'Email terkirim ke ' + reportEmail };
        } catch (e) {
          try { POS_logError_('sendReport[email]', e.message); } catch(le) {}
          results.email = { success: false, message: 'Gagal kirim email: ' + e.message };
        }
      }
    }

    if (channels.indexOf('telegram') !== -1) {
      if (!telegramToken || !telegramChatId) {
        const msg = 'Token atau Chat ID Telegram belum dikonfigurasi.';
        try { POS_logError_('sendReport[telegram]', msg); } catch(e) {}
        results.telegram = { success: false, message: msg };
      } else {
        try {
          POS_sendTelegram_(telegramToken, telegramChatId, telegramMessage);
          results.telegram = { success: true, message: 'Pesan Telegram terkirim.' };
        } catch (e) {
          try { POS_logError_('sendReport[telegram]', e.message); } catch(le) {}
          results.telegram = { success: false, message: 'Gagal kirim Telegram: ' + e.message };
        }
      }
    }

    const allOk = Object.values(results).every(function(r) { return r.success; });
    return {
      success: allOk,
      message: allOk ? 'Laporan berhasil dikirim.' : 'Sebagian laporan gagal dikirim.',
      data: results
    };
  } catch (error) { return POS_errorResponse_(error); }
}

/** Kirim pesan teks ke Telegram via Bot API. */
function POS_sendTelegram_(token, chatId, text) {
  const url = 'https://api.telegram.org/bot' + token + '/sendMessage';
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      chat_id: chatId,
      text: text,
      parse_mode: 'HTML'
    }),
    muteHttpExceptions: true
  });
  const body = JSON.parse(response.getContentText());
  if (!body.ok) throw new Error('Telegram API error: ' + (body.description || 'Unknown'));
}

/** Build HTML email laporan. */
function POS_buildReportEmailHtml_(storeName, rangeLabel, days, d) {
  function row(label, value, bold) {
    return '<tr><td style="padding:6px 12px;color:#6b7280;">' + label + '</td>'
      + '<td style="padding:6px 12px;text-align:right;' + (bold ? 'font-weight:700;' : '') + '">' + value + '</td></tr>';
  }
  function section(title) {
    return '<tr><td colspan="2" style="padding:10px 12px 4px;font-weight:700;background:#f3f4f6;color:#374151;">' + title + '</td></tr>';
  }

  let html = '<div style="font-family:sans-serif;max-width:560px;margin:0 auto;border:1px solid #e5e7eb;border-radius:12px;overflow:hidden;">';
  html += '<div style="background:#1d4ed8;color:#fff;padding:20px 24px;">';
  html += '<h2 style="margin:0;font-size:18px;">📊 Laporan ' + storeName + '</h2>';
  html += '<p style="margin:4px 0 0;opacity:.85;">📅 ' + rangeLabel + (days > 1 ? ' (' + days + ' hari)' : '') + '</p></div>';
  html += '<table style="width:100%;border-collapse:collapse;">';

  html += section('📈 Ringkasan Utama');
  html += row('Total Omzet', POS_formatRupiah_(d.omzet), true);
  html += row('Total Transaksi', d.transaksi + ' trx', false);

  html += section('💸 Pengeluaran');
  html += row('Operasional', POS_formatRupiah_(d.operasional), false);
  html += row('Kasbon Terapis', POS_formatRupiah_(d.kasbon), false);

  html += section('🤝 Bagi Hasil');
  html += row('Laba Terapis (50%)', POS_formatRupiah_(d.labaTerapis), true);
  html += row('Sisa Investor (50% - Operasional)', POS_formatRupiah_(d.sisaInvestor), false);
  html += row('Laba Investor 1 (60%)', POS_formatRupiah_(d.labaInvestor1), true);
  html += row('Laba Investor 2 (40%)', POS_formatRupiah_(d.labaInvestor2), true);

  html += section('💳 Metode Pembayaran');
  ['Cash', 'QRIS', 'Transfer'].forEach(function(m) {
    if (d.payment[m]) html += row(m, POS_formatRupiah_(d.payment[m]), false);
  });

  if (d.topTherapists && d.topTherapists.length > 0) {
    html += section('🏆 Top Peringkat Terapis');
    d.topTherapists.forEach(function(t, i) {
      html += row((i+1) + '. ' + t.therapistName, t.totalPoint + ' pt (' + POS_formatRupiah_(t.totalAmount) + ')', false);
    });
  }

  if (d.sharingExp && d.sharingExp.length > 0) {
    html += section('🏢 Detail Operasional');
    d.sharingExp.forEach(function(e) {
      html += row(e.category, POS_formatRupiah_(e.amount), false);
    });
  }

  if (d.personalExp && d.personalExp.length > 0) {
    html += section('💳 Detail Kasbon');
    d.personalExp.forEach(function(e) {
      html += row(e.cashierName, POS_formatRupiah_(e.amount), false);
    });
  }

  // Pembukuan tambahan (terpisah)
  const addon = d.addon || {};
  if ((addon.count || 0) > 0) {
    html += section('🥤 Pembukuan Tambahan (Terpisah)');
    html += row('Omzet Tambahan', POS_formatRupiah_(addon.revenue || 0), true);
    html += row('Modal', POS_formatRupiah_(addon.cost || 0), false);
    html += row('Laba Tambahan', POS_formatRupiah_(addon.profit || 0), true);
    html += row('Item Terjual', Number(addon.qty || 0) + ' pcs', false);
    (addon.topAddons || []).forEach(function(t, i) {
      html += row((i+1) + '. ' + t.name, t.qty + ' pcs (' + POS_formatRupiah_(t.amount) + ')', false);
    });
  }

  html += '</table>';
  html += '<div style="padding:12px 24px;background:#f9fafb;color:#9ca3af;font-size:12px;">Dikirim otomatis oleh ' + storeName + ' POS · ' + POS_nowString_() + '</div>';
  html += '</div>';
  return html;
}

/* =====================================================
   DAILY AUTO REPORT (Trigger 22:00 WIB)
===================================================== */

/** Dipanggil time-based trigger setiap hari jam 22:00. Kirim laporan hari berjalan. */
function POS_dailyReportTrigger_() {
  try {
    const today = POS_todayString_();
    const res = sendReport({ filter: { startDate: today, endDate: today }, channels: ['email', 'telegram'] });
    if (!res.success) {
      try { POS_logError_('POS_dailyReportTrigger_', res.message || 'Laporan harian sebagian gagal.'); } catch(le) {}
    }
  } catch (e) {
    try { POS_logError_('POS_dailyReportTrigger_', e.message); } catch(le) {}
    Logger.log('❌ Laporan harian gagal: ' + e.message);
  }
}

/** Setup trigger laporan harian otomatis jam 22:00 WIB. Idempotent. */
function setupDailyReportTrigger() {
  try {
    ScriptApp.getProjectTriggers().forEach(function(t) {
      if (t.getHandlerFunction() === 'POS_dailyReportTrigger_') ScriptApp.deleteTrigger(t);
    });
    // Timezone project = Asia/Jakarta, jadi atHour(22) = 22:00 WIB
    ScriptApp.newTrigger('POS_dailyReportTrigger_').timeBased().everyDays(1).atHour(22).create();
    return { success: true, message: 'Laporan harian otomatis aktif (setiap hari jam 22:00 WIB).' };
  } catch (e) { return POS_errorResponse_(e); }
}

/** Nonaktifkan trigger laporan harian. */
function disableDailyReportTrigger() {
  try {
    let removed = 0;
    ScriptApp.getProjectTriggers().forEach(function(t) {
      if (t.getHandlerFunction() === 'POS_dailyReportTrigger_') { ScriptApp.deleteTrigger(t); removed++; }
    });
    return { success: true, message: removed > 0 ? 'Laporan harian otomatis dinonaktifkan.' : 'Tidak ada trigger aktif.' };
  } catch (e) { return POS_errorResponse_(e); }
}

/** Status trigger laporan harian (untuk UI). */
function getDailyReportTriggerStatus() {
  try {
    const active = ScriptApp.getProjectTriggers().some(function(t) {
      return t.getHandlerFunction() === 'POS_dailyReportTrigger_';
    });
    return { success: true, data: { active: active } };
  } catch (e) { return { success: true, data: { active: false } }; }
}

/* =====================================================
   EXPORT DASHBOARD KE PDF
===================================================== */

/**
 * Generate PDF dashboard sesuai filter.
 * Isi: ringkasan keuangan + bagi hasil, rekap transaksi harian,
 * detail seluruh pengeluaran, dan ringkasan pembayaran.
 * Return base64 untuk di-download di browser.
 */
function exportDashboardPdf(filter) {
  try {
    const dashResult = getDashboardData(filter);
    if (!dashResult.success) throw new Error('Gagal ambil data: ' + dashResult.message);
    const data = dashResult.data;
    const range = data.range || {};

    // Detail seluruh pengeluaran dalam range (bukan grouped)
    const expenseDetail = POS_readObjects_(POS_SHEET.EXPENSES)
      .map(POS_normalizeExpenseRow_)
      .filter(r => POS_toString_(r.Expense_ID))
      .filter(r => POS_isDateInRange_(r.Date, range.startDate, range.endDate))
      .sort((a, b) => (POS_toString_(a.Date) + ' ' + POS_toString_(a.Time)).localeCompare(POS_toString_(b.Date) + ' ' + POS_toString_(b.Time)));

    const config = POS_getConfigMap_();
    const storeName = POS_toString_(config.STORE_NAME || 'POS System');

    const html = POS_buildDashboardPdfHtml_(storeName, data, expenseDetail);

    const rangeTag = range.startDate === range.endDate
      ? range.startDate
      : range.startDate + '_sd_' + range.endDate;
    const filename = 'Laporan_' + storeName.replace(/[^\w]+/g, '_') + '_' + rangeTag + '.pdf';

    const blob = Utilities.newBlob(html, 'text/html', 'report.html').getAs('application/pdf');
    const base64 = Utilities.base64Encode(blob.getBytes());

    return { success: true, message: 'PDF berhasil dibuat.', data: { base64: base64, filename: filename } };
  } catch (e) { return POS_errorResponse_(e); }
}

/** Build HTML untuk PDF dashboard (print-friendly A4). */
function POS_buildDashboardPdfHtml_(storeName, data, expenseDetail) {
  const cards = data.cards || {};
  const range = data.range || {};
  const summary = data.monthlySummary || {};
  const dailySales = data.monthlySales || [];
  const payment = data.paymentSummary || {};

  const omzet = Number(cards.revenueToday || 0);
  const operasional = Number(cards.sharingExpensesToday || 0);
  const kasbon = Number(cards.personalExpensesToday || 0);
  const labaTerapis = omzet / 2;
  const sisaInvestor = (omzet / 2) - operasional;
  const labaInvestor1 = sisaInvestor * 0.60;
  const labaInvestor2 = sisaInvestor * 0.40;
  const transaksi = Number(cards.transactionsToday || 0);

  const days = range.days || 1;
  const rangeLabel = range.startDate === range.endDate
    ? range.startDate
    : range.startDate + ' s/d ' + range.endDate;

  function esc(s) { return POS_toString_(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
  function rp(n) { return POS_formatRupiah_(n); }

  let html = ''
    + '<!DOCTYPE html><html><head><meta charset="utf-8"><style>'
    + '* { font-family: Arial, sans-serif; box-sizing: border-box; }'
    + 'body { margin: 0; padding: 24px; color: #1f2937; font-size: 12px; }'
    + 'h1 { font-size: 20px; margin: 0 0 4px; color: #1d4ed8; }'
    + 'h2 { font-size: 14px; margin: 20px 0 8px; padding-bottom: 4px; border-bottom: 2px solid #e5e7eb; color: #374151; }'
    + '.sub { color: #6b7280; font-size: 12px; margin-bottom: 4px; }'
    + 'table { width: 100%; border-collapse: collapse; margin-bottom: 8px; }'
    + 'th, td { padding: 6px 8px; text-align: left; border-bottom: 1px solid #e5e7eb; }'
    + 'th { background: #f3f4f6; font-size: 11px; text-transform: uppercase; color: #6b7280; }'
    + 'td.num, th.num { text-align: right; }'
    + '.cards { display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 8px; }'
    + '.card { flex: 1 1 30%; border: 1px solid #e5e7eb; border-radius: 8px; padding: 10px 12px; min-width: 150px; }'
    + '.card .label { color: #6b7280; font-size: 11px; }'
    + '.card .value { font-size: 16px; font-weight: 700; margin-top: 2px; }'
    + '.card.highlight { background: #eff6ff; border-color: #bfdbfe; }'
    + '.totrow td { font-weight: 700; background: #f9fafb; }'
    + '.footer { margin-top: 24px; color: #9ca3af; font-size: 10px; text-align: center; border-top: 1px solid #e5e7eb; padding-top: 8px; }'
    + '</style></head><body>';

  // Header
  html += '<h1>📊 Laporan ' + esc(storeName) + '</h1>';
  html += '<div class="sub">Periode: <strong>' + esc(rangeLabel) + '</strong>' + (days > 1 ? ' (' + days + ' hari)' : '') + '</div>';
  html += '<div class="sub">Dicetak: ' + esc(POS_nowString_()) + '</div>';

  // Ringkasan keuangan & bagi hasil (cards)
  html += '<h2>Ringkasan</h2><div class="cards">';
  html += '<div class="card highlight"><div class="label">Total Omzet</div><div class="value">' + rp(omzet) + '</div></div>';
  html += '<div class="card"><div class="label">Total Transaksi</div><div class="value">' + transaksi + ' trx</div></div>';
  html += '<div class="card"><div class="label">Pengeluaran Operasional</div><div class="value">' + rp(operasional) + '</div></div>';
  html += '<div class="card"><div class="label">Total Kasbon</div><div class="value">' + rp(kasbon) + '</div></div>';
  html += '<div class="card highlight"><div class="label">Laba Terapis (50%)</div><div class="value">' + rp(labaTerapis) + '</div></div>';
  html += '<div class="card"><div class="label">Laba Investor 1 (60%)</div><div class="value">' + rp(labaInvestor1) + '</div></div>';
  html += '<div class="card"><div class="label">Laba Investor 2 (40%)</div><div class="value">' + rp(labaInvestor2) + '</div></div>';
  html += '</div>';

  // Rekap transaksi harian
  html += '<h2>Rekap Transaksi Harian</h2>';
  html += '<table><thead><tr><th>Tanggal</th><th class="num">Transaksi</th><th class="num">Omzet</th><th class="num">Gross Profit</th><th class="num">Pengeluaran</th><th class="num">Net Profit</th></tr></thead><tbody>';
  if (dailySales.length === 0) {
    html += '<tr><td colspan="6">Tidak ada transaksi pada periode ini.</td></tr>';
  } else {
    dailySales.forEach(function(r) {
      html += '<tr><td>' + esc(r.date) + '</td>'
        + '<td class="num">' + Number(r.transactions || 0) + '</td>'
        + '<td class="num">' + rp(r.revenue) + '</td>'
        + '<td class="num">' + rp(r.grossProfit) + '</td>'
        + '<td class="num">' + rp(r.expenses) + '</td>'
        + '<td class="num">' + rp(r.netProfit) + '</td></tr>';
    });
    html += '<tr class="totrow"><td>TOTAL</td>'
      + '<td class="num">' + Number(summary.transactions || 0) + '</td>'
      + '<td class="num">' + rp(summary.revenue) + '</td>'
      + '<td class="num">' + rp(summary.grossProfit) + '</td>'
      + '<td class="num">' + rp(summary.expenses) + '</td>'
      + '<td class="num">' + rp(summary.netProfit) + '</td></tr>';
  }
  html += '</tbody></table>';

  // Detail seluruh pengeluaran
  html += '<h2>Detail Pengeluaran</h2>';
  html += '<table><thead><tr><th>Tanggal</th><th>Jenis</th><th>Penerima</th><th>Kategori</th><th>Keterangan</th><th class="num">Nilai</th></tr></thead><tbody>';
  if (expenseDetail.length === 0) {
    html += '<tr><td colspan="6">Tidak ada pengeluaran pada periode ini.</td></tr>';
  } else {
    let totalExp = 0;
    expenseDetail.forEach(function(e) {
      const type = POS_toString_(e.Expense_Type);
      let typeLabel = 'Operasional';
      if (type === 'Therapist') typeLabel = 'Kasbon Terapis';
      else if (type === 'Personal') typeLabel = 'Pribadi';
      const amount = POS_toNumber_(e.Amount);
      totalExp += amount;
      html += '<tr><td>' + esc(e.Date) + '<br><small style="color:#9ca3af;">' + esc(e.Time) + '</small></td>'
        + '<td>' + esc(typeLabel) + '</td>'
        + '<td>' + esc(e.Personal_Cashier || '-') + '</td>'
        + '<td>' + esc(e.Category || '-') + '</td>'
        + '<td>' + esc(e.Description || '-') + '</td>'
        + '<td class="num">' + rp(amount) + '</td></tr>';
    });
    html += '<tr class="totrow"><td colspan="5">TOTAL PENGELUARAN</td><td class="num">' + rp(totalExp) + '</td></tr>';
  }
  html += '</tbody></table>';

  // Ringkasan pembayaran
  html += '<h2>Ringkasan Metode Pembayaran</h2>';
  html += '<table><thead><tr><th>Metode</th><th class="num">Total</th></tr></thead><tbody>';
  ['Cash', 'QRIS', 'Transfer'].forEach(function(m) {
    html += '<tr><td>' + m + '</td><td class="num">' + rp(payment[m] || 0) + '</td></tr>';
  });
  html += '</tbody></table>';

  // Pembukuan tambahan (addon) — terpisah
  const addon = data.addonSummary || {};
  html += '<h2>Pembukuan Tambahan (Air Mineral / Snack) — Terpisah</h2>';
  html += '<div class="cards">';
  html += '<div class="card"><div class="label">Omzet Tambahan</div><div class="value">' + rp(addon.revenue || 0) + '</div></div>';
  html += '<div class="card"><div class="label">Modal</div><div class="value">' + rp(addon.cost || 0) + '</div></div>';
  html += '<div class="card highlight"><div class="label">Laba Tambahan</div><div class="value">' + rp(addon.profit || 0) + '</div></div>';
  html += '<div class="card"><div class="label">Item Terjual</div><div class="value">' + Number(addon.qty || 0) + ' pcs</div></div>';
  html += '</div>';
  const topAddons = addon.topAddons || [];
  if (topAddons.length > 0) {
    html += '<table><thead><tr><th>Item Tambahan Terlaris</th><th class="num">Qty</th><th class="num">Omzet</th></tr></thead><tbody>';
    topAddons.forEach(function(t) {
      html += '<tr><td>' + esc(t.name) + '</td><td class="num">' + Number(t.qty || 0) + '</td><td class="num">' + rp(t.amount) + '</td></tr>';
    });
    html += '</tbody></table>';
  } else {
    html += '<p style="color:#9ca3af;">Belum ada penjualan item tambahan pada periode ini.</p>';
  }

  html += '<div class="footer">Dibuat otomatis oleh ' + esc(storeName) + ' POS System · ' + esc(POS_nowString_()) + '</div>';
  html += '</body></html>';
  return html;
}

/* =====================================================
   ITEM TAMBAHAN (ADDON) — Pembukuan Terpisah
   Air mineral / makanan ringan yang dibeli customer.
   Tidak masuk bagi hasil terapis/investor.
===================================================== */

function POS_generateAddonId_() {
  const sheet = POS_getOrCreateSheet_(POS_SHEET.ADDON_MENU, ADDON_MENU_HEADERS_);
  const lastRow = sheet.getLastRow();
  let maxNum = 0;
  if (lastRow >= 2) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    ids.forEach(r => { const m = /^A(\d+)$/.exec(POS_toString_(r[0])); if (m) { const n = parseInt(m[1], 10); if (n > maxNum) maxNum = n; } });
  }
  return 'A' + ('000' + (maxNum + 1)).slice(-3);
}

function POS_generateAddonSaleId_(extra) {
  const now = POS_now_();
  const datePart = Utilities.formatDate(now, POS_getTimezone_(), 'yyyyMMdd');
  const timePart = Utilities.formatDate(now, POS_getTimezone_(), 'HHmmss');
  const rand = Math.floor(Math.random() * 900 + 100);
  return 'AS-' + datePart + '-' + timePart + '-' + rand + (extra !== undefined ? '-' + extra : '');
}

/** Bangun objek addon dari row sheet. */
function POS_buildAddonObject_(row) {
  const stock = POS_toNumber_(row.Stock);
  return {
    addonId: POS_toString_(row.Addon_ID),
    name: POS_toString_(row.Name),
    price: POS_toNumber_(row.Price),
    cost: POS_toNumber_(row.Cost),
    stock: stock,
    stockStatus: POS_computeStockStatus_(stock),
    active: POS_toBoolean_(row.Active),
    createdAt: POS_toString_(row.Created_At),
    updatedAt: POS_toString_(row.Updated_At)
  };
}

/** Ambil daftar item tambahan. activeOnly opsional. */
function getAddons(payload) {
  try {
    POS_getOrCreateSheet_(POS_SHEET.ADDON_MENU, ADDON_MENU_HEADERS_);
    const activeOnly = !!(payload && payload.activeOnly === true);
    const rows = POS_readObjects_(POS_SHEET.ADDON_MENU).filter(r => POS_toString_(r.Addon_ID));
    const addons = rows.map(POS_buildAddonObject_)
      .filter(a => activeOnly ? a.active : true)
      .sort((a, b) => a.name.localeCompare(b.name));
    return { success: true, message: 'Item tambahan dimuat.', data: { addons: addons } };
  } catch (e) { return POS_errorResponse_(e); }
}

/** Tambah / update item tambahan. */
function saveAddon(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const addonId = POS_toString_(payload && payload.addonId).trim();
    const name = POS_toString_(payload && payload.name).trim();
    const price = POS_toNumber_(payload && payload.price);
    const cost = POS_toNumber_(payload && payload.cost);
    const stock = POS_toNumber_(payload && payload.stock);
    const active = (payload && payload.active !== undefined) ? POS_toBoolean_(payload.active) : true;

    if (!name) throw new Error('Nama item tambahan wajib diisi.');
    if (price < 0) throw new Error('Harga tidak boleh negatif.');

    const sheet = POS_getOrCreateSheet_(POS_SHEET.ADDON_MENU, ADDON_MENU_HEADERS_);
    const headers = POS_getHeaders_(sheet);
    const headerMap = POS_getHeaderMap_(headers);
    const lastRow = sheet.getLastRow();
    const now = POS_nowString_();
    const status = POS_computeStockStatus_(stock);

    if (addonId) {
      let foundRow = -1;
      if (lastRow >= 2) {
        const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
        for (let i = 0; i < values.length; i++) {
          if (POS_toString_(values[i][headerMap.Addon_ID]) === addonId) { foundRow = i + 2; break; }
        }
      }
      if (foundRow === -1) throw new Error('Item tambahan tidak ditemukan: ' + addonId);
      // Edit item TIDAK mengubah stok — stok hanya via updateAddonStock (berpassword).
      sheet.getRange(foundRow, headerMap.Name + 1).setValue(name);
      sheet.getRange(foundRow, headerMap.Price + 1).setValue(price);
      sheet.getRange(foundRow, headerMap.Cost + 1).setValue(cost);
      sheet.getRange(foundRow, headerMap.Active + 1).setValue(active ? 'TRUE' : 'FALSE');
      sheet.getRange(foundRow, headerMap.Updated_At + 1).setValue(now);
      SpreadsheetApp.flush();
      return { success: true, message: 'Item tambahan diperbarui.', data: { addonId: addonId } };
    }

    const newId = POS_generateAddonId_();
    POS_appendObjects_(POS_SHEET.ADDON_MENU, [{
      Addon_ID: newId, Name: name, Price: price, Cost: cost, Stock: stock,
      Stock_Status: status, Active: active ? 'TRUE' : 'FALSE', Created_At: now, Updated_At: now
    }]);
    SpreadsheetApp.flush();
    return { success: true, message: 'Item tambahan ditambahkan.', data: { addonId: newId } };
  } catch (e) { return POS_errorResponse_(e); }
  finally { try { lock.releaseLock(); } catch (e) {} }
}

/** Update stock item tambahan (mode SET/ADD/REDUCE). */
function updateAddonStock(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const addonId = POS_toString_(payload && payload.addonId).trim();
    const mode = POS_toString_(payload && payload.mode).trim().toUpperCase();
    const stockValue = POS_toNumber_(payload && payload.stockValue);

    // Proteksi: ubah stok item tambahan butuh password dashboard
    const passcode = POS_toString_(payload && payload.passcode).trim();
    const correctPasscode = POS_getConfigValue_('DASHBOARD_PASSCODE', '');
    if (!passcode || passcode !== correctPasscode) {
      throw new Error('Password salah. Akses ditolak untuk mengubah stok item tambahan.');
    }

    if (!addonId) throw new Error('Addon ID wajib diisi.');
    if (!['SET', 'ADD', 'REDUCE'].includes(mode)) throw new Error('Mode stock tidak valid.');
    if (stockValue < 0) throw new Error('Nilai stock tidak boleh negatif.');

    const sheet = POS_getOrCreateSheet_(POS_SHEET.ADDON_MENU, ADDON_MENU_HEADERS_);
    const headers = POS_getHeaders_(sheet);
    const headerMap = POS_getHeaderMap_(headers);
    const lastRow = sheet.getLastRow();
    let foundRow = -1, current = 0;
    if (lastRow >= 2) {
      const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
      for (let i = 0; i < values.length; i++) {
        if (POS_toString_(values[i][headerMap.Addon_ID]) === addonId) {
          foundRow = i + 2; current = POS_toNumber_(values[i][headerMap.Stock]); break;
        }
      }
    }
    if (foundRow === -1) throw new Error('Item tambahan tidak ditemukan.');
    let newStock = current;
    if (mode === 'SET') newStock = stockValue;
    if (mode === 'ADD') newStock = current + stockValue;
    if (mode === 'REDUCE') newStock = Math.max(0, current - stockValue);
    const status = POS_computeStockStatus_(newStock);
    sheet.getRange(foundRow, headerMap.Stock + 1).setValue(newStock);
    sheet.getRange(foundRow, headerMap.Stock_Status + 1).setValue(status);
    sheet.getRange(foundRow, headerMap.Updated_At + 1).setValue(POS_nowString_());
    SpreadsheetApp.flush();
    return { success: true, message: 'Stock item tambahan diperbarui.', data: { addonId: addonId, newStock: newStock } };
  } catch (e) { return POS_errorResponse_(e); }
  finally { try { lock.releaseLock(); } catch (e) {} }
}

/** Soft-delete / aktifkan kembali item tambahan. */
function setAddonActive(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const addonId = POS_toString_(payload && payload.addonId).trim();
    const active = POS_toBoolean_(payload && payload.active);
    if (!addonId) throw new Error('Addon ID wajib diisi.');
    const sheet = POS_getOrCreateSheet_(POS_SHEET.ADDON_MENU, ADDON_MENU_HEADERS_);
    const headers = POS_getHeaders_(sheet);
    const headerMap = POS_getHeaderMap_(headers);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('Belum ada item tambahan.');
    const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    for (let i = 0; i < values.length; i++) {
      if (POS_toString_(values[i][headerMap.Addon_ID]) === addonId) {
        sheet.getRange(i + 2, headerMap.Active + 1).setValue(active ? 'TRUE' : 'FALSE');
        sheet.getRange(i + 2, headerMap.Updated_At + 1).setValue(POS_nowString_());
        SpreadsheetApp.flush();
        return { success: true, message: active ? 'Item tambahan diaktifkan.' : 'Item tambahan dinonaktifkan.' };
      }
    }
    throw new Error('Item tambahan tidak ditemukan.');
  } catch (e) { return POS_errorResponse_(e); }
  finally { try { lock.releaseLock(); } catch (e) {} }
}

/* =====================================================
   ERROR LOG
===================================================== */

const LOG_HEADERS_ = ['Date', 'Time', 'Function', 'Error', 'Logged_At'];

/**
 * Catat error ke sheet ERROR_LOG.
 * Sheet dibuat otomatis jika belum ada.
 * Maksimum 500 karakter per pesan agar sheet tidak bengkak.
 */
function POS_logError_(fnName, message) {
  const sheet = POS_getOrCreateSheet_(POS_SHEET.ERROR_LOG, LOG_HEADERS_);
  sheet.appendRow([
    POS_todayString_(),
    POS_timeString_(),
    POS_toString_(fnName).slice(0, 80),
    POS_toString_(message).slice(0, 500),
    POS_nowString_()
  ]);
}

/**
 * Hapus log lebih tua dari retentionDays (default 60, max 90, min 1).
 * Menggunakan bulk read → filter → rewrite agar O(1) write, bukan O(n) deleteRow.
 * Return: jumlah baris yang dihapus.
 */
function POS_purgeOldLogs_(retentionDays) {
  const days = Math.min(90, Math.max(1, Number(retentionDays) || 60));
  const cutoff = POS_formatDate_(POS_addDays_(POS_now_(), -days));
  const sheet = POS_getOrCreateSheet_(POS_SHEET.ERROR_LOG, LOG_HEADERS_);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0; // kosong atau hanya header

  const data = sheet.getRange(2, 1, lastRow - 1, LOG_HEADERS_.length).getValues();
  const kept  = data.filter(function(row) {
    return POS_toString_(row[0]).slice(0, 10) >= cutoff;
  });
  const deleted = data.length - kept.length;
  if (deleted === 0) return 0;

  // Kosongkan baris data lama, tulis kembali yang masih berlaku
  sheet.getRange(2, 1, lastRow - 1, LOG_HEADERS_.length).clearContent();
  if (kept.length > 0) {
    sheet.getRange(2, 1, kept.length, LOG_HEADERS_.length).setValues(kept);
  }
  Logger.log('🧹 Purge log: ' + deleted + ' baris dihapus (retensi ' + days + ' hari).');
  return deleted;
}

/** Trigger harian — dipanggil GAS time-based trigger. */
function POS_autoCleanupTrigger_() {
  try {
    const days = POS_toNumber_(POS_getConfigValue_('LOG_RETENTION_DAYS', '60')) || 60;
    POS_purgeOldLogs_(days);
  } catch (e) {
    try { POS_logError_('POS_autoCleanupTrigger_', e.message); } catch(le) {}
    Logger.log('❌ Auto-cleanup gagal: ' + e.message);
  }
}

/** Setup trigger harian jam 03:00 untuk auto-cleanup. Idempotent. */
function setupLogCleanupTrigger() {
  try {
    // Hapus trigger lama agar tidak duplikat
    ScriptApp.getProjectTriggers().forEach(function(t) {
      if (t.getHandlerFunction() === 'POS_autoCleanupTrigger_') ScriptApp.deleteTrigger(t);
    });
    ScriptApp.newTrigger('POS_autoCleanupTrigger_').timeBased().everyDays(1).atHour(3).create();
    return { success: true, message: 'Auto-cleanup trigger aktif (setiap hari jam 03:00).' };
  } catch (e) { return POS_errorResponse_(e); }
}

/** Ambil log terbaru (default 50 baris terakhir, max 200). */
function getErrorLogs(payload) {
  try {
    const limit = Math.min(200, Math.max(1, Number((payload && payload.limit) || 50)));
    const sheet = POS_getOrCreateSheet_(POS_SHEET.ERROR_LOG, LOG_HEADERS_);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, data: { logs: [], total: 0 } };

    const total = lastRow - 1;
    const startRow = Math.max(2, lastRow - limit + 1);
    const count = lastRow - startRow + 1;
    const values = sheet.getRange(startRow, 1, count, LOG_HEADERS_.length).getValues();

    const logs = values.reverse().map(function(row) {
      return {
        date: POS_toString_(row[0]),
        time: POS_toString_(row[1]),
        fn:   POS_toString_(row[2]),
        error: POS_toString_(row[3]),
        loggedAt: POS_toString_(row[4])
      };
    });
    return { success: true, data: { logs: logs, total: total } };
  } catch (e) { return POS_errorResponse_(e); }
}

/**
 * Terima error dari frontend (client-side) dan tulis ke ERROR_LOG.
 * Dipanggil via google.script.run — fire-and-forget dari browser.
 */
function logClientError(payload) {
  try {
    const context = POS_toString_(payload && payload.context).slice(0, 80) || '[CLIENT]';
    const message = POS_toString_(payload && payload.message).slice(0, 500) || 'unknown';
    POS_logError_(context, message);
    return { success: true };
  } catch (e) { return { success: false }; }
}

/** Purge manual dengan retensi yang bisa dikonfigurasi. */
function purgeErrorLogs(payload) {
  try {
    const days = Number((payload && payload.retentionDays) || 60);
    const deleted = POS_purgeOldLogs_(days);
    return {
      success: true,
      message: deleted + ' entri log dihapus (retensi ' + days + ' hari).',
      data: { deleted: deleted }
    };
  } catch (e) { return POS_errorResponse_(e); }
}

/* =====================================================
   MAINTENANCE MODE
===================================================== */

/**
 * Set maintenance mode. Pakai PropertiesService — tidak baca spreadsheet,
 * overhead ~microseconds vs ~200ms spreadsheet read.
 */
function setMaintenanceMode(payload) {
  try {
    const enabled = !!(payload && payload.enabled);
    const reason  = POS_toString_(payload && payload.reason).trim() || 'Sistem sedang dalam pemeliharaan.';
    const props = PropertiesService.getScriptProperties();
    props.setProperty('MAINTENANCE_MODE', enabled ? 'true' : 'false');
    props.setProperty('MAINTENANCE_REASON', reason);
    return { success: true, message: 'Maintenance mode ' + (enabled ? 'diaktifkan' : 'dinonaktifkan') + '.', data: { enabled: enabled, reason: reason } };
  } catch (e) { return POS_errorResponse_(e); }
}

/** Cek status maintenance — dipanggil frontend saat init. */
function getMaintenanceStatus() {
  try {
    const props = PropertiesService.getScriptProperties();
    const enabled = props.getProperty('MAINTENANCE_MODE') === 'true';
    const reason  = props.getProperty('MAINTENANCE_REASON') || '';
    return { success: true, data: { enabled: enabled, reason: reason } };
  } catch (e) {
    return { success: true, data: { enabled: false, reason: '' } };
  }
}

function testLogin() {
  const result = validateCashierLogin({
    cashierName: 'Farkhan',
    passcode: '123456'
  });

  Logger.log(JSON.stringify(result, null, 2));
}

function testCheckout() {
  const result = checkoutOrder({
    cashierName: 'Farkhan',
    paymentMethod: 'Cash',
    paidAmount: 100000,
    items: [
      {
        menuId: 'M001',
        qty: 2
      },
      {
        menuId: 'M009',
        qty: 1
      }
    ]
  });

  Logger.log(JSON.stringify(result, null, 2));
}

/* =====================================================
   THERAPIST MASTER (CRUD)
===================================================== */

/** Daftarkan terapis baru atau update yang ada. */
function saveTherapist(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const therapistId = POS_toString_(payload && payload.therapistId).trim();
    const name = POS_toString_(payload && payload.therapistName).trim();
    const phone = POS_toString_(payload && payload.phone).trim();
    const pointValue = POS_toNumber_(payload && payload.defaultPointValue);
    const active = (payload && payload.active !== undefined) ? POS_toBoolean_(payload.active) : true;
    const notes = POS_toString_(payload && payload.notes).trim();

    if (!name) throw new Error('Nama terapis wajib diisi.');
    if (pointValue < 0) throw new Error('Nilai point tidak boleh negatif.');

    const sheet = POS_getSheet_(POS_SHEET.THERAPISTS);
    const headers = POS_getHeaders_(sheet);
    const headerMap = POS_getHeaderMap_(headers);
    const lastRow = sheet.getLastRow();
    const now = POS_nowString_();

    // === UPDATE ===
    if (therapistId) {
      // cari row
      let foundRow = -1;
      if (lastRow >= 2) {
        const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
        for (let i = 0; i < values.length; i++) {
          if (POS_toString_(values[i][headerMap.Therapist_ID]) === therapistId) {
            foundRow = i + 2;
            break;
          }
        }
      }
      if (foundRow === -1) throw new Error('Terapis tidak ditemukan: ' + therapistId);

      sheet.getRange(foundRow, headerMap.Therapist_Name + 1).setValue(name);
      sheet.getRange(foundRow, headerMap.Phone + 1).setValue(phone);
      sheet.getRange(foundRow, headerMap.Default_Point_Value + 1).setValue(pointValue);
      sheet.getRange(foundRow, headerMap.Active + 1).setValue(active ? 'TRUE' : 'FALSE');
      sheet.getRange(foundRow, headerMap.Notes + 1).setValue(notes);
      sheet.getRange(foundRow, headerMap.Updated_At + 1).setValue(now);

      SpreadsheetApp.flush();
      return { success: true, message: 'Terapis berhasil diupdate.', data: { therapistId: therapistId } };
    }

    // === INSERT ===
    const newId = POS_generateTherapistId_();
    const row = {
      Therapist_ID: newId,
      Therapist_Name: name,
      Phone: phone,
      Default_Point_Value: pointValue,
      Active: active ? 'TRUE' : 'FALSE',
      Notes: notes,
      Created_At: now,
      Updated_At: now
    };
    POS_appendObjects_(POS_SHEET.THERAPISTS, [row]);
    SpreadsheetApp.flush();
    return { success: true, message: 'Terapis berhasil ditambahkan.', data: row };
  } catch (error) {
    return POS_errorResponse_(error);
  } finally { try { lock.releaseLock(); } catch (e) {} }
}

/** Soft-delete: set Active = FALSE. Aman karena data point lama tetap valid. */
function deactivateTherapist(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const therapistId = POS_toString_(payload && payload.therapistId).trim();
    if (!therapistId) throw new Error('Therapist ID wajib diisi.');
    const sheet = POS_getSheet_(POS_SHEET.THERAPISTS);
    const headers = POS_getHeaders_(sheet);
    const headerMap = POS_getHeaderMap_(headers);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('Belum ada data terapis.');
    const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    for (let i = 0; i < values.length; i++) {
      if (POS_toString_(values[i][headerMap.Therapist_ID]) === therapistId) {
        sheet.getRange(i + 2, headerMap.Active + 1).setValue('FALSE');
        sheet.getRange(i + 2, headerMap.Updated_At + 1).setValue(POS_nowString_());
        SpreadsheetApp.flush();
        return { success: true, message: 'Terapis dinonaktifkan.' };
      }
    }
    throw new Error('Terapis tidak ditemukan.');
  } catch (error) {
    return POS_errorResponse_(error);
  } finally { try { lock.releaseLock(); } catch (e) {} }
}

/** Aktifkan kembali terapis. */
function reactivateTherapist(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const therapistId = POS_toString_(payload && payload.therapistId).trim();
    if (!therapistId) throw new Error('Therapist ID wajib diisi.');
    const sheet = POS_getSheet_(POS_SHEET.THERAPISTS);
    const headers = POS_getHeaders_(sheet);
    const headerMap = POS_getHeaderMap_(headers);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('Belum ada data terapis.');
    const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    for (let i = 0; i < values.length; i++) {
      if (POS_toString_(values[i][headerMap.Therapist_ID]) === therapistId) {
        sheet.getRange(i + 2, headerMap.Active + 1).setValue('TRUE');
        sheet.getRange(i + 2, headerMap.Updated_At + 1).setValue(POS_nowString_());
        SpreadsheetApp.flush();
        return { success: true, message: 'Terapis diaktifkan kembali.' };
      }
    }
    throw new Error('Terapis tidak ditemukan.');
  } catch (error) {
    return POS_errorResponse_(error);
  } finally { try { lock.releaseLock(); } catch (e) {} }
}

/** Ambil semua data master terapis. includeInactive default true (untuk tampilan master). */
function getTherapists(payload) {
  try {
    const includeInactive = !(payload && payload.activeOnly === true);
    const rows = POS_readObjects_(POS_SHEET.THERAPISTS)
      .filter(row => POS_toString_(row.Therapist_ID));
    const therapists = rows.map(row => ({
      therapistId: POS_toString_(row.Therapist_ID),
      therapistName: POS_toString_(row.Therapist_Name),
      phone: POS_toString_(row.Phone),
      defaultPointValue: POS_toNumber_(row.Default_Point_Value),
      active: POS_toBoolean_(row.Active),
      notes: POS_toString_(row.Notes),
      createdAt: POS_toString_(row.Created_At),
      updatedAt: POS_toString_(row.Updated_At)
    })).filter(t => includeInactive ? true : t.active)
      .sort((a, b) => a.therapistName.localeCompare(b.therapistName));
    return { success: true, message: 'Data terapis berhasil dimuat.', data: { therapists: therapists } };
  } catch (error) { return POS_errorResponse_(error); }
}

/* =====================================================
   THERAPIST POINTS (kasir input point per transaksi)
===================================================== */

/**
 * Kasir mencatat point terapis (1 entry = 1 layanan untuk 1 terapis).
 * Bisa multi-line dengan payload.entries = [{therapistId, serviceName, qty, pointPerUnit, amount, notes}, ...]
 * atau single entry langsung pada payload (kompatibilitas).
 */
function saveTherapistPoint(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const cashierName = POS_toString_(payload && payload.cashierName).trim();
    if (!cashierName) throw new Error('Nama kasir tidak ditemukan. Silakan login ulang.');

    let entries = [];
    if (payload && Array.isArray(payload.entries) && payload.entries.length > 0) {
      entries = payload.entries;
    } else if (payload) {
      entries = [payload];
    }
    if (entries.length === 0) throw new Error('Belum ada data point yang akan disimpan.');

    // Build therapist lookup
    const therapistRows = POS_readObjects_(POS_SHEET.THERAPISTS);
    const therapistMap = {};
    therapistRows.forEach(t => {
      const id = POS_toString_(t.Therapist_ID);
      if (id) therapistMap[id] = t;
    });

    const today = POS_todayString_();
    const time = POS_timeString_();
    const now = POS_nowString_();
    const rowsToInsert = [];

    entries.forEach((e, idx) => {
      const therapistId = POS_toString_(e.therapistId).trim();
      const serviceName = POS_toString_(e.serviceName).trim() || 'Layanan';
      const qty = POS_toNumber_(e.qty) || 1;
      const pointPerUnit = POS_toNumber_(e.pointPerUnit);
      const amount = POS_toNumber_(e.amount);
      const notes = POS_toString_(e.notes).trim();
      const explicitDate = POS_toString_(e.date).trim(); // opsional, kalau kasir mau input untuk tanggal lalu

      if (!therapistId) throw new Error('Baris #' + (idx + 1) + ': pilih terapis terlebih dahulu.');
      if (!therapistMap[therapistId]) throw new Error('Baris #' + (idx + 1) + ': terapis tidak ditemukan.');
      if (qty <= 0) throw new Error('Baris #' + (idx + 1) + ': qty harus > 0.');
      if (pointPerUnit < 0) throw new Error('Baris #' + (idx + 1) + ': point per unit tidak boleh negatif.');

      const therapistName = POS_toString_(therapistMap[therapistId].Therapist_Name);
      const totalPoint = qty * pointPerUnit;
      const useDate = explicitDate && /^\d{4}-\d{2}-\d{2}$/.test(explicitDate) ? explicitDate : today;

      rowsToInsert.push({
        Point_ID: POS_generatePointId_(idx),
        Date: useDate,
        Time: time,
        Therapist_ID: therapistId,
        Therapist_Name: therapistName,
        Service_Name: serviceName,
        Qty: qty,
        Point_Per_Unit: pointPerUnit,
        Total_Point: totalPoint,
        Amount: amount,
        Cashier_Name: cashierName,
        Notes: notes,
        Created_At: now
      });
    });

    POS_appendObjects_(POS_SHEET.THERAPIST_POINTS, rowsToInsert);
    SpreadsheetApp.flush();

    return {
      success: true,
      message: 'Berhasil menyimpan ' + rowsToInsert.length + ' entri point terapis.',
      data: { count: rowsToInsert.length, rows: rowsToInsert }
    };
  } catch (error) {
    return POS_errorResponse_(error);
  } finally { try { lock.releaseLock(); } catch (e) {} }
}

/**
 * Ambil log point untuk ditampilkan di tab terapis (kasir).
 * Opsi filter: { startDate, endDate, therapistId }.
 */
function getTherapistPoints(payload) {
  try {
    const filter = payload || {};
    const today = POS_todayString_();
    const range = POS_resolveDashboardRange_(filter, today);
    const therapistId = POS_toString_(filter.therapistId).trim();

    const rows = POS_readObjects_(POS_SHEET.THERAPIST_POINTS)
      .filter(r => POS_toString_(r.Point_ID))
      .filter(r => POS_isDateInRange_(r.Date, range.startDate, range.endDate))
      .filter(r => therapistId ? POS_toString_(r.Therapist_ID) === therapistId : true)
      .map(r => ({
        pointId: POS_toString_(r.Point_ID),
        date: POS_toString_(r.Date),
        time: POS_toString_(r.Time),
        therapistId: POS_toString_(r.Therapist_ID),
        therapistName: POS_toString_(r.Therapist_Name),
        serviceName: POS_toString_(r.Service_Name),
        qty: POS_toNumber_(r.Qty),
        pointPerUnit: POS_toNumber_(r.Point_Per_Unit),
        totalPoint: POS_toNumber_(r.Total_Point),
        amount: POS_toNumber_(r.Amount),
        cashierName: POS_toString_(r.Cashier_Name),
        notes: POS_toString_(r.Notes)
      }))
      .sort((a, b) => (b.date + ' ' + b.time).localeCompare(a.date + ' ' + a.time));

    const summary = rows.reduce((s, r) => {
      s.totalPoint += r.totalPoint;
      s.totalAmount += r.amount;
      s.count += 1;
      return s;
    }, { totalPoint: 0, totalAmount: 0, count: 0 });

    return {
      success: true,
      message: 'Data point terapis berhasil dimuat.',
      data: {
        range: range,
        rows: rows,
        summary: summary
      }
    };
  } catch (error) { return POS_errorResponse_(error); }
}

/** Hapus 1 entry point (hanya jika diberi password — gunakan superadmin atau therapist). */
function deleteTherapistPoint(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const pointId = POS_toString_(payload && payload.pointId).trim();
    if (!pointId) throw new Error('Point ID wajib diisi.');
    const sheet = POS_getSheet_(POS_SHEET.THERAPIST_POINTS);
    const headers = POS_getHeaders_(sheet);
    const headerMap = POS_getHeaderMap_(headers);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('Belum ada data point.');
    const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    for (let i = 0; i < values.length; i++) {
      if (POS_toString_(values[i][headerMap.Point_ID]) === pointId) {
        sheet.deleteRow(i + 2);
        SpreadsheetApp.flush();
        return { success: true, message: 'Entri point dihapus.' };
      }
    }
    throw new Error('Point tidak ditemukan.');
  } catch (error) {
    return POS_errorResponse_(error);
  } finally { try { lock.releaseLock(); } catch (e) {} }
}

/* =====================================================
   THERAPIST CONSOLIDATION (Superadmin)
===================================================== */

/**
 * Konsolidasi settlement terapis pada range tertentu.
 *
 * Formula per terapis:
 *   total_point_nominal = SUM(Total_Point * Point_Per_Unit) ATAU SUM(Amount) bila dipakai
 *   share              = total_point_nominal * (THERAPIST_SHARE_PERCENT / 100)
 *   kasbon             = SUM(EXPENSES.Amount) WHERE Expense_Type='Therapist' AND Personal_Cashier=therapistName
 *   net_payout         = share - kasbon
 *
 * Catatan: "nilai nominal point" dihitung sebagai SUM(Qty * Point_Per_Unit) supaya tidak
 * tergantung field Amount (yang bisa kosong). Ini = Total_Point itu sendiri yang sudah
 * dianggap rupiah ekuivalen. Jika user butuh "point murni" tanpa rupiah, gunakan
 * Default_Point_Value dari master sebagai multiplier — tetapi paling fleksibel: pakai
 * Point_Per_Unit yang sudah diinput per entri.
 */
function getTherapistConsolidation(payload) {
  try {
    const filter = payload || {};
    const today = POS_todayString_();
    const range = POS_resolveDashboardRange_(filter, today);

    // Override share percent kalau dikirim, kalau tidak ambil dari config
    let sharePercent;
    if (filter.sharePercent !== undefined && filter.sharePercent !== null && filter.sharePercent !== '') {
      sharePercent = POS_toNumber_(filter.sharePercent);
    } else {
      sharePercent = POS_toNumber_(POS_getConfigValue_('THERAPIST_SHARE_PERCENT', '50'));
    }
    if (sharePercent < 0) sharePercent = 0;
    if (sharePercent > 100) sharePercent = 100;
    const shareRatio = sharePercent / 100;

    // 1. Ambil master terapis untuk header map nama
    const therapists = POS_readObjects_(POS_SHEET.THERAPISTS)
      .filter(r => POS_toString_(r.Therapist_ID));
    const masterById = {};
    const masterByName = {};
    therapists.forEach(t => {
      const id = POS_toString_(t.Therapist_ID);
      const nm = POS_toString_(t.Therapist_Name);
      masterById[id] = { id: id, name: nm, active: POS_toBoolean_(t.Active) };
      if (nm) masterByName[nm.toLowerCase()] = id;
    });

    // 2. Ambil semua point dalam range
    const points = POS_readObjects_(POS_SHEET.THERAPIST_POINTS)
      .filter(r => POS_toString_(r.Point_ID))
      .filter(r => POS_isDateInRange_(r.Date, range.startDate, range.endDate));

    // 3. Ambil kasbon terapis dalam range (Expense_Type=Therapist)
    const expenses = POS_readObjects_(POS_SHEET.EXPENSES).map(POS_normalizeExpenseRow_)
      .filter(r => POS_toString_(r.Expense_Type) === 'Therapist')
      .filter(r => POS_isDateInRange_(r.Date, range.startDate, range.endDate));

    // 4. Bucket per therapist
    const bucket = {}; // key = therapistId

    function ensureBucket(id, fallbackName) {
      if (!bucket[id]) {
        const master = masterById[id];
        bucket[id] = {
          therapistId: id,
          therapistName: master ? master.name : (fallbackName || id),
          totalPoint: 0,
          totalAmount: 0,
          entriesCount: 0,
          kasbon: 0,
          kasbonCount: 0,
          gross: 0,
          share: 0,
          netPayout: 0
        };
      }
      return bucket[id];
    }

    points.forEach(p => {
      const id = POS_toString_(p.Therapist_ID);
      if (!id) return;
      const b = ensureBucket(id, POS_toString_(p.Therapist_Name));
      const totalPoint = POS_toNumber_(p.Total_Point);
      const amount = POS_toNumber_(p.Amount);
      b.totalPoint += totalPoint;
      b.totalAmount += amount;
      b.entriesCount += 1;
    });

    // Kasbon di-link via nama (Personal_Cashier). Coba match nama → id.
    expenses.forEach(e => {
      const nm = POS_toString_(e.Personal_Cashier).trim();
      if (!nm) return;
      let id = masterByName[nm.toLowerCase()];
      if (!id) {
        // terapis sudah dihapus / nama tidak match — buat bucket virtual berdasarkan nama
        id = 'UNKNOWN__' + nm;
      }
      const b = ensureBucket(id, nm);
      b.kasbon += POS_toNumber_(e.Amount);
      b.kasbonCount += 1;
    });

    // 5. Hitung share & net
    const result = Object.keys(bucket).map(id => {
      const b = bucket[id];
      // gross = total nominal point yang dipakai sebagai dasar bagi hasil.
      // Pakai Total_Point (Qty * Point_Per_Unit) — yang sudah nominal rupiah ekuivalen.
      b.gross = b.totalPoint;
      b.share = Math.round(b.gross * shareRatio);
      b.netPayout = b.share - b.kasbon;
      return b;
    }).sort((a, b) => b.netPayout - a.netPayout);

    // 6. Summary global
    const summary = result.reduce((s, r) => {
      s.totalGross += r.gross;
      s.totalShare += r.share;
      s.totalKasbon += r.kasbon;
      s.totalNet += r.netPayout;
      s.totalEntries += r.entriesCount;
      s.totalKasbonCount += r.kasbonCount;
      return s;
    }, { totalGross: 0, totalShare: 0, totalKasbon: 0, totalNet: 0, totalEntries: 0, totalKasbonCount: 0, therapistCount: result.length });

    // 7. Breakdown harian (untuk chart)
    const dates = POS_getDateStringsBetween_(range.startDate, range.endDate);
    const dailyMap = {};
    dates.forEach(d => { dailyMap[d] = { date: d, gross: 0, share: 0, kasbon: 0, net: 0 }; });
    points.forEach(p => {
      const d = POS_toString_(p.Date).slice(0, 10);
      if (dailyMap[d]) {
        dailyMap[d].gross += POS_toNumber_(p.Total_Point);
      }
    });
    expenses.forEach(e => {
      const d = POS_toString_(e.Date).slice(0, 10);
      if (dailyMap[d]) {
        dailyMap[d].kasbon += POS_toNumber_(e.Amount);
      }
    });
    const daily = dates.map(d => {
      const row = dailyMap[d];
      row.share = Math.round(row.gross * shareRatio);
      row.net = row.share - row.kasbon;
      return row;
    });

    return {
      success: true,
      message: 'Konsolidasi terapis berhasil dimuat.',
      data: {
        range: range,
        sharePercent: sharePercent,
        therapists: result,
        summary: summary,
        daily: daily
      }
    };
  } catch (error) { return POS_errorResponse_(error); }
}

/* =====================================================
   ID GENERATORS (therapist & point)
===================================================== */

function POS_generateTherapistId_() {
  const sheet = POS_getSheet_(POS_SHEET.THERAPISTS);
  const lastRow = sheet.getLastRow();
  let maxNum = 0;
  if (lastRow >= 2) {
    const headers = POS_getHeaders_(sheet);
    const idx = headers.indexOf('Therapist_ID');
    if (idx !== -1) {
      const ids = sheet.getRange(2, idx + 1, lastRow - 1, 1).getValues();
      ids.forEach(r => {
        const m = /^T(\d+)$/.exec(POS_toString_(r[0]));
        if (m) {
          const n = parseInt(m[1], 10);
          if (n > maxNum) maxNum = n;
        }
      });
    }
  }
  const nextNum = maxNum + 1;
  return 'T' + ('000' + nextNum).slice(-3);
}

function POS_generatePointId_(extra) {
  const now = POS_now_();
  const datePart = Utilities.formatDate(now, POS_getTimezone_(), 'yyyyMMdd');
  const timePart = Utilities.formatDate(now, POS_getTimezone_(), 'HHmmss');
  const randomPart = Math.floor(Math.random() * 900 + 100);
  const suffix = extra !== undefined ? '-' + extra : '';
  return 'PT-' + datePart + '-' + timePart + '-' + randomPart + suffix;
}
