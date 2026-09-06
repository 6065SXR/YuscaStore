/**
 * ============================================================================
 * ZETTBOT 3.1 - BACKEND BUSINESS LOGIC & API (Yusca STORE)
 * File: Code.gs
 * Framework: Google Apps Script Web App
 * ============================================================================
 */

var TIMEZONE = 'Asia/Jakarta';

/**
 * Endpoint utama Web App untuk merender antarmuka HTML
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Yusca STORE - Management & E-Commerce System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Helper function untuk menyertakan file modul HTML/CSS/JS pendukung
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Generasi ID Sekuensial berdasarkan tanggal dan penghitung harian
 */
function generateSequentialId(prefix, dateKey, countKey) {
  var props = PropertiesService.getScriptProperties();
  var todayStr = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
  var lastDate = props.getProperty(dateKey);
  var currentCount = 1;

  if (lastDate === todayStr) {
    var storedCount = parseInt(props.getProperty(countKey) || '0', 10);
    currentCount = storedCount + 1;
  }

  props.setProperty(dateKey, todayStr);
  props.setProperty(countKey, currentCount.toString());

  var paddedNum = ('0000' + currentCount).slice(-4);
  return prefix + '-' + paddedNum;
}

/**
 * Mencatat log aktivitas admin ke sheet Admin_Logs
 */
function logAdminActivity(adminName, actionType, details) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Admin_Logs');
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName('Admin_Logs');
    }
    var logId = generateSequentialId('LOG', 'LOG_LAST_DATE', 'LOG_LAST_COUNT');
    var nowStr = Utilities.formatDate(new Date(), TIMEZONE, 'dd/MM/yyyy HH:mm:ss');
    sheet.appendRow([logId, adminName || 'Admin System', actionType, details, nowStr]);
    SpreadsheetApp.flush();
  } catch (err) {
    Logger.log('Error logAdminActivity: ' + err.toString());
  }
}

/**
 * Otentikasi login pengguna (Admin & Pembeli)
 */
function loginUser(phone, password) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Users');
    if (!sheet) return { success: false, message: 'Database tabel Users tidak ditemukan!' };

    var rows = sheet.getDataRange().getDisplayValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][1] === String(phone).trim() && rows[i][2] === String(password).trim()) {
        return {
          success: true,
          user: {
            name: rows[i][0],
            phone: rows[i][1],
            role: rows[i][3],
            userId: rows[i][4]
          }
        };
      }
    }
    return { success: false, message: 'Nomor telepon atau password salah!' };
  } catch (err) {
    return { success: false, message: 'Terjadi kesalahan sistem: ' + err.toString() };
  }
}

/**
 * Pendaftaran akun pembeli baru
 */
function registerBuyer(userData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Users');
    if (!sheet) return { success: false, message: 'Sheet Users belum tersedia.' };

    var rows = sheet.getDataRange().getDisplayValues();
    var cleanPhone = String(userData.phone).trim();
    
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][1] === cleanPhone) {
        return { success: false, message: 'Nomor telepon sudah terdaftar. Silakan login!' };
      }
    }

    var nextId = generateSequentialId('CUST', 'CUST_LAST_DATE', 'CUST_LAST_COUNT');
    var nowStr = Utilities.formatDate(new Date(), TIMEZONE, 'dd/MM/yyyy HH:mm:ss');
    var newRow = [
      String(userData.name).trim(),
      cleanPhone,
      String(userData.password).trim(),
      'Pembeli',
      nextId,
      nowStr
    ];

    sheet.appendRow(newRow);
    SpreadsheetApp.flush();

    return {
      success: true,
      user: {
        name: userData.name,
        phone: cleanPhone,
        role: 'Pembeli',
        userId: nextId
      }
    };
  } catch (err) {
    return { success: false, message: 'Gagal mendaftarkan akun: ' + err.toString() };
  }
}

/**
 * Mengambil metrik analitik dashboard toko dan stok kritis
 */
function getDashboardMetrics() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var orderSheet = ss.getSheetByName('Orders');
    var invSheet = ss.getSheetByName('Inventory');
    var masterSheet = ss.getSheetByName('Master_Stock');

    var orderRows = orderSheet ? orderSheet.getDataRange().getDisplayValues() : [];
    var invRows = invSheet ? invSheet.getDataRange().getDisplayValues() : [];
    var masterRows = masterSheet ? masterSheet.getDataRange().getDisplayValues() : [];

    var netSales = 0;
    var totalOrders = 0;
    var returnedGoods = 750000;
    var canceledOrders = 0;
    var uniqueCustomersMap = {};
    var pendingPayments = 0;
    var productSalesCounter = {};

    for (var i = 1; i < orderRows.length; i++) {
      var row = orderRows[i];
      var buyerId = row[1];
      var skus = row[2] ? row[2].split(',') : [];
      var qty = row[3] ? row[3].split(',') : [];
      var total = parseFloat(row[4]) || 0;
      var status = row[5];

      totalOrders++;
      if (buyerId) uniqueCustomersMap[buyerId] = true;

      if (status === 'Confirmed' || status === 'Shipped' || status === 'Delivered') {
        netSales += total;
        for (var k = 0; k < skus.length; k++) {
          var itemSku = skus[k].trim();
          var itemQty = parseInt(qty[k], 10) || 1;
          productSalesCounter[itemSku] = (productSalesCounter[itemSku] || 0) + itemQty;
        }
      } else if (status === 'Pending') {
        pendingPayments += total;
      } else if (status === 'Canceled') {
        canceledOrders++;
      }
    }

    var uniqueCustomers = Object.keys(uniqueCustomersMap).length;
    var avgOrderValue = totalOrders > 0 ? Math.round(netSales / Math.max(1, (totalOrders - canceledOrders))) : 0;

    var topSku = '';
    var maxSales = 0;
    for (var s in productSalesCounter) {
      if (productSalesCounter[s] > maxSales) {
        maxSales = productSalesCounter[s];
        topSku = s;
      }
    }

    var topSellingBrand = 'Yusca NMD R1';
    if (topSku) {
      for (var m = 1; m < masterRows.length; m++) {
        if (masterRows[m][0] === topSku) {
          topSellingBrand = masterRows[m][1];
          break;
        }
      }
    }

    var criticalAlerts = [];
    var masterNameMap = {};
    for (var j = 1; j < masterRows.length; j++) {
      masterNameMap[masterRows[j][0]] = masterRows[j][1];
    }

    for (var v = 1; v < invRows.length; v++) {
      var iRow = invRows[v];
      var iSku = iRow[0];
      var iQty = parseInt(iRow[1], 10) || 0;
      var iStatus = iRow[2];

      if (iQty <= 5 || iStatus === 'Low Stock') {
        criticalAlerts.push({
          sku: iSku,
          productName: masterNameMap[iSku] || iSku,
          quantity: iQty,
          status: 'Low Stock'
        });
      }
    }

    return {
      success: true,
      data: {
        netSales: netSales,
        totalOrders: totalOrders,
        returnedGoods: returnedGoods,
        canceledOrders: canceledOrders,
        uniqueCustomers: uniqueCustomers,
        avgOrderValue: avgOrderValue,
        pendingPayments: pendingPayments,
        topSellingBrand: topSellingBrand,
        criticalAlerts: criticalAlerts,
        salesGrowth: {
          labels: ['JAN', 'FEB', 'MAR', 'APR', 'JUL', 'AUG', 'SEP', 'NOV', 'DEC'],
          sales: [6800, 8100, 9400, 7500, 8300, 8200, 10700, 8800, 10100]
        },
        profitTrend: {
          labels: ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'DEC'],
          netProfit: [0, 58000, 32000, 86000, 52000, 110000, 90000, 138000],
          revenue: [10000, 70000, 45000, 105000, 68000, 130000, 115000, 160000]
        }
      }
    };
  } catch (err) {
    return { success: false, message: 'Gagal mengambil metrik dashboard: ' + err.toString() };
  }
}

/**
 * Mengambil katalog produk terpaginasi dengan promo flash sale
 */
function getCatalogProducts(page, pageSize, categoryFilter, searchQuery) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName('Master_Stock');
    var invSheet = ss.getSheetByName('Inventory');

    if (!masterSheet) return { success: true, items: [], total: 0 };

    var masterRows = masterSheet.getDataRange().getDisplayValues();
    var invRows = invSheet ? invSheet.getDataRange().getDisplayValues() : [];

    var stockMap = {};
    for (var i = 1; i < invRows.length; i++) {
      var s = invRows[i][0];
      var q = parseInt(invRows[i][1], 10) || 0;
      stockMap[s] = (stockMap[s] || 0) + q;
    }

    var flashConfig = {};
    try {
      var rawFlash = PropertiesService.getScriptProperties().getProperty('FLASH_SALE_CONFIG');
      if (rawFlash) flashConfig = JSON.parse(rawFlash);
    } catch (e) {}

    var allProducts = [];
    var search = (searchQuery || '').toLowerCase().trim();
    var category = (categoryFilter || 'All').trim();
    var todayStr = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');

    for (var m = 1; m < masterRows.length; m++) {
      var row = masterRows[m];
      var sku = row[0];
      var name = row[1];
      var desc = row[2];
      var cat = row[3];
      var brand = row[4] || '-';
      var img = row[5];
      var price = parseFloat(row[6]) || 0;
      var stock = stockMap[sku] !== undefined ? stockMap[sku] : 0;

      if (category !== 'All' && category !== 'Flash Sale' && cat !== category) continue;
      if (search && name.toLowerCase().indexOf(search) === -1 && sku.toLowerCase().indexOf(search) === -1) continue;

      var flashData = flashConfig[sku] || null;
      var isFlash = flashData ? !!flashData.isFlashSale : (m % 2 === 1);
      var discPrice = (flashData && flashData.discountPrice) ? parseFloat(flashData.discountPrice) : Math.round(price * 0.8);
      var sDate = (flashData && flashData.startDate) ? flashData.startDate : todayStr;
      var eDate = (flashData && flashData.endDate) ? flashData.endDate : '2026-12-31';

      if (category === 'Flash Sale' && !isFlash) continue;

      allProducts.push({
        sku: sku,
        productName: name,
        description: desc,
        category: cat,
        brand: brand,
        imageUrl: img,
        price: price,
        stock: stock,
        isFlashSale: isFlash,
        discountPrice: discPrice,
        startDate: sDate,
        endDate: eDate
      });
    }

    var p = parseInt(page, 10) || 1;
    var size = parseInt(pageSize, 10) || 30;
    var start = (p - 1) * size;

    return {
      success: true,
      items: allProducts.slice(start, start + size),
      total: allProducts.length,
      page: p,
      totalPages: Math.ceil(allProducts.length / size)
    };
  } catch (err) {
    return { success: false, message: 'Gagal memuat katalog: ' + err.toString() };
  }
}

/**
 * Menyimpan konfigurasi promo Flash Sale ke Script Properties
 */
function saveFlashSaleBackend(flashConfig, adminName) {
  try {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('FLASH_SALE_CONFIG', JSON.stringify(flashConfig || {}));
    SpreadsheetApp.flush();
    logAdminActivity(adminName || 'Admin', 'Simpan Flash Sale', 'Memperbarui konfigurasi promo Flash Sale');
    return { success: true, message: 'Konfigurasi Promo Flash Sale berhasil disimpan!' };
  } catch (err) {
    return { success: false, message: 'Gagal menyimpan Flash Sale: ' + err.toString() };
  }
}

/**
 * Memproses pesanan checkout pembeli dan pemotongan stok otomatis
 */
function submitOrder(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var orderSheet = ss.getSheetByName('Orders');
    var trxSheet = ss.getSheetByName('Transactions');
    var invSheet = ss.getSheetByName('Inventory');

    var orderId = generateSequentialId('ORD', 'ORD_LAST_DATE', 'ORD_LAST_COUNT');
    var trxId = generateSequentialId('TRX', 'TRX_LAST_DATE', 'TRX_LAST_COUNT');
    var nowStr = Utilities.formatDate(new Date(), TIMEZONE, 'dd/MM/yyyy HH:mm:ss');

    var skus = payload.items.map(function(item) { return item.sku; }).join(',');
    var quantities = payload.items.map(function(item) { return item.qty; }).join(',');
    var totalPrice = parseFloat(payload.totalPrice) || 0;
    var buyerUserId = String(payload.buyerUserId).trim();

    orderSheet.appendRow([orderId, buyerUserId, skus, quantities, totalPrice, 'Pending', '', nowStr, trxId]);
    trxSheet.appendRow([trxId, orderId, totalPrice, 'Order Payment', nowStr, 'Menunggu Transfer Bank via ' + (payload.bankName || 'BCA')]);

    if (invSheet) {
      var invRows = invSheet.getDataRange().getDisplayValues();
      for (var k = 0; k < payload.items.length; k++) {
        var cartItem = payload.items[k];
        for (var r = 1; r < invRows.length; r++) {
          if (invRows[r][0] === cartItem.sku) {
            var currQty = parseInt(invRows[r][1], 10) || 0;
            var newQty = Math.max(0, currQty - parseInt(cartItem.qty, 10));
            var newStatus = newQty <= 5 ? 'Low Stock' : 'In Stock';
            invSheet.getRange(r + 1, 2).setValue(newQty);
            invSheet.getRange(r + 1, 3).setValue(newStatus);
            invSheet.getRange(r + 1, 4).setValue(nowStr);
            break;
          }
        }
      }
    }

    SpreadsheetApp.flush();

    var adminPhone = '6281234567890';
    var waMessage = 'Halo Admin Yusca STORE,%0A%0ASaya ingin konfirmasi pesanan baru:%0A' +
      '- *Order ID:* ' + orderId + '%0A' +
      '- *Customer ID:* ' + buyerUserId + '%0A' +
      '- *Total Bayar:* Rp ' + totalPrice.toLocaleString('id-ID') + '%0A' +
      '- *Bank Transfer:* ' + (payload.bankName || 'BCA') + '%0A%0A' +
      'Mohon verifikasi pembayaran saya. Terima kasih!';

    return {
      success: true,
      orderId: orderId,
      transactionId: trxId,
      totalPrice: totalPrice,
      waUrl: 'https://api.whatsapp.com/send?phone=' + adminPhone + '&text=' + waMessage
    };
  } catch (err) {
    return { success: false, message: 'Gagal membuat pesanan: ' + err.toString() };
  }
}

/**
 * Mengambil daftar pesanan terpaginasi
 */
function getOrdersPaginated(page, pageSize, statusFilter, searchQuery) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Orders');
    if (!sheet) return { success: true, items: [], total: 0 };

    var rows = sheet.getDataRange().getDisplayValues();
    var all = [];
    var search = (searchQuery || '').toLowerCase().trim();
    var filterStatus = (statusFilter || 'All').trim();

    for (var i = rows.length - 1; i >= 1; i--) {
      var row = rows[i];
      if (filterStatus !== 'All' && row[5] !== filterStatus) continue;
      if (search && row[0].toLowerCase().indexOf(search) === -1 && row[1].toLowerCase().indexOf(search) === -1) continue;

      all.push({
        orderId: row[0],
        buyerUserId: row[1],
        skus: row[2],
        quantities: row[3],
        totalPrice: parseFloat(row[4]) || 0,
        status: row[5],
        paymentProof: row[6],
        dateOrdered: row[7],
        transactionId: row[8]
      });
    }

    var p = parseInt(page, 10) || 1;
    var size = parseInt(pageSize, 10) || 10;
    var start = (p - 1) * size;

    return { success: true, items: all.slice(start, start + size), total: all.length, page: p, totalPages: Math.ceil(all.length / size) };
  } catch (err) {
    return { success: false, message: 'Gagal memuat pesanan: ' + err.toString() };
  }
}

/**
 * Memperbarui status pesanan dan menghasilkan tautan notifikasi WhatsApp
 */
function updateOrderStatus(orderId, newStatus, adminName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Orders');
    var userSheet = ss.getSheetByName('Users');
    if (!sheet) return { success: false, message: 'Sheet Orders tidak ditemukan.' };

    var rows = sheet.getDataRange().getDisplayValues();
    var buyerPhone = '6281234567890';
    var buyerUserId = '';

    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === String(orderId).trim()) {
        sheet.getRange(i + 1, 6).setValue(newStatus);
        buyerUserId = rows[i][1];

        if (userSheet && buyerUserId) {
          var userRows = userSheet.getDataRange().getDisplayValues();
          for (var u = 1; u < userRows.length; u++) {
            if (userRows[u][4] === buyerUserId) {
              buyerPhone = userRows[u][1];
              if (buyerPhone.indexOf('0') === 0) buyerPhone = '62' + buyerPhone.slice(1);
              break;
            }
          }
        }

        SpreadsheetApp.flush();
        logAdminActivity(adminName || 'Admin', 'Update Order Status', 'Order ' + orderId + ' diubah ke status ' + newStatus);

        var waMsg = encodeURIComponent(
          'Halo Kak, kami dari *Yusca STORE*! ⚡%0A%0A' +
          'Status pesanan Anda dengan *Order ID: ' + orderId + '* telah diperbarui menjadi:%0A' +
          '👉 *' + newStatus.toUpperCase() + '*%0A%0A' +
          'Terima kasih telah berbelanja di Yusca STORE! Jika ada pertanyaan silakan hubungi kami.'
        );
        var waUrl = 'https://api.whatsapp.com/send?phone=' + buyerPhone + '&text=' + waMsg;

        return {
          success: true,
          message: 'Status pesanan ' + orderId + ' berhasil diubah menjadi ' + newStatus,
          waUrl: waUrl
        };
      }
    }
    return { success: false, message: 'Order ID tidak ditemukan!' };
  } catch (err) {
    return { success: false, message: 'Gagal mengupdate status: ' + err.toString() };
  }
}

/**
 * Mengambil data inventori stok produk
 */
function getInventoryPaginated(page, pageSize, searchQuery) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var invSheet = ss.getSheetByName('Inventory');
    var masterSheet = ss.getSheetByName('Master_Stock');

    var invRows = invSheet ? invSheet.getDataRange().getDisplayValues() : [];
    var masterRows = masterSheet ? masterSheet.getDataRange().getDisplayValues() : [];

    var masterMap = {};
    for (var m = 1; m < masterRows.length; m++) {
      masterMap[masterRows[m][0]] = { name: masterRows[m][1], category: masterRows[m][3], brand: masterRows[m][4] };
    }

    var all = [];
    var search = (searchQuery || '').toLowerCase().trim();

    for (var i = 1; i < invRows.length; i++) {
      var row = invRows[i];
      var prod = masterMap[row[0]] || { name: row[0], category: '-', brand: '-' };

      if (search && row[0].toLowerCase().indexOf(search) === -1 && prod.name.toLowerCase().indexOf(search) === -1) continue;

      all.push({
        sku: row[0],
        productName: prod.name,
        category: prod.category,
        brand: prod.brand,
        quantity: parseInt(row[1], 10) || 0,
        status: row[2],
        dateUpdated: row[3]
      });
    }

    var p = parseInt(page, 10) || 1;
    var size = parseInt(pageSize, 10) || 10;
    var start = (p - 1) * size;

    return { success: true, items: all.slice(start, start + size), total: all.length, page: p, totalPages: Math.ceil(all.length / size) };
  } catch (err) {
    return { success: false, message: 'Gagal memuat inventori: ' + err.toString() };
  }
}

/**
 * Memperbarui jumlah stok produk
 */
function updateInventoryQuantity(sku, newQuantity, adminName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Inventory');
    var nowStr = Utilities.formatDate(new Date(), TIMEZONE, 'dd/MM/yyyy HH:mm:ss');
    var qtyNum = parseInt(newQuantity, 10) || 0;
    var status = qtyNum <= 5 ? 'Low Stock' : 'In Stock';

    var rows = sheet.getDataRange().getDisplayValues();
    var found = false;

    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === String(sku).trim()) {
        sheet.getRange(i + 1, 2).setValue(qtyNum);
        sheet.getRange(i + 1, 3).setValue(status);
        sheet.getRange(i + 1, 4).setValue(nowStr);
        found = true;
        break;
      }
    }

    if (!found) sheet.appendRow([sku, qtyNum, status, nowStr]);

    SpreadsheetApp.flush();
    logAdminActivity(adminName || 'Admin', 'Update Stok', 'SKU ' + sku + ' stok diubah ke ' + qtyNum);

    return { success: true, message: 'Stok berhasil diperbarui!' };
  } catch (err) {
    return { success: false, message: 'Gagal memperbarui stok: ' + err.toString() };
  }
}

/**
 * Mengambil daftar metadata Jenis Produk
 */
function getJenisProdukList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Jenis_Produk');
    if (!sheet) return { success: true, items: [] };

    var rows = sheet.getDataRange().getDisplayValues();
    var all = [];
    for (var i = 1; i < rows.length; i++) {
      all.push({
        id: rows[i][0],
        type: rows[i][1],
        name: rows[i][2],
        description: rows[i][3],
        dateCreated: rows[i][4]
      });
    }
    return { success: true, items: all };
  } catch (err) {
    return { success: false, message: 'Gagal memuat Jenis Produk: ' + err.toString() };
  }
}

/**
 * Menyimpan atau memperbarui data Jenis Produk
 */
function saveJenisProduk(payload, adminName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Jenis_Produk');
    var nowStr = Utilities.formatDate(new Date(), TIMEZONE, 'dd/MM/yyyy HH:mm:ss');

    if (payload.id) {
      var rows = sheet.getDataRange().getDisplayValues();
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][0] === String(payload.id).trim()) {
          sheet.getRange(i + 1, 2).setValue(payload.type);
          sheet.getRange(i + 1, 3).setValue(payload.name);
          sheet.getRange(i + 1, 4).setValue(payload.description);
          SpreadsheetApp.flush();
          logAdminActivity(adminName || 'Admin', 'Edit Jenis Produk', 'Ubah ' + payload.id + ' (' + payload.name + ')');
          return { success: true, message: 'Jenis Produk berhasil diperbarui!' };
        }
      }
    }

    var nextId = generateSequentialId('JPN', 'JPN_LAST_DATE', 'JPN_LAST_COUNT');
    sheet.appendRow([nextId, payload.type, payload.name, payload.description, nowStr]);
    SpreadsheetApp.flush();
    logAdminActivity(adminName || 'Admin', 'Tambah Jenis Produk', 'Tambah baru ' + nextId + ' (' + payload.name + ')');

    return { success: true, message: 'Jenis Produk baru berhasil ditambahkan!' };
  } catch (err) {
    return { success: false, message: 'Gagal menyimpan Jenis Produk: ' + err.toString() };
  }
}

/**
 * Menghapus Jenis Produk berdasarkan ID
 */
function deleteJenisProduk(id, adminName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Jenis_Produk');
    var rows = sheet.getDataRange().getDisplayValues();

    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === String(id).trim()) {
        sheet.deleteRow(i + 1);
        SpreadsheetApp.flush();
        logAdminActivity(adminName || 'Admin', 'Hapus Jenis Produk', 'Hapus ID ' + id);
        return { success: true, message: 'Jenis Produk berhasil dihapus!' };
      }
    }
    return { success: false, message: 'ID Jenis Produk tidak ditemukan!' };
  } catch (err) {
    return { success: false, message: 'Gagal menghapus Jenis Produk: ' + err.toString() };
  }
}

/**
 * Mengambil daftar Master SKU terpaginasi
 */
function getMasterProductsPaginated(page, pageSize, searchQuery) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Master_Stock');
    if (!sheet) return { success: true, items: [], total: 0 };

    var rows = sheet.getDataRange().getDisplayValues();
    var all = [];
    var search = (searchQuery || '').toLowerCase().trim();

    for (var i = 1; i < rows.length; i++) {
      var row = rows[i];
      if (search && row[0].toLowerCase().indexOf(search) === -1 && row[1].toLowerCase().indexOf(search) === -1) continue;

      all.push({
        sku: row[0],
        productName: row[1],
        description: row[2],
        category: row[3],
        brand: row[4] || '-',
        imageUrl: row[5],
        price: parseFloat(row[6]) || 0
      });
    }

    var p = parseInt(page, 10) || 1;
    var size = parseInt(pageSize, 10) || 10;
    var start = (p - 1) * size;

    return { success: true, items: all.slice(start, start + size), total: all.length, page: p, totalPages: Math.ceil(all.length / size) };
  } catch (err) {
    return { success: false, message: 'Gagal memuat master produk: ' + err.toString() };
  }
}

/**
 * Menyimpan atau memperbarui data Master Produk
 */
function saveMasterProduct(productData, isEdit, adminName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Master_Stock');
    var sku = String(productData.sku).trim();
    var rows = sheet.getDataRange().getDisplayValues();

    if (isEdit) {
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][0] === sku) {
          sheet.getRange(i + 1, 2).setValue(productData.productName);
          sheet.getRange(i + 1, 3).setValue(productData.description);
          sheet.getRange(i + 1, 4).setValue(productData.category);
          sheet.getRange(i + 1, 5).setValue(productData.brand || 'Yusca');
          sheet.getRange(i + 1, 6).setValue(productData.imageUrl);
          sheet.getRange(i + 1, 7).setValue(parseFloat(productData.price) || 0);
          SpreadsheetApp.flush();
          logAdminActivity(adminName || 'Admin', 'Edit Master SKU', 'Perbarui produk ' + sku + ' (' + productData.productName + ')');
          return { success: true, message: 'Produk berhasil diperbarui!' };
        }
      }
      return { success: false, message: 'SKU tidak ditemukan untuk diedit!' };
    } else {
      for (var k = 1; k < rows.length; k++) {
        if (rows[k][0] === sku) return { success: false, message: 'SKU sudah digunakan! Harap pakai SKU unik.' };
      }
      sheet.appendRow([
        sku,
        productData.productName,
        productData.description,
        productData.category,
        productData.brand || 'Yusca',
        productData.imageUrl || 'https://images.unsplash.com/photo-1542291026-7eec264c27ff?w=500',
        parseFloat(productData.price) || 0
      ]);
      SpreadsheetApp.flush();
      logAdminActivity(adminName || 'Admin', 'Tambah Master SKU', 'Tambah produk baru ' + sku + ' (' + productData.productName + ')');
      return { success: true, message: 'Produk baru berhasil ditambahkan!' };
    }
  } catch (err) {
    return { success: false, message: 'Gagal menyimpan produk: ' + err.toString() };
  }
}

/**
 * Mengimpor produk dari file CSV secara masal dengan pengecekan SKU unik
 */
function importMasterProductsCSV(itemsArray, adminName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName('Master_Stock');
    var invSheet = ss.getSheetByName('Inventory');

    if (!masterSheet) return { success: false, message: 'Sheet Master_Stock tidak ditemukan!' };

    var masterRows = masterSheet.getDataRange().getDisplayValues();
    var existingSkus = {};
    for (var i = 1; i < masterRows.length; i++) {
      if (masterRows[i][0]) existingSkus[String(masterRows[i][0]).toUpperCase().trim()] = true;
    }

    var nowStr = Utilities.formatDate(new Date(), TIMEZONE, 'dd/MM/yyyy HH:mm:ss');
    var newMasterRows = [];
    var newInvRows = [];
    var addedCount = 0;
    var skippedCount = 0;

    for (var k = 0; k < itemsArray.length; k++) {
      var item = itemsArray[k];
      var rawSku = String(item.sku || '').trim();
      if (!rawSku) continue;

      var upperSku = rawSku.toUpperCase();
      if (existingSkus[upperSku]) {
        skippedCount++;
      } else {
        existingSkus[upperSku] = true;
        addedCount++;
        newMasterRows.push([
          rawSku,
          String(item.productName || 'Produk Baru').trim(),
          String(item.description || '').trim(),
          String(item.category || 'Sepatu Pria').trim(),
          String(item.brand || 'Yusca').trim(),
          String(item.imageUrl || 'https://images.unsplash.com/photo-1542291026-7eec264c27ff?w=500').trim(),
          parseFloat(item.price) || 0
        ]);

        newInvRows.push([
          rawSku,
          0,
          'Low Stock',
          nowStr
        ]);
      }
    }

    if (newMasterRows.length > 0) {
      masterSheet.getRange(masterSheet.getLastRow() + 1, 1, newMasterRows.length, newMasterRows[0].length).setValues(newMasterRows);
      if (invSheet) {
        invSheet.getRange(invSheet.getLastRow() + 1, 1, newInvRows.length, newInvRows[0].length).setValues(newInvRows);
      }
      SpreadsheetApp.flush();
      logAdminActivity(adminName || 'Admin', 'Import CSV Master', 'Import ' + addedCount + ' produk baru (' + skippedCount + ' SKU duplikat dilewati)');
    }

    return {
      success: true,
      addedCount: addedCount,
      skippedCount: skippedCount,
      message: 'Import Selesai: ' + addedCount + ' produk baru berhasil ditambahkan. ' + skippedCount + ' SKU duplikat dilewati!'
    };
  } catch (err) {
    return { success: false, message: 'Gagal mengimpor produk CSV: ' + err.toString() };
  }
}

/**
 * Menghapus Master Produk berdasarkan SKU
 */
function deleteMasterProduct(sku, adminName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Master_Stock');
    var rows = sheet.getDataRange().getDisplayValues();

    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === String(sku).trim()) {
        sheet.deleteRow(i + 1);
        SpreadsheetApp.flush();
        logAdminActivity(adminName || 'Admin', 'Hapus Master SKU', 'Hapus produk ' + sku);
        return { success: true, message: 'Produk berhasil dihapus!' };
      }
    }
    return { success: false, message: 'SKU tidak ditemukan!' };
  } catch (err) {
    return { success: false, message: 'Gagal menghapus produk: ' + err.toString() };
  }
}

/**
 * Mengkalkulasi data laporan laba rugi dan finansial berdasarkan rentang tanggal
 */
function getProfitReportData(startDateStr, endDateStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var orderSheet = ss.getSheetByName('Orders');
    var orderRows = orderSheet ? orderSheet.getDataRange().getDisplayValues() : [];

    var totalRevenue = 0;
    var totalOrdersCount = 0;

    for (var i = 1; i < orderRows.length; i++) {
      var status = orderRows[i][5];
      var total = parseFloat(orderRows[i][4]) || 0;
      var dateRaw = orderRows[i][7];

      var includeOrder = true;
      if (startDateStr || endDateStr) {
        var parts = dateRaw.split(' ')[0].split('/');
        if (parts.length === 3) {
          var orderISO = parts[2] + '-' + parts[1] + '-' + parts[0];
          if (startDateStr && orderISO < startDateStr) includeOrder = false;
          if (endDateStr && orderISO > endDateStr) includeOrder = false;
        }
      }

      if (includeOrder && (status === 'Confirmed' || status === 'Delivered')) {
        totalRevenue += total;
        totalOrdersCount++;
      }
    }

    var estimatedCOGS = Math.round(totalRevenue * 0.65);
    var grossProfit = totalRevenue - estimatedCOGS;
    var operationalCost = Math.round(totalRevenue * 0.10);
    var netProfit = grossProfit - operationalCost;

    return {
      success: true,
      data: {
        totalRevenue: totalRevenue,
        estimatedCOGS: estimatedCOGS,
        grossProfit: grossProfit,
        operationalCost: operationalCost,
        netProfit: netProfit,
        completedOrders: totalOrdersCount
      }
    };
  } catch (err) {
    return { success: false, message: 'Gagal membuat laporan profit: ' + err.toString() };
  }
}
