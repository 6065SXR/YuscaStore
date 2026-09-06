/**
 * ============================================================================
 * ZETTBOT 3.2 - BACKEND CORE & AUTHENTICATION (Yusca STORE)
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
    if (!sheet) return;

    var logId = generateSequentialId('LOG', 'LOG_LAST_DATE', 'LOG_LAST_COUNT');
    var nowStr = Utilities.formatDate(new Date(), TIMEZONE, 'dd/MM/yyyy HH:mm:ss');
    sheet.appendRow([logId, adminName || 'Admin System', actionType, details, nowStr]);
    SpreadsheetApp.flush();
  } catch (err) {
    Logger.log('Error logAdminActivity: ' + err.toString());
  }
}

/**
 * Helper function untuk menormalisasi format nomor HP
 */
function normalizePhone(phoneStr) {
  if (!phoneStr) return '';
  var clean = String(phoneStr).trim().replace(/\D/g, '');
  if (clean.indexOf('62') === 0) {
    clean = clean.slice(2);
  }
  while (clean.indexOf('0') === 0) {
    clean = clean.slice(1);
  }
  return clean;
}

/**
 * Otentikasi login pengguna
 */
function loginUser(phone, password) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Users');
    if (!sheet) return { success: false, message: 'Database tabel Users tidak ditemukan di Google Sheets!' };

    var displayRows = sheet.getDataRange().getDisplayValues();
    var rawRows = sheet.getDataRange().getValues();

    var inputNormPhone = normalizePhone(phone);
    var cleanInputPhone = String(phone).trim();
    var cleanPass = String(password).trim();

    if (!cleanInputPhone) {
      return { success: false, message: 'Nomor HP / Username tidak boleh kosong!' };
    }

    for (var i = 1; i < displayRows.length; i++) {
      var dRow = displayRows[i];
      var rRow = rawRows[i];

      var phoneBDisplay = dRow[1] ? String(dRow[1]).trim() : '';
      var phoneBRaw = rRow[1] ? String(rRow[1]).trim() : '';
      var phoneADisplay = dRow[0] ? String(dRow[0]).trim() : '';

      var normBDisplay = normalizePhone(phoneBDisplay);
      var normBRaw = normalizePhone(phoneBRaw);
      var normADisplay = normalizePhone(phoneADisplay);

      var phoneMatched = false;
      var phoneColIdx = 1;

      if ((inputNormPhone && normBDisplay === inputNormPhone) || 
          (inputNormPhone && normBRaw === inputNormPhone) || 
          phoneBDisplay === cleanInputPhone || 
          phoneBRaw === cleanInputPhone) {
        phoneMatched = true;
        phoneColIdx = 1;
      } else if ((inputNormPhone && normADisplay === inputNormPhone) || phoneADisplay === cleanInputPhone) {
        phoneMatched = true;
        phoneColIdx = 0;
      }

      if (phoneMatched) {
        var passColIdx = (phoneColIdx === 1) ? 2 : 1;
        var dbPassDisplay = dRow[passColIdx] ? String(dRow[passColIdx]).trim() : '';
        var dbPassRaw = rRow[passColIdx] ? String(rRow[passColIdx]).trim() : '';

        if (cleanPass === dbPassDisplay || cleanPass === dbPassRaw) {
          var userName = (phoneColIdx === 1) ? (dRow[0] || 'User') : (dRow[2] || 'User');
          var userPhone = dRow[phoneColIdx] || phone;
          var userRole = dRow[3] || 'Pembeli';
          var userId = dRow[4] || ('CUST-' + i);

          return {
            success: true,
            user: {
              name: userName,
              phone: userPhone,
              role: userRole,
              userId: userId,
              recipientName: dRow[6] || userName,
              address: dRow[7] || '',
              courierNotes: dRow[8] || '',
              rt: dRow[9] || '',
              rw: dRow[10] || '',
              kelurahan: dRow[11] || '',
              kecamatan: dRow[12] || '',
              city: dRow[13] || '',
              postalCode: dRow[14] || ''
            }
          };
        } else {
          return { success: false, message: 'Password salah untuk nomor HP: ' + cleanInputPhone };
        }
      }
    }
    return { success: false, message: 'Nomor HP / Username (' + cleanInputPhone + ') tidak ditemukan di database Users!' };
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
    var inputNormPhone = normalizePhone(cleanPhone);

    for (var i = 1; i < rows.length; i++) {
      var dbPhone = String(rows[i][1]).trim();
      var dbNormPhone = normalizePhone(dbPhone);

      if (dbPhone === cleanPhone || (inputNormPhone && dbNormPhone === inputNormPhone)) {
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
      nowStr,
      String(userData.name).trim(), '', '', '', '', '', '', '', ''
    ];

    sheet.appendRow(newRow);
    SpreadsheetApp.flush();

    return {
      success: true,
      user: {
        name: userData.name,
        phone: cleanPhone,
        role: 'Pembeli',
        userId: nextId,
        recipientName: userData.name,
        address: '', courierNotes: '', rt: '', rw: '', kelurahan: '', kecamatan: '', city: '', postalCode: ''
      }
    };
  } catch (err) {
    return { success: false, message: 'Gagal mendaftarkan akun: ' + err.toString() };
  }
}

/**
 * Menyimpan data profil & alamat pengiriman lengkap pelanggan
 */
function saveUserProfileAddress(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Users');
    if (!sheet) return { success: false, message: 'Sheet Users tidak ditemukan!' };

    var cleanPhone = String(payload.phone).trim();
    var normInputPhone = normalizePhone(cleanPhone);
    var rows = sheet.getDataRange().getDisplayValues();

    for (var i = 1; i < rows.length; i++) {
      var dbPhone = String(rows[i][1]).trim();
      var dbNormPhone = normalizePhone(dbPhone);

      if (dbPhone === cleanPhone || (normInputPhone && dbNormPhone === normInputPhone)) {
        sheet.getRange(i + 1, 1).setValue(payload.name || rows[i][0]);
        sheet.getRange(i + 1, 7).setValue(payload.recipientName || '');
        sheet.getRange(i + 1, 8).setValue(payload.address || '');
        sheet.getRange(i + 1, 9).setValue(payload.courierNotes || '');
        sheet.getRange(i + 1, 10).setValue(payload.rt || '');
        sheet.getRange(i + 1, 11).setValue(payload.rw || '');
        sheet.getRange(i + 1, 12).setValue(payload.kelurahan || '');
        sheet.getRange(i + 1, 13).setValue(payload.kecamatan || '');
        sheet.getRange(i + 1, 14).setValue(payload.city || '');
        sheet.getRange(i + 1, 15).setValue(payload.postalCode || '');

        SpreadsheetApp.flush();

        return {
          success: true,
          message: 'Data alamat pengiriman berhasil diperbarui!',
          user: {
            name: payload.name || rows[i][0],
            phone: cleanPhone,
            role: rows[i][3],
            userId: rows[i][4],
            recipientName: payload.recipientName || '',
            address: payload.address || '',
            courierNotes: payload.courierNotes || '',
            rt: payload.rt || '',
            rw: payload.rw || '',
            kelurahan: payload.kelurahan || '',
            kecamatan: payload.kecamatan || '',
            city: payload.city || '',
            postalCode: payload.postalCode || ''
          }
        };
      }
    }

    return { success: false, message: 'Data pengguna tidak ditemukan!' };
  } catch (err) {
    return { success: false, message: 'Gagal menyimpan profil: ' + err.toString() };
  }
}

/**
 * Pengambilan metrik dashboard analitik
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

    var invStockMap = {};
    for (var v = 1; v < invRows.length; v++) {
      var iSku = String(invRows[v][0]).trim();
      var iQty = parseInt(invRows[v][1], 10) || 0;
      if (iSku) invStockMap[iSku] = (invStockMap[iSku] || 0) + iQty;
    }

    var criticalAlerts = [];
    for (var j = 1; j < masterRows.length; j++) {
      var mSku = String(masterRows[j][0]).trim();
      if (!mSku) continue;
      var mQty = invStockMap[mSku] !== undefined ? invStockMap[mSku] : 0;

      if (mQty <= 5) {
        criticalAlerts.push({
          sku: mSku,
          productName: masterRows[j][1] || mSku,
          quantity: mQty,
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
 * Pengolahan laporan laba rugi berdasar rentang tanggal
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
