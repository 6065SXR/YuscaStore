/**
 * ============================================================================
 * ZETTBOT 3.2 - BACKEND STORE OPERATIONS & CRUD (Yusca STORE)
 * File: Code_Store.gs
 * Framework: Google Apps Script Web App
 * ============================================================================
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
      var s = String(invRows[i][0]).trim();
      var q = parseInt(invRows[i][1], 10) || 0;
      if (s) stockMap[s] = (stockMap[s] || 0) + q;
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
      var sku = String(row[0]).trim();
      if (!sku) continue;

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
          if (String(invRows[r][0]).trim() === cartItem.sku) {
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

function getOrdersPaginated(page, pageSize, statusFilter, searchQuery) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Orders');
    var userSheet = ss.getSheetByName('Users');
    if (!sheet) return { success: true, items: [], total: 0 };

    var userMap = {};
    if (userSheet) {
      var userRows = userSheet.getDataRange().getDisplayValues();
      for (var u = 1; u < userRows.length; u++) {
        var uId = userRows[u][4];
        if (uId) {
          userMap[uId] = {
            name: userRows[u][0],
            phone: userRows[u][1],
            recipientName: userRows[u][6] || userRows[u][0],
            address: userRows[u][7] || '',
            courierNotes: userRows[u][8] || '',
            rt: userRows[u][9] || '',
            rw: userRows[u][10] || '',
            kelurahan: userRows[u][11] || '',
            kecamatan: userRows[u][12] || '',
            city: userRows[u][13] || '',
            postalCode: userRows[u][14] || ''
          };
        }
      }
    }

    var rows = sheet.getDataRange().getDisplayValues();
    var all = [];
    var search = (searchQuery || '').toLowerCase().trim();
    var filterStatus = (statusFilter || 'All').trim();

    for (var i = rows.length - 1; i >= 1; i--) {
      var row = rows[i];
      if (filterStatus !== 'All' && row[5] !== filterStatus) continue;

      var bId = row[1] || '';
      var bInfo = userMap[bId] || { 
        name: 'Siti Aminah', 
        phone: '085712345678',
        recipientName: 'Siti Aminah',
        address: 'Jl. Pemuda No. 45',
        courierNotes: 'Pagar warna hitam',
        rt: '02',
        rw: '05',
        kelurahan: 'Rawamangun',
        kecamatan: 'Pulogadung',
        city: 'Jakarta Timur',
        postalCode: '13220'
      };

      if (search && row[0].toLowerCase().indexOf(search) === -1 && 
          bId.toLowerCase().indexOf(search) === -1 && 
          bInfo.name.toLowerCase().indexOf(search) === -1 && 
          bInfo.phone.toLowerCase().indexOf(search) === -1) {
        continue;
      }

      all.push({
        orderId: row[0],
        buyerUserId: bId,
        buyerName: bInfo.name,
        buyerPhone: bInfo.phone,
        buyerAddress: bInfo,
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

function getInventoryPaginated(page, pageSize, searchQuery) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName('Master_Stock');
    var invSheet = ss.getSheetByName('Inventory');

    if (!masterSheet) return { success: true, items: [], total: 0 };

    var masterRows = masterSheet.getDataRange().getDisplayValues();
    var invRows = invSheet ? invSheet.getDataRange().getDisplayValues() : [];

    var stockMap = {};
    for (var i = 1; i < invRows.length; i++) {
      var row = invRows[i];
      var skuKey = String(row[0]).trim();
      if (!skuKey) continue;

      var qtyVal = parseInt(row[1], 10);
      if (isNaN(qtyVal)) qtyVal = 0;

      var statusVal = row[2] || '';
      var dateVal = row[3] || '';

      if (!stockMap[skuKey]) {
        stockMap[skuKey] = { quantity: qtyVal, status: statusVal, dateUpdated: dateVal };
      } else {
        stockMap[skuKey].quantity += qtyVal;
        if (dateVal) stockMap[skuKey].dateUpdated = dateVal;
      }
    }

    var all = [];
    var search = (searchQuery || '').toLowerCase().trim();

    for (var m = 1; m < masterRows.length; m++) {
      var mRow = masterRows[m];
      var sku = String(mRow[0]).trim();
      if (!sku) continue;

      var productName = mRow[1] || sku;
      var category = mRow[3] || '-';
      var brand = mRow[4] || '-';

      if (search && sku.toLowerCase().indexOf(search) === -1 && productName.toLowerCase().indexOf(search) === -1) {
        continue;
      }

      var invData = stockMap[sku] || { quantity: 0, status: 'Low Stock', dateUpdated: '-' };
      var calcStatus = invData.quantity <= 5 ? 'Low Stock' : 'In Stock';

      all.push({
        sku: sku,
        productName: productName,
        category: category,
        brand: brand,
        quantity: invData.quantity,
        status: calcStatus,
        dateUpdated: invData.dateUpdated || '-'
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

function updateInventoryQuantity(sku, newQuantity, adminName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Inventory');
    if (!sheet) return { success: false, message: 'Sheet Inventory tidak ditemukan.' };

    var nowStr = Utilities.formatDate(new Date(), TIMEZONE, 'dd/MM/yyyy HH:mm:ss');
    var qtyNum = parseInt(newQuantity, 10) || 0;
    var status = qtyNum <= 5 ? 'Low Stock' : 'In Stock';
    var cleanSku = String(sku).trim();

    var rows = sheet.getDataRange().getDisplayValues();
    var matchingRowIndices = [];

    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === cleanSku) {
        matchingRowIndices.push(i + 1);
      }
    }

    if (matchingRowIndices.length > 0) {
      var primaryRow = matchingRowIndices[0];
      sheet.getRange(primaryRow, 2).setValue(qtyNum);
      sheet.getRange(primaryRow, 3).setValue(status);
      sheet.getRange(primaryRow, 4).setValue(nowStr);

      for (var d = matchingRowIndices.length - 1; d > 0; d--) {
        sheet.deleteRow(matchingRowIndices[d]);
      }
    } else {
      sheet.appendRow([cleanSku, qtyNum, status, nowStr]);
    }

    SpreadsheetApp.flush();
    logAdminActivity(adminName || 'Admin', 'Update Stok', 'SKU ' + cleanSku + ' stok diubah ke ' + qtyNum);

    return { success: true, message: 'Stok SKU ' + cleanSku + ' berhasil diperbarui!' };
  } catch (err) {
    return { success: false, message: 'Gagal memperbarui stok: ' + err.toString() };
  }
}

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

function deleteMasterProduct(sku, adminName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName('Master_Stock');
    var invSheet = ss.getSheetByName('Inventory');
    var cleanSku = String(sku).trim();

    if (masterSheet) {
      var mRows = masterSheet.getDataRange().getDisplayValues();
      for (var i = mRows.length - 1; i >= 1; i--) {
        if (String(mRows[i][0]).trim() === cleanSku) {
          masterSheet.deleteRow(i + 1);
        }
      }
    }

    if (invSheet) {
      var iRows = invSheet.getDataRange().getDisplayValues();
      for (var k = iRows.length - 1; k >= 1; k--) {
        if (String(iRows[k][0]).trim() === cleanSku) {
          invSheet.deleteRow(k + 1);
        }
      }
    }

    SpreadsheetApp.flush();
    logAdminActivity(adminName || 'Admin', 'Hapus Master SKU', 'Hapus produk ' + cleanSku);

    return { success: true, message: 'Produk SKU ' + cleanSku + ' berhasil dihapus dari Master & Inventori!' };
  } catch (err) {
    return { success: false, message: 'Gagal menghapus produk: ' + err.toString() };
  }
}
