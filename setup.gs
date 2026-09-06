/**
 * ============================================================================
 * ZETTBOT 3.1 - DATABASE SETUP & SAFE MIGRATION SYSTEM (SINGLE STORE)
 * File: setup.gs
 * Framework: Google Apps Script
 * ============================================================================
 */

function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var timezone = 'Asia/Jakarta';
  var nowStr = Utilities.formatDate(new Date(), timezone, 'dd/MM/yyyy HH:mm:ss');

  // Skema Tabel & Header (Single Store Inventory & Jenis Produk)
  var schema = {
    'Users': ['Name', 'Phone', 'Password', 'Role', 'UserID', 'DateCreated'],
    'Master_Stock': ['SKU', 'ProductName', 'Description', 'Category', 'Brand', 'ImageURL', 'Price'],
    'Inventory': ['SKU', 'Quantity', 'Status', 'DateUpdated'],
    'Jenis_Produk': ['ID', 'Type', 'Name', 'Description', 'DateCreated'],
    'Orders': ['OrderID', 'BuyerUserID', 'SKUs', 'Quantities', 'TotalPrice', 'Status', 'PaymentProofImage', 'DateOrdered', 'TransactionID'],
    'Transactions': ['TransactionID', 'OrderID', 'Amount', 'Type', 'Date', 'ReferenceNote']
  };

  for (var sheetName in schema) {
    var sheet = ss.getSheetByName(sheetName);
    var expectedHeaders = schema[sheetName];

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(expectedHeaders);
      sheet.getRange(1, 1, 1, expectedHeaders.length).setFontWeight('bold').setBackground('#006064').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    } else {
      var headerRange = sheet.getRange(1, 1, 1, expectedHeaders.length);
      headerRange.setValues([expectedHeaders]);
      headerRange.setFontWeight('bold').setBackground('#006064').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
  }

  var userSheet = ss.getSheetByName('Users');
  if (userSheet.getLastRow() <= 1) {
    var dummyUsers = [
      ['Erik Dermawan', '081234567890', 'admin123', 'Super Admin', 'USR-0001', nowStr],
      ['Budi Santoso', '081298765432', 'admin123', 'Admin', 'USR-0002', nowStr],
      ['Siti Aminah', '085712345678', 'buyer123', 'Pembeli', 'CUST-0001', nowStr]
    ];
    userSheet.getRange(2, 1, dummyUsers.length, dummyUsers[0].length).setValues(dummyUsers);
  }

  var jenisSheet = ss.getSheetByName('Jenis_Produk');
  if (jenisSheet.getLastRow() <= 1) {
    var dummyJenis = [
      ['JPN-0001', 'Brand', 'Adidas', 'Brand sportswear global premium', nowStr],
      ['JPN-0002', 'Brand', 'Nike', 'Brand sepatu & pakaian olahraga', nowStr],
      ['JPN-0003', 'Brand', 'Puma', 'Brand sepatu olahraga & lifestyle', nowStr],
      ['JPN-0004', 'Category', 'Sepatu Pria', 'Kategori sneakers & sepatu pria', nowStr],
      ['JPN-0005', 'Category', 'Sepatu Lari', 'Kategori sepatu lari & jogging', nowStr],
      ['JPN-0006', 'Category', 'Aksesoris', 'Kategori tas, kaus kaki, & aksesoris', nowStr],
      ['JPN-0007', 'Target / Size', 'Pria (Size 39-45)', 'Target pria dewasa ukuran standar EU', nowStr],
      ['JPN-0008', 'Target / Size', 'Wanita (Size 36-40)', 'Target wanita dewasa ukuran standar EU', nowStr]
    ];
    jenisSheet.getRange(2, 1, dummyJenis.length, dummyJenis[0].length).setValues(dummyJenis);
  }

  var masterSheet = ss.getSheetByName('Master_Stock');
  if (masterSheet.getLastRow() <= 1) {
    var dummyMaster = [
      ['SKU-SP-001', 'Adidas Superstar - Size 42 (Black)', 'Sepatu kasual klasik kulit sintetis dengan shell toe ikonik.', 'Sepatu Pria', 'Adidas', 'https://images.unsplash.com/photo-1582588678413-dbf45f4823e9?w=500', 1250000],
      ['SKU-SP-002', 'Adidas Superstar - Size 43 (Black)', 'Sepatu kasual klasik kulit sintetis ukuran 43.', 'Sepatu Pria', 'Adidas', 'https://images.unsplash.com/photo-1582588678413-dbf45f4823e9?w=500', 1250000],
      ['SKU-NMD-001', 'Adidas NMD R1 Primeknit', 'Sneakers running urban dengan bantalan boost responsif.', 'Sepatu Pria', 'Adidas', 'https://images.unsplash.com/photo-1584735935682-2f2b69dff9d2?w=500', 2100000],
      ['SKU-UB-001', 'Ultraboost Light Running Shoes', 'Sepatu lari jarak jauh paling ringan dengan teknologi energi terbarukan.', 'Sepatu Lari', 'Adidas', 'https://images.unsplash.com/photo-1542291026-7eec264c27ff?w=500', 2800000],
      ['SKU-ACC-001', 'Running Waist Bag Sporty', 'Tas pinggang olahraga anti-air dengan slot botol minum.', 'Aksesoris', 'Nike', 'https://images.unsplash.com/photo-1553062407-98eeb64c6a62?w=500', 185000],
      ['SKU-APP-001', 'AeroReady Training Jersey', 'Baju kaos dry-fit untuk gym dan lari dengan ventilasi optimal.', 'Aksesoris', 'Puma', 'https://images.unsplash.com/photo-1521572267360-ee0c2909d518?w=500', 350000]
    ];
    masterSheet.getRange(2, 1, dummyMaster.length, dummyMaster[0].length).setValues(dummyMaster);
  }

  var invSheet = ss.getSheetByName('Inventory');
  if (invSheet.getLastRow() <= 1) {
    var dummyInv = [
      ['SKU-SP-001', 3, 'Low Stock', nowStr],
      ['SKU-SP-002', 2, 'Low Stock', nowStr],
      ['SKU-NMD-001', 14, 'In Stock', nowStr],
      ['SKU-UB-001', 8, 'In Stock', nowStr],
      ['SKU-ACC-001', 25, 'In Stock', nowStr],
      ['SKU-APP-001', 18, 'In Stock', nowStr]
    ];
    invSheet.getRange(2, 1, dummyInv.length, dummyInv[0].length).setValues(dummyInv);
  }

  var props = PropertiesService.getScriptProperties();
  var todayDateStr = Utilities.formatDate(new Date(), timezone, 'yyyy-MM-dd');
  props.setProperty('ORD_LAST_DATE', todayDateStr);
  props.setProperty('ORD_LAST_COUNT', '4');
  props.setProperty('TRX_LAST_DATE', todayDateStr);
  props.setProperty('TRX_LAST_COUNT', '4');
  props.setProperty('JPN_LAST_DATE', todayDateStr);
  props.setProperty('JPN_LAST_COUNT', '8');

  SpreadsheetApp.flush();
  Logger.log('Inisialisasi Database ZETT STORE (Single Store) Selesai!');
}
