/**
 * GOOGLE APPS SCRIPT BACKEND UNTUK BUKU KAS FUTSAL (PRO TABLE EDITION)
 * 
 * Fitur:
 * - Format Tabel Excel Pro
 * - Judul Kolom Tebal (Bold Header)
 * - Format Mata Uang Otomatis "Rp #,##0"
 * - Baris Total Tebal (Bold Total Row)
 */

function doGet(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Rekap Kas Futsal") || ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();
    
    // Jika sheet baru/kosong, inisialisasi tabel pro
    if (data.length <= 1) {
      formatProSheet(sheet, []);
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        data: []
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[1] && !row[2] && !row[4] && !row[5]) continue;
      if (row[0] && row[0].toString().toUpperCase() === "TOTAL") continue;
      
      var tgl = row[1];
      if (tgl instanceof Date) {
        tgl = Utilities.formatDate(tgl, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      
      rows.push({
        id: row[0] || i,
        date: tgl ? tgl.toString() : "",
        desc: row[2] ? row[2].toString() : "",
        category: row[3] ? row[3].toString() : "Iuran Kas",
        income: Number(row[4]) || 0,
        expense: Number(row[5]) || 0,
        balance: Number(row[6]) || 0
      });
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      data: rows
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Rekap Kas Futsal") || ss.getSheets()[0];
    
    var payload;
    if (e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else if (e.parameter && e.parameter.data) {
      payload = JSON.parse(e.parameter.data);
    } else {
      payload = e.parameter;
    }
    
    if (Array.isArray(payload.transactions)) {
      formatProSheet(sheet, payload.transactions);
      
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "Seluruh data berhasil disinkronkan dengan format tabel pro",
        rowCount: payload.transactions.length
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "Data diterima"
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Fungsi pembantu untuk memformat sheet menjadi tabel profesional
function formatProSheet(sheet, items) {
  sheet.clear();
  
  // 1. Header Judul Kolom (BOLD & PRO COLOR)
  var headers = ["No", "Tanggal", "Keterangan Transaksi", "Kategori", "Pemasukan (Rp)", "Pengeluaran (Rp)", "Sisa Kas (Rp)"];
  sheet.appendRow(headers);
  
  var headerRange = sheet.getRange(1, 1, 1, 7);
  headerRange.setFontWeight("bold")
             .setBackground("#107C41")
             .setFontColor("#FFFFFF")
             .setHorizontalAlignment("center")
             .setVerticalAlignment("middle");
  sheet.setRowHeight(1, 32);
  
  var totalIncome = 0;
  var totalExpense = 0;
  
  // 2. Baris Data Transaksi
  if (items.length > 0) {
    var rowsData = [];
    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      var inc = Number(item.income) || 0;
      var exp = Number(item.expense) || 0;
      var bal = Number(item.balance) || 0;
      
      totalIncome += inc;
      totalExpense += exp;
      
      rowsData.push([
        item.id || (i + 1),
        item.date || "",
        item.desc || "",
        item.category || "Iuran Kas",
        inc,
        exp,
        bal
      ]);
    }
    
    sheet.getRange(2, 1, rowsData.length, 7).setValues(rowsData);
    
    // Zebra striping & Borders
    for (var r = 0; r < rowsData.length; r++) {
      var rowNum = 2 + r;
      var bg = (r % 2 === 1) ? "#F8FAF9" : "#FFFFFF";
      sheet.getRange(rowNum, 1, 1, 7).setBackground(bg).setBorder(true, true, true, true, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
      sheet.setRowHeight(rowNum, 24);
    }
    
    // Format Angka Rupiah (Kolom E, F, G)
    var numFormat = '"Rp "#,##0';
    sheet.getRange(2, 5, rowsData.length, 3).setNumberFormat(numFormat).setFontWeight("bold");
    sheet.getRange(2, 5, rowsData.length, 1).setFontColor("#15803D"); // Hijau Pemasukan
    sheet.getRange(2, 6, rowsData.length, 1).setFontColor("#B91C1C"); // Merah Pengeluaran
    sheet.getRange(2, 7, rowsData.length, 1).setFontColor("#1D4ED8"); // Biru Sisa
  }
  
  // 3. Baris Total di Bagian Bawah (BOLD & DOUBLE UNDERLINE)
  var lastRow = sheet.getLastRow() + 1;
  sheet.appendRow(["TOTAL KESELURUHAN", "", "", "", totalIncome, totalExpense, totalIncome - totalExpense]);
  
  var totalRange = sheet.getRange(lastRow, 1, 1, 7);
  totalRange.setFontWeight("bold")
            .setBackground("#E2E8F0")
            .setFontColor("#0F172A")
            .setVerticalAlignment("middle")
            .setBorder(true, true, true, true, null, null, "#475569", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  sheet.getRange(lastRow, 5, 1, 3).setNumberFormat('"Rp "#,##0');
  sheet.setRowHeight(lastRow, 28);
  
  // Auto-fit kolom agar pas dan rapi
  sheet.setColumnWidth(1, 55);   // No
  sheet.setColumnWidth(2, 110);  // Tanggal
  sheet.setColumnWidth(3, 260);  // Keterangan
  sheet.setColumnWidth(4, 140);  // Kategori
  sheet.setColumnWidth(5, 135);  // Pemasukan
  sheet.setColumnWidth(6, 135);  // Pengeluaran
  sheet.setColumnWidth(7, 135);  // Sisa Kas
}
