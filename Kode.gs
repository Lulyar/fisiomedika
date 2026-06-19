// ============================================
// KODE.GS — KLINIK FISIO MEDIKA
// Paste seluruh isi file ini ke Kode.gs 
// di Google Apps Script
// ============================================

// Fungsi utama: serve halaman website
function doGet(e) {
  // Jika ada query action, layani sebagai API JSON (untuk Vercel/hosting eksternal)
  if (e && e.parameter && e.parameter.action) {
    var action = e.parameter.action;
    var result;
    try {
      if (action === 'getAllData') {
        result = getAllData();
      } else if (action === 'upsertData') {
        var sheetName = e.parameter.sheetName;
        var rowData = JSON.parse(e.parameter.data);
        result = upsertData(sheetName, rowData);
      } else if (action === 'deleteData') {
        var sheetName = e.parameter.sheetName;
        var keyData = JSON.parse(e.parameter.data);
        result = deleteData(sheetName, keyData);
      } else if (action === 'getSheetStats') {
        result = getSheetStats();
      } else {
        result = { error: 'Unknown action: ' + action };
      }
    } catch (err) {
      result = { error: err.toString() };
    }
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Default: sajikan halaman HTML jika dibuka langsung di Google Script
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Klinik Fisio Medika — Sistem Manajemen Klinik Fisioterapi')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  var result;
  try {
    var postData = JSON.parse(e.postData.contents);
    var action = postData.action;
    if (action === 'upsertData') {
      result = upsertData(postData.sheetName, postData.data);
    } else if (action === 'deleteData') {
      result = deleteData(postData.sheetName, postData.data);
    } else {
      result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.toString() };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}


// ============================================
// KONFIGURASI SHEET & HEADER
// ============================================

var SHEET_CONFIG = {
  'Pasien': {
    key: 'ID_Pasien',
    headers: ['ID_Pasien','Nama','KTP_NIK','Tgl_Lahir','Jenis_Kelamin','No_HP','Email','Keluhan','Tgl_Daftar','Status']
  },
  'RekamMedis': {
    key: 'ID_Rekam',
    headers: ['ID_Rekam','ID_Pasien','Diagnosa','Rekomendasi_Terapi','Catatan','Pemeriksa','Tgl_Periksa','Status']
  },
  'Jadwal': {
    key: 'ID_Jadwal',
    headers: ['ID_Jadwal','ID_Pasien','Nama_Pasien','Jenis_Terapi','Terapis','Tanggal','Jam','Ruang','Status']
  },
  'Users': {
    key: 'Username',
    headers: ['Username','Password','Nama_Lengkap','Role']
  },
  'Terapis': {
    key: 'ID_Terapis',
    headers: ['ID_Terapis','Nama','Spesialisasi','Deskripsi','Shift_Mulai','Shift_Selesai','Total_Sesi','Pasien_Hari_Ini']
  },
  'JenisTerapi': {
    key: 'Kode_Terapi',
    headers: ['Kode_Terapi','Nama_Terapi','Deskripsi','Durasi_Menit']
  }
};

// ============================================
// INISIALISASI SHEET (FIX HEADER)
// ============================================

function initSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var report = [];
  
  for (var name in SHEET_CONFIG) {
    var config = SHEET_CONFIG[name];
    var sheet = ss.getSheetByName(name);
    
    if (!sheet) {
      sheet = ss.insertSheet(name);
      report.push(name + ': CREATED');
    }
    
    // Set headers di baris 1
    var headerRange = sheet.getRange(1, 1, 1, config.headers.length);
    headerRange.setValues([config.headers]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#0d9488');
    headerRange.setFontColor('#ffffff');
    
    report.push(name + ': Headers OK (' + config.headers.length + ' kolom)');
  }
  
  return report.join('\n');
}

// ============================================
// AMBIL SEMUA DATA
// ============================================

function getAllData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = {};
  
  for (var name in SHEET_CONFIG) {
    var sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) {
      result[name] = [];
      continue;
    }
    var data = sheet.getDataRange().getValues();
    var headers = SHEET_CONFIG[name].headers; // Pakai header dari config, bukan dari sheet
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var row = {};
      for (var j = 0; j < headers.length; j++) {
        row[headers[j]] = (j < data[i].length) ? String(data[i][j]) : '';
      }
      rows.push(row);
    }
    result[name] = rows;
  }
  
  return result;
}

// ============================================
// UPSERT DATA (Insert / Update)
// ============================================

function upsertData(sheetName, rowData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 'Sheet not found: ' + sheetName;
  
  var config = SHEET_CONFIG[sheetName];
  if (!config) return 'Config not found: ' + sheetName;
  
  var headers = config.headers;
  var keyColumnName = config.key;
  var keyIndex = headers.indexOf(keyColumnName);
  var keyValue = rowData[keyColumnName];
  var existingRowIndex = -1;
  
  if (keyIndex !== -1 && keyValue) {
    var values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][keyIndex]) === String(keyValue)) {
        existingRowIndex = i + 1;
        break;
      }
    }
  }
  
  var newRow = [];
  headers.forEach(function(h) {
    newRow.push(rowData[h] !== undefined ? rowData[h] : '');
  });
  
  if (existingRowIndex !== -1) {
    sheet.getRange(existingRowIndex, 1, 1, newRow.length).setValues([newRow]);
    return 'Updated';
  } else {
    sheet.appendRow(newRow);
    return 'Appended';
  }
}

// ============================================
// HAPUS DATA
// ============================================

function deleteData(sheetName, keyData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 'Sheet not found: ' + sheetName;
  
  var config = SHEET_CONFIG[sheetName];
  if (!config) return 'Config not found: ' + sheetName;
  
  var headers = config.headers;
  var keyColumnName = config.key;
  var keyIndex = headers.indexOf(keyColumnName);
  var keyVal = keyData[keyColumnName];
  
  if (keyIndex !== -1 && keyVal) {
    var values = sheet.getDataRange().getValues();
    for (var i = values.length - 1; i >= 1; i--) {
      if (String(values[i][keyIndex]) === String(keyVal)) {
        sheet.deleteRow(i + 1);
      }
    }
    return 'Deleted';
  }
  return 'Key not found';
}

// ============================================
// STATISTIK UNTUK DASHBOARD & LAPORAN
// ============================================

function getSheetStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var stats = {};
  
  // Total pasien
  var pSheet = ss.getSheetByName('Pasien');
  stats.totalPasien = pSheet ? Math.max(0, pSheet.getLastRow() - 1) : 0;
  
  // Jadwal hari ini
  var jSheet = ss.getSheetByName('Jadwal');
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  stats.sesiHariIni = 0;
  stats.menunggu = 0;
  stats.selesai = 0;
  stats.dalamProses = 0;
  
  if (jSheet && jSheet.getLastRow() >= 2) {
    var jData = jSheet.getDataRange().getValues();
    for (var i = 1; i < jData.length; i++) {
      var tgl = '';
      if (jData[i][5] instanceof Date) {
        tgl = Utilities.formatDate(jData[i][5], Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        tgl = String(jData[i][5]);
      }
      if (tgl === today) {
        stats.sesiHariIni++;
        var status = String(jData[i][8]).toLowerCase();
        if (status === 'menunggu') stats.menunggu++;
        else if (status === 'dalam proses') stats.dalamProses++;
        else if (status === 'selesai') stats.selesai++;
      }
    }
  }
  
  return stats;
}
