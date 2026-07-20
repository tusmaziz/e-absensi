/**
 * Sistem Absensi Barcode Code 128 - Google Apps Script Backend
 * SMA Al Aziz Bangil
 */

// Menampilkan halaman utama aplikasi web saat diakses
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Sistem Absensi Barcode SMA Al Aziz Bangil')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * SETUP DATABASE OTOMATIS
 * Jalankan fungsi ini sekali di Google Apps Script Editor untuk membuat sheet & data awal.
 */
function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Setup Sheet Anggota
  var sheetAnggota = ss.getSheetByName("Anggota");
  if (!sheetAnggota) {
    sheetAnggota = ss.insertSheet("Anggota");
    sheetAnggota.appendRow(["ID", "Nama", "Peran", "Detail", "Barcode"]);
    // Data awal demo
    sheetAnggota.appendRow(["M001", "Ahmad Fauzi", "Guru", "Wali Kelas XII-MIPA-1", "GR-001"]);
    sheetAnggota.appendRow(["M002", "Nur Hidayah, S.Pd", "Guru", "Guru Kimia", "GR-002"]);
    sheetAnggota.appendRow(["M003", "Muhammad Syarif", "Siswa", "XII-MIPA-1", "SW-101"]);
    sheetAnggota.appendRow(["M004", "Lailatul Qodriyah", "Siswa", "XI-IPS-2", "SW-102"]);
    sheetAnggota.appendRow(["M005", "Deni Prasetyo", "Siswa", "X-Fase-E", "SW-103"]);
    Logger.log("Sheet 'Anggota' berhasil dibuat.");
  }

  // 2. Setup Sheet Kehadiran (Logs)
  var sheetKehadiran = ss.getSheetByName("Kehadiran");
  if (!sheetKehadiran) {
    sheetKehadiran = ss.insertSheet("Kehadiran");
    sheetKehadiran.appendRow(["ID_Absen", "Barcode", "Nama", "Peran", "Tipe", "Timestamp", "Status"]);
    Logger.log("Sheet 'Kehadiran' berhasil dibuat.");
  }

  // 3. Setup Sheet Kebijakan
  var sheetKebijakan = ss.getSheetByName("Kebijakan");
  if (!sheetKebijakan) {
    sheetKebijakan = ss.insertSheet("Kebijakan");
    sheetKebijakan.appendRow(["Key", "Value"]);
    sheetKebijakan.appendRow(["timeStart", "06:00"]);
    sheetKebijakan.appendRow(["timeLate", "07:15"]);
    sheetKebijakan.appendRow(["timeLeave", "13:30"]);
    Logger.log("Sheet 'Kebijakan' berhasil dibuat.");
  }
  
  return "Database Berhasil Dikonfigurasi! Silakan muat ulang halaman.";
}

/**
 * Mengambil semua data gabungan untuk inisialisasi aplikasi di sisi klien
 */
function getInitialData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Tarik data Anggota
    var sheetAnggota = ss.getSheetByName("Anggota");
    var members = [];
    if (sheetAnggota) {
      var dataAnggota = sheetAnggota.getDataRange().getValues();
      for (var i = 1; i < dataAnggota.length; i++) {
        members.push({
          id: dataAnggota[i][0],
          name: dataAnggota[i][1],
          role: dataAnggota[i][2],
          detail: dataAnggota[i][3],
          barcode: dataAnggota[i][4]
        });
      }
    }

    // Tarik data Kehadiran
    var sheetKehadiran = ss.getSheetByName("Kehadiran");
    var attendance = [];
    if (sheetKehadiran) {
      var dataKehadiran = sheetKehadiran.getDataRange().getValues();
      for (var j = 1; j < dataKehadiran.length; j++) {
        attendance.push({
          id: dataKehadiran[j][0],
          barcode: dataKehadiran[j][1],
          name: dataKehadiran[j][2],
          role: dataKehadiran[j][3],
          type: dataKehadiran[j][4],
          timestamp: Number(dataKehadiran[j][5]), // Simpan dalam format epoch ms
          status: dataKehadiran[j][6]
        });
      }
    }

    // Tarik data Kebijakan
    var sheetKebijakan = ss.getSheetByName("Kebijakan");
    var config = { timeStart: "06:00", timeLate: "07:15", timeLeave: "13:30" };
    if (sheetKebijakan) {
      var dataKebijakan = sheetKebijakan.getDataRange().getValues();
      for (var k = 1; k < dataKebijakan.length; k++) {
        var key = dataKebijakan[k][0];
        var val = dataKebijakan[k][1];
        if (key === "timeStart") config.timeStart = val;
        if (key === "timeLate") config.timeLate = val;
        if (key === "timeLeave") config.timeLeave = val;
      }
    }

    return {
      success: true,
      members: members,
      attendance: attendance,
      config: config
    };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Menyimpan Anggota Baru ke Google Sheet
 */
function addMemberGS(member) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Anggota");
    if (!sheet) return { success: false, error: "Sheet Anggota tidak ditemukan." };
    
    sheet.appendRow([member.id, member.name, member.role, member.detail, member.barcode]);
    return { success: true };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Menghapus Anggota berdasarkan ID khusus
 */
function deleteMemberGS(id) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Anggota");
    if (!sheet) return { success: false, error: "Sheet Anggota tidak ditemukan." };
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        sheet.deleteRow(i + 1); // Indeks baris Google Sheets dimulai dari 1
        return { success: true };
      }
    }
    return { success: false, error: "ID anggota tidak ditemukan." };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Menambahkan Baris Log Presensi Baru
 */
function addAttendanceLogGS(log) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Kehadiran");
    if (!sheet) return { success: false, error: "Sheet Kehadiran tidak ditemukan." };
    
    sheet.appendRow([log.id, log.barcode, log.name, log.role, log.type, log.timestamp, log.status]);
    return { success: true };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Menghapus log hari ini (mengosongkan sheet selain header)
 */
function clearLogsGS() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Kehadiran");
    if (!sheet) return { success: false, error: "Sheet Kehadiran tidak ditemukan." };
    
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
    return { success: true };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Memperbarui kebijakan waktu masuk dan pulang
 */
function saveConfigGS(config) {
  try {
    var ss = SpreadsheetApp.openById("1j9AKmvVvoZMhyFUxxCeSIjmYMq-unM3-w-gQX244smA");
    var sheet = ss.getSheetByName("Kebijakan");
    if (!sheet) return { success: false, error: "Sheet Kebijakan tidak ditemukan." };
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var key = data[i][0];
      if (key === "timeStart") sheet.getRange(i + 1, 2).setValue(config.timeStart);
      if (key === "timeLate") sheet.getRange(i + 1, 2).setValue(config.timeLate);
      if (key === "timeLeave") sheet.getRange(i + 1, 2).setValue(config.timeLeave);
    }
    return { success: true };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}