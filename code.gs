/**
 * SIKEMAS - Sistem Keterangan Keluar Masuk Pegawai
 * Copyright Tim Umum dan Humas BPS Provinsi Kalimantan Barat @2026
 * BACKEND LOGIC (Code.gs)
 */

function doGet(e) {
  let page = e && e.parameter && e.parameter.page ? e.parameter.page : 'Index';
  let title = 'SIKEMAS - BPS Prov. Kalbar';
  let templateName = 'Index';
  
  if (page.toLowerCase() === 'monitoring') {
    templateName = 'monitoring';
    title = 'Dashboard Monitoring SIKEMAS - BPS Kalbar';
  }
  
  return HtmlService.createTemplateFromFile(templateName)
    .evaluate()
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function getSheet() {
  let doc;
  try {
    doc = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) { doc = null; }
  
  if (!doc) {
    const properties = PropertiesService.getScriptProperties();
    let sheetId = properties.getProperty("SPREADSHEET_ID");
    if (sheetId) {
      try { doc = SpreadsheetApp.openById(sheetId); } catch(e) { doc = null; }
    }
    if (!doc) {
      doc = SpreadsheetApp.create("Database SIKEMAS - BPS Kalbar");
      properties.setProperty("SPREADSHEET_ID", doc.getId());
    }
  }
  
  let sheet = doc.getSheetByName("Log_SIKEMAS");
  if (!sheet) {
    sheet = doc.insertSheet("Log_SIKEMAS");
    sheet.appendRow(["Nama Pegawai", "Hari", "Tanggal", "Waktu Keluar", "Waktu Kembali", "Keterangan", "Timestamp", "SessionID"]);
    const headerRange = sheet.getRange("A1:H1");
    headerRange.setFontWeight("bold").setBackground("#005aa9").setFontColor("#ffffff").setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Mengambil daftar nama pegawai, tim kerja, NIP, dan role dari sheet "User"
 * (Kolom A: Nama, Kolom B: Tim Kerja, Kolom C: NIP, Kolom D: Password, Kolom E: Role)
 */
function getEmployeeNames() {
  try {
    let doc;
    try { doc = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { doc = null; }
    
    if (!doc) {
      const properties = PropertiesService.getScriptProperties();
      let sheetId = properties.getProperty("SPREADSHEET_ID");
      if (sheetId) {
        try { doc = SpreadsheetApp.openById(sheetId); } catch(e) { doc = null; }
      }
    }
    
    if (!doc) return { success: true, data: [], names: [] };

    let sheet = doc.getSheetByName("User");
    if (!sheet) {
      sheet = doc.insertSheet("User");
      sheet.appendRow(["Nama", "Tim Kerja", "NIP", "Password", "Role"]);
      sheet.appendRow(["Azhari", "Tim IPDS", "199501012020011001", "admin123", "Admin"]);
      sheet.appendRow(["Budi Santoso", "Tim Stat. Distribusi", "199002022015021002", "supervisor123", "Supervisor"]);
      sheet.appendRow(["Dewi Lestari", "Tim Stat. Sosial", "199203032018032003", "ketua123", "Ketua Tim"]);
      sheet.getRange("A1:E1").setFontWeight("bold").setBackground("#e2e8f0");
    } else {
      // Pastikan header 5 kolom
      const headerValues = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 5)).getValues()[0];
      if (!headerValues[2] || String(headerValues[2]).trim() === "") {
        sheet.getRange(1, 1, 1, 5).setValues([["Nama", "Tim Kerja", "NIP", "Password", "Role"]]);
        sheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#e2e8f0");
      }
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, data: [], names: [] };

    const maxCols = Math.max(sheet.getLastColumn(), 5);
    const range = sheet.getRange(2, 1, lastRow - 1, maxCols);
    const values = range.getValues();

    const users = values
      .map(row => {
        const nama = String(row[0] || "").trim();
        const timKerja = String(row[1] || "").trim() || "Lainnya";
        const nip = String(row[2] || "").trim();
        const role = String(row[4] || "").trim() || "Ketua Tim";
        return { nama: nama, timKerja: timKerja, nip: nip, role: role };
      })
      .filter(item => item.nama.length > 0);

    const names = users.map(u => u.nama);
      
    return { success: true, data: users, names: names };
  } catch (error) {
    Logger.log("Error getEmployeeNames: " + error.toString());
    return { success: false, message: error.toString(), data: [], names: [] };
  }
}

/**
 * Autentikasi Login Pengguna berdasarkan NIP (Kolom C) dan Password (Kolom D)
 * Role (Kolom E) didapatkan secara otomatis dari database
 */
function loginUser(data) {
  try {
    const inputNip = String(data.nip || "").trim();
    const inputPassword = String(data.password || "").trim();
    
    if (!inputNip || !inputPassword) {
      return { success: false, message: "NIP dan Password wajib diisi!" };
    }

    let doc;
    try { doc = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { doc = null; }
    if (!doc) {
      const properties = PropertiesService.getScriptProperties();
      let sheetId = properties.getProperty("SPREADSHEET_ID");
      if (sheetId) {
        try { doc = SpreadsheetApp.openById(sheetId); } catch(e) { doc = null; }
      }
    }
    if (!doc) return { success: false, message: "Database Spreadsheet tidak ditemukan." };

    let sheet = doc.getSheetByName("User");
    if (!sheet) {
      getEmployeeNames();
      sheet = doc.getSheetByName("User");
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: "Data pengguna tidak ditemukan di database." };

    const range = sheet.getRange(2, 1, lastRow - 1, 5);
    const values = range.getValues();

    for (let i = 0; i < values.length; i++) {
      const nama = String(values[i][0] || "").trim();
      const timKerja = String(values[i][1] || "").trim() || "Lainnya";
      const nip = String(values[i][2] || "").trim();
      const password = String(values[i][3] || "").trim();
      const role = String(values[i][4] || "").trim() || "Ketua Tim";

      if (nip === inputNip && password === inputPassword) {
        return {
          success: true,
          message: "Login berhasil!",
          user: {
            nama: nama,
            timKerja: timKerja,
            nip: nip,
            role: role
          }
        };
      }
    }

    return { success: false, message: "NIP atau Password tidak sesuai!" };
  } catch (error) {
    Logger.log("Error loginUser: " + error.toString());
    return { success: false, message: "Gagal login: " + error.toString() };
  }
}

// Menyimpan data rencana awal dan mengembalikan Row Index dan Session ID unik
function saveRecord(data) {
  try {
    const sheet = getSheet();
    const timestamp = new Date();
    const sessionId = Utilities.getUuid();
    
    sheet.appendRow([
      data.nama,
      data.hari,
      data.tanggal,
      data.waktuKeluar,
      data.waktuKembali,
      data.keterangan || "-",
      timestamp,
      sessionId
    ]);
    
    const lastRow = sheet.getLastRow();
    return { 
      success: true, 
      message: "Rencana kegiatan berhasil disimpan! Sesi presensi Anda kini aktif.", 
      rowIndex: lastRow,
      sessionId: sessionId,
      autoJamKeluar: isAutoJamKeluarEnabled()
    };
  } catch (error) {
    return { success: false, message: "Gagal menyimpan: " + error.toString() };
  }
}

// Fungsi pembantu mencari baris data sesi aktif secara dinamis dan aman
function findRowBySessionOrEmployee(sessionId, employeeName) {
  try {
    if (!employeeName) return null;
    const sheet = getSheet();
    const values = sheet.getDataRange().getValues();
    
    // 1. Cari berdasarkan Session ID (kolom H / indeks 7)
    if (sessionId) {
      for (let i = 1; i < values.length; i++) {
        if (values[i].length > 7 && String(values[i][7]).trim() === String(sessionId).trim()) {
          if (String(values[i][0]).trim() === employeeName.trim()) {
            return i + 1; // 1-indexed row number
          }
        }
      }
    }
    
    // 2. Fallback: Cari baris terakhir milik pegawai tersebut yang belum mencatat Jam Kembali ("-")
    for (let i = values.length - 1; i >= 1; i--) {
      if (String(values[i][0]).trim() === employeeName.trim() && String(values[i][4]).trim() === "-") {
        return i + 1;
      }
    }
  } catch (e) {
    Logger.log("Error findRowBySessionOrEmployee: " + e.toString());
  }
  return null;
}

// Memperbarui waktu keluar/kembali berdasarkan Session ID/Nama secara aman dan instan
function updateTimeSafe(sessionId, employeeName, type, waktu) {
  try {
    const sheet = getSheet();
    const rowIndex = findRowBySessionOrEmployee(sessionId, employeeName);
    
    if (!rowIndex) {
      return { success: false, message: "Sesi aktif tidak ditemukan di database." };
    }
    
    const colIndex = (type === 'keluar') ? 4 : 5; // Kolom D atau E
    sheet.getRange(rowIndex, colIndex).setValue("'" + waktu);
    return { success: true, message: `Berhasil mencatat Jam ${type === 'keluar' ? 'Keluar' : 'Kembali'}: ${waktu}` };
  } catch (error) {
    return { success: false, message: "Gagal memperbarui waktu: " + error.toString() };
  }
}

// Menghapus baris data di database ketika sesi dibatalkan / direset oleh pengguna secara aman (presisi hanya baris pegawai tersebut)
function deleteRecordSafe(sessionId, employeeName) {
  try {
    const sheet = getSheet();
    const rowIndex = findRowBySessionOrEmployee(sessionId, employeeName);
    
    if (!rowIndex) {
      return { success: false, message: "Sesi aktif tidak ditemukan atau sudah dihapus." };
    }
    
    sheet.deleteRow(rowIndex);
    return { success: true, message: "Sesi berhasil dibatalkan dan rencana kegiatan dihapus dari database SIKEMAS secara permanen." };
  } catch (error) {
    return { success: false, message: "Gagal menghapus data dari database: " + error.toString() };
  }
}

// Memeriksa status sesi aktif pegawai di server secara real-time
function checkSessionActive(sessionId, employeeName) {
  try {
    // Jalankan auto jam keluar jika aktif sebelum memeriksa sesi
    processAutoJamKeluar();

    const rowIndex = findRowBySessionOrEmployee(sessionId, employeeName);
    if (!rowIndex) {
      return { found: false, autoJamKeluar: isAutoJamKeluarEnabled() };
    }
    
    const sheet = getSheet();
    const values = sheet.getRange(rowIndex, 1, 1, 7).getValues()[0];
    const displayValues = sheet.getRange(rowIndex, 1, 1, 7).getDisplayValues()[0];
    
    return {
      found: true,
      nama: String(values[0] || "-"),
      waktuKeluar: String(displayValues[3] || "-"),
      waktuKembali: String(displayValues[4] || "-"),
      keterangan: String(values[5] || "-"),
      autoJamKeluar: isAutoJamKeluarEnabled()
    };
  } catch (error) {
    return { found: false, error: error.toString() };
  }
}

// Mengambil riwayat khusus pegawai tertentu (terisolasi) dengan konversi string aman
function getRecordsInternal(employeeName) {
  try {
    if (!employeeName) return [];
    // Jalankan auto jam keluar jika aktif sebelum mengambil data
    processAutoJamKeluar();

    const sheet = getSheet();
    const range = sheet.getDataRange();
    const rows = range.getValues();
    const displayRows = range.getDisplayValues();
    
    if (rows.length <= 1) return [];

    // Auto fix unreturned sessions past 24:00 WIB
    autoFixExpiredSessions(sheet, rows, displayRows);
    
    const dataRows = [];
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === employeeName.trim()) {
        dataRows.push({
          rawRow: rows[i],
          displayRow: displayRows[i]
        });
      }
    }
    
    dataRows.sort((a, b) => new Date(b.rawRow[6]) - new Date(a.rawRow[6]));
    
    return dataRows.map(item => {
      const raw = item.rawRow;
      const disp = item.displayRow;
      
      let rawDate = raw[2];
      let formattedDate = rawDate;
      if (rawDate instanceof Date) {
        const d = rawDate.getDate().toString().padStart(2, '0');
        const m = (rawDate.getMonth() + 1).toString().padStart(2, '0');
        const y = rawDate.getFullYear();
        formattedDate = `${d}-${m}-${y}`;
      }
      
      return {
        nama: String(raw[0] || "-"),
        hari: String(raw[1] || "-"),
        tanggal: String(formattedDate || "-"),
        waktuKeluar: String(disp[3] || "-"),
        waktuKembali: String(disp[4] || "-"),
        keterangan: String(raw[5] || "-")
      };
    });
  } catch (error) { 
    return []; 
  }
}

function getRecords(data) {
  let employeeName = "";
  if (data && typeof data === 'object') {
    employeeName = data.nama;
  } else if (typeof data === 'string') {
    employeeName = data;
  }
  try {
    const records = getRecordsInternal(employeeName);
    return { success: true, data: records };
  } catch (e) {
    return { success: false, message: e.toString(), data: [] };
  }
}

function updateTime(data) {
  try {
    const sessionId = data.sessionId;
    const employeeName = data.nama;
    const type = data.type;
    const waktu = data.time;
    return updateTimeSafe(sessionId, employeeName, type, waktu);
  } catch (error) {
    return { success: false, message: "Gagal memperbarui waktu: " + error.toString() };
  }
}

function deleteRecord(data) {
  try {
    const sessionId = data.sessionId;
    const employeeName = data.nama;
    return deleteRecordSafe(sessionId, employeeName);
  } catch (error) {
    return { success: false, message: "Gagal menghapus data: " + error.toString() };
  }
}

function checkSession(data) {
  try {
    const sessionId = data.sessionId;
    const employeeName = data.nama;
    return checkSessionActive(sessionId, employeeName);
  } catch (error) {
    return { found: false, error: error.toString() };
  }
}

function doPost(e) {
  let requestData;
  try {
    if (e && e.postData && e.postData.contents) {
      const contents = e.postData.contents;
      try {
        requestData = JSON.parse(contents);
      } catch (err) {
        if (e.parameter && e.parameter.data) {
          requestData = JSON.parse(e.parameter.data);
        } else {
          requestData = e.parameter;
        }
      }
    } else if (e && e.parameter && e.parameter.data) {
      requestData = JSON.parse(e.parameter.data);
    } else if (e && e.parameter) {
      requestData = e.parameter;
    } else {
      requestData = {};
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: "Gagal memproses request payload: " + err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }

  const action = requestData.action;
  let result;

  try {
    switch (action) {
      case 'loginUser':
        result = loginUser(requestData);
        break;
      case 'getEmployeeNames':
      case 'getUsersData':
        result = getEmployeeNames();
        break;
      case 'saveRecord':
        result = saveRecord(requestData);
        break;
      case 'updateTime':
        result = updateTime(requestData);
        break;
      case 'deleteRecord':
        result = deleteRecord(requestData);
        break;
      case 'checkSession':
        result = checkSession(requestData);
        break;
      case 'getRecords':
        result = getRecords(requestData);
        break;
      case 'getAllRecords':
        result = getAllRecords();
        break;
      case 'getSpreadsheetUrl':
        result = { success: true, url: getSpreadsheetUrl() };
        break;
      case 'getScriptUrl':
        result = { success: true, url: getScriptUrl() };
        break;
      case 'adminForceCompleteSession':
        result = adminForceCompleteSession(requestData.sessionId, requestData.nama, requestData.time);
        break;
      case 'adminDeleteRecord':
        result = adminDeleteRecord(requestData.sessionId, requestData.nama);
        break;
      case 'updatePassword':
        result = updatePassword(requestData);
        break;
      case 'getAutoJamKeluarSetting':
        result = getAutoJamKeluarSetting();
        break;
      case 'setAutoJamKeluarSetting':
        result = setAutoJamKeluarSetting(requestData);
        break;
      case 'triggerAutoJamKeluar':
        result = { success: true, processed: processAutoJamKeluar() };
        break;
      default:
        result = { success: false, message: "Aksi tidak dikenal: " + action };
    }
  } catch (err) {
    result = { success: false, message: "Terjadi kesalahan di server: " + err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * =========================================================================
 * FITUR AUTO JAM KELUAR (10 MENIT SETELAH RENCANA DISIMPAN)
 * Pengaturan khusus role Admin & proses otomatisasi pencatatan jam keluar
 * =========================================================================
 */

/**
 * Memeriksa apakah fitur Auto Jam Keluar diaktifkan oleh Admin
 */
function isAutoJamKeluarEnabled() {
  try {
    const properties = PropertiesService.getScriptProperties();
    return properties.getProperty("AUTO_JAM_KELUAR_ENABLED") === "true";
  } catch (e) {
    return false;
  }
}

/**
 * Mengambil status pengaturan Auto Jam Keluar
 */
function getAutoJamKeluarSetting() {
  try {
    const enabled = isAutoJamKeluarEnabled();
    return { success: true, enabled: enabled };
  } catch (e) {
    return { success: false, enabled: false, message: e.toString() };
  }
}

/**
 * Memvalidasi apakah NIP terdaftar sebagai Admin pada sheet User
 */
function verifyAdminRole(nip) {
  try {
    if (!nip) return false;
    let doc;
    try { doc = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { doc = null; }
    if (!doc) {
      const properties = PropertiesService.getScriptProperties();
      const sheetId = properties.getProperty("SPREADSHEET_ID");
      if (sheetId) {
        try { doc = SpreadsheetApp.openById(sheetId); } catch(e) { doc = null; }
      }
    }
    if (!doc) return false;
    const sheet = doc.getSheetByName("User");
    if (!sheet) return false;
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return false;
    const values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    for (let i = 0; i < values.length; i++) {
      const rowNip = String(values[i][2] || "").trim();
      const rowRole = String(values[i][4] || "").trim();
      if (rowNip === String(nip).trim() && rowRole === "Admin") {
        return true;
      }
    }
    return false;
  } catch (e) {
    Logger.log("Error verifyAdminRole: " + e.toString());
    return false;
  }
}

/**
 * Menyimpan pengaturan Auto Jam Keluar (Khusus Role Admin)
 */
function setAutoJamKeluarSetting(data) {
  try {
    const role = String((data && data.role) || "").trim();
    const nip = String((data && data.nip) || "").trim();

    // 1. Validasi hak akses Admin
    if (role !== "Admin") {
      return { 
        success: false, 
        message: "Akses ditolak! Hanya role Admin yang berhak mengakses dan mengubah pengaturan ini." 
      };
    }

    // 2. Validasi NIP di sheet User jika nip disediakan
    if (nip) {
      const isAdmin = verifyAdminRole(nip);
      if (!isAdmin) {
        return { 
          success: false, 
          message: "Akses ditolak! Akun dengan NIP tersebut tidak memiliki hak akses Admin." 
        };
      }
    }

    const enabled = data.enabled === true || data.enabled === "true" || data.enabled === 1 || data.enabled === "1";
    const properties = PropertiesService.getScriptProperties();
    properties.setProperty("AUTO_JAM_KELUAR_ENABLED", enabled ? "true" : "false");

    // Kelola trigger terjadwal di Google Apps Script jika memungkinkan
    setupAutoJamKeluarTrigger(enabled);

    // Jika diaktifkan, langsung proses antrean sesi yang sudah melewati 10 menit
    let processedCount = 0;
    if (enabled) {
      processedCount = processAutoJamKeluar();
    }

    return {
      success: true,
      enabled: enabled,
      processedCount: processedCount,
      message: `Fitur Auto Jam Keluar berhasil ${enabled ? "diaktifkan" : "dinonaktifkan"}.`
    };
  } catch (error) {
    Logger.log("Error setAutoJamKeluarSetting: " + error.toString());
    return { success: false, message: "Gagal menyimpan pengaturan: " + error.toString() };
  }
}

/**
 * Mengelola Google Apps Script time-driven trigger untuk menjalankan background job secara periodik
 */
function setupAutoJamKeluarTrigger(enabled) {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === "runAutoJamKeluarJob") {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
    if (enabled) {
      ScriptApp.newTrigger("runAutoJamKeluarJob")
        .timeBased()
        .everyMinutes(1)
        .create();
    }
  } catch (e) {
    Logger.log("Notice trigger setup (fallback to on-request handling): " + e.toString());
  }
}

/**
 * Handler yang dipanggil oleh time-driven trigger Google Apps Script
 */
function runAutoJamKeluarJob() {
  processAutoJamKeluar();
}

/**
 * Menjalankan proses Auto Jam Keluar:
 * Memeriksa seluruh baris di Log_SIKEMAS yang memiliki waktuKeluar === "-" dan waktuKembali === "-"
 * serta sudah melewati 10 menit (600.000 ms) sejak rencana kegiatan disimpan (timestamp kolom G).
 * Waktu Jam Keluar diisi dengan waktu tepat 10 menit setelah rencana disimpan tanpa duplikasi data.
 */
function processAutoJamKeluar() {
  try {
    if (!isAutoJamKeluarEnabled()) {
      return 0;
    }

    // Gunakan LockService untuk mencegah duplikasi dan race conditions
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(5000)) {
      return 0;
    }

    let modifiedCount = 0;
    try {
      const sheet = getSheet();
      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) return 0;

      const dataRange = sheet.getRange(1, 1, lastRow, Math.max(sheet.getLastColumn(), 8));
      const values = dataRange.getValues();
      const displayValues = dataRange.getDisplayValues();
      const now = new Date();
      const TEN_MINUTES_MS = 10 * 60 * 1000;

      for (let i = 1; i < values.length; i++) {
        const raw = values[i];
        const disp = displayValues[i];
        const waktuKeluar = String(disp[3] || "-").trim();
        const waktuKembali = String(disp[4] || "-").trim();
        const rawTimestamp = raw[6];

        if (waktuKeluar === "-" && waktuKembali === "-" && rawTimestamp) {
          let recDate = null;
          if (rawTimestamp instanceof Date) {
            recDate = rawTimestamp;
          } else {
            const parsed = new Date(rawTimestamp);
            if (!isNaN(parsed.getTime())) {
              recDate = parsed;
            }
          }

          if (recDate) {
            const elapsedMs = now.getTime() - recDate.getTime();
            if (elapsedMs >= TEN_MINUTES_MS) {
              const autoKeluarDate = new Date(recDate.getTime() + TEN_MINUTES_MS);
              let keluarStr = "";
              try {
                keluarStr = Utilities.formatDate(autoKeluarDate, Session.getScriptTimeZone() || "GMT+7", "HH:mm");
              } catch (e) {
                const hh = String(autoKeluarDate.getHours()).padStart(2, '0');
                const mm = String(autoKeluarDate.getMinutes()).padStart(2, '0');
                keluarStr = `${hh}:${mm}`;
              }

              sheet.getRange(i + 1, 4).setValue("'" + keluarStr);
              values[i][3] = keluarStr;
              displayValues[i][3] = keluarStr;
              modifiedCount++;
              Logger.log(`[Auto Jam Keluar] Baris ${i + 1} (${raw[0]}) berhasil dicatat keluar otomatis: ${keluarStr}`);
            }
          }
        }
      }
    } finally {
      lock.releaseLock();
    }

    return modifiedCount;
  } catch (error) {
    Logger.log("Error processAutoJamKeluar: " + error.toString());
    return 0;
  }
}

/**
 * Otomatis set Jam Kembali ke 16:00 bagi sesi yang lupa ditekan tombol kembali setelah 24.00 WIB hari tersebut
 */
function autoFixExpiredSessions(sheet, rows, displayRows) {
  if (!rows || rows.length <= 1) return;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  for (let i = 1; i < rows.length; i++) {
    const raw = rows[i];
    const disp = displayRows ? displayRows[i] : raw;
    const waktuKeluar = String(disp[3] || "-").trim();
    const waktuKembali = String(disp[4] || "-").trim();
    
    if (waktuKeluar !== "-" && waktuKembali === "-") {
      let rawDate = raw[2];
      let recDate = null;
      
      if (rawDate instanceof Date) {
        recDate = new Date(rawDate);
        recDate.setHours(0, 0, 0, 0);
      } else if (typeof rawDate === 'string' && rawDate.includes('-')) {
        const cleanDate = rawDate.split('T')[0];
        const parts = cleanDate.split('-');
        if (parts.length === 3) {
          if (parts[0].length === 4) {
            recDate = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
          } else {
            recDate = new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
          }
        }
      }
      
      if (recDate && recDate.getTime() < today.getTime()) {
        try {
          sheet.getRange(i + 1, 5).setValue("'16:00");
          rows[i][4] = "16:00";
          if (displayRows && displayRows[i]) {
            displayRows[i][4] = "16:00";
          }
        } catch(e) {
          Logger.log("Error autoFixExpiredSessions row " + (i+1) + ": " + e.toString());
        }
      }
    }
  }
}

// Mengambil semua riwayat dari database untuk Dashboard Monitoring
function getAllRecords() {
  try {
    // Jalankan auto jam keluar jika aktif sebelum mengambil data
    processAutoJamKeluar();

    const sheet = getSheet();
    const range = sheet.getDataRange();
    const rows = range.getValues();
    const displayRows = range.getDisplayValues();
    
    if (rows.length <= 1) return { success: true, data: [] };

    autoFixExpiredSessions(sheet, rows, displayRows);

    // Ambil data tim kerja pegawai dari sheet User untuk pemetaan tim
    let doc;
    try { doc = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { doc = null; }
    if (!doc) {
      const properties = PropertiesService.getScriptProperties();
      let sheetId = properties.getProperty("SPREADSHEET_ID");
      if (sheetId) {
        try { doc = SpreadsheetApp.openById(sheetId); } catch(e) { doc = null; }
      }
    }
    const userTeamMap = {};
    if (doc) {
      const userSheet = doc.getSheetByName("User");
      if (userSheet) {
        const uLastRow = userSheet.getLastRow();
        if (uLastRow > 1) {
          const uVals = userSheet.getRange(2, 1, uLastRow - 1, Math.max(userSheet.getLastColumn(), 2)).getValues();
          uVals.forEach(uRow => {
            const uNama = String(uRow[0] || "").trim().toLowerCase();
            const uTim = String(uRow[1] || "").trim();
            if (uNama) userTeamMap[uNama] = uTim;
          });
        }
      }
    }
    
    const records = [];
    for (let i = 1; i < rows.length; i++) {
      const raw = rows[i];
      const disp = displayRows[i];
      
      let rawDate = raw[2];
      let formattedDate = rawDate;
      if (rawDate instanceof Date) {
        const d = rawDate.getDate().toString().padStart(2, '0');
        const m = (rawDate.getMonth() + 1).toString().padStart(2, '0');
        const y = rawDate.getFullYear();
        formattedDate = `${d}-${m}-${y}`;
      } else if (typeof rawDate === 'string' && rawDate.includes('-')) {
        const cleanDate = rawDate.split('T')[0];
        const parts = cleanDate.split('-');
        if (parts.length === 3) {
          if (parts[0].length === 4) {
            formattedDate = `${parts[2]}-${parts[1]}-${parts[0]}`;
          } else {
            formattedDate = cleanDate;
          }
        }
      }
      
      const waktuKeluar = String(disp[3] || "-").trim();
      const waktuKembali = String(disp[4] || "-").trim();
      
      let durasiMenit = null;
      let durasiTeks = "-";
      if (waktuKeluar !== "-" && waktuKembali !== "-") {
        durasiMenit = hitungDurasiMenit(waktuKeluar, waktuKembali);
        if (durasiMenit !== null) {
          durasiTeks = formatDurasi(durasiMenit);
        }
      } else if (waktuKeluar !== "-" && waktuKembali === "-") {
        durasiTeks = "Sedang Keluar";
      }
      
      const empName = String(raw[0] || "-").trim();
      const empTeam = userTeamMap[empName.toLowerCase()] || "-";

      records.push({
        nama: empName,
        timKerja: empTeam,
        hari: String(raw[1] || "-"),
        tanggal: String(formattedDate || "-"),
        waktuKeluar: waktuKeluar,
        waktuKembali: waktuKembali,
        keterangan: String(raw[5] || "-"),
        timestamp: raw[6] ? new Date(raw[6]).toISOString() : "",
        sessionId: raw.length > 7 ? String(raw[7] || "") : "",
        durasiMenit: durasiMenit,
        durasiTeks: durasiTeks
      });
    }
    
    records.sort((a, b) => {
      const timeA = a.timestamp ? new Date(a.timestamp).getTime() : 0;
      const timeB = b.timestamp ? new Date(b.timestamp).getTime() : 0;
      return timeB - timeA;
    });
    
    return { success: true, data: records };
  } catch (error) {
    Logger.log("Error getAllRecords: " + error.toString());
    return { success: false, message: error.toString(), data: [] };
  }
}

function hitungDurasiMenit(keluar, kembali) {
  try {
    const kParts = keluar.split(':');
    const bParts = kembali.split(':');
    if (kParts.length >= 2 && bParts.length >= 2) {
      const kMin = parseInt(kParts[0], 10) * 60 + parseInt(kParts[1], 10);
      const bMin = parseInt(bParts[0], 10) * 60 + parseInt(bParts[1], 10);
      let diff = bMin - kMin;
      if (diff < 0) diff += 24 * 60;
      return diff;
    }
  } catch(e) {}
  return null;
}

function formatDurasi(menit) {
  if (menit === null || menit < 0) return "-";
  const jam = Math.floor(menit / 60);
  const sisaMenit = menit % 60;
  if (jam > 0) {
    return `${jam} jam ${sisaMenit} menit`;
  }
  return `${sisaMenit} menit`;
}

function getSpreadsheetUrl() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet().getUrl();
  } catch (e) {
    const properties = PropertiesService.getScriptProperties();
    let sheetId = properties.getProperty("SPREADSHEET_ID");
    if (sheetId) {
      return "https://docs.google.com/spreadsheets/d/" + sheetId + "/edit";
    }
  }
  return "#";
}

function getScriptUrl() {
  try {
    return ScriptApp.getService().getUrl();
  } catch (e) {
    return "";
  }
}

function adminForceCompleteSession(sessionId, employeeName, kembaliTime) {
  try {
    const sheet = getSheet();
    const rowIndex = findRowBySessionOrEmployee(sessionId, employeeName);
    
    if (!rowIndex) {
      return { success: false, message: "Sesi aktif tidak ditemukan di database." };
    }
    
    sheet.getRange(rowIndex, 5).setValue("'" + kembaliTime);
    return { success: true, message: `Sesi ${employeeName} berhasil diselesaikan pada pukul ${kembaliTime}.` };
  } catch (error) {
    return { success: false, message: "Gagal menyelesaikan sesi: " + error.toString() };
  }
}

function adminDeleteRecord(sessionId, employeeName) {
  try {
    return deleteRecordSafe(sessionId, employeeName);
  } catch (error) {
    return { success: false, message: "Gagal menghapus log: " + error.toString() };
  }
}

/**
 * Memperbarui Password Akun Pegawai pada sheet "User"
 * data: { nip, oldPassword, newPassword }
 */
function updatePassword(data) {
  try {
    const inputNip = String((data && data.nip) || "").trim();
    const oldPassword = String((data && data.oldPassword) || "").trim();
    const newPassword = String((data && data.newPassword) || "").trim();

    if (!inputNip) {
      return { success: false, message: "NIP pengguna tidak valid!" };
    }
    if (!oldPassword || !newPassword) {
      return { success: false, message: "Password lama dan password baru wajib diisi!" };
    }
    if (newPassword.length < 4) {
      return { success: false, message: "Password baru minimal 4 karakter!" };
    }

    let doc;
    try { doc = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { doc = null; }
    if (!doc) {
      const properties = PropertiesService.getScriptProperties();
      let sheetId = properties.getProperty("SPREADSHEET_ID");
      if (sheetId) {
        try { doc = SpreadsheetApp.openById(sheetId); } catch(e) { doc = null; }
      }
    }
    if (!doc) return { success: false, message: "Database Spreadsheet tidak ditemukan." };

    let sheet = doc.getSheetByName("User");
    if (!sheet) return { success: false, message: "Sheet User tidak ditemukan di database." };

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: "Data pengguna tidak ditemukan di database." };

    const values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    for (let i = 0; i < values.length; i++) {
      const nip = String(values[i][2] || "").trim();
      const currentPassword = String(values[i][3] || "").trim();

      if (nip === inputNip) {
        if (currentPassword !== oldPassword) {
          return { success: false, message: "Password lama tidak sesuai!" };
        }
        // Kolom D (ke-4) adalah Password
        sheet.getRange(i + 2, 4).setValue(newPassword);
        return { 
          success: true, 
          message: "Password berhasil diperbarui! Silakan gunakan password baru pada login berikutnya." 
        };
      }
    }

    return { success: false, message: "Akun pengguna dengan NIP tersebut tidak ditemukan." };
  } catch (error) {
    Logger.log("Error updatePassword: " + error.toString());
    return { success: false, message: "Gagal memperbarui password: " + error.toString() };
  }
}