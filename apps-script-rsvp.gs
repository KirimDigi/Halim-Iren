// Google Apps Script untuk RSVP Halim & Iren
// 1. Buka spreadsheet: https://docs.google.com/spreadsheets/d/1pQJrFVEhobOQoZZz0onkFzYGHvkmuiRMKEM_e308O-g/edit
// 2. Buat sheet baru bernama "RSVP" (jika belum ada) dengan header di baris 1: Timestamp | Nama | Ucapan | Kehadiran | Jumlah_Tamu | Pasangan
// 3. Extensions > Apps Script > paste code ini > Deploy > New deployment > Web app > Execute as Me, Who has access: Anyone
// 4. Copy URL Web App dan ganti di index.html: const SHEET_URL = "https://script.google.com/macros/s/XXXX/exec"

const SHEET_ID = "1pQJrFVEhobOQoZZz0onkFzYGHvkmuiRMKEM_e308O-g";
const SHEET_NAME = "RSVP";

function doPost(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(["Timestamp", "Nama", "Ucapan", "Kehadiran", "Jumlah_Tamu", "Pasangan", "UserAgent"]);
    }
    // e.postData.contents bisa berupa JSON string (karena fetch dengan text/plain)
    let data = {};
    if (e.postData && e.postData.contents) {
      try { data = JSON.parse(e.postData.contents); } catch (err) {
        // fallback form-encoded
        data = e.parameter || {};
      }
    } else {
      data = e.parameter || {};
    }

    // Simpan timestamp sebagai Date agar di Sheets muncul jam & tanggal (WIB)
    let ts = new Date();
    if (data.timestamp) {
      try { ts = new Date(data.timestamp); } catch(e2) {}
    }
    const row = [
      ts,
      data.nama || "",
      data.ucapan || "",
      data.kehadiran || "",
      data.jumlah_tamu || "",
      data.pasangan || "Halim & Iren",
      e.parameter && e.parameter.userAgent || ""
    ];
    sheet.appendRow(row);
    // Format kolom A sebagai tanggal jam
    try { sheet.getRange(sheet.getLastRow(), 1).setNumberFormat("dd/MM/yyyy HH:mm:ss"); } catch(e3) {}
    
    return ContentService.createTextOutput(JSON.stringify({status:"ok"}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status:"error", message: String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    // Jika ?action=read maka kembalikan JSON daftar RSVP untuk ditampilkan setelah reload
    if (e.parameter && e.parameter.action === "read") {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const sheet = ss.getSheetByName(SHEET_NAME);
      if (!sheet) {
        return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
      }
      const values = sheet.getDataRange().getValues();
      // values[0] = header
      const data = [];
      for (let i = 1; i < values.length; i++) {
        const r = values[i];
        if (!r[1] && !r[2]) continue; // skip baris kosong (nama & ucapan kosong)
        data.push({
          timestamp: r[0] ? (r[0] instanceof Date ? r[0].toISOString() : String(r[0])) : "",
          nama: String(r[1] || ""),
          ucapan: String(r[2] || ""),
          kehadiran: String(r[3] || ""),
          jumlah_tamu: String(r[4] || ""),
          pasangan: String(r[5] || "")
        });
      }
      // Terbaru dulu
      data.reverse();
      return ContentService.createTextOutput(JSON.stringify(data))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput("RSVP Halim & Iren ready").setMimeType(ContentService.MimeType.TEXT);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({status:"error", message:String(err)})).setMimeType(ContentService.MimeType.JSON);
  }
}

function doOptions(e) {
  return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.TEXT);
}
