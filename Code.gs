// ============================================================
//  PORTOFOLIO PEGAWAI — Google Apps Script
//  Updated: mendukung field unitKerja, divisi, region, jabatanList
//  Updated: mendukung login admin (sheet USERS)
// ============================================================

const SHEET_NAME = "DATA_PEGAWAI";
const USERS_SHEET_NAME = "USERS"; // Sheet baru: kolom A = username, kolom B = password

function doGet() {
  return ContentService
    .createTextOutput(
      JSON.stringify(getAllData())
    )
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);

  // Jika request adalah percobaan login admin
  if (body.action === "login") {
    return handleLogin(body);
  }

  // Selain itu, dianggap sebagai penyimpanan data portofolio (perilaku lama, tidak berubah)
  return simpanDataPegawai(body);
}

function handleLogin(body) {
  const sh = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(USERS_SHEET_NAME);

  if (!sh) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: "Sheet USERS belum dibuat. Buat sheet bernama 'USERS' dengan kolom: username | password"
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const data = sh.getDataRange().getValues();
  data.shift(); // hapus baris header

  const username = String(body.username || "").trim();
  const password = String(body.password || "").trim();

  const found = data.some(function (row) {
    return String(row[0]).trim() === username && String(row[1]).trim() === password;
  });

  return ContentService
    .createTextOutput(JSON.stringify({ success: found }))
    .setMimeType(ContentService.MimeType.JSON);
}

function simpanDataPegawai(data) {
  const sh = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEET_NAME);

  // Urutan kolom (sesuai index yang dibaca dashboard):
  // [0]  timestamp
  // [1]  nik
  // [2]  nama
  // [3]  unitKerja
  // [4]  divisi
  // [5]  region
  // [6]  masaKerja
  // [7]  jabatanCount
  // [8]  jabatanList
  // [9]  pendidikan
  // [10] pelatihanExt
  // [11] ketExt
  // [12] pelatihanInt
  // [13] ketInt
  // [14] sertifikasi
  // [15] ketSertifikasi
  // [16] inovasi
  // [17] ketInovasi
  // [18] project
  // [19] ketProject
  // [20] total
  // [21] rating

  sh.appendRow([
    new Date(),           // [0]
    data.nik,             // [1]
    data.nama,            // [2]
    data.unitKerja,       // [3]
    data.divisi,          // [4]
    data.region,          // [5]
    data.masaKerja,       // [6]
    data.jabatanCount,    // [7]
    data.jabatanList,     // [8]
    data.pendidikan,      // [9]
    data.pelatihanExt,    // [10]
    data.ketExt,          // [11]
    data.pelatihanInt,    // [12]
    data.ketInt,          // [13]
    data.sertifikasi,     // [14]
    data.ketSertifikasi,  // [15]
    data.inovasi,         // [16]
    data.ketInovasi,      // [17]
    data.project,         // [18]
    data.ketProject,      // [19]
    data.total,           // [20]
    data.rating           // [21]
  ]);

  return ContentService
    .createTextOutput(
      JSON.stringify({ status: "success" })
    )
    .setMimeType(ContentService.MimeType.JSON);
}

function getAllData() {
  const sh = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEET_NAME);

  const data = sh.getDataRange().getValues();

  // Hapus baris header
  data.shift();

  return data;
}
