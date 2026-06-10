// ============================================================
//  PORTOFOLIO PEGAWAI — Google Apps Script
//  Updated: mendukung field unitKerja, divisi, region, jabatanList
// ============================================================

const SHEET_NAME = "DATA_PEGAWAI";

function doGet() {
  return ContentService
    .createTextOutput(
      JSON.stringify(getAllData())
    )
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);

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
