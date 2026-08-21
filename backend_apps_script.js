// Kode ini untuk dijalankan di Google Apps Script (extensions -> Apps Script dari Google Sheets)

function doGet(e) {
  return HtmlService.createHtmlOutput('Sistem Buku Induk Backend Aktif');
}

function setupDatabase() {
  var sheetName = "DataSiswa";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  // Base headers tanpa nilai
  var baseHeaders = [
    "id_unik", "nis", "nama", "jk", "kelas", "nik", "nokk", "anak_ke", "tempat_lahir", "tgl_lahir", 
    "no_akte", "asal_sd", "nisn", "no_ijazah_sd", "no_ujian_sd", "agama",
    "alamat", "rt", "rw", "kelurahan", "kecamatan", "kode_pos", "tempat_tinggal", "transportasi", "no_hp", "no_telp", "email",
    "tinggi_badan", "berat_badan", "jarak", "waktu_tempuh", "jml_saudara",
    "nama_ayah", "tgl_lahir_ayah", "pendidikan_ayah", "pekerjaan_ayah", "penghasilan_ayah",
    "nama_ibu", "tgl_lahir_ibu", "pendidikan_ibu", "pekerjaan_ibu", "penghasilan_ibu",
    "nama_wali", "tgl_lahir_wali", "pendidikan_wali", "pekerjaan_wali", "penghasilan_wali",
    "mutasi_masuk_sekolah", "mutasi_masuk_alamat", "mutasi_masuk_surat", "mutasi_masuk_tgl", "mutasi_masuk_kelas",
    "mutasi_keluar_surat", "mutasi_keluar_tgl", "mutasi_keluar_kelas", "mutasi_keluar_alasan", "mutasi_keluar_tujuan", "mutasi_keluar_alamat",
    "lulus_tgl", "lulus_ijazah", "lulus_peserta", "lulus_lanjut",
    "putus_tgl", "putus_kelas", "putus_alasan",
    "catatan"
  ];
  
  // Mengenerate otomatis 60 kolom mapel sesuai format template JS
  var subjects = ['pabp', 'pancasila', 'bindo', 'mtk', 'ipa', 'ips', 'binggris', 'pjok', 'informatika', 'seni'];
  for(var i = 1; i <= 6; i++) {
    for(var j = 0; j < subjects.length; j++) {
      baseHeaders.push("nilai_smt" + i + "_" + subjects[j]);
    }
  }
  
  baseHeaders.push("timestamp");
  
  sheet.getRange(1, 1, 1, baseHeaders.length).setValues([baseHeaders]);
  sheet.getRange(1, 1, 1, baseHeaders.length).setFontWeight("bold").setBackground("#d9ead3");
  
  // Sheet Pengaturan untuk Password Admin
  var configSheet = ss.getSheetByName("Pengaturan");
  if(!configSheet) {
    configSheet = ss.insertSheet("Pengaturan");
    configSheet.getRange("A1").setValue("Password Admin");
    configSheet.getRange("B1").setValue("admin123");
  }
}

function simpanData(dataObj) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("DataSiswa");
  
  var id = dataObj.id_unik || Utilities.getUuid();
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var rowData = [];
  
  for (var i = 0; i < headers.length; i++) {
    var header = headers[i];
    if (header === 'id_unik') {
      rowData.push(id);
    } else if (header === 'timestamp') {
      rowData.push(new Date());
    } else {
      rowData.push(dataObj[header] || "");
    }
  }
  
  var found = false;
  if(dataObj.id_unik) {
    var dataIds = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), 1).getValues();
    for(var j=0; j<dataIds.length; j++) {
      if(dataIds[j][0] === id) {
        sheet.getRange(j + 2, 1, 1, rowData.length).setValues([rowData]);
        found = true;
        break;
      }
    }
  }
  
  if(!found) {
    sheet.appendRow(rowData);
  }
  
  return JSON.stringify({status: "success", id: id});
}

function getDataSiswa() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("DataSiswa");
  if(sheet.getLastRow() < 2) return JSON.stringify([]);
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var result = [];
  
  for(var i=1; i<data.length; i++) {
    var obj = {};
    for(var j=0; j<headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    result.push(obj);
  }
  return JSON.stringify(result);
}
