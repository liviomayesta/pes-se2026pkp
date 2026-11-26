// =====================================================================
// KODE GOOGLE APPS SCRIPT: BACKEND DATABASE REAL-TIME PES SE2026
// Tempelkan kode ini di Google Spreadsheet: Ekstensi -> Apps Script
// =====================================================================


// ---------------------------------------------------------------------
// KONFIGURASI SHEET Data_SLS (format baru dari Excel FASIH per petugas)
// Target = ASSIGNED (beban), Realisasi = DIOLAH (selesai)
// ---------------------------------------------------------------------
var SLS_SHEET = "Data_SLS";
var SLS_STATUS_KEYS = ["open", "draft", "dicacah", "submit", "approve", "complete", "edit", "revoke", "reject"];
var SLS_HEADERS = ["Kode PPL", "Nama", "Target", "Realisasi", "Open", "Draft", "Dicacah", "Submit", "Approve", "Complete", "Edit", "Revoke", "Reject", "Terakhir Diupdate"];
var SLS_UNASSIGNED_CODE = "BELUM-ADA-PETUGAS";

// Ambil nilai status dari item kiriman web (mendukung nama lama: completed/submitted/rejected)
function slsStatusValue(it, key) {
  var fd = it.fasihDetail || {};
  var legacy = { complete: "completed", submit: "submitted", reject: "rejected" };
  var v = fd[key];
  if (v === undefined || v === null) v = it[key];
  if ((v === undefined || v === null) && legacy[key]) {
    v = (fd[legacy[key]] !== undefined) ? fd[legacy[key]] : it[legacy[key]];
  }
  return Number(v) || 0;
}

function buildSlsRow(it, fallbackCode, now) {
  var row = [
    it.kodePpl || fallbackCode,
    it.nama || "-",
    Number(it.target) || 0,
    Number(it.realisasi) || 0
  ];
  for (var i = 0; i < SLS_STATUS_KEYS.length; i++) row.push(slsStatusValue(it, SLS_STATUS_KEYS[i]));
  row.push(now);
  return row;
}

// Tulis ulang seluruh sheet Data_SLS (22 petugas + 1 baris "belum ada petugas" bila ada)
function writeSlsSheet(ss, items, unassigned, now) {
  var sheet = getOrCreateSheet(ss, SLS_SHEET, SLS_HEADERS);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, SLS_HEADERS.length).setValues([SLS_HEADERS]).setFontWeight("bold").setBackground("#e2e8f0");
  sheet.setFrozenRows(1);
  var rows = [];
  for (var i = 0; i < items.length; i++) rows.push(buildSlsRow(items[i], "PPL-" + (i + 1), now));
  if (unassigned && ((Number(unassigned.target) || 0) > 0 || (Number(unassigned.realisasi) || 0) > 0)) {
    rows.push(buildSlsRow({
      kodePpl: SLS_UNASSIGNED_CODE,
      nama: "(Belum ada petugas di FASIH)",
      target: unassigned.target,
      realisasi: unassigned.realisasi,
      fasihDetail: unassigned.detail || unassigned
    }, SLS_UNASSIGNED_CODE, now));
  }
  if (rows.length > 0) sheet.getRange(2, 1, rows.length, SLS_HEADERS.length).setValues(rows);
  return rows.length;
}

// Baca Data_SLS berdasarkan NAMA KOLOM (aman walau urutan kolom berubah / sheet masih format lama)
function readSlsSheet(ss) {
  var sheet = getOrCreateSheet(ss, SLS_SHEET, SLS_HEADERS);
  var values = sheet.getDataRange().getValues();
  var out = { list: [], unassigned: null };
  if (values.length < 2) return out;
  var head = [];
  for (var h = 0; h < values[0].length; h++) head.push(String(values[0][h]).trim().toLowerCase());
  var col = function (name) { return head.indexOf(name.toLowerCase()); };
  var cKode = col("Kode PPL"), cNama = col("Nama"), cTarget = col("Target"), cReal = col("Realisasi"), cUpd = col("Terakhir Diupdate");

  for (var r = 1; r < values.length; r++) {
    var v = values[r];
    if (!v[cKode]) continue;
    var obj = {
      kodePpl: String(v[cKode]),
      nama: v[cNama],
      target: Number(v[cTarget]) || 0,
      realisasi: Number(v[cReal]) || 0,
      updated_at: cUpd >= 0 ? v[cUpd] : ""
    };
    // Kolom status hanya ikut dikirim bila ada di sheet (sheet lama tidak punya Dicacah/Approve)
    for (var k = 0; k < SLS_STATUS_KEYS.length; k++) {
      var c = col(SLS_STATUS_KEYS[k]);
      if (c >= 0) obj[SLS_STATUS_KEYS[k]] = Number(v[c]) || 0;
    }
    if (obj.kodePpl === SLS_UNASSIGNED_CODE) out.unassigned = obj;
    else out.list.push(obj);
  }
  return out;
}

// FUNGSI INSTAN: BISA DIJALANKAN SEKALI UNTUK MENGISI SELURUH SHEET
function isiDataOtomatis() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var now = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss WIB");

  // 1. ISI SHEET Data_SLS (22 PPL - Update 26 Nov 2026 11:47, sumber Excel FASIH)
  // Kolom: Kode PPL, Nama, Target(ASSIGNED), Realisasi(DIOLAH), Open, Draft, Dicacah, Submit, Approve, Complete, Edit, Revoke, Reject
  var rawSls = [
    ["PPL-01", "Amelia", 460, 320, 0, 139, 321, 6, 189, 124, 0, 1, 1],
     ["PPL-02", "Fitri Ayu", 491, 339, 53, 97, 341, 13, 128, 198, 0, 0, 2],
     ["PPL-03", "Siti Khodijah", 449, 400, 0, 44, 405, 16, 291, 93, 0, 0, 5],
     ["PPL-04", "Sawarita", 266, 254, 0, 8, 258, 12, 153, 89, 0, 0, 4],
     ["PPL-05", "Sintia", 284, 226, 55, 0, 229, 153, 53, 20, 0, 0, 3],
     ["PPL-06", "Renny Anggraini", 431, 431, 0, 0, 431, 0, 80, 351, 0, 0, 0],
     ["PPL-07", "Retno Kharismawati", 422, 422, 0, 0, 422, 0, 3, 419, 0, 0, 0],
     ["PPL-08", "Rosmita", 480, 480, 0, 0, 480, 0, 214, 266, 0, 0, 0],
     ["PPL-09", "Ria Mulyanti", 313, 219, 22, 55, 236, 60, 107, 52, 0, 0, 17],
     ["PPL-10", "Nabila Lestari", 521, 521, 0, 0, 521, 0, 167, 354, 0, 0, 0],
     ["PPL-11", "Ninda Juwita", 312, 195, 74, 36, 202, 14, 163, 18, 0, 0, 7],
     ["PPL-12", "Pirdania", 453, 453, 0, 0, 453, 0, 150, 303, 0, 0, 0],
     ["PPL-13", "Husnawati", 349, 346, 0, 3, 346, 0, 146, 200, 0, 0, 0],
     ["PPL-14", "Ferrawati", 442, 433, 0, 9, 433, 0, 198, 235, 0, 0, 0],
     ["PPL-15", "Dahlia", 445, 352, 3, 44, 398, 79, 150, 123, 0, 0, 46],
     ["PPL-16", "Nopi Astutik", 586, 585, 0, 1, 585, 0, 201, 384, 0, 0, 0],
     ["PPL-17", "Sri Megawana", 487, 487, 0, 0, 487, 0, 223, 264, 0, 0, 0],
     ["PPL-18", "Sudalia", 450, 450, 0, 0, 450, 0, 6, 444, 0, 0, 0],
     ["PPL-19", "Romadina", 336, 326, 0, 8, 328, 0, 64, 262, 0, 0, 2],
     ["PPL-20", "Piska Aryani", 398, 372, 0, 25, 373, 17, 110, 245, 0, 0, 1],
     ["PPL-21", "Zsa Zsa Putri", 440, 287, 126, 0, 314, 0, 46, 241, 0, 0, 27],
     ["PPL-22", "Gepa", 405, 355, 7, 33, 365, 2, 115, 238, 0, 0, 10]
  ];
  var itemsSls = [];
  for (var i = 0; i < rawSls.length; i++) {
    var r = rawSls[i];
    var detail = {};
    for (var k = 0; k < SLS_STATUS_KEYS.length; k++) detail[SLS_STATUS_KEYS[k]] = r[4 + k];
    itemsSls.push({ kodePpl: r[0], nama: r[1], target: r[2], realisasi: r[3], fasihDetail: detail });
  }
  // Baris "-" di Excel: beban yang belum punya petugas di FASIH
  var unassignedSls = {"target": 81, "realisasi": 46, "detail": {"open": 0, "draft": 33, "dicacah": 48, "submit": 33, "approve": 13, "complete": 0, "edit": 0, "revoke": 0, "reject": 2}};
  writeSlsSheet(ss, itemsSls, unassignedSls, now);

  // 2. ISI SHEET Data_UB_Petugas (5 PPL UB)
  var sheetUbP = getOrCreateSheet(ss, "Data_UB_Petugas", ["Nama", "Target", "Realisasi", "Dicacah", "Submit", "Diolah", "Draft", "Open", "Pengawas", "Terakhir Diupdate"]);
  sheetUbP.clearContents();
  sheetUbP.appendRow(["Nama", "Target", "Realisasi", "Dicacah", "Submit", "Diolah", "Draft", "Open", "Pengawas", "Terakhir Diupdate"]);
  sheetUbP.getRange(1, 1, 1, 10).setFontWeight("bold").setBackground("#fef3c7");
  
  var rawUbP = [
  {
    "nama": "Adi Dirgantara",
    "email": "dirgantara@bps.go.id",
    "pml": "Jawani Eka P.",
    "target": 7,
    "realisasi": 0,
    "open": 7,
    "draft": 0,
    "dicacah": 0,
    "submit": 0,
    "approve": 0,
    "diolah": 0,
    "catatan": "7 perusahaan masih Open."
  },
  {
    "nama": "Farel Aristo H.",
    "email": "farel.aristo@bps.go.id",
    "pml": "Jawani Eka P.",
    "target": 8,
    "realisasi": 0,
    "open": 8,
    "draft": 0,
    "dicacah": 0,
    "submit": 0,
    "approve": 0,
    "diolah": 0,
    "catatan": "8 perusahaan masih Open."
  },
  {
    "nama": "Shafira Husna",
    "email": "shafirahusna@bps.go.id",
    "pml": "Sarah Putri Utami",
    "target": 8,
    "realisasi": 0,
    "open": 7,
    "draft": 1,
    "dicacah": 0,
    "submit": 0,
    "approve": 0,
    "diolah": 0,
    "catatan": "1 perusahaan Draft, 7 perusahaan Open."
  },
  {
    "nama": "Soraya Afkarina",
    "email": "sorayaafkarina@bps.go.id",
    "pml": "Sarah Putri Utami",
    "target": 8,
    "realisasi": 2,
    "open": 5,
    "draft": 1,
    "dicacah": 2,
    "submit": 2,
    "approve": 0,
    "diolah": 2,
    "catatan": "2 perusahaan selesai dicacah & masuk tahap pengolahan (Submit/Diolah: 2, Draft: 1, Open: 5)."
  },
  {
    "nama": "Maryani",
    "email": "marian@bps.go.id",
    "pml": "Sarah Putri Utami",
    "target": 7,
    "realisasi": 0,
    "open": 6,
    "draft": 1,
    "dicacah": 0,
    "submit": 0,
    "approve": 0,
    "diolah": 0,
    "catatan": "1 perusahaan Draft, 6 perusahaan Open."
  }
];
  var rowsUbP = [];
  for (var j = 0; j < rawUbP.length; j++) {
    var u = rawUbP[j];
    rowsUbP.push([u.nama, u.target, u.realisasi, u.dicacah||0, u.submit||0, u.diolah||0, u.draft||0, u.open||0, u.pml||"-", now]);
  }
  sheetUbP.getRange(2, 1, rowsUbP.length, 10).setValues(rowsUbP);

  // (Data_UB 38 sampel diisi dari web: Pengaturan Cloud -> Unggah Semua Data / sync_all)

  SpreadsheetApp.getUi().alert("Sukses! Seluruh data (Data_SLS 22 PPL dan Data_UB_Petugas) berhasil diisi ke spreadsheet!");
}

// FUNGSI API GET (DIBACA OLEH WEB PORTAL)
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "get_all";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (action === "get_all" || action === "get_ub" || action === "get_sls") {
    // 1. Data Sampel 38 UB
    var sheetUb = getOrCreateSheet(ss, "Data_UB", ["IDSBR", "Nama Perusahaan", "Status", "PPL", "Pengawas", "Terakhir Diupdate"]);
    var dataUb = sheetUb.getDataRange().getValues();
    var ubList = [];
    for (var i = 1; i < dataUb.length; i++) {
      if (dataUb[i][0]) {
        ubList.push({
          idsbr: String(dataUb[i][0]),
          usaha: dataUb[i][1],
          status: dataUb[i][2],
          ppl: dataUb[i][3],
          pml: dataUb[i][4],
          updated_at: dataUb[i][5]
        });
      }
    }
    
    // 2. Data Petugas SLS (22 PPL) - dibaca berdasarkan nama kolom
    var slsRead = readSlsSheet(ss);
    var slsList = slsRead.list;

    // 3. Data Progres Petugas UB (5 PPL)
    var sheetUbPetugas = getOrCreateSheet(ss, "Data_UB_Petugas", ["Nama", "Target", "Realisasi", "Dicacah", "Submit", "Diolah", "Draft", "Open", "Pengawas", "Terakhir Diupdate"]);
    var dataUbP = sheetUbPetugas.getDataRange().getValues();
    var ubPetugasList = [];
    for (var k = 1; k < dataUbP.length; k++) {
      if (dataUbP[k][0]) {
        ubPetugasList.push({
          nama: dataUbP[k][0],
          target: Number(dataUbP[k][1]) || 0,
          realisasi: Number(dataUbP[k][2]) || 0,
          dicacah: Number(dataUbP[k][3]) || 0,
          submit: Number(dataUbP[k][4]) || 0,
          diolah: Number(dataUbP[k][5]) || 0,
          draft: Number(dataUbP[k][6]) || 0,
          open: Number(dataUbP[k][7]) || 0,
          pml: dataUbP[k][8],
          updated_at: dataUbP[k][9]
        });
      }
    }

    var result = {
      status: "success",
      timestamp: new Date().toISOString(),
      ub_data: ubList,
      sls_data: slsList,
      sls_unassigned: slsRead.unassigned,
      ub_petugas_data: ubPetugasList
    };

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: "ready" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// FUNGSI API POST (DISIMPAN DARI IMPORT EXCEL / WEB PORTAL)
function doPost(e) {
  // Kunci agar dua penyimpanan bersamaan tidak saling menimpa (clear + tulis ulang sheet)
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var raw = e.postData.contents;
    var payload = JSON.parse(raw);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var action = payload.action;
    var now = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss WIB");

    // 1. UPDATE STATUS UB SATUAN
    if (action === "update_ub") {
      var sheetUb = getOrCreateSheet(ss, "Data_UB", ["IDSBR", "Nama Perusahaan", "Status", "PPL", "Pengawas", "Terakhir Diupdate"]);
      var data = sheetUb.getDataRange().getValues();
      var foundRow = -1;
      var targetIdsbr = String(payload.idsbr);

      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]) === targetIdsbr) {
          foundRow = i + 1;
          break;
        }
      }

      if (foundRow !== -1) {
        sheetUb.getRange(foundRow, 3).setValue(payload.status);
        sheetUb.getRange(foundRow, 6).setValue(now);
      } else {
        sheetUb.appendRow([targetIdsbr, payload.usaha || "-", payload.status || "Belum Berhasil Didata", payload.ppl || "-", payload.pml || "-", now]);
      }

      return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Status UB berhasil disimpan!" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 2. SINKRONISASI DATA SLS DARI EXCEL FASIH (22 PPL + baris "belum ada petugas")
    if (action === "sync_sls") {
      var countSls = writeSlsSheet(ss, payload.items || [], payload.unassigned || null, now);
      return ContentService.createTextOutput(JSON.stringify({ status: "success", count: countSls }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 3. SINKRONISASI DATA PETUGAS UB (5 PPL)
    if (action === "sync_ub_petugas") {
      var sheetUbP = getOrCreateSheet(ss, "Data_UB_Petugas", ["Nama", "Target", "Realisasi", "Dicacah", "Submit", "Diolah", "Draft", "Open", "Pengawas", "Terakhir Diupdate"]);
      var pList = payload.petugas || payload.items || [];
      sheetUbP.clearContents();
      sheetUbP.appendRow(["Nama", "Target", "Realisasi", "Dicacah", "Submit", "Diolah", "Draft", "Open", "Pengawas", "Terakhir Diupdate"]);
      sheetUbP.getRange(1, 1, 1, 10).setFontWeight("bold").setBackground("#fef3c7");
      
      var rowsUbP = [];
      for (var u = 0; u < pList.length; u++) {
        var p = pList[u];
        rowsUbP.push([
          p.nama || "-",
          p.target || 0,
          p.realisasi || 0,
          p.dicacah || 0,
          p.submit || 0,
          p.diolah || 0,
          p.draft || 0,
          p.open || 0,
          p.pml || "-",
          now
        ]);
      }
      if (rowsUbP.length > 0) {
        sheetUbP.getRange(2, 1, rowsUbP.length, 10).setValues(rowsUbP);
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "success", count: rowsUbP.length }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 4. BULK SYNC 38 SAMPEL UB
    if (action === "bulk_sync_ub") {
      var sheetUb = getOrCreateSheet(ss, "Data_UB", ["IDSBR", "Nama Perusahaan", "Status", "PPL", "Pengawas", "Terakhir Diupdate"]);
      var items = payload.items || [];
      sheetUb.clearContents();
      sheetUb.appendRow(["IDSBR", "Nama Perusahaan", "Status", "PPL", "Pengawas", "Terakhir Diupdate"]);
      sheetUb.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#e2e8f0");
      
      var rowsToInsert = [];
      for (var k = 0; k < items.length; k++) {
        var itm = items[k];
        rowsToInsert.push([String(itm.idsbr), itm.usaha || "-", itm.status || "Belum Berhasil Didata", itm.ppl || "-", itm.pml || "-", now]);
      }
      if (rowsToInsert.length > 0) {
        sheetUb.getRange(2, 1, rowsToInsert.length, 6).setValues(rowsToInsert);
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "success", count: rowsToInsert.length }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 5. SINKRONISASI SEMUA DATA SEKALIGUS
    if (action === "sync_all") {
      if (payload.sls_items) {
        writeSlsSheet(ss, payload.sls_items, payload.unassigned || null, now);
      }
      if (payload.ub_items) {
        var shU = getOrCreateSheet(ss, "Data_UB", ["IDSBR", "Nama Perusahaan", "Status", "PPL", "Pengawas", "Terakhir Diupdate"]);
        shU.clearContents();
        shU.appendRow(["IDSBR", "Nama Perusahaan", "Status", "PPL", "Pengawas", "Terakhir Diupdate"]);
        var rU = [];
        for (var i2 = 0; i2 < payload.ub_items.length; i2++) {
          var y = payload.ub_items[i2];
          rU.push([String(y.idsbr), y.usaha, y.status, y.ppl, y.pml, now]);
        }
        if (rU.length > 0) shU.getRange(2, 1, rU.length, 6).setValues(rU);
      }
      if (payload.ub_petugas) {
        var shP = getOrCreateSheet(ss, "Data_UB_Petugas", ["Nama", "Target", "Realisasi", "Dicacah", "Submit", "Diolah", "Draft", "Open", "Pengawas", "Terakhir Diupdate"]);
        shP.clearContents();
        shP.appendRow(["Nama", "Target", "Realisasi", "Dicacah", "Submit", "Diolah", "Draft", "Open", "Pengawas", "Terakhir Diupdate"]);
        var rP = [];
        for (var i3 = 0; i3 < payload.ub_petugas.length; i3++) {
          var z = payload.ub_petugas[i3];
          rP.push([z.nama, z.target, z.realisasi, z.dicacah||0, z.submit||0, z.diolah||0, z.draft||0, z.open||0, z.pml||"-", now]);
        }
        if (rP.length > 0) shP.getRange(2, 1, rP.length, 10).setValues(rP);
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Aksi tidak dikenali" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function getOrCreateSheet(ss, sheetName, headers) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (headers && headers.length > 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e2e8f0");
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}
