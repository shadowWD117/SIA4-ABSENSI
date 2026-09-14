// =============================================
// GOOGLE APPS SCRIPT — BACKEND ABSENSI v2
// =============================================

var SHEET_ABSEN = "Absensi";
var SHEET_SESI  = "Sesi Aktif";

// ── Utilitas ──────────────────────────────────
function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers) {
      sheet.appendRow(headers);
      var hr = sheet.getRange(1, 1, 1, headers.length);
      hr.setFontWeight("bold");
      hr.setBackground("#1a1a2e");
      hr.setFontColor("#ffffff");
    }
  }
  return sheet;
}

// ── POST: simpan absensi ──────────────────────
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var sheet = getOrCreateSheet(ss, SHEET_ABSEN, [
      "Waktu","Nama","NIM/NIS","Kelas","Sesi","Status","Keterangan","Latitude","Longitude"
    ]);

    // Cek duplikat NIM di sesi yang sama hari ini
    var rows = sheet.getDataRange().getValues();
    var sekarang = new Date();
    var hariIni = sekarang.getFullYear() + '-'
      + String(sekarang.getMonth() + 1).padStart(2, '0') + '-'
      + String(sekarang.getDate()).padStart(2, '0');
    var duplikat = rows.slice(1).some(function(row) {
      var tglRow = new Date(row[0]);
      var tglRowStr = tglRow.getFullYear() + '-'
        + String(tglRow.getMonth() + 1).padStart(2, '0') + '-'
        + String(tglRow.getDate()).padStart(2, '0');
      return String(row[2]) === String(payload.nim)
          && row[4] === payload.sesi
          && tglRowStr === hariIni;
    });

    if (duplikat) {
      return jsonOut({ status: "error", message: "NIM ini sudah absen di sesi yang sama hari ini." });
    }

    sheet.appendRow([
      payload.waktu,
      payload.nama,
      payload.nim,
      payload.kelas,
      payload.sesi,
      payload.status,
      payload.keterangan || "",
      payload.lat || "",
      payload.lng || ""
    ]);
    sheet.autoResizeColumns(1, 9);

    return jsonOut({ status: "ok", message: "Absensi berhasil dicatat." });

  } catch(err) {
    return jsonOut({ status: "error", message: err.toString() });
  }
}

// ── GET: berbagai action ──────────────────────
function doGet(e) {
  try {
    var action = e.parameter.action;
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Ambil semua data absensi
    if (action === "getAll") {
      var sheet = getOrCreateSheet(ss, SHEET_ABSEN, [
        "Waktu","Nama","NIM/NIS","Kelas","Sesi","Status","Keterangan","Latitude","Longitude"
      ]);
      var rows = sheet.getDataRange().getValues();
      var totalRows = rows.length - 1; // tidak hitung header
      if (totalRows <= 0) return jsonOut({ status: "ok", records: [], total: 0 });

      var since = parseInt(e.parameter.since) || 0;
      var dataRows = rows.slice(1); // buang header

      // Jika client sudah punya sebagian data, kirim sisanya saja
      var newRows = since > 0 && since < dataRows.length
        ? dataRows.slice(since)
        : dataRows;

      var records = newRows.map(function(row) {
        return {
          waktu: row[0], nama: row[1], nim: row[2],
          kelas: row[3], sesi: row[4], status: row[5],
          keterangan: row[6], lat: row[7], lng: row[8]
        };
      }).reverse();

      return jsonOut({ status: "ok", records: records, total: totalRows });
    }

    // Ambil sesi aktif (untuk halaman absensi siswa)
    if (action === "getSesi") {
      var sesiSheet = getOrCreateSheet(ss, SHEET_SESI, [
        "Kelas","Sesi","Jam Mulai","Menit Mulai","Jam Selesai","Menit Selesai",
        "Lat Pusat","Lng Pusat","Radius (m)","Status","Tanggal"
      ]);
      var sesiRows = sesiSheet.getDataRange().getValues();
      // Cari baris dengan Status = "Aktif"
      var aktif = null;
      for (var i = 1; i < sesiRows.length; i++) {
        if (sesiRows[i][9] === "Aktif") {
          aktif = {
            kelas:       sesiRows[i][0],
            sesi:        sesiRows[i][1],
            jamMulai:    sesiRows[i][2],
            menitMulai:  sesiRows[i][3],
            jamSelesai:  sesiRows[i][4],
            menitSelesai:sesiRows[i][5],
            latPusat:    sesiRows[i][6],
            lngPusat:    sesiRows[i][7],
            radius:      sesiRows[i][8],
            tanggal:     sesiRows[i][10]
              ? Utilities.formatDate(new Date(sesiRows[i][10]), Session.getScriptTimeZone(), "yyyy-MM-dd")
              : ""
          };
          break;
        }
      }
      return jsonOut({ status: "ok", sesi: aktif });
    }

    // Simpan / update sesi aktif dari admin
    if (action === "setSesi") {
      var sesiSheet = getOrCreateSheet(ss, SHEET_SESI, [
        "Kelas","Sesi","Jam Mulai","Menit Mulai","Jam Selesai","Menit Selesai",
        "Lat Pusat","Lng Pusat","Radius (m)","Status","Tanggal"
      ]);

      // Nonaktifkan semua sesi dulu
      var lastRow = sesiSheet.getLastRow();
      if (lastRow > 1) {
        sesiSheet.getRange(2, 10, lastRow - 1, 1).setValue("Nonaktif");
      }

      // Jika flag nonaktif=1, hanya matikan semua sesi tanpa tambah baru
      if (e.parameter.nonaktif === '1') {
        return jsonOut({ status: "ok", message: "Sesi dinonaktifkan." });
      }
      // Tambah sesi baru sebagai Aktif
      sesiSheet.appendRow([
        e.parameter.kelas,
        e.parameter.sesi,
        parseInt(e.parameter.jamMulai),
        parseInt(e.parameter.menitMulai),
        parseInt(e.parameter.jamSelesai),
        parseInt(e.parameter.menitSelesai),
        parseFloat(e.parameter.latPusat),
        parseFloat(e.parameter.lngPusat),
        parseInt(e.parameter.radius),
        "Aktif",
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd")
      ]);
      sesiSheet.autoResizeColumns(1, 11);

      return jsonOut({ status: "ok", message: "Sesi berhasil diperbarui." });
    }

    return jsonOut({ status: "error", message: "Action tidak dikenal." });

  } catch(err) {
    return jsonOut({ status: "error", message: err.toString() });
  }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
