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

    // Validasi sesi aktif di server
    var sesiSheet = getOrCreateSheet(ss, SHEET_SESI, [
      "Kelas","Sesi","Jam Mulai","Menit Mulai","Jam Selesai","Menit Selesai",
      "Lat Pusat","Lng Pusat","Radius (m)","Status","Tanggal"
    ]);
    var sesiRows = sesiSheet.getDataRange().getValues();
    var sesiAktif = null;
    var now = new Date();
    var tglHariIni = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
    var nowMenit = now.getHours() * 60 + now.getMinutes();

    for (var s = 1; s < sesiRows.length; s++) {
      if (sesiRows[s][9] === "Aktif") {
        var tglSesi = sesiRows[s][10]
          ? Utilities.formatDate(new Date(sesiRows[s][10]), Session.getScriptTimeZone(), "yyyy-MM-dd")
          : "";
        var mulaiMenit   = parseInt(sesiRows[s][2]) * 60 + parseInt(sesiRows[s][3]);
        var selesaiMenit = parseInt(sesiRows[s][4]) * 60 + parseInt(sesiRows[s][5]);
        if (tglSesi === tglHariIni && nowMenit >= mulaiMenit && nowMenit <= selesaiMenit) {
          sesiAktif = sesiRows[s];
        }
        break;
      }
    }

    if (!sesiAktif) {
      return jsonOut({ status: "error", message: "Sesi absen sudah berakhir atau tidak aktif. Hubungi admin." });
    }

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

    // Bersihkan duplikat dari Sheets
    if (action === "bersihkanDuplikat") {
      var sheet = getOrCreateSheet(ss, SHEET_ABSEN, [
        "Waktu","Nama","NIM/NIS","Kelas","Sesi","Status","Keterangan","Latitude","Longitude"
      ]);
      var rows = sheet.getDataRange().getValues();
      if (rows.length <= 1) return jsonOut({ status: "ok", dihapus: 0 });

      var seen = {};
      var toDelete = [];
      for (var i = 1; i < rows.length; i++) {
        var tgl = "";
        try {
          tgl = Utilities.formatDate(new Date(rows[i][0]), Session.getScriptTimeZone(), "yyyy-MM-dd");
        } catch(ex) { tgl = String(rows[i][0]).split(' ')[0]; }
        var key = String(rows[i][2]) + '|' + String(rows[i][4]) + '|' + tgl;
        if (seen[key]) {
          toDelete.push(i + 1); // +1 karena index sheet mulai dari 1
        } else {
          seen[key] = true;
        }
      }

      // Hapus dari bawah ke atas agar index tidak bergeser
      for (var d = toDelete.length - 1; d >= 0; d--) {
        sheet.deleteRow(toDelete[d]);
      }

      return jsonOut({ status: "ok", dihapus: toDelete.length });
    }

    // Cek sesi aktif real-time (untuk validasi dari halaman absensi)
    if (action === "cekSesi") {
      var sesiSheet = getOrCreateSheet(ss, SHEET_SESI, [
        "Kelas","Sesi","Jam Mulai","Menit Mulai","Jam Selesai","Menit Selesai",
        "Lat Pusat","Lng Pusat","Radius (m)","Status","Tanggal"
      ]);
      var sesiRows = sesiSheet.getDataRange().getValues();
      var now = new Date();
      var tglHariIni = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
      var nowMenit = now.getHours() * 60 + now.getMinutes();
      var aktif = false;

      for (var i = 1; i < sesiRows.length; i++) {
        if (sesiRows[i][9] === "Aktif") {
          var tglSesi = sesiRows[i][10]
            ? Utilities.formatDate(new Date(sesiRows[i][10]), Session.getScriptTimeZone(), "yyyy-MM-dd")
            : "";
          var mulai   = parseInt(sesiRows[i][2]) * 60 + parseInt(sesiRows[i][3]);
          var selesai = parseInt(sesiRows[i][4]) * 60 + parseInt(sesiRows[i][5]);
          if (tglSesi === tglHariIni && nowMenit >= mulai && nowMenit <= selesai) {
            aktif = true;
          }
          break;
        }
      }
      return jsonOut({ status: "ok", aktif: aktif });
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
