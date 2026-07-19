// ============================================================
// SISTEM INFORMASI PEGAWAI PUSKESMAS — Code.gs
// Pure ES5 — aman di semua versi Google Apps Script
// Versi 2.0 — Perbaikan + Fitur Absensi GPS
// ============================================================

// ─── KONFIGURASI GPS KANTOR ─────────────────────────────────
var GPS_CONFIG = {
  lat:    -5.152878096461292,
  lng:    119.40723207043266,
  radius: 105   // meter — setengah dari estimasi lebar area ~19.593 m²
                // √(19593/π) ≈ 79 m, pakai 105 m agar cukup longgar
};

// ─── KONFIGURASI SESI ────────────────────────────────────────
var SESI_EXPIRY_JAM = 10; // token otomatis expire setelah N jam

// ─── ENTRY POINT ─────────────────────────────────────────────
function doGet(e) {
  return HtmlService.createTemplateFromFile('UI')
    .evaluate()
    .setTitle('SIPP')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function handleRequest(action, data) {
  // PENTING: google.script.run hanya bisa terima/kirim string
  // Selalu return JSON.stringify — jangan pernah return object langsung
  var result;
  try {
    if (!action) {
      result = { ok: false, msg: 'Action kosong.' };
    } else if (action === 'login') {
      result = handleLogin(data);
    } else if (!data || !data._token) {
      result = { ok: false, msg: 'Token tidak ada.' };
    } else {
      var sesi = getSesi(data._token);
      if (!sesi) {
        result = { ok: false, msg: 'Sesi tidak valid atau sudah expired. Silakan login ulang.' };
      } else if (action === 'logout')              { result = handleLogout(data._token); }
      else if (action === 'getPegawai')            { result = handleGetPegawai(sesi); }
      else if (action === 'savePegawai')           { result = handleSavePegawai(data, sesi); }
      else if (action === 'nonaktifkanPegawai')    { result = handleNonaktifkanPegawai(data, sesi); }
      else if (action === 'hapusPegawai')          { result = handleHapusPegawai(data, sesi); }
      else if (action === 'checkIn')               { result = handleCheckIn(sesi, data); }
      else if (action === 'checkOut')              { result = handleCheckOut(sesi, data); }
      else if (action === 'getKehadiran')          { result = handleGetKehadiran(data, sesi); }
      else if (action === 'saveJurnal')            { result = handleSaveJurnal(data, sesi); }
      else if (action === 'getJurnal')             { result = handleGetJurnal(data, sesi); }
      else if (action === 'ajukanIzin')            { result = handleAjukanIzin(data, sesi); }
      else if (action === 'getIzin')               { result = handleGetIzin(data, sesi); }
      else if (action === 'approveIzin')           { result = handleApproveIzin(data, sesi); }
      else if (action === 'getDashboard')          { result = handleGetDashboard(sesi); }
      else if (action === 'exportCsv')             { result = handleExportCsv(data, sesi); }
      else if (action === 'getGpsConfig')          { result = handleGetGpsConfig(sesi); }
      else { result = { ok: false, msg: 'Action tidak dikenal: ' + action }; }
    }
  } catch (err) {
    result = { ok: false, msg: 'Server error: ' + err.message };
  }

  try {
    var str = JSON.stringify(result);
    return str || '{"ok":false,"msg":"JSON.stringify gagal"}';
  } catch(e) {
    return '{"ok":false,"msg":"Serialisasi gagal: ' + e.message + '"}';
  }
}

// ─── GPS CONFIG ──────────────────────────────────────────────
function handleGetGpsConfig(sesi) {
  // Hanya kembalikan radius & koordinat, bukan data sensitif
  return {
    ok:     true,
    lat:    GPS_CONFIG.lat,
    lng:    GPS_CONFIG.lng,
    radius: GPS_CONFIG.radius
  };
}

// Hitung jarak Haversine antara dua titik (meter)
function hitungJarak(lat1, lng1, lat2, lng2) {
  var R  = 6371000; // radius bumi dalam meter
  var dL = (lat2 - lat1) * Math.PI / 180;
  var dG = (lng2 - lng1) * Math.PI / 180;
  var a  = Math.sin(dL/2) * Math.sin(dL/2) +
           Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
           Math.sin(dG/2) * Math.sin(dG/2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

// ─── SETUP ───────────────────────────────────────────────────
function setupSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var defs = [
    ['Pegawai',     ['ID','NIP','Nama','Jabatan','UnitKerja','Email','Telepon','Peran','PasswordHash','Aktif']],
    ['Kehadiran',   ['ID','PegawaiID','Tanggal','JamMasuk','JamKeluar','Status','LatMasuk','LngMasuk','JarakMasuk','MetodeMasuk']],
    ['JurnalHarian',['ID','PegawaiID','Tanggal','Kegiatan','Hasil','Kendala']],
    ['IzinCuti',    ['ID','PegawaiID','Jenis','TanggalMulai','TanggalSelesai','Alasan','LampiranUrl','Status','ApprovedBy','CatatanApproval']],
    ['Sesi',        ['Token','PegawaiID','Nama','Peran','LoginAt']]
  ];

  for (var i = 0; i < defs.length; i++) {
    var name = defs[i][0];
    var cols = defs[i][1];
    var sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.appendRow(cols);
    // Perbaiki header jika kolom baru (Kehadiran v2)
    if (name === 'Kehadiran' && sh.getLastRow() >= 1) {
      var existingHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      if (existingHeaders.length < cols.length) {
        // Tambahkan header kolom baru yang belum ada
        for (var c = existingHeaders.length; c < cols.length; c++) {
          sh.getRange(1, c + 1).setValue(cols[c]);
        }
      }
    }
  }

  var pSh   = ss.getSheetByName('Pegawai');
  var isNew = pSh.getLastRow() <= 1;
  if (isNew) {
    pSh.appendRow([1,'000','Administrator','Admin','Semua',
      'admin@puskesmas.id','','admin', hashPassword('admin123'), true]);
  }

  var msg = isNew
    ? 'Setup selesai!\nEmail: admin@puskesmas.id\nPassword: admin123\n\n⚠️ Segera ganti password default!'
    : 'Setup selesai! Data existing tidak diubah.';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch(e) { /* dari editor, lihat Execution log */ }
}

// ─── SHEET HELPERS ───────────────────────────────────────────
function getSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function getSheetData(name) {
  var sh = getSheet(name);
  if (sh.getLastRow() <= 1) return [];
  var vals    = sh.getDataRange().getValues();
  var headers = vals[0];
  var tz      = Session.getScriptTimeZone();
  var dateColumns = { Tanggal:true, TanggalMulai:true, TanggalSelesai:true, LoginAt:false };
  var result  = [];
  for (var i = 1; i < vals.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var key = String(headers[j]);
      var val = vals[i][j];
      if (val instanceof Date && dateColumns[key] === true) {
        val = Utilities.formatDate(val, tz, 'yyyy-MM-dd');
      }
      obj[key] = val;
    }
    result.push(obj);
  }
  return result;
}

function getNextId(name) {
  var sh   = getSheet(name);
  var last = sh.getLastRow();
  if (last <= 1) return 1;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  var max = 0;
  for (var i = 0; i < ids.length; i++) {
    var n = Number(ids[i][0]);
    if (n > max) max = n;
  }
  return max + 1;
}

// Cari index kolom berdasarkan nama header (aman terhadap perubahan urutan kolom)
function getColIndex(sh, headerName) {
  if (sh.getLastRow() < 1) return -1;
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]) === headerName) return i + 1; // 1-based
  }
  return -1;
}

// ─── AUTH ────────────────────────────────────────────────────
function handleLogin(data) {
  if (!data) return { ok: false, msg: 'Data login kosong.' };
  var email    = String(data.email    || '').trim().toLowerCase();
  var password = String(data.password || '');
  if (!email)    return { ok: false, msg: 'Email wajib diisi.' };
  if (!password) return { ok: false, msg: 'Password wajib diisi.' };

  var rows = getSheetData('Pegawai');
  if (rows.length === 0) {
    return { ok: false, msg: 'Belum ada data pegawai. Jalankan Setup Spreadsheet terlebih dahulu.' };
  }

  var pegawai = null;
  for (var i = 0; i < rows.length; i++) {
    var r         = rows[i];
    var emailSheet = String(r.Email || '').trim().toLowerCase();
    var aktif      = (r.Aktif === true || r.Aktif === 'TRUE' || r.Aktif === 1 || r.Aktif === 'true');
    if (emailSheet === email && aktif) { pegawai = r; break; }
  }

  if (!pegawai) return { ok: false, msg: 'Email tidak ditemukan atau akun nonaktif.' };

  var hashDiSheet = String(pegawai.PasswordHash || '');
  if (!verifyPassword(password, hashDiSheet)) {
    return { ok: false, msg: 'Password salah.' };
  }

  // Hapus sesi lama milik pegawai ini sebelum buat baru (single session)
  bersihkanSesiLamaPegawai(String(pegawai.ID));

  var token = Utilities.getUuid();
  getSheet('Sesi').appendRow([token, pegawai.ID, pegawai.Nama, pegawai.Peran, new Date().toISOString()]);

  return {
    ok:    true,
    token: token,
    nama:  String(pegawai.Nama),
    peran: String(pegawai.Peran),
    id:    String(pegawai.ID)
  };
}

function handleLogout(token) {
  var sh   = getSheet('Sesi');
  var vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (vals[i][0] === token) { sh.deleteRow(i + 1); break; }
  }
  return { ok: true };
}

function getSesi(token) {
  if (!token) return null;
  var sh = getSheet('Sesi');
  if (sh.getLastRow() <= 1) return null;
  var vals = sh.getDataRange().getValues();
  var now  = new Date();
  for (var i = 1; i < vals.length; i++) {
    if (vals[i][0] === token) {
      // Cek expiry sesi
      var loginAt = new Date(vals[i][4]);
      var selisihJam = (now - loginAt) / 1000 / 3600;
      if (selisihJam > SESI_EXPIRY_JAM) {
        sh.deleteRow(i + 1); // Hapus token expired
        return null;
      }
      return { token: token, id: String(vals[i][1]), nama: String(vals[i][2]), peran: String(vals[i][3]) };
    }
  }
  return null;
}

// Hapus semua sesi lama milik satu pegawai (untuk single-session)
function bersihkanSesiLamaPegawai(pegawaiId) {
  var sh = getSheet('Sesi');
  if (sh.getLastRow() <= 1) return;
  var vals = sh.getDataRange().getValues();
  // Hapus dari bawah ke atas agar row index tidak bergeser
  for (var i = vals.length - 1; i >= 1; i--) {
    if (String(vals[i][1]) === pegawaiId) {
      sh.deleteRow(i + 1);
    }
  }
}

// Bersihkan semua token expired (bisa dijadwalkan lewat Triggers)
function bersihkanSesiExpired() {
  var sh = getSheet('Sesi');
  if (sh.getLastRow() <= 1) return;
  var vals = sh.getDataRange().getValues();
  var now  = new Date();
  for (var i = vals.length - 1; i >= 1; i--) {
    var loginAt    = new Date(vals[i][4]);
    var selisihJam = (now - loginAt) / 1000 / 3600;
    if (selisihJam > SESI_EXPIRY_JAM) sh.deleteRow(i + 1);
  }
}

// ─── PEGAWAI ─────────────────────────────────────────────────
function handleGetPegawai(sesi) {
  if (sesi.peran !== 'admin' && sesi.peran !== 'atasan') return { ok: false, msg: 'Akses ditolak.' };
  var rows = getSheetData('Pegawai');
  for (var i = 0; i < rows.length; i++) rows[i].PasswordHash = '***';
  return { ok: true, data: rows };
}

function handleSavePegawai(data, sesi) {
  if (sesi.peran !== 'admin') return { ok: false, msg: 'Hanya admin.' };
  if (!data.Nama)  return { ok: false, msg: 'Nama wajib diisi.' };
  if (!data.Email) return { ok: false, msg: 'Email wajib diisi.' };

  var sh   = getSheet('Pegawai');
  var vals = sh.getDataRange().getValues();

  if (data.ID) {
    // Update
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][0]) === String(data.ID)) {
        var newHash = data.Password ? hashPassword(data.Password) : vals[i][8];
        sh.getRange(i+1,1,1,10).setValues([[
          data.ID, data.NIP||'', data.Nama, data.Jabatan||'',
          data.UnitKerja||'', data.Email, data.Telepon||'', data.Peran,
          newHash,
          data.Aktif !== undefined ? data.Aktif : vals[i][9]
        ]]);
        return { ok: true, msg: 'Data pegawai diperbarui.' };
      }
    }
    return { ok: false, msg: 'Pegawai tidak ditemukan.' };
  }

  // Cek duplikat email
  for (var j = 1; j < vals.length; j++) {
    if (String(vals[j][5]).trim().toLowerCase() === String(data.Email).trim().toLowerCase()) {
      return { ok: false, msg: 'Email sudah digunakan pegawai lain.' };
    }
  }

  // Insert baru — password default: puskesmas123
  var defaultPass = 'puskesmas123';
  sh.appendRow([getNextId('Pegawai'), data.NIP||'', data.Nama, data.Jabatan||'',
    data.UnitKerja||'', data.Email, data.Telepon||'', data.Peran,
    hashPassword(data.Password || defaultPass), true]);
  var passMsg = data.Password
    ? 'Pegawai ditambahkan. Password sesuai yang dimasukkan.'
    : 'Pegawai ditambahkan. Password default: ' + defaultPass + ' (segera ganti!)';
  return { ok: true, msg: passMsg };
}

function handleNonaktifkanPegawai(data, sesi) {
  if (sesi.peran !== 'admin') return { ok: false, msg: 'Hanya admin.' };
  var sh    = getSheet('Pegawai');
  var col   = getColIndex(sh, 'Aktif');
  var vals  = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(data.id)) {
      sh.getRange(i+1, col).setValue(false);
      // Paksa logout sesi aktif pegawai ini
      bersihkanSesiLamaPegawai(String(data.id));
      return { ok: true, msg: 'Pegawai dinonaktifkan dan sesi dihapus.' };
    }
  }
  return { ok: false, msg: 'Tidak ditemukan.' };
}

function handleHapusPegawai(data, sesi) {
  if (sesi.peran !== 'admin') return { ok: false, msg: 'Hanya admin yang bisa menghapus.' };
  if (String(data.id) === String(sesi.id)) return { ok: false, msg: 'Tidak bisa menghapus akun sendiri.' };

  var shPeg  = getSheet('Pegawai');
  var vals   = shPeg.getDataRange().getValues();
  var rowIdx = -1;
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(data.id)) { rowIdx = i + 1; break; }
  }
  if (rowIdx === -1) return { ok: false, msg: 'Pegawai tidak ditemukan.' };

  shPeg.deleteRow(rowIdx);

  // Cascade delete data relasional
  var sheets = ['Kehadiran','JurnalHarian','IzinCuti'];
  for (var s = 0; s < sheets.length; s++) {
    var shX = getSheet(sheets[s]);
    if (shX.getLastRow() > 1) {
      var xVals = shX.getDataRange().getValues();
      for (var k = xVals.length - 1; k >= 1; k--) {
        if (String(xVals[k][1]) === String(data.id)) shX.deleteRow(k + 1);
      }
    }
  }

  // Hapus sesi aktif
  bersihkanSesiLamaPegawai(String(data.id));

  return { ok: true, msg: 'Pegawai dan seluruh datanya berhasil dihapus.' };
}

// ─── KEHADIRAN (dengan GPS) ───────────────────────────────────
function handleCheckIn(sesi, data) {
  var today = todayStr();
  var rows  = getSheetData('Kehadiran');

  // Cek sudah check-in hari ini
  for (var i = 0; i < rows.length; i++) {
    var tgl = rows[i].Tanggal;
    if (tgl instanceof Date) {
      tgl = Utilities.formatDate(tgl, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else {
      tgl = String(tgl).substring(0, 10);
    }
    if (String(rows[i].PegawaiID) === sesi.id && tgl === today) {
      return { ok: false, msg: 'Sudah check-in hari ini pukul ' + rows[i].JamMasuk };
    }
  }

  // Validasi GPS
  var latMasuk  = data.lat  ? parseFloat(data.lat)  : null;
  var lngMasuk  = data.lng  ? parseFloat(data.lng)  : null;
  var metode    = 'manual';
  var jarak     = null;
  var gpsValid  = false;

  if (latMasuk !== null && lngMasuk !== null && !isNaN(latMasuk) && !isNaN(lngMasuk)) {
    jarak    = hitungJarak(latMasuk, lngMasuk, GPS_CONFIG.lat, GPS_CONFIG.lng);
    metode   = 'gps';
    gpsValid = jarak <= GPS_CONFIG.radius;
    if (!gpsValid) {
      return {
        ok:    false,
        msg:   'Lokasi Anda di luar area kantor. Jarak: ' + Math.round(jarak) + ' m (maks. ' + GPS_CONFIG.radius + ' m).',
        jarak: Math.round(jarak)
      };
    }
  } else {
    // Tidak ada koordinat — tolak jika GPS wajib
    // Uncomment baris berikut jika GPS wajib mutlak:
    // return { ok: false, msg: 'Lokasi GPS diperlukan untuk absensi.' };
    metode = 'manual';
  }

  var now    = new Date();
  var jam    = timeStr(now);
  var status = now.getHours() >= 9 ? 'terlambat' : 'hadir';

  getSheet('Kehadiran').appendRow([
    getNextId('Kehadiran'), sesi.id, today, jam, '', status,
    latMasuk || '', lngMasuk || '',
    jarak !== null ? Math.round(jarak) : '',
    metode
  ]);

  return {
    ok:      true,
    jamMasuk: jam,
    status:  status,
    jarak:   jarak !== null ? Math.round(jarak) : null,
    metode:  metode,
    msg:     'Check-in berhasil pukul ' + jam + (jarak !== null ? ' (GPS ✓ ' + Math.round(jarak) + ' m dari kantor)' : '')
  };
}

function handleCheckOut(sesi, data) {
  var today = todayStr();
  var sh    = getSheet('Kehadiran');
  if (sh.getLastRow() <= 1) return { ok: false, msg: 'Belum ada data kehadiran.' };

  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var vals    = sh.getDataRange().getValues();

  // Cari index kolom secara dinamis
  var iJamKeluar = headers.indexOf('JamKeluar')  + 1; // 1-based
  var iStatus    = headers.indexOf('Status')      + 1;
  var iLatM      = headers.indexOf('LatMasuk')    + 1;
  var iLngM      = headers.indexOf('LngMasuk')    + 1;
  var iJarak     = headers.indexOf('JarakMasuk')  + 1;
  var iMetode    = headers.indexOf('MetodeMasuk') + 1;

  for (var i = 1; i < vals.length; i++) {
    var rowPegId   = String(vals[i][1]);
    var rowTanggal = vals[i][2];
    if (rowTanggal instanceof Date) {
      rowTanggal = Utilities.formatDate(rowTanggal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else {
      rowTanggal = String(rowTanggal).substring(0, 10);
    }

    if (rowPegId === sesi.id && rowTanggal === today) {
      if (vals[i][iJamKeluar - 1]) return { ok: false, msg: 'Sudah check-out pukul ' + vals[i][iJamKeluar - 1] };

      // Validasi GPS untuk check-out
      var latKeluar = data.lat  ? parseFloat(data.lat)  : null;
      var lngKeluar = data.lng  ? parseFloat(data.lng)  : null;
      var jarak     = null;

      if (latKeluar !== null && lngKeluar !== null && !isNaN(latKeluar) && !isNaN(lngKeluar)) {
        jarak = hitungJarak(latKeluar, lngKeluar, GPS_CONFIG.lat, GPS_CONFIG.lng);
        if (jarak > GPS_CONFIG.radius) {
          return {
            ok:    false,
            msg:   'Lokasi Anda di luar area kantor untuk check-out. Jarak: ' + Math.round(jarak) + ' m.',
            jarak: Math.round(jarak)
          };
        }
      }

      var now    = new Date();
      var jam    = timeStr(now);
      var status = now.getHours() < 16 ? 'pulang_cepat' : (vals[i][iStatus - 1] || 'hadir');

      sh.getRange(i+1, iJamKeluar).setValue(jam);
      sh.getRange(i+1, iStatus).setValue(status);

      return {
        ok:       true,
        jamKeluar: jam,
        jarak:    jarak !== null ? Math.round(jarak) : null,
        msg:      'Check-out berhasil pukul ' + jam + (jarak !== null ? ' (GPS ✓ ' + Math.round(jarak) + ' m dari kantor)' : '')
      };
    }
  }
  return { ok: false, msg: 'Belum check-in hari ini.' };
}

function handleGetKehadiran(data, sesi) {
  var rows = getSheetData('Kehadiran');
  var out  = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (sesi.peran === 'pegawai' && String(r.PegawaiID) !== sesi.id) continue;
    if (data.pegawaiId && String(r.PegawaiID) !== String(data.pegawaiId)) continue;
    if (data.dari   && r.Tanggal < data.dari)   continue;
    if (data.sampai && r.Tanggal > data.sampai) continue;
    out.push(r);
  }
  return { ok: true, data: out };
}

// ─── JURNAL ──────────────────────────────────────────────────
function handleSaveJurnal(data, sesi) {
  if (!data.kegiatan || !data.hasil) return { ok: false, msg: 'Kegiatan dan Hasil wajib diisi.' };
  var today = todayStr();
  var sh    = getSheet('JurnalHarian');
  var vals  = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][1]) === sesi.id && vals[i][2] === today) {
      sh.getRange(i+1, 4, 1, 3).setValues([[data.kegiatan, data.hasil, data.kendala||'']]);
      return { ok: true, msg: 'Jurnal hari ini diperbarui.' };
    }
  }
  sh.appendRow([getNextId('JurnalHarian'), sesi.id, today, data.kegiatan, data.hasil, data.kendala||'']);
  return { ok: true, msg: 'Jurnal berhasil disimpan.' };
}

function handleGetJurnal(data, sesi) {
  var rows = getSheetData('JurnalHarian');
  var out  = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (sesi.peran === 'pegawai' && String(r.PegawaiID) !== sesi.id) continue;
    if (data.pegawaiId && String(r.PegawaiID) !== String(data.pegawaiId)) continue;
    if (data.dari   && r.Tanggal < data.dari)   continue;
    if (data.sampai && r.Tanggal > data.sampai) continue;
    out.push(r);
  }
  return { ok: true, data: out };
}

// ─── IZIN/CUTI ───────────────────────────────────────────────
function handleAjukanIzin(data, sesi) {
  if (!data.tanggalMulai || !data.tanggalSelesai || !data.alasan)
    return { ok: false, msg: 'Semua field wajib diisi.' };
  if (data.tanggalMulai > data.tanggalSelesai)
    return { ok: false, msg: 'Tanggal mulai tidak boleh lebih dari tanggal selesai.' };

  getSheet('IzinCuti').appendRow([
    getNextId('IzinCuti'), sesi.id, data.jenis||'izin',
    data.tanggalMulai, data.tanggalSelesai, data.alasan, '', 'menunggu', '', ''
  ]);
  return { ok: true, msg: 'Pengajuan berhasil dikirim.' };
}

function handleGetIzin(data, sesi) {
  var rows = getSheetData('IzinCuti');
  var pMap = getPegawaiMap();
  var out  = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (sesi.peran === 'pegawai' && String(r.PegawaiID) !== sesi.id) continue;
    if (data.pegawaiId && String(r.PegawaiID) !== String(data.pegawaiId)) continue;
    if (data.status && r.Status !== data.status) continue;
    r.NamaPegawai = pMap[String(r.PegawaiID)] || 'N/A';
    out.push(r);
  }
  return { ok: true, data: out };
}

function handleApproveIzin(data, sesi) {
  if (sesi.peran !== 'atasan' && sesi.peran !== 'admin') return { ok: false, msg: 'Akses ditolak.' };
  var sh   = getSheet('IzinCuti');
  var vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(data.id)) {
      sh.getRange(i+1, 8).setValue(data.aksi === 'setuju' ? 'disetujui' : 'ditolak');
      sh.getRange(i+1, 9).setValue(sesi.id);
      sh.getRange(i+1,10).setValue(data.catatan||'');
      return { ok: true, msg: 'Pengajuan ' + (data.aksi === 'setuju' ? 'disetujui' : 'ditolak') + '.' };
    }
  }
  return { ok: false, msg: 'Data tidak ditemukan.' };
}

// ─── DASHBOARD ───────────────────────────────────────────────
function handleGetDashboard(sesi) {
  var today    = todayStr();
  var bulan    = today.substring(0, 7);
  var kehadiran = getSheetData('Kehadiran');
  var izin      = getSheetData('IzinCuti');
  var jurnal    = getSheetData('JurnalHarian');

  if (sesi.peran === 'pegawai') {
    var absen = null;
    for (var i = 0; i < kehadiran.length; i++) {
      if (String(kehadiran[i].PegawaiID) === sesi.id && kehadiran[i].Tanggal === today) {
        absen = kehadiran[i]; break;
      }
    }
    var izinPend = 0, jurBulan = 0;
    for (var j = 0; j < izin.length; j++) {
      if (String(izin[j].PegawaiID) === sesi.id && izin[j].Status === 'menunggu') izinPend++;
    }
    for (var k = 0; k < jurnal.length; k++) {
      if (String(jurnal[k].PegawaiID) === sesi.id && String(jurnal[k].Tanggal).substring(0,7) === bulan) jurBulan++;
    }
    return { ok: true, data: {
      absenHariIni:  absen,
      sudahCheckIn:  !!absen,
      sudahCheckOut: !!(absen && absen.JamKeluar),
      izinPending:   izinPend,
      jurnalBulan:   jurBulan,
      metodeMasuk:   absen ? (absen.MetodeMasuk || '') : '',
      jarakMasuk:    absen ? (absen.JarakMasuk  || '') : ''
    }};
  }

  var totalPeg = 0, hadirHariIni = 0, izinPend2 = 0, jurHariIni = 0;
  var rekapBulan = {};
  var pegawaiRows = getSheetData('Pegawai');
  for (var a = 0; a < pegawaiRows.length; a++) {
    var ak = pegawaiRows[a].Aktif;
    if (ak === true || ak === 'TRUE' || ak === 1 || ak === 'true') totalPeg++;
  }
  for (var b = 0; b < kehadiran.length; b++) {
    if (kehadiran[b].Tanggal === today) hadirHariIni++;
    if (String(kehadiran[b].Tanggal).substring(0,7) === bulan) {
      var st = kehadiran[b].Status;
      rekapBulan[st] = (rekapBulan[st] || 0) + 1;
    }
  }
  for (var c = 0; c < izin.length; c++) {
    if (izin[c].Status === 'menunggu') izinPend2++;
  }
  for (var d = 0; d < jurnal.length; d++) {
    if (jurnal[d].Tanggal === today) jurHariIni++;
  }

  return { ok: true, data: {
    totalPegawai:  totalPeg,
    hadirHariIni:  hadirHariIni,
    izinPending:   izinPend2,
    jurnalHariIni: jurHariIni,
    rekapBulan:    rekapBulan
  }};
}

// ─── EXPORT CSV ──────────────────────────────────────────────
function handleExportCsv(data, sesi) {
  if (sesi.peran !== 'admin' && sesi.peran !== 'atasan') return { ok: false, msg: 'Akses ditolak.' };

  var rows, headers;
  var pMap = getPegawaiMap();

  if (data.tipe === 'kehadiran') {
    headers = ['ID','NamaPegawai','NIP','Tanggal','JamMasuk','JamKeluar','Status','JarakMasuk','MetodeMasuk'];
    rows    = getSheetData('Kehadiran');
    // Join nama pegawai
    rows = rows.map(function(r) {
      var info = getPegawaiById(String(r.PegawaiID));
      r.NamaPegawai = info ? info.Nama : 'N/A';
      r.NIP         = info ? info.NIP  : '';
      return r;
    });
  } else if (data.tipe === 'izin') {
    headers = ['ID','NamaPegawai','Jenis','TanggalMulai','TanggalSelesai','Alasan','Status','CatatanApproval'];
    rows    = getSheetData('IzinCuti');
    rows = rows.map(function(r) {
      r.NamaPegawai = pMap[String(r.PegawaiID)] || 'N/A';
      return r;
    });
  } else {
    return { ok: false, msg: 'Tipe tidak valid.' };
  }

  if (data.dari)   rows = rows.filter(function(r){ return String(r.Tanggal||r.TanggalMulai) >= data.dari; });
  if (data.sampai) rows = rows.filter(function(r){ return String(r.Tanggal||r.TanggalSelesai) <= data.sampai; });

  var lines = [headers.join(',')];
  for (var i = 0; i < rows.length; i++) {
    var cols = [];
    for (var j = 0; j < headers.length; j++) cols.push('"' + String(rows[i][headers[j]]||'').replace(/"/g,'""') + '"');
    lines.push(cols.join(','));
  }

  return { ok: true, csv: lines.join('\n'), filename: data.tipe + '_' + todayStr() + '.csv' };
}

// ─── HELPERS ─────────────────────────────────────────────────
function getPegawaiMap() {
  var map  = {};
  var rows = getSheetData('Pegawai');
  for (var i = 0; i < rows.length; i++) map[String(rows[i].ID)] = rows[i].Nama;
  return map;
}

function getPegawaiById(id) {
  var rows = getSheetData('Pegawai');
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].ID) === id) return rows[i];
  }
  return null;
}

function todayStr() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function timeStr(d) {
  return Utilities.formatDate(d || new Date(), Session.getScriptTimeZone(), 'HH:mm');
}

function hashPassword(plain) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    plain + 'sipp_salt_2024',
    Utilities.Charset.UTF_8
  );
  var hex = '';
  for (var i = 0; i < bytes.length; i++) {
    var b = ('0' + (bytes[i] & 0xFF).toString(16)).slice(-2);
    hex += b;
  }
  return hex;
}

function verifyPassword(plain, hash) {
  return hashPassword(plain) === hash;
}

// ─── MENU & DEBUG ────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi().createMenu('SIPP Admin')
    .addItem('1. Setup Spreadsheet', 'setupSpreadsheet')
    .addItem('2. Lihat URL Aplikasi', 'showAppUrl')
    .addItem('3. Test Login Admin',   'testLogin')
    .addItem('4. Debug Info',         'debugInfo')
    .addItem('5. Bersihkan Sesi Expired', 'bersihkanSesiExpired')
    .addToUi();
}

function showAppUrl() {
  try {
    var url = ScriptApp.getService().getUrl();
    Logger.log('URL: ' + url);
    try { SpreadsheetApp.getUi().alert('URL Aplikasi SIPP:\n' + url); } catch(e2) {}
  } catch(e) {
    Logger.log('Belum di-deploy sebagai Web App.');
  }
}

function testLogin() {
  var r = handleLogin({ email: 'admin@puskesmas.id', password: 'admin123' });
  Logger.log('=== testLogin ===');
  Logger.log('ok    : ' + r.ok);
  Logger.log('msg   : ' + r.msg);
  Logger.log('nama  : ' + r.nama);
  Logger.log('peran : ' + r.peran);
  var s = handleRequest('login', { email: 'admin@puskesmas.id', password: 'admin123', _token: null });
  var msg2 = r.ok
    ? 'LOGIN OK | Nama: ' + r.nama + ' | Peran: ' + r.peran
    : 'LOGIN GAGAL: ' + r.msg;
  Logger.log(msg2);
  try { SpreadsheetApp.getUi().alert(msg2); } catch(e) {}
}

function debugInfo() {
  var rows = getSheetData('Pegawai');
  var msg  = 'Jumlah baris Pegawai: ' + rows.length + '\n\n';
  if (rows.length > 0) {
    var r = rows[0];
    msg += 'Email    : [' + r.Email + ']\n';
    msg += 'Peran    : [' + r.Peran + ']\n';
    msg += 'Aktif    : [' + r.Aktif + '] (' + typeof r.Aktif + ')\n';
    msg += 'Hash len : ' + String(r.PasswordHash||'').length + ' char\n\n';
    msg += 'Hash admin123 sekarang:\n' + hashPassword('admin123') + '\n\n';
    msg += 'Cocok?   : ' + verifyPassword('admin123', String(r.PasswordHash||''));
  } else {
    msg += 'Sheet Pegawai KOSONG. Jalankan Setup terlebih dahulu.';
  }
  msg += '\n\n--- GPS Config ---\n';
  msg += 'Lat: ' + GPS_CONFIG.lat + '\n';
  msg += 'Lng: ' + GPS_CONFIG.lng + '\n';
  msg += 'Radius: ' + GPS_CONFIG.radius + ' m';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch(e) {}
}
