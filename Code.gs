/**
 * @NotOnlyCurrentDoc
 * SISTEM INFORMASI MANAJEMEN SANGGAR (SIMS)
 * Backend Script - Google Apps Script (Code.gs)
 * Aplikasi Web: Happy Drawing
 * Subjudul: "Belajar Bermain Berkarya Penuh Warna"
 */

const CONFIG = {
  SPREADSHEET_ID: "1vzBWF1Sbr9_0Ra-PnSZotuFPBG1wlgr_m1_bEAVSKA8",
  FOLDER_FOTO_ID: "1Bk0SspPGvrvZNSWur-Mqc0E_pTpC5j42",
  FOLDER_GALERI_ID: "1M23qJ-DO_rT2Yaz5zFpiwRgTyerEr2jM",
  FOLDER_BUKTI_ID: "1wKPIaaNWUYjMeQc8Nq9lCHlKda6VjWcW",
  LOGO_ID: "1NZ_5ZsZHVsZnV1cTJhXobElmyCYXzntV",
  APP_NAME: "Happy Drawing",
  VERSION: "2.2.0"
};

/**
 * Helper function to deeply sanitize data objects before returning them
 * over google.script.run to prevent mae_html_user.js RPC serialization errors (Uncaught Gr).
 */
function sanitizeForRPC(data) {
  if (data === undefined || data === null) return "";
  
  if (typeof data === 'number') {
    if (isNaN(data) || !isFinite(data)) return 0;
    return data;
  }
  
  if (typeof data === 'boolean') {
    return data;
  }
  
  const type = Object.prototype.toString.call(data);
  
  if (type === '[object Date]') {
    if (isNaN(data.getTime())) return "";
    const yyyy = data.getFullYear();
    const mm = String(data.getMonth() + 1).padStart(2, '0');
    const dd = String(data.getDate()).padStart(2, '0');
    const hours = String(data.getHours()).padStart(2, '0');
    const mins = String(data.getMinutes()).padStart(2, '0');
    const secs = String(data.getSeconds()).padStart(2, '0');
    if (hours === '00' && mins === '00' && secs === '00') {
      return dd + '/' + mm + '/' + yyyy;
    }
    return dd + '/' + mm + '/' + yyyy + ' ' + hours + ':' + mins + ':' + secs;
  }

  if (Array.isArray(data)) {
    return data.map(function(item) { return sanitizeForRPC(item); });
  }

  if (typeof data === 'object') {
    const cleanObj = {};
    for (let key in data) {
      if (Object.prototype.hasOwnProperty.call(data, key)) {
        const val = data[key];
        if (typeof val !== 'function' && typeof val !== 'symbol') {
          cleanObj[String(key || 'key')] = sanitizeForRPC(val);
        }
      }
    }
    return cleanObj;
  }

  return String(data);
}

/**
 * Memeriksa dan menginisialisasi seluruh Sheet database
 */
function setupDatabase() {
  try {
    try {
      const checkFolder = function(folderId, defaultName) {
        try {
          if (folderId) {
            DriveApp.getFolderById(folderId);
            return true;
          }
        } catch (e) {}
        const folders = DriveApp.getFoldersByName(defaultName);
        if (!folders.hasNext()) {
          DriveApp.createFolder(defaultName);
        }
        return false;
      };

      checkFolder(CONFIG.FOLDER_FOTO_ID, 'FOTO');
      checkFolder(CONFIG.FOLDER_GALERI_ID, 'GALERI');
      checkFolder(CONFIG.FOLDER_BUKTI_ID, 'BUKTI_TRANSFER');
    } catch (driveErr) {
      Logger.log("Peringatan Otorisasi Drive: " + driveErr.message);
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

    // 1. Sheet ROLE
    let roleSheet = ss.getSheetByName('ROLE');
    if (!roleSheet) {
      roleSheet = ss.insertSheet('ROLE');
      roleSheet.appendRow(['id_role', 'nama_role', 'permissions', 'created_at', 'updated_at', 'is_deleted']);
    }
    if (roleSheet.getLastRow() === 1) {
      roleSheet.appendRow(['ROL-001', 'SUPER_ADMIN', '{"all": true}', new Date().toISOString(), new Date().toISOString(), 'FALSE']);
      roleSheet.appendRow(['ROL-002', 'ADMIN', '{"dashboard": true, "master": true, "pendaftaran": true, "pembayaran": true, "absensi": true, "laporan": true}', new Date().toISOString(), new Date().toISOString(), 'FALSE']);
      roleSheet.appendRow(['ROL-003', 'SISWA', '{"portal_siswa": true}', new Date().toISOString(), new Date().toISOString(), 'FALSE']);
    }

    // 2. Sheet USER
    let userSheet = ss.getSheetByName('USER') || ss.getSheetByName('USERS');
    if (!userSheet) {
      userSheet = ss.insertSheet('USER');
      userSheet.appendRow(['id_user', 'email', 'password', 'id_role', 'nama', 'status', 'created_at', 'updated_at', 'is_deleted', 'permissions']);
    }
    if (userSheet.getLastRow() === 1) {
      userSheet.appendRow(['USR-001', 'superadmin@sanggarku.com', hashPasswordV2('admin123'), 'SUPER_ADMIN', 'Super Administrator', 'AKTIF', new Date().toISOString(), new Date().toISOString(), 'FALSE']);
      userSheet.appendRow(['USR-002', 'admin@sanggarku.com', hashPasswordV2('admin123'), 'ADMIN', 'Administrator Sanggar', 'AKTIF', new Date().toISOString(), new Date().toISOString(), 'FALSE']);
      userSheet.appendRow(['USR-003', 'siswa@sanggarku.com', hashPasswordV2('siswa123'), 'SISWA', 'Budi Santoso (Siswa)', 'AKTIF', new Date().toISOString(), new Date().toISOString(), 'FALSE']);
      userSheet.appendRow(['USR-004', 'denluckyman@gmail.com', hashPasswordV2('123456789'), 'SUPER_ADMIN', 'Denny (Super Admin)', 'AKTIF', new Date().toISOString(), new Date().toISOString(), 'FALSE']);
    }

    // 3. Sheet SETTING
    let settingSheet = ss.getSheetByName('SETTING');
    if (!settingSheet) {
      settingSheet = ss.insertSheet('SETTING');
      settingSheet.appendRow(['id_setting', 'nama_sanggar', 'alamat', 'kontak', 'logo_url', 'theme_mode', 'created_at', 'updated_at', 'is_deleted']);
    }
    if (settingSheet.getLastRow() === 1) {
      settingSheet.appendRow(['SET-001', 'Happy Drawing Management', 'Jl. Seni Budaya No. 123, Kota Subang', '+62 812-3456-7890', 'https://lh3.googleusercontent.com/d/' + CONFIG.LOGO_ID, 'light', new Date().toISOString(), new Date().toISOString(), 'FALSE']);
    }

    // 4. Sheet KEUANGAN
    let keuSheet = ss.getSheetByName('KEUANGAN');
    if (!keuSheet) {
      keuSheet = ss.insertSheet('KEUANGAN');
      keuSheet.appendRow([
        'ID TRANSAKSI', 'NO INVOICE', 'ID SISWA', 'NAMA SISWA', 'KELAS', 'KATEGORI',
        'JENIS TRANSAKSI', 'PERIODE BULAN', 'PERIODE TAHUN', 'NOMINAL', 'TANGGAL BAYAR', 
        'METODE PEMBAYARAN', 'BUKTI TRANSFER', 'STATUS PEMBAYARAN', 'KETERANGAN', 'VERIFIKATOR', 
        'WAKTU VERIFIKASI', 'IS DELETED', 'CATATAN ADMIN'
      ]);
    }

    // 4b. Sheet PENGELUARAN (Pengeluaran Operasional Sanggar)
    let pengSheet = ss.getSheetByName('PENGELUARAN');
    const pengHeaders = [
      'ID PENGELUARAN', 'TANGGAL', 'KATEGORI', 'KETERANGAN', 'NOMINAL',
      'METODE', 'BUKTI NOTA', 'DICATAT OLEH', 'WAKTU CATAT', 'IS DELETED'
    ];
    if (!pengSheet) {
      pengSheet = ss.insertSheet('PENGELUARAN');
      pengSheet.appendRow(pengHeaders);
    } else if (pengSheet.getLastRow() === 0) {
      pengSheet.appendRow(pengHeaders);
    }

    // 5. Sheet GALERI
    let galSheet = ss.getSheetByName('GALERI');
    const galHeaders = [
      'ID GALERI', 'JUDUL MEDIA', 'KATEGORI', 'DESKRIPSI', 'URL FOTO',
      'ID SISWA PENGUNGGAH', 'NAMA SISWA', 'STATUS TAYANG', 'WAKTU UPLOAD',
      'VERIFIKATOR', 'IS DELETED'
    ];
    if (!galSheet) {
      galSheet = ss.insertSheet('GALERI');
      galSheet.appendRow(galHeaders);
    } else if (galSheet.getLastRow() === 0) {
      galSheet.appendRow(galHeaders);
    }

    // 6. Sheet JADWAL
    let jadwalSheet = ss.getSheetByName('JADWAL');
    const jadwalHeaders = ['ID JADWAL', 'TANGGAL', 'HARI', 'JAM', 'PERTEMUAN', 'KELAS', 'STATUS', 'CREATED_AT', 'IS DELETED'];
    if (!jadwalSheet) {
      jadwalSheet = ss.insertSheet('JADWAL');
      jadwalSheet.appendRow(jadwalHeaders);
    } else if (jadwalSheet.getLastRow() === 0) {
      jadwalSheet.appendRow(jadwalHeaders);
    }

    // 7. Sheet ABSENSI
    let absensiSheet = ss.getSheetByName('ABSENSI');
    const absensiHeaders = [
      'ID ABSENSI', 'ID JADWAL', 'TANGGAL', 'ID SISWA', 'NAMA SISWA',
      'KELAS', 'STATUS KEHADIRAN', 'CATATAN', 'VERIFIKATOR', 'WAKTU RECORD', 'IS DELETED'
    ];
    if (!absensiSheet) {
      absensiSheet = ss.insertSheet('ABSENSI');
      absensiSheet.appendRow(absensiHeaders);
    } else if (absensiSheet.getLastRow() === 0) {
      absensiSheet.appendRow(absensiHeaders);
    }

    // 8. Sheet PENGUMUMAN
    let pengumumanSheet = ss.getSheetByName('PENGUMUMAN');
    const pengumumanHeaders = ['ID PENGUMUMAN', 'JUDUL', 'KATEGORI', 'TANGGAL', 'ISI', 'STATUS', 'CREATED_AT', 'IS DELETED'];
    if (!pengumumanSheet) {
      pengumumanSheet = ss.insertSheet('PENGUMUMAN');
      pengumumanSheet.appendRow(pengumumanHeaders);
    } else if (pengumumanSheet.getLastRow() === 0) {
      pengumumanSheet.appendRow(pengumumanHeaders);
    }

    // 9. Sheet BIODATA
    let bioSheet = ss.getSheetByName('BIODATA');
    const bioHeaders = [
      'NO', 'ID SISWA', 'NAMA SISWA', 'TEMPAT LAHIR', 'TANGGAL LAHIR', 'USIA', 'ASAL SEKOLAH',
      'KELAS', 'KATEGORI', 'NAMA ORANGTUA', 'NOMOR HP', 'EMAIL', 'ALAMAT',
      'PROVINSI', 'KABUPATEN', 'KECAMATAN', 'KELURAHAN', 'FOTO', 'WAKTU DAFTAR', 'IS DELETED'
    ];
    if (!bioSheet) {
      bioSheet = ss.insertSheet('BIODATA');
      bioSheet.appendRow(bioHeaders);
    } else if (bioSheet.getLastRow() === 0) {
      bioSheet.appendRow(bioHeaders);
    }
    if (bioSheet.getLastRow() === 1) {
      bioSheet.appendRow([
        1, '26072606290001', 'Budi Santoso', 'Subang', '2016-05-14', 10, 'SDN Subang 01',
        'Tari Tradisional', 'Smart', 'Santoso', '08123456789', 'siswa@sanggarku.com',
        'Jl. Merdeka No. 12', 'Jawa Barat', 'Subang', 'Subang', 'Karanganyar',
        'https://images.unsplash.com/photo-1534528741775-53994a69daeb?w=150', new Date().toISOString(), 'FALSE'
      ]);
    }

    // 10. Sheet DAFTAR
    let daftarSheet = ss.getSheetByName('DAFTAR');
    const daftarHeaders = [
      'NO', 'NAMA SISWA', 'TEMPAT LAHIR', 'TANGGAL LAHIR', 'USIA', 'ASAL SEKOLAH',
      'KELAS', 'KATEGORI', 'NAMA ORANGTUA', 'NOMOR HP', 'EMAIL', 'PASSWORD',
      'ALAMAT', 'PROVINSI', 'KABUPATEN', 'KECAMATAN', 'KELURAHAN', 'FOTO', 'WAKTU DAFTAR', 'STATUS'
    ];
    if (!daftarSheet) {
      daftarSheet = ss.insertSheet('DAFTAR');
      daftarSheet.appendRow(daftarHeaders);
    } else if (daftarSheet.getLastRow() === 0) {
      daftarSheet.appendRow(daftarHeaders);
    }

    // 11. Sheet REFERENSI
    let refSheet = ss.getSheetByName('REFERENSI');
    const refHeaders = ['KELAS', 'KATEGORI', 'JENIS TRANSAKSI', 'NOMINAL PEMBAYARAN'];
    if (!refSheet) {
      refSheet = ss.insertSheet('REFERENSI');
      refSheet.appendRow(refHeaders);
    } else if (refSheet.getLastRow() === 0) {
      refSheet.appendRow(refHeaders);
    }
    if (refSheet.getLastRow() === 1) {
      refSheet.appendRow(['Tari Tradisional', 'Tari', 'PENDAFTARAN', 85000]);
      refSheet.appendRow(['Lukis Dasar', 'Seni Rupa', 'INFAK BULANAN', 150000]);
      refSheet.appendRow(['Mewarnai & Menggambar', 'Seni Rupa', 'SUKARELA', 50000]);
      refSheet.appendRow(['Manga & Anime', 'Seni Rupa', '', '']);
      refSheet.appendRow(['Digital Art', 'Desain', '', '']);
    }

    // 12. Sheet ALAMAT
    let alSheet = ss.getSheetByName('ALAMAT');
    const alHeaders = ['PROVINSI', 'KABUPATEN', 'KECAMATAN', 'KELURAHAN'];
    if (!alSheet) {
      alSheet = ss.insertSheet('ALAMAT');
      alSheet.appendRow(alHeaders);
    } else if (alSheet.getLastRow() === 0) {
      alSheet.appendRow(alHeaders);
    }
    if (alSheet.getLastRow() === 1) {
      alSheet.appendRow(['Jawa Barat', 'Subang', 'Subang', 'Karanganyar']);
      alSheet.appendRow(['Jawa Barat', 'Subang', 'Subang', 'Sipasung']);
      alSheet.appendRow(['Jawa Barat', 'Subang', 'Subang', 'Pasirkareumbi']);
      alSheet.appendRow(['Jawa Barat', 'Subang', 'Subang', 'Soklat']);
      alSheet.appendRow(['Jawa Barat', 'Subang', 'Cibogo', 'Padaasih']);
      alSheet.appendRow(['Jawa Barat', 'Subang', 'Cibogo', 'Sumurbarang']);
    }

    // 13. Sheet LOG
    let logSheet = ss.getSheetByName('LOG');
    const logHeaders = ['id_log', 'timestamp', 'user_id', 'action', 'module', 'detail'];
    if (!logSheet) {
      logSheet = ss.insertSheet('LOG');
      logSheet.appendRow(logHeaders);
    } else if (logSheet.getLastRow() === 0) {
      logSheet.appendRow(logHeaders);
    }

    return sanitizeForRPC({ success: true, message: "Database dan seluruh 13 Sheet berhasil dikonfigurasi!" });
  } catch (error) {
    return sanitizeForRPC({ success: false, message: "Gagal membuat database: " + error.toString() });
  }
}

function hashPassword(password) {
  if (!password) return '';
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(password), Utilities.Charset.UTF_8);
  let txtHash = '';
  for (let i = 0; i < rawHash.length; i++) {
    let byte = rawHash[i];
    if (byte < 0) byte += 256;
    let byteStr = byte.toString(16);
    if (byteStr.length === 1) byteStr = '0' + byteStr;
    txtHash += byteStr;
  }
  return txtHash;
}

function hashPasswordV2(password, salt) {
  if (!password) return '';
  const cleanPass = String(password).trim();
  if (!cleanPass) return '';
  if (!salt) {
    const saltBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, Utilities.getUuid() + Math.random().toString(), Utilities.Charset.UTF_8);
    salt = '';
    for (let i = 0; i < 8; i++) {
      let b = saltBytes[i];
      if (b < 0) b += 256;
      let s = b.toString(16);
      salt += (s.length === 1 ? '0' + s : s);
    }
  }
  const saltedPass = salt + ':' + cleanPass;
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, saltedPass, Utilities.Charset.UTF_8);
  let txtHash = '';
  for (let i = 0; i < rawHash.length; i++) {
    let byte = rawHash[i];
    if (byte < 0) byte += 256;
    let byteStr = byte.toString(16);
    if (byteStr.length === 1) byteStr = '0' + byteStr;
    txtHash += byteStr;
  }
  return 'v2$' + salt + '$' + txtHash;
}

function verifyPassword_(inputPassword, storedHash) {
  if (!inputPassword || !storedHash) return false;
  const cleanInput = String(inputPassword).trim();
  const cleanStored = String(storedHash).trim();
  if (!cleanInput || !cleanStored) return false;

  if (cleanStored.indexOf('v2$') === 0) {
    const parts = cleanStored.split('$');
    if (parts.length === 3) {
      const salt = parts[1];
      const expected = hashPasswordV2(cleanInput, salt);
      return expected === cleanStored;
    }
    return false;
  }

  // Legacy SHA-256 (unsalted)
  const legacyHash = hashPassword(cleanInput);
  if (cleanStored.toLowerCase() === legacyHash.toLowerCase()) {
    return true;
  }

  // Legacy plaintext
  if (cleanStored === cleanInput) {
    return true;
  }

  return false;
}

// ----------------------------------------------------
// HELPER OTORISASI & MANAJEMEN SESI (P0.2)
// ----------------------------------------------------
function getSession_(token) {
  if (!token) return null;
  const cleanToken = String(token).trim();
  if (!cleanToken) return null;
  try {
    const cache = CacheService.getScriptCache();
    const raw = cache.get('SESS_' + cleanToken);
    if (!raw) return null;
    const sess = JSON.parse(raw);
    // Sliding TTL: perpanjang masa aktif 24 jam (86400 detik)
    cache.put('SESS_' + cleanToken, raw, 86400);
    return sess;
  } catch (e) {
    return null;
  }
}

function requireAuth_(token, allowedRoles) {
  const sess = getSession_(token);
  if (!sess) {
    return {
      authorized: false,
      error: sanitizeForRPC({
        success: false,
        code: 'SESSION_EXPIRED',
        message: "Sesi Anda telah berakhir. Silakan login kembali."
      })
    };
  }
  if (allowedRoles && Array.isArray(allowedRoles) && allowedRoles.length > 0) {
    const roleNorm = String(sess.role || '').toUpperCase();
    const isAllowed = allowedRoles.some(r => roleNorm === String(r).toUpperCase());
    if (!isAllowed) {
      return {
        authorized: false,
        error: sanitizeForRPC({
          success: false,
          code: 'UNAUTHORIZED',
          message: "Anda tidak memiliki hak akses untuk melakukan aksi ini."
        })
      };
    }
  }
  return { authorized: true, session: sess };
}

function requireOwnership_(sess, targetIdSiswa) {
  if (!sess) return false;
  const roleNorm = String(sess.role || '').toUpperCase();
  if (roleNorm === 'ADMIN' || roleNorm === 'SUPER_ADMIN') return true;
  const sUid = String(sess.uid || '').trim().toLowerCase();
  const tId = String(targetIdSiswa || '').trim().toLowerCase();
  return (sUid !== '' && sUid === tId);
}

function canEditModule_(sess, moduleName) {
  if (!sess) return false;
  const roleNorm = String(sess.role || '').toUpperCase();
  if (roleNorm === 'SUPER_ADMIN') return true;
  if (!sess.permissions) return true;
  try {
    const p = typeof sess.permissions === 'string' ? JSON.parse(sess.permissions) : sess.permissions;
    if (p && p[moduleName] !== undefined) {
      const perm = String(p[moduleName]).toLowerCase().trim();
      return perm === 'editor' || perm === 'true';
    }
  } catch (e) {}
  return true;
}

// ----------------------------------------------------
// HELPER MANAJEMEN INDEKS SESI PER-PENGGUNA (BLOK E)
// ----------------------------------------------------
function registerSessionToken_(identifier, token) {
  if (!identifier || !token) return;
  try {
    const cache = CacheService.getScriptCache();
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, String(identifier).trim().toLowerCase());
    const key = 'SESSUID_' + Utilities.base64EncodeWebSafe(digest);
    const raw = cache.get(key);
    let list = [];
    if (raw) {
      try { list = JSON.parse(raw); } catch (e) { list = []; }
    }
    if (!Array.isArray(list)) list = [];
    if (!list.includes(token)) {
      list.push(token);
    }
    if (list.length > 20) list = list.slice(-20);
    cache.put(key, JSON.stringify(list), 86400);
  } catch (err) {}
}

function unregisterSessionToken_(identifier, token) {
  if (!identifier || !token) return;
  try {
    const cache = CacheService.getScriptCache();
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, String(identifier).trim().toLowerCase());
    const key = 'SESSUID_' + Utilities.base64EncodeWebSafe(digest);
    const raw = cache.get(key);
    if (raw) {
      try {
        let list = JSON.parse(raw);
        if (Array.isArray(list)) {
          list = list.filter(t => t !== token);
          if (list.length > 0) {
            cache.put(key, JSON.stringify(list), 86400);
          } else {
            cache.remove(key);
          }
        }
      } catch (e) {}
    }
  } catch (err) {}
}

function revokeUserSessions_(identifier) {
  if (!identifier) return;
  try {
    const cache = CacheService.getScriptCache();
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, String(identifier).trim().toLowerCase());
    const key = 'SESSUID_' + Utilities.base64EncodeWebSafe(digest);
    const raw = cache.get(key);
    if (raw) {
      try {
        const list = JSON.parse(raw);
        if (Array.isArray(list)) {
          list.forEach(function(tok) {
            if (tok) cache.remove('SESS_' + tok);
          });
        }
      } catch (e) {}
      cache.remove(key);
    }
  } catch (err) {}
}

function validateSession(token) {
  try {
    const sess = getSession_(token);
    if (!sess) {
      return sanitizeForRPC({ success: false, code: 'SESSION_EXPIRED', message: "Sesi tidak valid atau telah berakhir." });
    }
    // Periksa status akun di sheet USER/USERS
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const userSheet = ss.getSheetByName('USER') || ss.getSheetByName('USERS');
    if (userSheet) {
      const data = userSheet.getDataRange().getValues();
      if (data.length > 1) {
        const headers = data[0].map(h => String(h || '').toLowerCase().replace(/[^a-z0-9]/g, ''));
        let emailIdx = headers.findIndex(h => h === 'email' || h === 'surel');
        let statusIdx = headers.findIndex(h => h.includes('status') || h.includes('aktif'));
        let delIdx = headers.findIndex(h => h.includes('deleted') || h.includes('hapus'));
        let roleIdx = headers.findIndex(h => h.includes('role') || h.includes('hak'));
        let permIdx = headers.findIndex(h => h.includes('permis') || h.includes('hakakses') || h.includes('otorisasi'));
        
        for (let i = 1; i < data.length; i++) {
          let uEmail = emailIdx > -1 ? String(data[i][emailIdx] || '').trim().toLowerCase() : '';
          let uId = String(data[i][0] || '').trim().toLowerCase();
          if (uEmail === String(sess.email || '').toLowerCase() || uId === String(sess.uid || '').toLowerCase()) {
            let status = statusIdx > -1 ? String(data[i][statusIdx] || 'AKTIF').trim().toUpperCase() : 'AKTIF';
            let isDel = delIdx > -1 ? String(data[i][delIdx] || 'FALSE').trim().toUpperCase() : 'FALSE';
            if (isDel === 'TRUE' || status === 'NONAKTIF' || status === 'INACTIVE' || status === 'FALSE' || status === 'DITOLAK') {
              CacheService.getScriptCache().remove('SESS_' + token);
              return sanitizeForRPC({ success: false, code: 'ACCOUNT_DISABLED', message: "Akun Anda telah dinonaktifkan." });
            }
            if (roleIdx > -1) {
              let r = String(data[i][roleIdx] || '').trim().toUpperCase();
              let cleanR = r.replace(/[^A-Z0-9]/g, '').replace(/^ROLE/, 'ROL');
              if (cleanR.includes('SUPER') || cleanR.includes('ROL001') || cleanR === '1') sess.role = 'SUPER_ADMIN';
              else if (cleanR.includes('ADMIN') || cleanR.includes('ROL002') || cleanR === '2') sess.role = 'ADMIN';
              else if (cleanR.includes('SISWA') || cleanR.includes('ROL003') || cleanR === '3') sess.role = 'SISWA';
            }
            if (permIdx > -1) {
              sess.permissions = String(data[i][permIdx] || '').trim();
            }
            break;
          }
        }
      }
    }
    return sanitizeForRPC({
      success: true,
      user: {
        id: sess.uid,
        email: sess.email,
        nama: sess.nama,
        role: sess.role,
        permissions: sess.permissions || '',
        token: token
      }
    });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: err.toString() });
  }
}

function logoutSession(token) {
  try {
    if (token) {
      const cleanTok = String(token).trim();
      const sess = getSession_(cleanTok);
      if (sess) {
        unregisterSessionToken_(sess.email, cleanTok);
        unregisterSessionToken_(sess.uid, cleanTok);
      }
      CacheService.getScriptCache().remove('SESS_' + cleanTok);
    }
    return sanitizeForRPC({ success: true, message: "Sesi berhasil dihapus." });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: err.toString() });
  }
}

function ensurePengeluaranSheet_() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let pengSheet = ss.getSheetByName('PENGELUARAN');
    const headers = [
      'ID PENGELUARAN', 'TANGGAL', 'KATEGORI', 'KETERANGAN', 'NOMINAL',
      'METODE', 'BUKTI NOTA', 'DICATAT OLEH', 'WAKTU CATAT', 'IS DELETED'
    ];
    if (!pengSheet) {
      pengSheet = ss.insertSheet('PENGELUARAN');
      pengSheet.appendRow(headers);
    } else if (pengSheet.getLastRow() === 0) {
      pengSheet.appendRow(headers);
    }
    return pengSheet;
  } catch (err) {
    return null;
  }
}

function doGet(e) {
  ensurePengeluaranSheet_();
  var output = HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle(CONFIG.APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');

  try {
    output.setFaviconUrl('https://happydrawing.web.app/favicon.ico');
  } catch (favErr) {
    try {
      output.setFaviconUrl('https://happydrawing.web.app/favicon.png');
    } catch (favErr2) {
      Logger.log('Favicon error: ' + favErr2.message);
    }
  }

  return output;
}

function doPost(e) {
  try {
    const request = JSON.parse(e.postData.contents);
    const action = request.action;
    const payload = request.payload;

    // Hoist token otorisasi agar dapat diteruskan ke setiap fungsi handler di bawah
    const token = request.authToken || (payload && (payload.authToken || payload.__token || payload.token)) || '';

    // Allowlist endpoint publik tanpa token otorisasi
    const PUBLIC_ACTIONS = [
      'login', 'resetAndSendPassword', 'submitBiodataPendaftaran',
      'getFormOptions', 'fetchFormOptions', 'fetchGaleriData', 'fetchJadwalLes'
    ];

    if (!PUBLIC_ACTIONS.includes(action)) {
      const ADMIN_ACTIONS = [
        'getDashboardStats', 'getDashboardOverview', 'fetchDashboardOverview', 'getSiswaList', 'fetchMasterSiswa', 'fetchBiodata',
        'fetchUsersAdmin', 'softDeleteSiswa', 'reactivateSiswa', 'updateStatusPendaftaranDaftar',
        'approvePendaftaranSiswa', 'rejectPendaftaranSiswa', 'resetNomorUrutDanIdSiswa', 'sendEmailSiswa',
        'updateStatusGaleri', 'softDeleteGaleri', 'updateStatusPembayaran', 'saveJadwalLes',
        'saveBatchJadwalLes', 'deleteJadwalLes', 'saveBatchAbsensi', 'fetchAbsensiByJadwal',
        'savePengumuman', 'deletePengumuman', 'saveUserAdmin', 'deleteUserAdmin',
        'softDeleteUserAdmin', 'reactivateUserAdmin', 'rejectPengajuanIzin',
        'createTransaksiManual', 'softDeleteTransaksi', 'logLaporanAccess',
        'savePengeluaran', 'softDeletePengeluaran'
      ];

      if (ADMIN_ACTIONS.includes(action)) {
        const auth = requireAuth_(token, ['ADMIN', 'SUPER_ADMIN']);
        if (!auth.authorized) {
          return ContentService.createTextOutput(JSON.stringify(auth.error)).setMimeType(ContentService.MimeType.JSON);
        }
      } else if (action === 'fetchTableData') {
        const targetSheet = String((payload && (payload.sheetName || payload.table)) || payload || '').toUpperCase();
        const ADMIN_SHEETS = ['KEUANGAN', 'PENGELUARAN', 'LOG', 'USERS', 'USER', 'DAFTAR', 'BIODATA', 'ABSENSI', 'ROLE', 'GALERI'];
        const requiredRoles = ADMIN_SHEETS.includes(targetSheet) ? ['ADMIN', 'SUPER_ADMIN'] : null;
        const auth = requireAuth_(token, requiredRoles);
        if (!auth.authorized) {
          return ContentService.createTextOutput(JSON.stringify(auth.error)).setMimeType(ContentService.MimeType.JSON);
        }
      } else {
        const auth = requireAuth_(token, null);
        if (!auth.authorized) {
          return ContentService.createTextOutput(JSON.stringify(auth.error)).setMimeType(ContentService.MimeType.JSON);
        }
        // Validasi kepemilikan data (ownership) untuk aksi siswa
        let targetId = '';
        if (action === 'fetchKeuanganSiswa' || action === 'submitBuktiPembayaran' || action === 'submitPengajuanIzinSakit' || action === 'batalkanPengajuanIzin') {
          targetId = (payload && payload.idSiswa) || '';
        } else if (action === 'fetchAbsensiSiswa') {
          targetId = (payload && (payload.idSiswa || payload)) || '';
        } else if (action === 'updateAkunSiswa' || action === 'getInfoAkunSiswa') {
          targetId = (payload && (payload.id || payload.idSiswa)) || '';
        } else if (action === 'fetchBiodataSiswa' || action === 'updateBiodataSiswa') {
          targetId = (payload && (payload.idSiswa || payload.id || payload)) || '';
        } else if (action === 'submitUploadGaleriSiswa' || action === 'updateGaleriSiswa' || action === 'deleteGaleriSiswa' || action === 'ajukanUlangGaleri') {
          targetId = (payload && (payload.idSiswa || payload.idSiswaPengunggah)) || '';
        }

        if (targetId && !requireOwnership_(auth.session, targetId)) {
          return ContentService.createTextOutput(JSON.stringify(sanitizeForRPC({ success: false, message: "Akses data pengguna lain tidak diizinkan." }))).setMimeType(ContentService.MimeType.JSON);
        }
      }
    }

    let response = { success: false, message: "Aksi tidak dikenali." };

    switch (action) {
      case 'login':
        response = processLogin(payload.email, payload.password);
        break;
      case 'getDashboardStats':
        response = fetchDashboardStats(token, payload && (payload.filterYear || payload.year), payload && (payload.filterMonth || payload.month));
        break;
      case 'getDashboardOverview':
      case 'fetchDashboardOverview':
        response = fetchDashboardOverview(token, payload && (payload.filterYear || payload.year), payload && (payload.filterMonth || payload.month));
        break;
      case 'getSiswaList':
      case 'fetchMasterSiswa':
      case 'fetchBiodata':
        response = fetchMasterSiswa(token);
        break;
      case 'fetchUsersAdmin':
        response = fetchUsersAdmin(token);
        break;
      case 'softDeleteSiswa':
        response = softDeleteSiswa(payload.idSiswa || payload.id || payload, token);
        break;
      case 'reactivateSiswa':
        response = reactivateSiswa(payload.idSiswa || payload.id || payload, token);
        break;
      case 'getFormOptions':
      case 'fetchFormOptions':
        response = fetchFormOptions();
        break;
      case 'submitBiodataPendaftaran':
        response = submitBiodataPendaftaran(payload);
        break;
      case 'updateStatusPendaftaranDaftar':
        response = updateStatusPendaftaranDaftar(payload.email, payload.statusBaru, token);
        break;
      case 'approvePendaftaranSiswa':
        response = approvePendaftaranSiswa(payload.idSiswa || payload.idUser || payload, payload.verifikator, token);
        break;
      case 'rejectPendaftaranSiswa':
        response = rejectPendaftaranSiswa(payload.idSiswa || payload.idUser || payload, payload.verifikator, payload.alasan, token);
        break;
      case 'updateBiodataSiswa':
        response = updateBiodataSiswa(payload);
        break;
      case 'fetchBiodataSiswa':
        response = fetchBiodataSiswa(payload && (payload.idSiswa || payload.id) ? (payload.idSiswa || payload.id) : payload, token);
        break;
      case 'updateAkunSiswa':
        response = updateAkunSiswa(payload, token);
        break;
      case 'getInfoAkunSiswa':
        response = getInfoAkunSiswa(payload, token);
        break;
      case 'resetAndSendPassword':
        response = resetAndSendPassword(payload.email);
        break;
      case 'fetchKeuanganSiswa':
        response = fetchKeuanganSiswa(payload.idSiswa, token);
        break;
      case 'submitBuktiPembayaran':
        response = submitBuktiPembayaran(payload, token);
        break;
      case 'fetchGaleriData':
        response = fetchGaleriData(token);
        break;
      case 'fetchTableData':
        response = fetchTableData(payload && (payload.sheetName || payload.table) ? (payload.sheetName || payload.table) : payload, token);
        break;
      case 'resetNomorUrutDanIdSiswa':
        response = resetNomorUrutDanIdSiswa(token);
        break;
      case 'sendEmailSiswa':
        response = sendEmailSiswa(payload, token);
        break;
      case 'submitUploadGaleriSiswa':
        response = submitUploadGaleriSiswa(payload, token);
        break;
      case 'updateGaleriSiswa':
        response = updateGaleriSiswa(payload.idGaleri, payload.judulMedia, payload.kategori, payload.deskripsi, payload.idSiswa || payload.idSiswaPengunggah, token);
        break;
      case 'deleteGaleriSiswa':
        response = deleteGaleriSiswa(payload.idGaleri, payload.idSiswa || payload.idSiswaPengunggah, token);
        break;
      case 'ajukanUlangGaleri':
        response = ajukanUlangGaleri(payload.idGaleri, payload.idSiswa || payload.idSiswaPengunggah, token);
        break;
      case 'updateStatusGaleri':
        response = updateStatusGaleri(payload.idGaleri, payload.statusTayang, payload.verifikator, payload.catatanAdmin, token);
        break;
      case 'softDeleteGaleri':
        response = softDeleteGaleri(payload.idGaleri, token);
        break;
      case 'updateStatusPembayaran':
        response = updateStatusPembayaran(payload.idTransaksi, payload.statusPembayaran, payload.verifikator, payload.waktuVerifikasi, payload.catatanAdmin, token);
        break;
      case 'createTransaksiManual':
        response = createTransaksiManual(payload, token);
        break;
      case 'softDeleteTransaksi':
        response = softDeleteTransaksi(payload && (payload.idTransaksi || payload), token);
        break;
      case 'logLaporanAccess':
        response = logLaporanAccess(payload && (payload.action || payload), token);
        break;
      case 'savePengeluaran':
        response = savePengeluaran(payload, token);
        break;
      case 'softDeletePengeluaran':
        response = softDeletePengeluaran(payload && (payload.idPengeluaran || payload), token);
        break;
      case 'fetchJadwalLes':
        response = fetchJadwalLes(token);
        break;
      case 'saveJadwalLes':
        response = saveJadwalLes(payload, token);
        break;
      case 'saveBatchJadwalLes':
        response = saveBatchJadwalLes(payload.schedulesArray || payload, token);
        break;
      case 'deleteJadwalLes':
        response = deleteJadwalLes(payload.idJadwal || payload, token);
        break;
      case 'saveBatchAbsensi':
        response = saveBatchAbsensi(payload.batchData || payload, token);
        break;
      case 'fetchAbsensiByJadwal':
        response = fetchAbsensiByJadwal(payload.idJadwal || payload, token);
        break;
      case 'fetchAbsensiSiswa':
        response = fetchAbsensiSiswa(payload.idSiswa || payload, token);
        break;
      case 'submitPengajuanIzinSakit':
        response = submitPengajuanIzinSakit(payload, token);
        break;
      case 'batalkanPengajuanIzin':
        response = batalkanPengajuanIzin(payload.idAbsensi || payload.id, payload.idSiswa, token);
        break;
      case 'rejectPengajuanIzin':
        response = rejectPengajuanIzin(payload, token);
        break;
      case 'savePengumuman':
        response = savePengumuman(payload, token);
        break;
      case 'deletePengumuman':
        response = deletePengumuman(payload.idPengumuman || payload, token);
        break;
      case 'saveUserAdmin':
        response = saveUserAdmin(payload, token);
        break;
      case 'deleteUserAdmin':
        response = deleteUserAdmin(payload.idUser || payload, token);
        break;
      case 'softDeleteUserAdmin':
        response = softDeleteUserAdmin(payload.idUser || payload, token);
        break;
      case 'reactivateUserAdmin':
        response = reactivateUserAdmin(payload.idUser || payload, token);
        break;
      case 'getSyncStamp':
        response = getSyncStamp(token);
        break;
      default:
        response = { success: false, message: "Endpoint API tidak valid atau ditolak." };
    }

    return ContentService.createTextOutput(JSON.stringify(sanitizeForRPC(response))).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify(sanitizeForRPC({ success: false, message: error.toString() }))).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Helper pengiriman email multi-tier:
 * 1. GmailApp.sendEmail (Native Apps Script - scope https://www.googleapis.com/auth/gmail.send)
 * 2. MailApp.sendEmail (Native Apps Script - scope https://www.googleapis.com/auth/script.send_mail)
 * 3. Gmail REST API via UrlFetchApp (endpoint https://gmail.googleapis.com/gmail/v1/users/me/messages/send)
 * @param {Object} opts { to, subject, body, htmlBody }
 */
function kirimEmailGmailApi_(opts) {
  const to = String((opts && opts.to) || '').trim();
  const subject = String((opts && opts.subject) || '');
  const bodyText = String((opts && opts.body) || '');
  const htmlBody = String((opts && opts.htmlBody) || '');
  const senderName = (typeof CONFIG !== 'undefined' && CONFIG.APP_NAME) ? CONFIG.APP_NAME : 'Happy Drawing';

  if (!to || to.indexOf('@') === -1) {
    throw new Error('Alamat email tujuan tidak valid: ' + to);
  }

  const errors = [];

  // Jalur 1: GmailApp (Native Apps Script service - menggunakan scope https://www.googleapis.com/auth/gmail.send)
  try {
    const gmailOpts = { name: senderName };
    if (htmlBody) gmailOpts.htmlBody = htmlBody;
    GmailApp.sendEmail(to, subject, bodyText, gmailOpts);
    return 'Sent via GmailApp';
  } catch (errGmail) {
    errors.push('GmailApp: ' + ((errGmail && errGmail.message) ? errGmail.message : String(errGmail)));
  }

  // Jalur 2: MailApp (Native Apps Script service - menggunakan scope https://www.googleapis.com/auth/script.send_mail)
  try {
    const mailOpts = {
      to: to,
      subject: subject,
      body: bodyText,
      name: senderName
    };
    if (htmlBody) mailOpts.htmlBody = htmlBody;
    MailApp.sendEmail(mailOpts);
    return 'Sent via MailApp';
  } catch (errMail) {
    errors.push('MailApp: ' + ((errMail && errMail.message) ? errMail.message : String(errMail)));
  }

  // Jalur 3: Gmail REST API via UrlFetchApp
  try {
    const token = ScriptApp.getOAuthToken();
    const boundary = 'happy_drawing_' + Utilities.getUuid();
    const lines = [];
    lines.push('To: ' + to);
    lines.push('Subject: =?UTF-8?B?' + Utilities.base64Encode(Utilities.newBlob(subject).getBytes()) + '?=');
    lines.push('MIME-Version: 1.0');
    lines.push('Content-Type: multipart/alternative; boundary="' + boundary + '"');
    lines.push('');
    lines.push('--' + boundary);
    lines.push('Content-Type: text/plain; charset="UTF-8"');
    lines.push('');
    lines.push(bodyText);
    lines.push('--' + boundary);
    lines.push('Content-Type: text/html; charset="UTF-8"');
    lines.push('');
    lines.push(htmlBody);
    lines.push('--' + boundary + '--');

    const raw = Utilities.base64EncodeWebSafe(Utilities.newBlob(lines.join('\r\n')).getBytes());

    const res = UrlFetchApp.fetch('https://gmail.googleapis.com/gmail/v1/users/me/messages/send', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({ raw: raw }),
      muteHttpExceptions: true
    });

    const code = res.getResponseCode();
    if (code >= 200 && code < 300) {
      return res.getContentText();
    }
    errors.push('Gmail REST API HTTP ' + code + ': ' + res.getContentText());
  } catch (errRest) {
    errors.push('Gmail REST API: ' + ((errRest && errRest.message) ? errRest.message : String(errRest)));
  }

  throw new Error('Gagal mengirim email: ' + errors.join(' | '));
}

/**
 * Fungsi helper untuk pengujian dan pemicu dialog otorisasi izin pengiriman email di script editor Apps Script.
 * Jalankan fungsi ini dari Apps Script Editor jika muncul pesan izin belum disetujui.
 */
function testOtorisasiEmail() {
  const targetEmail = Session.getEffectiveUser().getEmail() || Session.getActiveUser().getEmail() || 'rumahkhitandamai@gmail.com';
  Logger.log('Menjalankan tes pengiriman email ke: ' + targetEmail);
  const hasil = kirimEmailGmailApi_({
    to: targetEmail,
    subject: '[Happy Drawing] Tes Otorisasi Pengiriman Email Berhasil',
    body: 'Halo,\n\nIzin pengiriman email sistem Happy Drawing telah berhasil disetujui dan berfungsi normal.',
    htmlBody: '<div style="font-family:Arial,sans-serif;padding:15px;border-left:4px solid #4CAF50;"><h3>Happy Drawing - Tes Otorisasi</h3><p>Izin pengiriman email sistem Happy Drawing telah berhasil disetujui dan berfungsi normal.</p></div>'
  });
  Logger.log('Hasil: ' + hasil);
  return hasil;
}

function kirimEmailPersetujuan_(targetEmail, valNama, idSiswa, valKelas) {
  try {
    if (!targetEmail || targetEmail.indexOf('@') === -1) return;
    const emailSubject = "[Happy Drawing] Selamat! Pendaftaran Siswa Baru Telah Disetujui";
    const emailBodyText = "Halo " + (valNama || 'Siswa') + ",\n\n" +
      "Selamat! Pendaftaran Anda di Sanggar Happy Drawing telah DISETUJUI oleh Admin.\n\n" +
      "Berikut informasi akun Portal Siswa Anda:\n" +
      "• ID Siswa        : " + (idSiswa || '-') + "\n" +
      "• Nama Siswa      : " + (valNama || 'Siswa') + "\n" +
      "• Program / Kelas : " + (valKelas || '-') + "\n" +
      "• Email Akun      : " + targetEmail + "\n" +
      "• Password        : Sesuai password yang Anda buat saat pendaftaran\n\n" +
      "Silakan gunakan Email atau ID Siswa Anda untuk login ke aplikasi Happy Drawing.\n\n" +
      "Salam,\nHappy Drawing Management";

    const emailHtmlBody = `
      <div style="font-family: 'Poppins', Arial, sans-serif; max-width: 540px; margin: 0 auto; padding: 28px; border: 1px solid #e2e8f0; border-radius: 16px; background-color: #ffffff;">
        <div style="text-align: center; margin-bottom: 24px;">
          <h2 style="color: #0d6efd; margin: 0; font-size: 24px; font-weight: 700;">Happy Drawing</h2>
          <p style="color: #64748b; font-size: 13px; margin-top: 4px;">Belajar Bermain Berkarya Penuh Warna</p>
        </div>
        <div style="background: #e8f5e9; border: 1px solid #c8e6c9; border-radius: 12px; padding: 14px 18px; margin-bottom: 20px; text-align: center;">
          <span style="color: #2e7d32; font-weight: 700; font-size: 16px;">✓ Pendaftaran Berhasil Disetujui</span>
        </div>
        <p style="color: #1e293b; font-size: 15px; font-weight: 600; margin-bottom: 12px;">Halo ${valNama || 'Siswa'},</p>
        <p style="color: #475569; font-size: 14px; line-height: 1.6; margin-bottom: 20px;">
          Selamat! Pendaftaran Anda di <strong>Sanggar Happy Drawing</strong> telah diverifikasi dan <strong>DISETUJUI</strong> oleh Admin. Anda kini telah resmi terdaftar sebagai siswa aktif.
        </p>
        <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 12px; padding: 18px; margin-bottom: 24px;">
          <table style="width: 100%; border-collapse: collapse; font-size: 13.5px;">
            <tr>
              <td style="padding: 7px 0; color: #64748b; width: 38%;">ID Siswa</td>
              <td style="padding: 7px 0; color: #0d6efd; font-weight: 700; font-size: 16px; letter-spacing: 0.5px;">${idSiswa || '-'}</td>
            </tr>
            <tr>
              <td style="padding: 7px 0; color: #64748b;">Nama Siswa</td>
              <td style="padding: 7px 0; color: #1e293b; font-weight: 600;">${valNama || 'Siswa'}</td>
            </tr>
            <tr>
              <td style="padding: 7px 0; color: #64748b;">Program / Kelas</td>
              <td style="padding: 7px 0; color: #1e293b;">${valKelas || '-'}</td>
            </tr>
            <tr>
              <td style="padding: 7px 0; color: #64748b;">Email Login</td>
              <td style="padding: 7px 0; color: #1e293b;">${targetEmail}</td>
            </tr>
            <tr>
              <td style="padding: 7px 0; color: #64748b;">Password</td>
              <td style="padding: 7px 0; color: #64748b; font-style: italic;">Sesuai password yang dibuat saat mendaftar</td>
            </tr>
          </table>
        </div>
        <p style="color: #475569; font-size: 13.5px; line-height: 1.6; margin-bottom: 20px;">
          Silakan masuk ke Portal Siswa menggunakan <strong>Email</strong> atau <strong>ID Siswa</strong> di atas untuk mengakses informasi kelas, jadwal, pembayaran, dan presensi.
        </p>
        <div style="text-align: center; margin: 25px 0 10px 0; border-top: 1px solid #f1f5f9; padding-top: 15px;">
          <p style="color: #94a3b8; font-size: 12px; margin: 0;">Email ini dikirim otomatis oleh Sistem Informasi Manajemen Sanggar Happy Drawing.</p>
        </div>
      </div>
    `;

    kirimEmailGmailApi_({
      to: targetEmail,
      subject: emailSubject,
      body: emailBodyText,
      htmlBody: emailHtmlBody
    });
  } catch (err) {
    Logger.log("Peringatan: Gagal mengirim email persetujuan ke " + targetEmail + ": " + err.toString());
  }
}

function kirimEmailPenolakan_(targetEmail, rejectedNama, alasan) {
  try {
    if (!targetEmail || targetEmail.indexOf('@') === -1) return;
    const emailSubject = "[Happy Drawing] Informasi Status Pendaftaran Siswa";
    const emailBodyText = "Halo " + (rejectedNama || 'Pendaftar') + ",\n\n" +
      "Terima kasih atas ketertarikan Anda untuk mendaftar di Sanggar Happy Drawing.\n\n" +
      "Setelah proses verifikasi data, mohon maaf saat ini pendaftaran Anda belum dapat kami setujui." +
      (alasan ? " Alasan: " + alasan + "\n\n" : " Hal ini dapat disebabkan oleh kelengkapan berkas atau kuota kelas yang saat ini telah terpenuhi.\n\n") +
      "Jika Anda memiliki pertanyaan lebih lanjut, silakan hubungi Admin Sanggar Happy Drawing.\n\n" +
      "Salam,\nHappy Drawing Management";

    const emailHtmlBody = `
      <div style="font-family: 'Poppins', Arial, sans-serif; max-width: 540px; margin: 0 auto; padding: 28px; border: 1px solid #e2e8f0; border-radius: 16px; background-color: #ffffff;">
        <div style="text-align: center; margin-bottom: 24px;">
          <h2 style="color: #0d6efd; margin: 0; font-size: 24px; font-weight: 700;">Happy Drawing</h2>
          <p style="color: #64748b; font-size: 13px; margin-top: 4px;">Belajar Bermain Berkarya Penuh Warna</p>
        </div>
        <div style="background: #fff3f3; border: 1px solid #ffcdd2; border-radius: 12px; padding: 14px 18px; margin-bottom: 20px; text-align: center;">
          <span style="color: #c62828; font-weight: 700; font-size: 16px;">Status Pendaftaran: Belum Disetujui</span>
        </div>
        <p style="color: #1e293b; font-size: 15px; font-weight: 600; margin-bottom: 12px;">Halo ${rejectedNama || 'Pendaftar'},</p>
        <p style="color: #475569; font-size: 14px; line-height: 1.6; margin-bottom: 16px;">
          Terima kasih atas ketertarikan Anda untuk bergabung dengan <strong>Sanggar Happy Drawing</strong>.
        </p>
        <p style="color: #475569; font-size: 14px; line-height: 1.6; margin-bottom: 16px;">
          Setelah melalui proses verifikasi data, mohon maaf saat ini pendaftaran Anda <strong>belum dapat kami setujui</strong>.
          ${alasan ? '<br><strong>Catatan:</strong> ' + String(alasan).replace(/</g, '&lt;').replace(/>/g, '&gt;') : 'Hal ini dapat disebabkan oleh kelengkapan berkas pendaftaran atau kuota kelas yang saat ini telah terpenuhi.'}
        </p>
        <p style="color: #475569; font-size: 13.5px; line-height: 1.6; margin-bottom: 20px;">
          Jika Anda membutuhkan informasi lebih lanjut mengenai pendaftaran atau ingin mendaftar pada periode berikutnya, silakan hubungi tim Admin kami.
        </p>
        <div style="text-align: center; margin: 25px 0 10px 0; border-top: 1px solid #f1f5f9; padding-top: 15px;">
          <p style="color: #94a3b8; font-size: 12px; margin: 0;">Email ini dikirim otomatis oleh Sistem Informasi Manajemen Sanggar Happy Drawing.</p>
        </div>
      </div>
    `;

    kirimEmailGmailApi_({
      to: targetEmail,
      subject: emailSubject,
      body: emailBodyText,
      htmlBody: emailHtmlBody
    });
  } catch (err) {
    Logger.log("Peringatan: Gagal mengirim email penolakan ke " + targetEmail + ": " + err.toString());
  }
}

/**
 * Memperbarui status pendaftaran khusus di Sheet DAFTAR
 * Jika statusBaru === 'DISETUJUI', menyalin data pendaftar ke Sheet USER dan Sheet BIODATA secara otomatis.
 */
function updateStatusPendaftaranDaftar(email, statusBaru, authToken) {
  const lock = LockService.getScriptLock();
  try {
    const lockAcquired = lock.tryLock(15000);
    if (!lockAcquired) {
      return sanitizeForRPC({ success: false, message: "Sistem sedang sibuk memproses pendaftaran lain. Silakan coba beberapa saat lagi." });
    }

    const auth = requireAuth_(authToken, ['ADMIN', 'SUPER_ADMIN']);
    if (!auth.authorized) return auth.error;
    const actorId = (auth.session && (auth.session.uid || auth.session.email)) || 'ADMIN';

    const normalizedStatus = String(statusBaru || '').trim().toUpperCase();
    if (normalizedStatus !== 'DISETUJUI' && normalizedStatus !== 'DITOLAK') {
      return sanitizeForRPC({ success: false, message: "Status pendaftaran tidak valid. Hanya menerima DISETUJUI atau DITOLAK." });
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let daftarSheet = ss.getSheetByName('DAFTAR');
    if (!daftarSheet) return sanitizeForRPC({ success: false, message: "Sheet DAFTAR tidak ditemukan!" });

    const data = daftarSheet.getDataRange().getValues();
    if (data.length <= 1) return sanitizeForRPC({ success: false, message: "Sheet DAFTAR kosong." });

    const headersClean = data[0].map(h => String(h || '').toUpperCase().trim());
    const emailIdx = headersClean.indexOf('EMAIL');
    if (emailIdx === -1) {
      return sanitizeForRPC({ success: false, message: "Format Sheet DAFTAR tidak valid: kolom EMAIL tidak ditemukan." });
    }
    let statusIdx = headersClean.indexOf('STATUS');

    if (statusIdx === -1) {
      statusIdx = headersClean.length;
      daftarSheet.getRange(1, statusIdx + 1).setValue('STATUS');
    }

    const targetEmail = String(email || '').trim().toLowerCase();
    let foundRow = -1;
    let studentData = {};

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][emailIdx]).trim().toLowerCase() === targetEmail) {
        foundRow = i + 1;
        headersClean.forEach((h, idx) => {
          studentData[h] = data[i][idx];
        });
        break;
      }
    }

    // A9: Jika tidak ditemukan di DAFTAR, periksa draf akun di USER untuk aksi dari modul Users
    if (foundRow === -1) {
      let userSheet = ss.getSheetByName('USER') || ss.getSheetByName('USERS');
      if (userSheet && userSheet.getLastRow() > 1) {
        const uData = userSheet.getDataRange().getValues();
        const uHeaders = uData[0].map(h => String(h || '').toLowerCase().replace(/[^a-z0-9]/g, ''));
        let uEmailIdx = uHeaders.findIndex(h => h === 'email' || h === 'surel');
        if (uEmailIdx === -1) uEmailIdx = 1;
        let uRoleIdx = uHeaders.findIndex(h => h.includes('role'));
        if (uRoleIdx === -1) uRoleIdx = 3;
        let uStatusIdx = uHeaders.findIndex(h => h === 'status');
        if (uStatusIdx === -1) uStatusIdx = 5;
        let uUpdatedIdx = uHeaders.findIndex(h => h.includes('update'));

        for (let u = 1; u < uData.length; u++) {
          if (String(uData[u][uEmailIdx] || '').trim().toLowerCase() === targetEmail) {
            const rowRole = String(uData[u][uRoleIdx] || '').toUpperCase();
            const rowStatus = String(uData[u][uStatusIdx] || '').toUpperCase();
            if ((rowRole === 'SISWA' || rowRole === 'ROL-003') && (rowStatus === 'PENDING' || rowStatus === 'PENDING_VERIFIKASI')) {
              const targetStatus = (normalizedStatus === 'DISETUJUI') ? 'AKTIF' : 'DITOLAK';
              userSheet.getRange(u + 1, uStatusIdx + 1).setValue(targetStatus);
              if (uUpdatedIdx > -1) {
                userSheet.getRange(u + 1, uUpdatedIdx + 1).setValue(new Date().toISOString());
              }

              let uNamaIdx = uHeaders.findIndex(h => h === 'nama' || h === 'namalengkap');
              let uIdIdx = uHeaders.findIndex(h => h === 'iduser' || h === 'id' || h === 'id_user');
              const studentNama = (uNamaIdx > -1) ? String(uData[u][uNamaIdx] || 'Siswa').trim() : 'Siswa';
              const studentId = (uIdIdx > -1) ? String(uData[u][uIdIdx] || '').trim() : '';

              if (targetStatus === 'AKTIF') {
                kirimEmailPersetujuan_(targetEmail, studentNama, studentId, '-');
              } else {
                kirimEmailPenolakan_(targetEmail, studentNama, '');
              }

              logActivity(actorId, 'VERIFIKASI_USER_PENDING', 'USER', 'Ubah status user draf ' + targetEmail + ' menjadi ' + targetStatus);
              return sanitizeForRPC({
                success: true,
                message: "Data DAFTAR tidak ditemukan, tetapi akun draf USER berhasil diubah menjadi " + targetStatus + "."
              });
            }
          }
        }
      }
      return sanitizeForRPC({ success: false, message: "Email pendaftar tidak ditemukan di Sheet DAFTAR (mungkin sudah pernah diproses)." });
    }

    const currentStatus = String(studentData['STATUS'] || '').trim().toUpperCase();
    if (currentStatus !== '' && currentStatus !== 'PENDING') {
      return sanitizeForRPC({
        success: false,
        message: "Pendaftaran sudah diproses sebelumnya (Status saat ini: " + currentStatus + ")."
      });
    }

    if (normalizedStatus === 'DISETUJUI') {
      const getStudentVal = function(keys) {
        for (let k of keys) {
          if (studentData[k] !== undefined && studentData[k] !== null && String(studentData[k]).trim() !== '') {
            return studentData[k];
          }
        }
        return '';
      };

      // 1. Ekstrak Tanggal Lahir untuk Format ID SISWA: YY(Birth) + MM(Birth) + DD(Birth)
      let rawTglLahir = getStudentVal(['TANGGAL LAHIR', 'TGL LAHIR', 'TGLLAHIR']);
      let birthYY = "00", birthMM = "01", birthDD = "01";

      if (rawTglLahir) {
        let bDate = null;
        if (rawTglLahir instanceof Date) {
          bDate = rawTglLahir;
        } else {
          let strTgl = String(rawTglLahir).trim();
          let ymd = strTgl.match(/^(\d{4})[--\/](\d{1,2})[--\/](\d{1,2})/);
          let dmy = strTgl.match(/^(\d{1,2})[--\/](\d{1,2})[--\/](\d{4})/);
          if (ymd) {
            bDate = new Date(parseInt(ymd[1], 10), parseInt(ymd[2], 10) - 1, parseInt(ymd[3], 10));
          } else if (dmy) {
            bDate = new Date(parseInt(dmy[3], 10), parseInt(dmy[2], 10) - 1, parseInt(dmy[1], 10));
          } else {
            let p = Date.parse(strTgl);
            if (!isNaN(p)) bDate = new Date(p);
          }
        }
        if (bDate && !isNaN(bDate.getTime())) {
          birthYY = String(bDate.getFullYear()).slice(-2);
          birthMM = String(bDate.getMonth() + 1).padStart(2, '0');
          birthDD = String(bDate.getDate()).padStart(2, '0');
        }
      }

      // 2. Ekstrak Tanggal Daftar untuk Format ID SISWA: YY(Reg) + MM(Reg)
      let rawWaktuDaftar = getStudentVal(['WAKTU DAFTAR', 'TIMESTAMP', 'WAKTU']);
      let rDate = new Date();
      if (rawWaktuDaftar) {
        if (rawWaktuDaftar instanceof Date) {
          rDate = rawWaktuDaftar;
        } else {
          let p = Date.parse(String(rawWaktuDaftar));
          if (!isNaN(p)) rDate = new Date(p);
        }
      }
      let regYY = String(rDate.getFullYear()).slice(-2);
      let regMM = String(rDate.getMonth() + 1).padStart(2, '0');

      // 3. Ekstrak & Hitung Nomor Urut Siswa (4 Digit Terakhir)
      // Selalu hitung nomor urut berikutnya dari Sheet BIODATA agar selalu unik, konsisten, & berurutan (0001, 0002, dst.)
      let bioSheet = ss.getSheetByName('BIODATA');
      let highestSeq = 0;
      let existingBioIds = new Set();

      if (bioSheet && bioSheet.getLastRow() > 1) {
        let bData = bioSheet.getDataRange().getValues();
        let bHeaders = bData[0].map(h => String(h || '').toUpperCase().trim());
        let bIdIdx = bHeaders.findIndex(h => h === 'ID SISWA' || h === 'IDSISWA' || h === 'ID');
        let bNoIdx = bHeaders.indexOf('NO');

        for (let b = 1; b < bData.length; b++) {
          let bId = bIdIdx > -1 ? String(bData[b][bIdIdx] || '').trim() : '';
          if (bId) {
            existingBioIds.add(bId);
            if (bId.length >= 4) {
              let suffixNum = parseInt(bId.slice(-4), 10);
              if (!isNaN(suffixNum) && suffixNum > highestSeq) {
                highestSeq = suffixNum;
              }
            }
          }
          if (bNoIdx > -1) {
            let rowNo = parseInt(bData[b][bNoIdx], 10);
            if (!isNaN(rowNo) && rowNo > highestSeq) {
              highestSeq = rowNo;
            }
          }
        }
      }

      const props = PropertiesService.getScriptProperties();
      // Nomor urut diturunkan langsung dari Sheet BIODATA (suffix 4 digit terbesar + 1)
      // agar tidak terjadi double-increment dari GLOBAL_SEQ_SISWA yang sudah di-reserve saat pendaftaran
      let targetSeq = highestSeq + 1;
      let seqStr = String(targetSeq).padStart(4, '0');
      let candidateId = birthYY + birthMM + birthDD + regYY + regMM + seqStr;

      // Validasi mutlak anti-duplikasi: jika ID kandidat sudah ada di BIODATA, increment sampai benar-benar unik
      while (existingBioIds.has(candidateId)) {
        targetSeq++;
        seqStr = String(targetSeq).padStart(4, '0');
        candidateId = birthYY + birthMM + birthDD + regYY + regMM + seqStr;
      }

      // Format ID SISWA: [YY_Lahir][MM_Lahir][DD_Lahir][YY_Daftar][MM_Daftar][Seq4Digit]
      const idSiswa = candidateId;
      props.setProperty('GLOBAL_SEQ_SISWA', targetSeq.toString());
      props.setProperty('GLOBAL_SEQ_DAFTAR', targetSeq.toString());
      const valNama = String(getStudentVal(['NAMA SISWA', 'NAMA', 'NAMASISWA', 'NAMA LENGKAP']) || 'Siswa Sanggar').trim();
      const valTempatLahir = getStudentVal(['TEMPAT LAHIR', 'TEMPATLAHIR', 'TEMPAT']);
      const valTglLahir = rawTglLahir;
      const valUsia = getStudentVal(['USIA', 'UMUR']);
      const valAsalSekolah = getStudentVal(['ASAL SEKOLAH', 'ASALSEKOLAH', 'SEKOLAH']);
      const valKelas = getStudentVal(['KELAS', 'PROGRAM']);
      const valKategori = getStudentVal(['KATEGORI', 'LEVEL']);
      const valNamaOrtu = getStudentVal(['NAMA ORANGTUA', 'NAMA ORTU', 'NAMAORTU', 'ORANGTUA']);
      const valNoHp = getStudentVal(['NOMOR HP', 'NO HP', 'NOHP', 'WA', 'NOMOR HP/WA']);
      const valEmail = targetEmail;
      const valAlamat = getStudentVal(['ALAMAT', 'ALAMAT LENGKAP']);
      const valProvinsi = getStudentVal(['PROVINSI']);
      const valKabupaten = getStudentVal(['KABUPATEN', 'KABUPATEN/KOTA', 'KOTA']);
      const valKecamatan = getStudentVal(['KECAMATAN']);
      const valKelurahan = getStudentVal(['KELURAHAN', 'KELURAHAN/DESA', 'DESA']);
      const valFoto = getStudentVal(['FOTO', 'PAS FOTO', 'PASFOTO', 'URL FOTO']);
      const nowIso = new Date().toISOString();
      const valWaktuDaftar = rawWaktuDaftar || nowIso;

      // Ambil password yang dibuat siswa saat pendaftaran, bukan me-reset ke password default 'siswa123'
      let regPass = getStudentVal(['PASSWORD', 'PASS', 'SANDI']);
      let finalStudentPassword = '';
      if (regPass) {
        let strPass = String(regPass).trim();
        if (strPass.indexOf('v2$') === 0) {
          finalStudentPassword = strPass;
        } else {
          finalStudentPassword = hashPasswordV2(strPass);
        }
      } else {
        finalStudentPassword = hashPasswordV2('siswa123');
      }

      // 4. Simpan / Sync ke Sheet USER
      let userSheet = ss.getSheetByName('USER') || ss.getSheetByName('USERS');
      if (!userSheet) {
        userSheet = ss.insertSheet('USER');
        userSheet.appendRow(['id_user', 'email', 'password', 'id_role', 'nama', 'status', 'created_at', 'updated_at', 'is_deleted']);
      }

      const uData = userSheet.getDataRange().getValues();
      const uHeaders = uData[0].map(h => String(h || '').toLowerCase().replace(/[^a-z0-9]/g, ''));
      let uIdIdx = uHeaders.findIndex(h => h === 'iduser' || h === 'id' || h === 'id_user');
      if (uIdIdx === -1) uIdIdx = 0;
      let uEmailIdx = uHeaders.findIndex(h => h === 'email' || h === 'surel');
      if (uEmailIdx === -1) uEmailIdx = 1;
      let uPassIdx = uHeaders.findIndex(h => h.includes('pass') || h.includes('pwd') || h.includes('sandi'));
      if (uPassIdx === -1) uPassIdx = 2;
      let uRoleIdx = uHeaders.findIndex(h => h.includes('role'));
      if (uRoleIdx === -1) uRoleIdx = 3;
      let uNamaIdx = uHeaders.findIndex(h => h === 'nama' || h === 'namalengkap');
      if (uNamaIdx === -1) uNamaIdx = 4;
      let uStatusIdx = uHeaders.findIndex(h => h === 'status');
      if (uStatusIdx === -1) uStatusIdx = 5;
      let uUpdatedIdx = uHeaders.findIndex(h => h.includes('update'));
      if (uUpdatedIdx === -1) uUpdatedIdx = 7;
      let uDelIdx = uHeaders.findIndex(h => h.includes('delete') || h.includes('hapus'));
      if (uDelIdx === -1 && uData[0].length >= 9) uDelIdx = 8;

      let existingUserRow = -1;
      for (let u = 1; u < uData.length; u++) {
        if (String(uData[u][uEmailIdx] || '').trim().toLowerCase() === targetEmail) {
          existingUserRow = u + 1;
          break;
        }
      }

      // A6: Guard anti-overwriting akun admin
      if (existingUserRow > -1) {
        const existingRole = String(uData[existingUserRow - 1][uRoleIdx] || '').toUpperCase();
        if (existingRole === 'SUPER_ADMIN' || existingRole === 'ADMIN' || existingRole === 'ROL-001' || existingRole === 'ROL-002') {
          return sanitizeForRPC({
            success: false,
            message: "Email pendaftar '" + targetEmail + "' terdaftar sebagai akun Admin/Super Admin. Persetujuan pendaftaran dibatalkan demi keamanan."
          });
        }

        userSheet.getRange(existingUserRow, uIdIdx + 1).setValue(idSiswa);
        if (uPassIdx > -1 && finalStudentPassword) {
          userSheet.getRange(existingUserRow, uPassIdx + 1).setValue(finalStudentPassword);
        }
        userSheet.getRange(existingUserRow, uRoleIdx + 1).setValue('SISWA');
        userSheet.getRange(existingUserRow, uNamaIdx + 1).setValue(valNama);
        userSheet.getRange(existingUserRow, uStatusIdx + 1).setValue('AKTIF');
        if (uUpdatedIdx > -1) {
          userSheet.getRange(existingUserRow, uUpdatedIdx + 1).setValue(nowIso);
        }
        if (uDelIdx > -1) {
          userSheet.getRange(existingUserRow, uDelIdx + 1).setValue('FALSE');
        }
      } else {
        const newRow = uHeaders.map((h, idx) => {
          if (idx === uIdIdx) return idSiswa;
          if (idx === uEmailIdx) return targetEmail;
          if (idx === uPassIdx) return finalStudentPassword;
          if (idx === uRoleIdx) return 'SISWA';
          if (idx === uNamaIdx) return valNama;
          if (idx === uStatusIdx) return 'AKTIF';
          if (h.includes('create')) return nowIso;
          if (idx === uUpdatedIdx) return nowIso;
          if (idx === uDelIdx) return 'FALSE';
          return '';
        });
        userSheet.appendRow(newRow);
      }

      // 5. Pindahkan Data ke Sheet BIODATA Sesuai Struktur Kolomnya:
      // [NO, ID SISWA, NAMA SISWA, TEMPAT LAHIR, TANGGAL LAHIR, USIA, ASAL SEKOLAH, KELAS, KATEGORI, NAMA ORANGTUA, NOMOR HP, EMAIL, ALAMAT, PROVINSI, KABUPATEN, KECAMATAN, KELURAHAN, FOTO, WAKTU DAFTAR, IS DELETED]
      if (!bioSheet) bioSheet = ss.getSheetByName('BIODATA');
      const targetBioHeaders = [
        'NO', 'ID SISWA', 'NAMA SISWA', 'TEMPAT LAHIR', 'TANGGAL LAHIR', 'USIA', 'ASAL SEKOLAH',
        'KELAS', 'KATEGORI', 'NAMA ORANGTUA', 'NOMOR HP', 'EMAIL', 'ALAMAT',
        'PROVINSI', 'KABUPATEN', 'KECAMATAN', 'KELURAHAN', 'FOTO', 'WAKTU DAFTAR', 'IS DELETED'
      ];

      if (!bioSheet) {
        bioSheet = ss.insertSheet('BIODATA');
        bioSheet.appendRow(targetBioHeaders);
      } else if (bioSheet.getLastRow() === 0) {
        bioSheet.appendRow(targetBioHeaders);
      }

      const bioData = bioSheet.getDataRange().getValues();
      const currentBioHeaders = bioData[0].map(h => String(h || '').toUpperCase().trim());

      let bioEmailIdx = currentBioHeaders.findIndex(h => h === 'EMAIL' || h === 'SUREL');
      let bioRow = -1;
      if (bioEmailIdx > -1 && bioData.length > 1) {
        for (let b = 1; b < bioData.length; b++) {
          if (String(bioData[b][bioEmailIdx] || '').trim().toLowerCase() === targetEmail) {
            bioRow = b + 1;
            break;
          }
        }
      }

      let calculatedNo = bioRow > -1 ? (parseInt(bioData[bioRow - 1][0], 10) || (bioRow - 1)) : targetSeq;

      const getBioValForHeader = function(headerName) {
        const h = String(headerName || '').toUpperCase().trim();
        if (h === 'NO' || h === 'NO.' || h === 'NOMOR') return calculatedNo;
        if (h === 'ID SISWA' || h === 'IDSISWA' || h === 'ID') return idSiswa;
        if (h === 'NAMA SISWA' || h === 'NAMASISWA' || h === 'NAMA') return valNama;
        if (h === 'TEMPAT LAHIR' || h === 'TEMPATLAHIR') return valTempatLahir;
        if (h === 'TANGGAL LAHIR' || h === 'TGLLAHIR') return valTglLahir;
        if (h === 'USIA' || h === 'UMUR') return valUsia;
        if (h === 'ASAL SEKOLAH' || h === 'ASALSEKOLAH') return valAsalSekolah;
        if (h === 'KELAS' || h === 'PROGRAM') return valKelas;
        if (h === 'KATEGORI' || h === 'LEVEL') return valKategori;
        if (h === 'NAMA ORANGTUA' || h === 'NAMA ORTU' || h === 'NAMAORTU') return valNamaOrtu;
        if (h === 'NOMOR HP' || h === 'NO HP' || h === 'NOHP') return valNoHp;
        if (h === 'EMAIL' || h === 'SUREL') return valEmail;
        if (h === 'ALAMAT') return valAlamat;
        if (h === 'PROVINSI') return valProvinsi;
        if (h === 'KABUPATEN' || h === 'KOTA') return valKabupaten;
        if (h === 'KECAMATAN') return valKecamatan;
        if (h === 'KELURAHAN' || h === 'DESA') return valKelurahan;
        if (h === 'FOTO' || h === 'PAS FOTO') return valFoto;
        if (h === 'WAKTU DAFTAR' || h === 'TIMESTAMP') return valWaktuDaftar;
        if (h === 'IS DELETED' || h === 'IS_DELETED') return 'FALSE';
        return '';
      };

      if (bioRow > -1) {
        currentBioHeaders.forEach((h, colIdx) => {
          bioSheet.getRange(bioRow, colIdx + 1).setValue(getBioValForHeader(h));
        });
      } else {
        const newBioRow = currentBioHeaders.map(h => getBioValForHeader(h));
        bioSheet.appendRow(newBioRow);
      }

      // 6. Hapus Baris Data Pendaftar di Sheet DAFTAR
      daftarSheet.deleteRow(foundRow);

      // 7. Kirim Notifikasi Email Otomatis ke Pendaftar
      try {
        if (targetEmail && targetEmail.indexOf('@') !== -1) {
          const emailSubject = "[Happy Drawing] Selamat! Pendaftaran Siswa Baru Telah Disetujui";
          const emailBodyText = "Halo " + valNama + ",\n\n" +
            "Selamat! Pendaftaran Anda di Sanggar Happy Drawing telah DISETUJUI oleh Admin.\n\n" +
            "Berikut informasi akun Portal Siswa Anda:\n" +
            "• ID Siswa        : " + idSiswa + "\n" +
            "• Nama Siswa      : " + valNama + "\n" +
            "• Program / Kelas : " + (valKelas || '-') + "\n" +
            "• Email Akun      : " + targetEmail + "\n" +
            "• Password        : Sesuai password yang Anda buat saat pendaftaran\n\n" +
            "Silakan gunakan Email atau ID Siswa Anda untuk login ke aplikasi Happy Drawing.\n\n" +
            "Salam,\nHappy Drawing Management";

          const emailHtmlBody = `
            <div style="font-family: 'Poppins', Arial, sans-serif; max-width: 540px; margin: 0 auto; padding: 28px; border: 1px solid #e2e8f0; border-radius: 16px; background-color: #ffffff;">
              <div style="text-align: center; margin-bottom: 24px;">
                <h2 style="color: #0d6efd; margin: 0; font-size: 24px; font-weight: 700;">Happy Drawing</h2>
                <p style="color: #64748b; font-size: 13px; margin-top: 4px;">Belajar Bermain Berkarya Penuh Warna</p>
              </div>
              <div style="background: #e8f5e9; border: 1px solid #c8e6c9; border-radius: 12px; padding: 14px 18px; margin-bottom: 20px; text-align: center;">
                <span style="color: #2e7d32; font-weight: 700; font-size: 16px;">✓ Pendaftaran Berhasil Disetujui</span>
              </div>
              <p style="color: #1e293b; font-size: 15px; font-weight: 600; margin-bottom: 12px;">Halo ${valNama},</p>
              <p style="color: #475569; font-size: 14px; line-height: 1.6; margin-bottom: 20px;">
                Selamat! Pendaftaran Anda di <strong>Sanggar Happy Drawing</strong> telah diverifikasi dan <strong>DISETUJUI</strong> oleh Admin. Anda kini telah resmi terdaftar sebagai siswa aktif.
              </p>
              <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 12px; padding: 18px; margin-bottom: 24px;">
                <table style="width: 100%; border-collapse: collapse; font-size: 13.5px;">
                  <tr>
                    <td style="padding: 7px 0; color: #64748b; width: 38%;">ID Siswa</td>
                    <td style="padding: 7px 0; color: #0d6efd; font-weight: 700; font-size: 16px; letter-spacing: 0.5px;">${idSiswa}</td>
                  </tr>
                  <tr>
                    <td style="padding: 7px 0; color: #64748b;">Nama Siswa</td>
                    <td style="padding: 7px 0; color: #1e293b; font-weight: 600;">${valNama}</td>
                  </tr>
                  <tr>
                    <td style="padding: 7px 0; color: #64748b;">Program / Kelas</td>
                    <td style="padding: 7px 0; color: #1e293b;">${valKelas || '-'}</td>
                  </tr>
                  <tr>
                    <td style="padding: 7px 0; color: #64748b;">Email Login</td>
                    <td style="padding: 7px 0; color: #1e293b;">${targetEmail}</td>
                  </tr>
                  <tr>
                    <td style="padding: 7px 0; color: #64748b;">Password</td>
                    <td style="padding: 7px 0; color: #64748b; font-style: italic;">Sesuai password yang dibuat saat mendaftar</td>
                  </tr>
                </table>
              </div>
              <p style="color: #475569; font-size: 13.5px; line-height: 1.6; margin-bottom: 20px;">
                Silakan masuk ke Portal Siswa menggunakan <strong>Email</strong> atau <strong>ID Siswa</strong> di atas untuk mengakses informasi kelas, jadwal, pembayaran, dan presensi.
              </p>
              <div style="text-align: center; margin: 25px 0 10px 0; border-top: 1px solid #f1f5f9; pt: 15px;">
                <p style="color: #94a3b8; font-size: 12px; margin: 0;">Email ini dikirim otomatis oleh Sistem Informasi Manajemen Sanggar Happy Drawing.</p>
              </div>
            </div>
          `;

          kirimEmailGmailApi_({
            to: targetEmail,
            subject: emailSubject,
            body: emailBodyText,
            htmlBody: emailHtmlBody
          });
        }
      } catch (mailErr) {
        Logger.log("Peringatan: Gagal mengirim email persetujuan ke " + targetEmail + ": " + mailErr.toString());
      }

      logActivity(actorId, 'APPROVE_PENDAFTARAN', 'DAFTAR', 'Setujui & pindahkan pendaftar ' + email + ' ke BIODATA (ID SISWA: ' + idSiswa + ')');
      return sanitizeForRPC({
        success: true,
        message: "Pendaftaran disetujui! Data berhasil dipindahkan ke Sheet BIODATA & USER dengan ID SISWA: " + idSiswa
      });

    } else {
      daftarSheet.getRange(foundRow, statusIdx + 1).setValue(normalizedStatus);

      // A1: Sinkronkan ke Sheet USER jika ada draf akun pendaftar (role SISWA & status PENDING)
      let userSheet = ss.getSheetByName('USER') || ss.getSheetByName('USERS');
      if (userSheet && userSheet.getLastRow() > 1) {
        const uData = userSheet.getDataRange().getValues();
        const uHeaders = uData[0].map(h => String(h || '').toLowerCase().replace(/[^a-z0-9]/g, ''));
        let uEmailIdx = uHeaders.findIndex(h => h === 'email' || h === 'surel');
        if (uEmailIdx === -1) uEmailIdx = 1;
        let uRoleIdx = uHeaders.findIndex(h => h.includes('role'));
        if (uRoleIdx === -1) uRoleIdx = 3;
        let uStatusIdx = uHeaders.findIndex(h => h === 'status');
        if (uStatusIdx === -1) uStatusIdx = 5;
        let uUpdatedIdx = uHeaders.findIndex(h => h.includes('update'));

        for (let u = 1; u < uData.length; u++) {
          const rowEmail = String(uData[u][uEmailIdx] || '').trim().toLowerCase();
          if (rowEmail === targetEmail) {
            const rowRole = String(uData[u][uRoleIdx] || '').toUpperCase();
            const rowStatus = String(uData[u][uStatusIdx] || '').toUpperCase();
            if ((rowRole === 'SISWA' || rowRole === 'ROL-003') && (rowStatus === 'PENDING' || rowStatus === 'PENDING_VERIFIKASI')) {
              userSheet.getRange(u + 1, uStatusIdx + 1).setValue('DITOLAK');
              if (uUpdatedIdx > -1) {
                userSheet.getRange(u + 1, uUpdatedIdx + 1).setValue(new Date().toISOString());
              }
            }
            break;
          }
        }
      }

      // Kirim Notifikasi Email Otomatis saat pendaftaran DITOLAK
      if (normalizedStatus === 'DITOLAK') {
        try {
          if (targetEmail && targetEmail.indexOf('@') !== -1) {
            const rejectedNama = String(studentData['NAMA SISWA'] || studentData['NAMA'] || studentData['NAMASISWA'] || studentData['NAMA LENGKAP'] || 'Pendaftar').trim();
            const emailSubject = "[Happy Drawing] Informasi Status Pendaftaran Siswa";
            const emailBodyText = "Halo " + rejectedNama + ",\n\n" +
              "Terima kasih atas ketertarikan Anda untuk mendaftar di Sanggar Happy Drawing.\n\n" +
              "Setelah proses verifikasi data, mohon maaf saat ini pendaftaran Anda belum dapat kami setujui. Hal ini dapat disebabkan oleh kelengkapan berkas atau kuota kelas yang saat ini telah terpenuhi.\n\n" +
              "Jika Anda memiliki pertanyaan lebih lanjut, silakan hubungi Admin Sanggar Happy Drawing.\n\n" +
              "Salam,\nHappy Drawing Management";

            const emailHtmlBody = `
              <div style="font-family: 'Poppins', Arial, sans-serif; max-width: 540px; margin: 0 auto; padding: 28px; border: 1px solid #e2e8f0; border-radius: 16px; background-color: #ffffff;">
                <div style="text-align: center; margin-bottom: 24px;">
                  <h2 style="color: #0d6efd; margin: 0; font-size: 24px; font-weight: 700;">Happy Drawing</h2>
                  <p style="color: #64748b; font-size: 13px; margin-top: 4px;">Belajar Bermain Berkarya Penuh Warna</p>
                </div>
                <div style="background: #fff3f3; border: 1px solid #ffcdd2; border-radius: 12px; padding: 14px 18px; margin-bottom: 20px; text-align: center;">
                  <span style="color: #c62828; font-weight: 700; font-size: 16px;">Status Pendaftaran: Belum Disetujui</span>
                </div>
                <p style="color: #1e293b; font-size: 15px; font-weight: 600; margin-bottom: 12px;">Halo ${rejectedNama},</p>
                <p style="color: #475569; font-size: 14px; line-height: 1.6; margin-bottom: 16px;">
                  Terima kasih atas ketertarikan Anda untuk bergabung dengan <strong>Sanggar Happy Drawing</strong>.
                </p>
                <p style="color: #475569; font-size: 14px; line-height: 1.6; margin-bottom: 16px;">
                  Setelah melalui proses verifikasi data, mohon maaf saat ini pendaftaran Anda <strong>belum dapat kami setujui</strong>. Hal ini dapat disebabkan oleh kelengkapan berkas pendaftaran atau kuota kelas yang saat ini telah terpenuhi.
                </p>
                <p style="color: #475569; font-size: 13.5px; line-height: 1.6; margin-bottom: 20px;">
                  Jika Anda membutuhkan informasi lebih lanjut mengenai pendaftaran atau ingin mendaftar pada periode berikutnya, silakan hubungi tim Admin kami.
                </p>
                <div style="text-align: center; margin: 25px 0 10px 0; border-top: 1px solid #f1f5f9; pt: 15px;">
                  <p style="color: #94a3b8; font-size: 12px; margin: 0;">Email ini dikirim otomatis oleh Sistem Informasi Manajemen Sanggar Happy Drawing.</p>
                </div>
              </div>
            `;

            kirimEmailGmailApi_({
              to: targetEmail,
              subject: emailSubject,
              body: emailBodyText,
              htmlBody: emailHtmlBody
            });
          }
        } catch (mailErr) {
          Logger.log("Peringatan: Gagal mengirim email penolakan ke " + targetEmail + ": " + mailErr.toString());
        }
      }

      logActivity(actorId, 'VERIFIKASI_PENDAFTARAN', 'DAFTAR', 'Ubah status pendaftaran ' + email + ' menjadi ' + normalizedStatus);
      return sanitizeForRPC({ success: true, message: "Status pendaftaran berhasil diubah menjadi " + normalizedStatus + "." });
    }
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal memproses: " + err.toString() });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function processLogin(email, password) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const userSheet = ss.getSheetByName('USER') || ss.getSheetByName('USERS');
    
    if (!userSheet) return sanitizeForRPC({ success: false, message: "Sheet USER/USERS tidak ditemukan!" });

    const data = userSheet.getDataRange().getValues();
    if (data.length <= 1) return sanitizeForRPC({ success: false, message: "Data pengguna masih kosong!" });
    
    const rawHeaders = data[0];
    const headers = rawHeaders.map(h => String(h || '').toLowerCase().replace(/[^a-z0-9]/g, ''));
    
    let idIdx = headers.findIndex(h => h.includes('iduser') || h.includes('id_user') || h === 'id');
    if (idIdx === -1) {
      let candidate = headers.findIndex(h => h.includes('id'));
      idIdx = candidate > -1 ? candidate : 0;
    }

    let emailIdx = headers.findIndex(h => h === 'email' || h === 'surel');
    if (emailIdx === -1) emailIdx = 1;

    let passIdx = headers.findIndex(h => h.includes('pass') || h.includes('pwd') || h.includes('sandi'));
    if (passIdx === -1) passIdx = 2;

    let roleIdx = headers.findIndex(h => h.includes('role') || h.includes('hak') || h.includes('akses') || h.includes('level'));
    if (roleIdx === -1) roleIdx = 3;

    let namaIdx = headers.findIndex(h => h.includes('nama') || h.includes('name'));
    if (namaIdx === -1) namaIdx = 4;

    let statusIdx = headers.findIndex(h => h.includes('status') || h.includes('aktif'));
    if (statusIdx === -1) statusIdx = 5;

    let delIdx = headers.findIndex(h => h.includes('deleted') || h.includes('hapus'));
    if (delIdx === -1) delIdx = headers.length - 1;

    let permIdx = headers.findIndex(h => h.includes('permis') || h.includes('hakakses') || h.includes('otorisasi'));

    const inputEmail = String(email || '').trim().toLowerCase();
    const inputPass = String(password || '').trim();
    if (!inputEmail || !inputPass) {
      return sanitizeForRPC({ success: false, message: "Email/ID dan Password wajib diisi!" });
    }

    const cache = CacheService.getScriptCache();
    const emailHashBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, inputEmail, Utilities.Charset.UTF_8);
    let emailHash = '';
    for (let h = 0; h < emailHashBytes.length; h++) {
      let b = emailHashBytes[h];
      if (b < 0) b += 256;
      let s = b.toString(16);
      emailHash += (s.length === 1 ? '0' + s : s);
    }
    const rlKey = 'LOGIN_FAIL_' + emailHash;
    const currentFails = parseInt(cache.get(rlKey) || '0', 10);
    if (currentFails >= 5) {
      return sanitizeForRPC({
        success: false,
        message: "Terlalu banyak percobaan login yang gagal. Akun dikunci sementara, silakan coba lagi dalam 15 menit."
      });
    }

    let mappedEmail = inputEmail;
    try {
      const bioSheet = ss.getSheetByName('BIODATA');
      if (bioSheet && bioSheet.getLastRow() > 1) {
        const bioData = bioSheet.getDataRange().getValues();
        const bioHeadersClean = bioData[0].map(h => String(h || '').toUpperCase().replace(/[^A-Z0-9]/g, ''));
        
        let idIdxBio = bioHeadersClean.findIndex(h => h === 'IDSISWA' || h === 'ID' || h === 'NOID');
        let emailIdxBio = bioHeadersClean.findIndex(h => h === 'EMAIL' || h === 'SUREL');

        if (idIdxBio > -1 && emailIdxBio > -1) {
          for (let b = bioData.length - 1; b >= 1; b--) {
            const bId = String(bioData[b][idIdxBio] || '').trim().toLowerCase();
            if (bId !== '' && bId === inputEmail) {
              mappedEmail = String(bioData[b][emailIdxBio] || '').trim().toLowerCase();
              break;
            }
          }
        }
      }
    } catch (e) {}

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const uId = String(row[idIdx] || '').trim().toLowerCase();
      const uEmail = String(row[emailIdx] || '').trim().toLowerCase();
      const uUsername = uEmail.includes('@') ? uEmail.split('@')[0] : uEmail;
      const uPass = String(row[passIdx] || '').trim();
      
      let rawRole = String(row[roleIdx] || '').trim().toUpperCase();

      let normalizedRole = 'SISWA';
      const cleanRole = rawRole.replace(/[^A-Z0-9]/g, '');
      const stdRole = cleanRole.replace(/^ROLE/, 'ROL');
      if (stdRole.includes('SUPER') || stdRole.includes('ROL001') || stdRole === '1') {
        normalizedRole = 'SUPER_ADMIN';
      } else if (stdRole.includes('ADMIN') || stdRole.includes('ROL002') || stdRole === '2') {
        normalizedRole = 'ADMIN';
      } else if (stdRole.includes('SISWA') || stdRole.includes('STUDENT') || stdRole.includes('ROL003') || stdRole === '3') {
        normalizedRole = 'SISWA';
      } else if (rawRole) {
        normalizedRole = rawRole;
      }

      const nama = String(row[namaIdx] || 'Pengguna').trim();
      const status = String(row[statusIdx] || 'AKTIF').trim().toUpperCase();
      const isDeleted = String(row[delIdx] || 'FALSE').trim().toUpperCase();

      const isMatchUser = (uEmail === inputEmail) || (uId === inputEmail) || (uUsername === inputEmail) || (uEmail === mappedEmail);
      const isMatchPass = verifyPassword_(inputPass, uPass);

      if (isMatchUser && isMatchPass) {
        // Reset counter kegagalan saat login berhasil
        cache.remove(rlKey);

        // Migrasi otomatis ke format v2$salt$hash jika akun masih menggunakan hash lama / plaintext
        if (uPass.indexOf('v2$') !== 0) {
          try {
            userSheet.getRange(i + 1, passIdx + 1).setValue(hashPasswordV2(inputPass));
          } catch (migErr) {}
        }
        if (status === 'PENDING' || status === 'PENDING_VERIFIKASI') {
          return sanitizeForRPC({ success: false, message: "Pendaftaran Anda sedang menunggu verifikasi dari Admin Sanggar. Silakan hubungi Admin untuk informasi lebih lanjut." });
        }

        // Catatan: status === 'FALSE' dipertahankan untuk kompatibilitas jika kolom status bertipe boolean spreadsheet (AKTIF)
        if (isDeleted === 'TRUE' || status === 'NONAKTIF' || status === 'INACTIVE' || status === 'FALSE' || status === 'DITOLAK') {
          return sanitizeForRPC({ success: false, message: "Akun Anda telah dinonaktifkan atau ditolak. Silahkan Hubungi Admin Happy Drawing." });
        }

        let returnedId = (uId && uId !== '-') ? String(row[idIdx]).trim() : String(row[0] || ('USR-' + i)).trim();
        let returnedFoto = ''; 
        let returnedKategori = '-'; 

        let biodataObj = {
          id: returnedId, nama: nama, tempatLahir: '-', tglLahir: '-', usia: '-',
          asalSekolah: '-', kelas: '-', kategori: '-', namaOrtu: '-', noHp: '-',
          email: uEmail, alamat: '-', provinsi: '-', kabupaten: '-',
          kecamatan: '-', kelurahan: '-', foto: '-'
        };

        if (normalizedRole === 'SISWA') {
          try {
            const bioSheet = ss.getSheetByName('BIODATA');
            if (bioSheet && bioSheet.getLastRow() > 1) {
              const bioData = bioSheet.getDataRange().getValues();
              const bioHeadersClean = bioData[0].map(h => String(h || '').toUpperCase().replace(/[^A-Z0-9]/g, ''));
              
              let emailIdxBio = bioHeadersClean.findIndex(h => h === 'EMAIL' || h === 'SUREL');

              let matchRow = -1;
              const targetId = returnedId.toLowerCase();
              const targetInput = inputEmail.toLowerCase();
              const targetEmailString = uEmail.toLowerCase();

              for (let b = bioData.length - 1; b >= 1; b--) {
                for (let c = 0; c < bioData[b].length; c++) {
                  let cellVal = String(bioData[b][c] || '').trim().toLowerCase();
                  if (cellVal !== '' && (cellVal === targetId || cellVal === targetInput)) {
                    matchRow = b;
                    break;
                  }
                }
                if (matchRow > -1) break;
              }

              if (matchRow === -1 && emailIdxBio > -1) {
                for (let b = bioData.length - 1; b >= 1; b--) {
                  let bEmail = String(bioData[b][emailIdxBio] || '').trim().toLowerCase();
                  if (bEmail !== '' && (bEmail === targetEmailString || bEmail === targetInput || bEmail === mappedEmail)) {
                    matchRow = b;
                    break;
                  }
                }
              }

              if (matchRow > -1) {
                const b = matchRow;
                const findCol = (exacts, includesArr) => {
                  let idx = bioHeadersClean.findIndex(h => exacts.includes(h));
                  if (idx === -1 && includesArr) {
                    idx = bioHeadersClean.findIndex(h => includesArr.some(inc => h.includes(inc)));
                  }
                  return idx;
                };

                const getVal = (exacts, includesArr) => {
                  let idx = findCol(exacts, includesArr);
                  if (idx > -1 && bioData[b][idx] !== undefined && bioData[b][idx] !== null && String(bioData[b][idx]).trim() !== '') {
                    let val = bioData[b][idx];
                    if (val instanceof Date) {
                      const yyyy = val.getFullYear();
                      const mm = String(val.getMonth() + 1).padStart(2, '0');
                      const dd = String(val.getDate()).padStart(2, '0');
                      return yyyy + '-' + mm + '-' + dd;
                    }
                    return String(val).trim();
                  }
                  return '-';
                };

                biodataObj.id = getVal(['IDSISWA', 'ID SISWA', 'ID', 'NOID'], []) !== '-' ? getVal(['IDSISWA', 'ID SISWA', 'ID', 'NOID'], []) : returnedId;
                biodataObj.nama = getVal(['NAMASISWA', 'NAMA SISWA', 'NAMA', 'NAMALENGKAP'], ['NAMA']) !== '-' ? getVal(['NAMASISWA', 'NAMA SISWA', 'NAMA', 'NAMALENGKAP'], ['NAMA']) : nama;
                biodataObj.tempatLahir = getVal(['TEMPATLAHIR', 'TEMPAT'], ['TEMPAT']);
                biodataObj.tglLahir = getVal(['TANGGALLAHIR', 'TGLLAHIR'], ['TANGGAL', 'LAHIR']);
                biodataObj.usia = getVal(['USIA', 'UMUR'], ['USIA', 'UMUR']);
                biodataObj.asalSekolah = getVal(['ASALSEKOLAH', 'SEKOLAHASAL', 'SEKOLAH'], ['SEKOLAH']);
                biodataObj.kelas = getVal(['KELAS', 'PROGRAM'], ['KELAS']);
                biodataObj.kategori = getVal(['KATEGORI', 'LEVEL'], ['KATEGORI']);
                biodataObj.namaOrtu = getVal(['NAMAORANGTUA', 'NAMAORTU', 'ORANGTUA', 'NAMAAYAH', 'NAMAIBU'], ['ORTU', 'ORANGTUA']);
                biodataObj.noHp = getVal(['NOMORHP', 'NOHP', 'WHATSAPP', 'WA', 'HP'], ['HP', 'TELEPON', 'WA']);
                biodataObj.email = getVal(['EMAIL', 'SUREL'], ['EMAIL']) !== '-' ? getVal(['EMAIL', 'SUREL'], ['EMAIL']) : uEmail;
                biodataObj.alamat = getVal(['ALAMAT', 'ALAMATLENGKAP'], ['ALAMAT']);
                biodataObj.provinsi = getVal(['PROVINSI'], ['PROVINSI']);
                biodataObj.kabupaten = getVal(['KABUPATEN', 'KOTA', 'KABKOTA'], ['KABUPATEN', 'KOTA']);
                biodataObj.kecamatan = getVal(['KECAMATAN'], ['KECAMATAN']);
                biodataObj.kelurahan = getVal(['KELURAHAN', 'DESA'], ['KELURAHAN', 'DESA']);
                biodataObj.foto = getVal(['FOTO', 'PASFOTO', 'FOTOPROFIL'], ['FOTO']);

                returnedId = biodataObj.id;
                returnedFoto = biodataObj.foto !== '-' ? biodataObj.foto : '';
                returnedKategori = biodataObj.kategori !== '-' ? biodataObj.kategori : '-';
              }
            }
          } catch (bioErr) {}
        }

        logActivity(returnedId, 'LOGIN', 'AUTH', 'User berhasil login dengan role: ' + normalizedRole);

        let rawPermissions = permIdx > -1 ? String(row[permIdx] || '').trim() : '';

        const sessionToken = Utilities.getUuid().replace(/-/g, '');
        const sessionData = {
          uid: returnedId,
          role: normalizedRole,
          email: uEmail,
          nama: nama,
          permissions: rawPermissions,
          ts: Date.now()
        };
        cache.put('SESS_' + sessionToken, JSON.stringify(sessionData), 86400); // TTL 24 jam

        registerSessionToken_(uEmail, sessionToken);
        registerSessionToken_(returnedId, sessionToken);

        return sanitizeForRPC({
          success: true,
          message: "Login Berhasil!",
          token: sessionToken,
          user: { 
            id: returnedId, 
            email: uEmail, 
            role: normalizedRole, 
            nama: nama, 
            foto: returnedFoto, 
            kategori: returnedKategori,
            biodata: biodataObj,
            permissions: rawPermissions,
            token: sessionToken
          }
        });
      }
    }

    const nextFails = currentFails + 1;
    cache.put(rlKey, String(nextFails), 900); // Kunci 15 menit

    // Catat log audit ke sheet LOG langsung tanpa memicu invalidasi sync stamp berlebih, maksimal 5 entri per siklus
    if (nextFails <= 5) {
      try {
        const logSheet = ss.getSheetByName('LOG');
        if (logSheet) {
          logSheet.appendRow([
            'LOG-' + Date.now(),
            new Date().toISOString(),
            inputEmail,
            'LOGIN_FAILED',
            'AUTH',
            'Gagal login (percobaan ke-' + nextFails + '/5). Akun/Email: ' + inputEmail
          ]);
        }
      } catch (logErr) {}
    }

    if (nextFails >= 5) {
      return sanitizeForRPC({
        success: false,
        message: "Terlalu banyak percobaan login yang gagal. Akun Anda dikunci sementara selama 15 menit."
      });
    }

    return sanitizeForRPC({ success: false, message: "Email/ID atau Password tidak cocok dengan database!" });
  } catch (error) {
    return sanitizeForRPC({ success: false, message: "Terjadi kesalahan server: " + error.toString() });
  }
}

function approvePendaftaranSiswa(idSiswa, verifikator, authToken) {
  try {
    const auth = requireAuth_(authToken, ['ADMIN', 'SUPER_ADMIN']);
    if (!auth.authorized) return auth.error;

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const userSheet = ss.getSheetByName('USER') || ss.getSheetByName('USERS');
    if (!userSheet) return sanitizeForRPC({ success: false, message: "Sheet USER tidak ditemukan!" });

    const uData = userSheet.getDataRange().getValues();
    const uHeadersClean = uData[0].map(h => String(h || '').toLowerCase().replace(/[^a-z0-9]/g, ''));
    let idIdx = uHeadersClean.findIndex(h => h.includes('id'));
    let statusIdx = uHeadersClean.findIndex(h => h.includes('status'));

    const targetId = String(idSiswa || '').trim().toLowerCase();

    for (let u = 1; u < uData.length; u++) {
      let uId = String(uData[u][idIdx] || '').trim().toLowerCase();
      if (uId === targetId) {
        if (statusIdx > -1) userSheet.getRange(u + 1, statusIdx + 1).setValue('AKTIF');
        let uEmailIdx = uHeadersClean.findIndex(h => h === 'email' || h === 'surel');
        let uNamaIdx = uHeadersClean.findIndex(h => h.includes('nama'));
        let targetEmail = (uEmailIdx > -1) ? String(uData[u][uEmailIdx] || '').trim().toLowerCase() : '';
        let studentNama = (uNamaIdx > -1) ? String(uData[u][uNamaIdx] || 'Siswa').trim() : 'Siswa';
        if (targetEmail) {
          kirimEmailPersetujuan_(targetEmail, studentNama, idSiswa, '-');
        }
        break;
      }
    }

    logActivity(verifikator || 'ADMIN', 'APPROVE_PENDAFTARAN', 'BIODATA', 'Menyetujui pendaftaran siswa ID: ' + idSiswa);
    return sanitizeForRPC({ success: true, message: "Pendaftaran siswa ID " + idSiswa + " berhasil disetujui (Status: AKTIF)." });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal menyetujui pendaftaran: " + err.toString() });
  }
}

function rejectPendaftaranSiswa(idSiswa, verifikator, alasan, authToken) {
  try {
    const auth = requireAuth_(authToken, ['ADMIN', 'SUPER_ADMIN']);
    if (!auth.authorized) return auth.error;

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const userSheet = ss.getSheetByName('USER') || ss.getSheetByName('USERS');
    if (!userSheet) return sanitizeForRPC({ success: false, message: "Sheet USER tidak ditemukan!" });

    const uData = userSheet.getDataRange().getValues();
    const uHeadersClean = uData[0].map(h => String(h || '').toLowerCase().replace(/[^a-z0-9]/g, ''));
    let idIdx = uHeadersClean.findIndex(h => h.includes('id'));
    let statusIdx = uHeadersClean.findIndex(h => h.includes('status'));

    const targetId = String(idSiswa || '').trim().toLowerCase();

    for (let u = 1; u < uData.length; u++) {
      let uId = String(uData[u][idIdx] || '').trim().toLowerCase();
      if (uId === targetId) {
        if (statusIdx > -1) userSheet.getRange(u + 1, statusIdx + 1).setValue('DITOLAK');
        let uEmailIdx = uHeadersClean.findIndex(h => h === 'email' || h === 'surel');
        let uNamaIdx = uHeadersClean.findIndex(h => h.includes('nama'));
        let targetEmail = (uEmailIdx > -1) ? String(uData[u][uEmailIdx] || '').trim().toLowerCase() : '';
        let studentNama = (uNamaIdx > -1) ? String(uData[u][uNamaIdx] || 'Siswa').trim() : 'Siswa';
        if (targetEmail) {
          kirimEmailPenolakan_(targetEmail, studentNama, alasan);
        }
        break;
      }
    }

    softDeleteSiswa(idSiswa);

    logActivity(verifikator || 'ADMIN', 'REJECT_PENDAFTARAN', 'BIODATA', 'Menolak pendaftaran siswa ID: ' + idSiswa + ' dengan alasan: ' + (alasan || '-'));
    return sanitizeForRPC({ success: true, message: "Pendaftaran siswa ID " + idSiswa + " telah ditolak." });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal menolak pendaftaran: " + err.toString() });
  }
}

function resetAndSendPassword(email) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const userSheet = ss.getSheetByName('USER') || ss.getSheetByName('USERS');
    if (!userSheet) return sanitizeForRPC({ success: false, message: "Sheet USER tidak ditemukan!" });

    const targetEmail = String(email || '').trim().toLowerCase();
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!targetEmail || !emailRegex.test(targetEmail)) {
      return sanitizeForRPC({ success: false, message: "Format alamat email tidak valid." });
    }

    const GENERIC_MSG = "Jika email terdaftar dan aktif dalam database, instruksi dan password baru akan dikirimkan ke kotak masuk email Anda.";

    const cache = CacheService.getScriptCache();
    const emailHashBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, targetEmail, Utilities.Charset.UTF_8);
    let emailHash = '';
    for (let h = 0; h < emailHashBytes.length; h++) {
      let b = emailHashBytes[h];
      if (b < 0) b += 256;
      let s = b.toString(16);
      emailHash += (s.length === 1 ? '0' + s : s);
    }
    const rlKey = 'PWD_RESET_' + emailHash;
    const currentResets = parseInt(cache.get(rlKey) || '0', 10);
    if (currentResets >= 3) {
      return sanitizeForRPC({
        success: false,
        message: "Terlalu banyak permintaan reset password untuk email ini. Silakan coba lagi dalam 1 jam."
      });
    }
    cache.put(rlKey, String(currentResets + 1), 3600); // Batas 1 jam

    const data = userSheet.getDataRange().getValues();
    if (data.length <= 1) return sanitizeForRPC({ success: true, message: GENERIC_MSG });

    const headersClean = data[0].map(h => String(h || '').toLowerCase().replace(/[^a-z0-9]/g, ''));
    let emailIdx = headersClean.findIndex(h => h === 'email' || h === 'surel');
    if (emailIdx === -1) emailIdx = 1;

    let passIdx = headersClean.findIndex(h => h.includes('pass') || h.includes('pwd') || h.includes('sandi'));
    if (passIdx === -1) passIdx = 2;

    let namaIdx = headersClean.findIndex(h => h.includes('nama') || h.includes('name'));
    if (namaIdx === -1) namaIdx = 4;

    let statusIdx = headersClean.findIndex(h => h.includes('status') || h.includes('aktif'));
    if (statusIdx === -1) statusIdx = 5;

    let delIdx = headersClean.findIndex(h => h.includes('deleted') || h.includes('hapus'));
    if (delIdx === -1) delIdx = headersClean.length - 1;

    let userRow = -1;
    let userName = '';
    let userId = '';

    for (let i = 1; i < data.length; i++) {
      let uEmail = String(data[i][emailIdx] || '').trim().toLowerCase();
      let isDeleted = String(data[i][delIdx] || 'FALSE').trim().toUpperCase();
      let status = String(data[i][statusIdx] || 'AKTIF').trim().toUpperCase();
      if (uEmail === targetEmail) {
        if (isDeleted === 'TRUE' || status === 'NONAKTIF' || status === 'INACTIVE' || status === 'PENDING' || status === 'PENDING_VERIFIKASI' || status === 'FALSE' || status === 'DITOLAK') {
          // Akun tidak aktif / pending: jangan ubah password & jangan kirim email, tapi kembalikan pesan generik
          return sanitizeForRPC({ success: true, message: GENERIC_MSG });
        }
        userRow = i;
        userName = String(data[i][namaIdx] || 'Pengguna').trim();
        userId = String(data[i][0] || '').trim();
        break;
      }
    }

    if (userRow === -1) {
      return sanitizeForRPC({ success: true, message: GENERIC_MSG });
    }

    const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
    let newPassword = 'HD-';
    for (let c = 0; c < 5; c++) {
      newPassword += chars.charAt(Math.floor(Math.random() * chars.length));
    }

    const hashedNewPass = hashPasswordV2(newPassword);
    userSheet.getRange(userRow + 1, passIdx + 1).setValue(hashedNewPass);

    const subject = "[Happy Drawing] Password Baru Akun Anda";
    const bodyText = "Halo " + userName + ",\n\n" +
      "Kami telah menerima permintaan reset password untuk akun Anda dengan email: " + targetEmail + ".\n\n" +
      "Berikut adalah Password Baru Anda:\n" +
      "Password: " + newPassword + "\n\n" +
      "Silakan gunakan password baru ini untuk login ke aplikasi Happy Drawing.\n\n" +
      "Salam,\nHappy Drawing Management";

    const htmlBody = `
      <div style="font-family: 'Poppins', Arial, sans-serif; max-width: 520px; margin: 0 auto; padding: 25px; border: 1px solid #e2e8f0; border-radius: 16px; background-color: #ffffff;">
        <div style="text-align: center; margin-bottom: 20px;">
          <h2 style="color: #0d6efd; margin: 0; font-size: 24px;">Happy Drawing</h2>
          <p style="color: #64748b; font-size: 13px; margin-top: 4px;">Belajar Bermain Berkarya Penuh Warna</p>
        </div>
        <p style="color: #1e293b; font-size: 15px; font-weight: 600;">Halo ${userName},</p>
        <p style="color: #475569; font-size: 14px;">Berikut adalah password baru untuk akun Portal Siswa Anda:</p>
        <div style="background: #f0f7ff; border: 2px dashed #0d6efd; border-radius: 12px; padding: 18px; text-align: center; margin: 25px 0;">
          <strong style="font-size: 26px; color: #0d6efd; letter-spacing: 3px;">${newPassword}</strong>
        </div>
      </div>
    `;

    kirimEmailGmailApi_({
      to: targetEmail,
      subject: subject,
      body: bodyText,
      htmlBody: htmlBody
    });

    logActivity(userId, 'RESET_PASSWORD', 'AUTH', 'Password direset dan dikirim via Email ke: ' + targetEmail);

    return sanitizeForRPC({
      success: true,
      message: GENERIC_MSG
    });

  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal memproses reset password: " + err.toString() });
  }
}

function fetchTableData(sheetName, authToken, allowNoAuth) {
  try {
    const targetSheet = String(sheetName || '').toUpperCase();
    if (!allowNoAuth) {
      const ADMIN_SHEETS = ['KEUANGAN', 'PENGELUARAN', 'LOG', 'USERS', 'USER', 'DAFTAR', 'BIODATA', 'ABSENSI', 'ROLE', 'GALERI'];
      const requiredRoles = ADMIN_SHEETS.includes(targetSheet) ? ['ADMIN', 'SUPER_ADMIN'] : null;
      const auth = requireAuth_(authToken, requiredRoles);
      if (!auth.authorized) return auth.error;
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet && (sheetName === 'USER' || sheetName === 'USERS')) {
      sheet = ss.getSheetByName('USER') || ss.getSheetByName('USERS');
    }
    if (!sheet && targetSheet === 'PENGELUARAN') {
      sheet = ensurePengeluaranSheet_();
    }
    if (!sheet) return sanitizeForRPC({ success: false, data: [] });

    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return sanitizeForRPC({ success: true, data: [] });

    const headers = values[0];
    const headersClean = headers.map(h => String(h || '').toUpperCase().replace(/[^A-Z0-9]/g, ''));
    let isDeletedIdx = headersClean.indexOf('ISDELETED');

    const result = [];

    for (let i = 1; i < values.length; i++) {
      let row = values[i];
      let isDeleted = isDeletedIdx > -1 ? String(row[isDeletedIdx] || '').toUpperCase() : 'FALSE';

      if (isDeleted !== 'TRUE') {
        let item = {};
        headers.forEach((h, idx) => {
          const hStr = String(h || '');
          if (/pass|pwd|sandi/i.test(hStr)) return;
          let val = row[idx];
          
          if (val instanceof Date) {
            const yyyy = val.getFullYear();
            const mm = String(val.getMonth() + 1).padStart(2, '0');
            const dd = String(val.getDate()).padStart(2, '0');
            const hours = String(val.getHours()).padStart(2, '0');
            const mins = String(val.getMinutes()).padStart(2, '0');
            const secs = String(val.getSeconds()).padStart(2, '0');
            
            if (hours === '00' && mins === '00' && secs === '00') {
              val = dd + '/' + mm + '/' + yyyy;
            } else {
              val = dd + '/' + mm + '/' + yyyy + ' ' + hours + ':' + mins + ':' + secs;
            }
          }
          
          item[hStr] = (val !== undefined && val !== null) ? String(val) : '';
        });
        result.push(item);
      }
    }

    return sanitizeForRPC({ success: true, data: result });
  } catch (error) {
    return sanitizeForRPC({ success: false, message: error.toString(), data: [] });
  }
}

function fetchMasterSiswa(authToken) {
  try {
    const auth = requireAuth_(authToken, ['ADMIN', 'SUPER_ADMIN']);
    if (!auth.authorized) return auth.error;

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName('BIODATA');
    if (!sheet) return sanitizeForRPC({ success: false, data: [] });

    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return sanitizeForRPC({ success: true, data: [] });

    const headers = values[0];
    const result = [];

    for (let i = 1; i < values.length; i++) {
      let row = values[i];
      let item = {};
      headers.forEach((h, idx) => {
        const hStr = String(h || '');
        if (/pass|pwd|sandi/i.test(hStr)) return;
        let val = row[idx];
        if (val instanceof Date) {
          const yyyy = val.getFullYear();
          const mm = String(val.getMonth() + 1).padStart(2, '0');
          const dd = String(val.getDate()).padStart(2, '0');
          const hours = String(val.getHours()).padStart(2, '0');
          const mins = String(val.getMinutes()).padStart(2, '0');
          const secs = String(val.getSeconds()).padStart(2, '0');
          
          if (hours === '00' && mins === '00' && secs === '00') {
            val = dd + '/' + mm + '/' + yyyy;
          } else {
            val = dd + '/' + mm + '/' + yyyy + ' ' + hours + ':' + mins + ':' + secs;
          }
        }
        item[String(h)] = (val !== undefined && val !== null) ? String(val) : '';
      });
      result.push(item);
    }

    return sanitizeForRPC({ success: true, data: result });
  } catch (error) {
    return sanitizeForRPC({ success: false, message: error.toString(), data: [] });
  }
}

function fetchBiodata(authToken) {
  return fetchMasterSiswa(authToken);
}

function fetchUsersAdmin(authToken) {
  try {
    const auth = requireAuth_(authToken, ['SUPER_ADMIN']);
    if (!auth.authorized) return auth.error;

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName('USER') || ss.getSheetByName('USERS');
    if (!sheet) return sanitizeForRPC({ success: false, data: [] });

    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return sanitizeForRPC({ success: true, data: [] });

    const headers = values[0];
    const result = [];

    for (let i = 1; i < values.length; i++) {
      let row = values[i];
      let item = {};
      headers.forEach((h, idx) => {
        const hStr = String(h || '').trim().toLowerCase();
        // Cegah kebocoran kolom password / hash sandi ke sisi klien
        if (hStr.includes('pass') || hStr.includes('pwd') || hStr.includes('sandi')) {
          return;
        }
        let val = row[idx];
        if (val instanceof Date) {
          const yyyy = val.getFullYear();
          const mm = String(val.getMonth() + 1).padStart(2, '0');
          const dd = String(val.getDate()).padStart(2, '0');
          const hours = String(val.getHours()).padStart(2, '0');
          const mins = String(val.getMinutes()).padStart(2, '0');
          const secs = String(val.getSeconds()).padStart(2, '0');
          
          if (hours === '00' && mins === '00' && secs === '00') {
            val = dd + '/' + mm + '/' + yyyy;
          } else {
            val = dd + '/' + mm + '/' + yyyy + ' ' + hours + ':' + mins + ':' + secs;
          }
        }
        item[String(h)] = (val !== undefined && val !== null) ? String(val) : '';
      });
      result.push(item);
    }

    return sanitizeForRPC({ success: true, data: result });
  } catch (error) {
    return sanitizeForRPC({ success: false, message: error.toString(), data: [] });
  }
}

function softDeleteSiswa(idSiswa, authToken) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(15000)) {
      return sanitizeForRPC({ success: false, message: "Sistem sedang sibuk memproses data lain. Harap coba lagi beberapa saat lagi." });
    }

    const auth = requireAuth_(authToken, ['ADMIN', 'SUPER_ADMIN']);
    if (!auth.authorized) return auth.error;

    if (String(auth.session.role || '').toUpperCase() === 'ADMIN' && !canEditModule_(auth.session, 'master')) {
      return sanitizeForRPC({ success: false, message: "Hak akses Anda untuk modul Master Siswa hanya baca (Readonly)." });
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const bioSheet = ss.getSheetByName('BIODATA');
    if (!bioSheet) return sanitizeForRPC({ success: false, message: "Sheet BIODATA tidak ditemukan!" });

    const bioData = bioSheet.getDataRange().getValues();
    if (bioData.length <= 1) return sanitizeForRPC({ success: false, message: "Data BIODATA kosong." });

    const headersClean = bioData[0].map(h => String(h || '').toUpperCase().replace(/[^A-Z0-9]/g, ''));
    let isDeletedIdx = headersClean.indexOf('ISDELETED');
    if (isDeletedIdx === -1) isDeletedIdx = headersClean.indexOf('IS_DELETED');
    if (isDeletedIdx === -1) {
      isDeletedIdx = headersClean.length;
      bioSheet.getRange(1, isDeletedIdx + 1).setValue('IS DELETED');
    }

    let emailIdxBio = headersClean.findIndex(h => h === 'EMAIL' || h === 'SUREL');
    const targetId = String(idSiswa || '').trim().toLowerCase();

    for (let b = bioData.length - 1; b >= 1; b--) {
      let isMatch = false;
      for (let c = 0; c < bioData[b].length; c++) {
        let cellVal = String(bioData[b][c] || '').trim().toLowerCase();
        if (cellVal !== '' && cellVal === targetId) {
          isMatch = true;
          break;
        }
      }
      if (isMatch) {
        bioSheet.getRange(b + 1, isDeletedIdx + 1).setValue('TRUE');

        let studentEmail = (emailIdxBio > -1) ? String(bioData[b][emailIdxBio] || '').trim().toLowerCase() : '';

        const userSheet = ss.getSheetByName('USER') || ss.getSheetByName('USERS');
        if (userSheet && userSheet.getLastRow() > 1) {
          const uData = userSheet.getDataRange().getValues();
          const uHeadersClean = uData[0].map(h => String(h || '').toLowerCase().replace(/[^a-z0-9]/g, ''));
          let uIdIdx = uHeadersClean.findIndex(h => h.includes('id'));
          let uEmailIdx = uHeadersClean.findIndex(h => h === 'email' || h === 'surel');
          let uDelIdx = uHeadersClean.findIndex(h => h.includes('deleted') || h.includes('hapus'));
          let uStatusIdx = uHeadersClean.findIndex(h => h.includes('status'));

          for (let u = 1; u < uData.length; u++) {
            let uId = uIdIdx > -1 ? String(uData[u][uIdIdx] || '').trim().toLowerCase() : '';
            let uEmail = uEmailIdx > -1 ? String(uData[u][uEmailIdx] || '').trim().toLowerCase() : '';
            
            let isUserMatch = (uId !== '' && uId === targetId) || 
                              (uEmail !== '' && uEmail === targetId) || 
                              (studentEmail !== '' && studentEmail !== '-' && uEmail === studentEmail);

            if (isUserMatch) {
              if (uDelIdx > -1) userSheet.getRange(u + 1, uDelIdx + 1).setValue('TRUE');
              if (uStatusIdx > -1) userSheet.getRange(u + 1, uStatusIdx + 1).setValue('NONAKTIF');
              if (uId) revokeUserSessions_(uId);
              if (uEmail) revokeUserSessions_(uEmail);
            }
          }
        }

        if (targetId) revokeUserSessions_(targetId);
        if (studentEmail && studentEmail !== '-') revokeUserSessions_(studentEmail);

        const actorId = auth.session ? (auth.session.uid || auth.session.email || 'ADMIN') : 'ADMIN';
        logActivity(actorId, 'SOFT_DELETE_SISWA', 'BIODATA', 'Nonaktifkan/Soft delete siswa: ' + idSiswa);
        return sanitizeForRPC({ success: true, message: "Siswa dengan ID " + idSiswa + " berhasil dinonaktifkan (IS DELETED = 'TRUE')." });
      }
    }

    return sanitizeForRPC({ success: false, message: "ID Siswa " + idSiswa + " tidak ditemukan di Sheet BIODATA." });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal menonaktifkan siswa: " + err.toString() });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function reactivateSiswa(idSiswa, authToken) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(15000)) {
      return sanitizeForRPC({ success: false, message: "Sistem sedang sibuk memproses data lain. Harap coba lagi beberapa saat lagi." });
    }

    const auth = requireAuth_(authToken, ['ADMIN', 'SUPER_ADMIN']);
    if (!auth.authorized) return auth.error;

    if (String(auth.session.role || '').toUpperCase() === 'ADMIN' && !canEditModule_(auth.session, 'master')) {
      return sanitizeForRPC({ success: false, message: "Hak akses Anda untuk modul Master Siswa hanya baca (Readonly)." });
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const bioSheet = ss.getSheetByName('BIODATA');
    if (!bioSheet) return sanitizeForRPC({ success: false, message: "Sheet BIODATA tidak ditemukan!" });

    const bioData = bioSheet.getDataRange().getValues();
    if (bioData.length <= 1) return sanitizeForRPC({ success: false, message: "Data BIODATA kosong." });

    const headersClean = bioData[0].map(h => String(h || '').toUpperCase().replace(/[^A-Z0-9]/g, ''));
    let isDeletedIdx = headersClean.indexOf('ISDELETED');
    if (isDeletedIdx === -1) isDeletedIdx = headersClean.indexOf('IS_DELETED');

    let emailIdxBio = headersClean.findIndex(h => h === 'EMAIL' || h === 'SUREL');
    const targetId = String(idSiswa || '').trim().toLowerCase();

    for (let b = bioData.length - 1; b >= 1; b--) {
      let isMatch = false;
      for (let c = 0; c < bioData[b].length; c++) {
        let cellVal = String(bioData[b][c] || '').trim().toLowerCase();
        if (cellVal !== '' && cellVal === targetId) {
          isMatch = true;
          break;
        }
      }
      if (isMatch) {
        if (isDeletedIdx > -1) {
          bioSheet.getRange(b + 1, isDeletedIdx + 1).setValue('FALSE');
        }

        let studentEmail = (emailIdxBio > -1) ? String(bioData[b][emailIdxBio] || '').trim().toLowerCase() : '';

        const userSheet = ss.getSheetByName('USER') || ss.getSheetByName('USERS');
        if (userSheet && userSheet.getLastRow() > 1) {
          const uData = userSheet.getDataRange().getValues();
          const uHeadersClean = uData[0].map(h => String(h || '').toLowerCase().replace(/[^a-z0-9]/g, ''));
          let uIdIdx = uHeadersClean.findIndex(h => h.includes('id'));
          let uEmailIdx = uHeadersClean.findIndex(h => h === 'email' || h === 'surel');
          let uDelIdx = uHeadersClean.findIndex(h => h.includes('deleted') || h.includes('hapus'));
          let uStatusIdx = uHeadersClean.findIndex(h => h.includes('status'));

          for (let u = 1; u < uData.length; u++) {
            let uId = uIdIdx > -1 ? String(uData[u][uIdIdx] || '').trim().toLowerCase() : '';
            let uEmail = uEmailIdx > -1 ? String(uData[u][uEmailIdx] || '').trim().toLowerCase() : '';
            
            let isUserMatch = (uId !== '' && uId === targetId) || 
                              (uEmail !== '' && uEmail === targetId) || 
                              (studentEmail !== '' && studentEmail !== '-' && uEmail === studentEmail);

            if (isUserMatch) {
              if (uDelIdx > -1) userSheet.getRange(u + 1, uDelIdx + 1).setValue('FALSE');
              if (uStatusIdx > -1) userSheet.getRange(u + 1, uStatusIdx + 1).setValue('AKTIF');
            }
          }
        }

        const actorId = auth.session ? (auth.session.uid || auth.session.email || 'ADMIN') : 'ADMIN';
        logActivity(actorId, 'REACTIVATE_SISWA', 'BIODATA', 'Mengaktifkan kembali siswa: ' + idSiswa);
        return sanitizeForRPC({ success: true, message: "Siswa dengan ID " + idSiswa + " berhasil diaktifkan kembali." });
      }
    }

    return sanitizeForRPC({ success: false, message: "ID Siswa " + idSiswa + " tidak ditemukan di Sheet BIODATA." });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal mengaktifkan kembali siswa: " + err.toString() });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * Mengirim email ke siswa dari menu Master Siswa (Akses Admin / Super Admin)
 * Alamat email diambil langsung dari Sheet BIODATA berdasarkan idSiswa untuk mencegah abuse.
 */
function sendEmailSiswa(payload, authToken) {
  try {
    const auth = requireAuth_(authToken, ['ADMIN', 'SUPER_ADMIN']);
    if (!auth.authorized) return auth.error;

    if (String(auth.session.role || '').toUpperCase() === 'ADMIN' && !canEditModule_(auth.session, 'master')) {
      return sanitizeForRPC({ success: false, message: "Hak akses Anda untuk modul Master Siswa hanya baca (Readonly)." });
    }

    if (!payload || typeof payload !== 'object') {
      return sanitizeForRPC({ success: false, message: "Payload tidak valid." });
    }

    const idSiswa = String(payload.idSiswa || payload.id || '').trim();
    if (!idSiswa) {
      return sanitizeForRPC({ success: false, message: "ID Siswa wajib disertakan." });
    }

    const subjek = String(payload.subjek || payload.subject || '').trim();
    if (!subjek) {
      return sanitizeForRPC({ success: false, message: "Subjek email wajib diisi." });
    }
    if (subjek.length > 150) {
      return sanitizeForRPC({ success: false, message: "Subjek email terlalu panjang (maksimal 150 karakter)." });
    }

    const pesan = String(payload.pesan || payload.message || payload.body || '').trim();
    if (!pesan) {
      return sanitizeForRPC({ success: false, message: "Isi pesan email wajib diisi." });
    }
    if (pesan.length > 5000) {
      return sanitizeForRPC({ success: false, message: "Isi pesan terlalu panjang (maksimal 5000 karakter)." });
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const bioSheet = ss.getSheetByName('BIODATA');
    if (!bioSheet) return sanitizeForRPC({ success: false, message: "Sheet BIODATA tidak ditemukan!" });

    const bioData = bioSheet.getDataRange().getValues();
    if (bioData.length <= 1) return sanitizeForRPC({ success: false, message: "Data BIODATA kosong." });

    const headersClean = bioData[0].map(h => String(h || '').toUpperCase().replace(/[^A-Z0-9]/g, ''));
    const idIdx = headersClean.findIndex(h => h === 'IDSISWA' || h === 'ID' || h === 'NOID');
    const emailIdx = headersClean.findIndex(h => h === 'EMAIL' || h === 'SUREL');
    const namaIdx = headersClean.findIndex(h => h.includes('NAMA'));

    if (idIdx === -1 || emailIdx === -1) {
      return sanitizeForRPC({ success: false, message: "Struktur Sheet BIODATA tidak lengkap (kolom ID atau EMAIL tidak ditemukan)." });
    }

    const targetId = idSiswa.toLowerCase();
    let targetEmail = '';
    let targetNama = 'Siswa';

    for (let r = 1; r < bioData.length; r++) {
      let rowId = String(bioData[r][idIdx] || '').trim().toLowerCase();
      if (rowId === targetId) {
        targetEmail = String(bioData[r][emailIdx] || '').trim().toLowerCase();
        if (namaIdx > -1) targetNama = String(bioData[r][namaIdx] || 'Siswa').trim();
        break;
      }
    }

    if (!targetEmail || targetEmail.indexOf('@') === -1) {
      return sanitizeForRPC({ success: false, message: "Alamat email untuk siswa ID '" + idSiswa + "' tidak ditemukan atau tidak valid di Sheet BIODATA." });
    }

    const safeSubject = "[Happy Drawing] " + subjek.replace(/^\[Happy Drawing\]\s*/i, '');
    const safeBodyHtml = pesan
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/\n/g, '<br>');

    const emailHtmlBody = `
      <div style="font-family: 'Poppins', Arial, sans-serif; max-width: 540px; margin: 0 auto; padding: 28px; border: 1px solid #e2e8f0; border-radius: 16px; background-color: #ffffff;">
        <div style="text-align: center; margin-bottom: 24px;">
          <h2 style="color: #0d6efd; margin: 0; font-size: 24px; font-weight: 700;">Happy Drawing</h2>
          <p style="color: #64748b; font-size: 13px; margin-top: 4px;">Belajar Bermain Berkarya Penuh Warna</p>
        </div>
        <div style="background: #f0f7ff; border: 1px solid #d0e7ff; border-radius: 12px; padding: 14px 18px; margin-bottom: 20px;">
          <span style="color: #0d6efd; font-weight: 700; font-size: 14px;">Pemberitahuan untuk: ${targetNama} (ID: ${idSiswa})</span>
        </div>
        <div style="color: #1e293b; font-size: 14px; line-height: 1.7; margin-bottom: 24px;">
          ${safeBodyHtml}
        </div>
        <div style="text-align: center; margin: 25px 0 10px 0; border-top: 1px solid #f1f5f9; padding-top: 15px;">
          <p style="color: #94a3b8; font-size: 12px; margin: 0;">Email ini dikirim oleh Admin Sanggar Happy Drawing.</p>
        </div>
      </div>
    `;

    kirimEmailGmailApi_({
      to: targetEmail,
      subject: safeSubject,
      body: pesan,
      htmlBody: emailHtmlBody
    });

    const actorId = (auth.session && (auth.session.uid || auth.session.email)) || 'ADMIN';
    logActivity(actorId, 'SEND_EMAIL_SISWA', 'BIODATA', 'Kirim email ke ' + targetEmail + ' (Siswa: ' + targetNama + ', ID: ' + idSiswa + ')');

    return sanitizeForRPC({
      success: true,
      message: "Email berhasil dikirim ke " + targetEmail + "!"
    });
  } catch (err) {
    return sanitizeForRPC({
      success: false,
      message: "Gagal mengirim email: " + err.toString()
    });
  }
}

function getKeuanganCacheVer() {
  try {
    const v = CacheService.getScriptCache().get('KEU_CACHE_VER');
    return v ? parseInt(v, 10) || 1 : 1;
  } catch (e) {
    return 1;
  }
}

function bumpKeuanganCacheVer() {
  try {
    const cur = getKeuanganCacheVer();
    const nextVer = (cur + 1) % 1000000;
    CacheService.getScriptCache().put('KEU_CACHE_VER', String(nextVer), 21600);
  } catch (e) {}
}

function fetchKeuanganSiswa(idSiswa, authToken) {
  try {
    const auth = requireAuth_(authToken, null);
    if (!auth.authorized) return auth.error;
    if (!requireOwnership_(auth.session, idSiswa)) {
      return sanitizeForRPC({ success: false, message: "Akses data keuangan siswa lain tidak diizinkan." });
    }

    const targetId = String(idSiswa || '').trim().toLowerCase();
    if (!targetId) {
      return sanitizeForRPC({ success: true, data: [] });
    }

    // 1. Cek server-side cache (TTL 60 detik)
    const ver = getKeuanganCacheVer();
    const cacheKey = 'KEU_V' + ver + '_' + targetId.replace(/[^a-z0-9_-]/gi, '_').slice(0, 50);
    try {
      const cached = CacheService.getScriptCache().get(cacheKey);
      if (cached) {
        const parsed = JSON.parse(cached);
        if (Array.isArray(parsed)) {
          return sanitizeForRPC({ success: true, data: parsed, fromCache: true });
        }
      }
    } catch (eCache) {}

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let keuSheet = ss.getSheetByName('KEUANGAN');
    if (!keuSheet || keuSheet.getLastRow() <= 1) {
      return sanitizeForRPC({ success: true, data: [] });
    }

    const lastRow = keuSheet.getLastRow();
    const lastCol = keuSheet.getLastColumn();
    if (lastCol <= 0) {
      return sanitizeForRPC({ success: true, data: [] });
    }

    // Baca hanya range data yang ada (hindari pembacaan kolom kosong di luar used-range)
    const values = keuSheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headersClean = values[0].map(h => String(h || '').trim().toUpperCase());

    let idSiswaIdx = headersClean.indexOf('ID SISWA');
    if (idSiswaIdx === -1) idSiswaIdx = 2;

    let isDeletedIdx = headersClean.indexOf('IS DELETED');
    if (isDeletedIdx === -1) isDeletedIdx = headersClean.length - 1;

    const result = [];

    for (let i = 1; i < values.length; i++) {
      let row = values[i];
      let rowIdSiswa = String(row[idSiswaIdx] || '').trim().toLowerCase();
      let isDeleted = String(row[isDeletedIdx] || 'FALSE').trim().toUpperCase();

      if (rowIdSiswa === targetId && isDeleted !== 'TRUE') {
        let item = {};
        headersClean.forEach((h, idx) => {
          let val = row[idx];
          if (val instanceof Date) {
            const yyyy = val.getFullYear();
            const mm = String(val.getMonth() + 1).padStart(2, '0');
            const dd = String(val.getDate()).padStart(2, '0');
            val = dd + '/' + mm + '/' + yyyy;
          }
          item[String(h)] = (val !== undefined && val !== null) ? String(val) : '';
        });
        result.push(item);
      }
    }

    // Simpan ke CacheService server (TTL 60 detik)
    try {
      CacheService.getScriptCache().put(cacheKey, JSON.stringify(result), 60);
    } catch (eSave) {}

    return sanitizeForRPC({ success: true, data: result });
  } catch (error) {
    return sanitizeForRPC({ success: false, message: error.toString(), data: [] });
  }
}

function submitBuktiPembayaran(payload, authToken) {
  try {
    const token = authToken || (payload && (payload.__token || payload.token));
    const auth = requireAuth_(token, null);
    if (!auth.authorized) return auth.error;

    if (!payload || typeof payload !== 'object') {
      return sanitizeForRPC({ success: false, message: "Payload pembayaran tidak valid." });
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const idSiswa = String(payload.idSiswa || '').trim();
    if (!idSiswa) {
      return sanitizeForRPC({ success: false, message: "ID Siswa wajib diisi." });
    }

    if (!requireOwnership_(auth.session, idSiswa)) {
      return sanitizeForRPC({ success: false, message: "Tidak diizinkan mengirim bukti pembayaran untuk siswa lain." });
    }

    // 1. Validasi Keberadaan Siswa di Sheet BIODATA
    let bioSheet = ss.getSheetByName('BIODATA');
    let studentFound = false;
    let studentNama = String(payload.namaSiswa || '').trim();
    let studentKelas = String(payload.kelas || '').trim();
    let studentKategori = String(payload.kategori || '').trim();

    if (bioSheet && bioSheet.getLastRow() > 1) {
      const bioData = bioSheet.getDataRange().getValues();
      const bioHeaders = bioData[0].map(h => String(h || '').toUpperCase().trim());
      let idCol = bioHeaders.findIndex(h => h === 'ID SISWA' || h === 'IDSISWA' || h === 'ID');
      let delCol = bioHeaders.indexOf('IS DELETED');
      let namaCol = bioHeaders.indexOf('NAMA LENGKAP') > -1 ? bioHeaders.indexOf('NAMA LENGKAP') : bioHeaders.indexOf('NAMA');
      let kelasCol = bioHeaders.indexOf('KELAS');
      let katCol = bioHeaders.indexOf('KATEGORI');

      for (let b = 1; b < bioData.length; b++) {
        let rowId = String(bioData[b][idCol] || '').trim().toLowerCase();
        let isDel = delCol > -1 ? String(bioData[b][delCol] || '').trim().toUpperCase() : 'FALSE';
        if (rowId === idSiswa.toLowerCase() && isDel !== 'TRUE') {
          studentFound = true;
          if (!studentNama && namaCol > -1) studentNama = String(bioData[b][namaCol] || '').trim();
          if (!studentKelas && kelasCol > -1) studentKelas = String(bioData[b][kelasCol] || '').trim();
          if (!studentKategori && katCol > -1) studentKategori = String(bioData[b][katCol] || '').trim();
          break;
        }
      }
    }

    if (!studentFound) {
      return sanitizeForRPC({ success: false, message: "ID Siswa tidak ditemukan atau status siswa sudah nonaktif." });
    }

    // 2. Validasi Nominal Pembayaran (harus berupa angka positif)
    let rawNom = payload.nominal;
    let nominal = typeof rawNom === 'number' ? rawNom : parseInt(String(rawNom || '').replace(/[^0-9]/g, ''), 10);
    if (isNaN(nominal) || nominal <= 0 || nominal > 100000000) {
      return sanitizeForRPC({ success: false, message: "Nominal pembayaran tidak valid (harus berupa angka antara Rp 1 s.d. Rp 100.000.000)." });
    }

    // 3. Validasi Jenis Transaksi & Periode
    let jenis = String(payload.jenisTransaksi || 'INFAK BULANAN').trim();
    if (!jenis || jenis.length > 60) {
      return sanitizeForRPC({ success: false, message: "Jenis transaksi wajib diisi dan maksimal 60 karakter." });
    }

    const now = new Date();
    const yyyy = now.getFullYear();
    const mm = String(now.getMonth() + 1).padStart(2, '0');
    const dd = String(now.getDate()).padStart(2, '0');

    const monthsList = [
      'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
      'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
    ];
    let periodeBulan = String(payload.periodeBulan || monthsList[now.getMonth()]).trim();
    let periodeTahun = parseInt(payload.periodeTahun, 10);
    if (isNaN(periodeTahun) || periodeTahun < 2020 || periodeTahun > 2099) {
      periodeTahun = yyyy;
    }

    let keuSheet = ss.getSheetByName('KEUANGAN');
    if (!keuSheet) {
      keuSheet = ss.insertSheet('KEUANGAN');
      keuSheet.appendRow([
        'ID TRANSAKSI', 'NO INVOICE', 'ID SISWA', 'NAMA SISWA', 'KELAS', 'KATEGORI',
        'JENIS TRANSAKSI', 'PERIODE BULAN', 'PERIODE TAHUN', 'NOMINAL', 'TANGGAL BAYAR', 
        'METODE PEMBAYARAN', 'BUKTI TRANSFER', 'STATUS PEMBAYARAN', 'KETERANGAN', 'VERIFIKATOR', 
        'WAKTU VERIFIKASI', 'IS DELETED', 'CATATAN ADMIN'
      ]);
    }

    const kData = keuSheet.getDataRange().getValues();
    const kHeaders = kData[0].map(h => String(h || '').toUpperCase().trim());

    // 4. Validasi Anti-Duplikat (kecuali jenis donasi / sukarela)
    if (!jenis.toUpperCase().includes('SUKARELA') && kData.length > 1) {
      let idSiswaCol = kHeaders.indexOf('ID SISWA');
      let jenisCol = kHeaders.indexOf('JENIS TRANSAKSI');
      let bulanCol = kHeaders.indexOf('PERIODE BULAN');
      let tahunCol = kHeaders.indexOf('PERIODE TAHUN');
      let statusCol = kHeaders.indexOf('STATUS PEMBAYARAN');
      let delCol = kHeaders.indexOf('IS DELETED');

      for (let r = 1; r < kData.length; r++) {
        let rowId = String(kData[r][idSiswaCol] || '').trim().toLowerCase();
        let rowJenis = String(kData[r][jenisCol] || '').trim().toUpperCase();
        let rowBulan = String(kData[r][bulanCol] || '').trim().toLowerCase();
        let rowTahun = String(kData[r][tahunCol] || '').trim();
        let rowStatus = String(kData[r][statusCol] || '').trim().toUpperCase();
        let rowDel = delCol > -1 ? String(kData[r][delCol] || '').trim().toUpperCase() : 'FALSE';

        if (rowDel === 'TRUE') continue;

        if (rowId === idSiswa.toLowerCase() &&
            rowJenis === jenis.toUpperCase() &&
            rowBulan === periodeBulan.toLowerCase() &&
            rowTahun === String(periodeTahun)) {
          if (rowStatus === 'LUNAS') {
            return sanitizeForRPC({
              success: false,
              message: "Pembayaran untuk " + jenis + " periode " + periodeBulan + " " + periodeTahun + " sudah berstatus LUNAS."
            });
          }
          if (rowStatus.includes('PENDING')) {
            return sanitizeForRPC({
              success: false,
              message: "Transaksi untuk " + jenis + " periode " + periodeBulan + " " + periodeTahun + " sedang dalam proses PENDING VERIFIKASI. Harap tunggu verifikasi admin."
            });
          }
        }
      }
    }

    // 5. Validasi & Unggah Foto Bukti Transfer (Wajib & Ukuran <= 2MB)
    if (!payload.buktiBase64) {
      return sanitizeForRPC({ success: false, message: "Foto bukti transfer wajib diunggah." });
    }

    let cleanBase64 = String(payload.buktiBase64);
    if (cleanBase64.indexOf(',') !== -1) {
      cleanBase64 = cleanBase64.split(',')[1];
    }
    cleanBase64 = cleanBase64.replace(/[^A-Za-z0-9+/=]/g, "");
    if (!cleanBase64) {
      return sanitizeForRPC({ success: false, message: "Data foto bukti transfer kosong atau rusak." });
    }

    let approxBytes = Math.ceil(cleanBase64.length * 3 / 4);
    if (approxBytes > 2202009) {
      return sanitizeForRPC({ success: false, message: "Ukuran file bukti transfer melebihi batas 2.0 MB." });
    }

    let decodedBytes = Utilities.base64Decode(cleanBase64);
    if (!decodedBytes || decodedBytes.length === 0) {
      return sanitizeForRPC({ success: false, message: "Gagal memproses data gambar bukti transfer." });
    }

    // Cek magic bytes: JPEG (FF D8) atau PNG (89 50 4E 47)
    let isJpeg = decodedBytes.length > 2 && (decodedBytes[0] & 0xFF) === 0xFF && (decodedBytes[1] & 0xFF) === 0xD8;
    let isPng = decodedBytes.length > 4 && (decodedBytes[0] & 0xFF) === 0x89 && (decodedBytes[1] & 0xFF) === 0x50 && (decodedBytes[2] & 0xFF) === 0x4E && (decodedBytes[3] & 0xFF) === 0x47;
    if (!isJpeg && !isPng) {
      return sanitizeForRPC({ success: false, message: "Format gambar tidak didukung. Harap unggah foto bukti dalam format JPG atau PNG." });
    }

    let ext = isPng ? ".png" : ".jpg";
    let mime = isPng ? "image/png" : "image/jpeg";

    let folderConfig = null;
    try {
      folderConfig = DriveApp.getFolderById(CONFIG.FOLDER_BUKTI_ID);
    } catch (fErr) {
      return sanitizeForRPC({ success: false, message: "Gagal mengakses folder BUKTI TRANSFER di Google Drive: " + fErr.message });
    }

    if (!folderConfig) {
      return sanitizeForRPC({ success: false, message: "Folder BUKTI TRANSFER tidak ditemukan di Google Drive." });
    }

    let fileName = "BUKTI_" + idSiswa + "_" + new Date().getTime() + ext;
    let blob = Utilities.newBlob(decodedBytes, mime, fileName);
    let file = folderConfig.createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    let fileUrl = "https://lh3.googleusercontent.com/d/" + file.getId();

    // 6. Generate ID Transaksi & No Invoice Unik (Anti Tabrakan)
    let idTrxIdx = kHeaders.indexOf('ID TRANSAKSI');
    if (idTrxIdx === -1) idTrxIdx = 0;
    let existingTrxIds = {};
    for (let r = 1; r < kData.length; r++) {
      let tId = String(kData[r][idTrxIdx] || '').trim();
      if (tId) existingTrxIds[tId] = true;
    }

    let candidateUrut = kData.length;
    let idTransaksi = "";
    let noInvoice = "";
    let tries = 0;
    while (tries < 100) {
      let urutStr = String(candidateUrut).padStart(3, '0');
      let testTrxId = "TRX-" + yyyy + mm + "-" + urutStr;
      if (!existingTrxIds[testTrxId]) {
        idTransaksi = testTrxId;
        noInvoice = "INV/HD/" + yyyy + "/" + mm + "/" + urutStr;
        break;
      }
      candidateUrut++;
      tries++;
    }

    if (!idTransaksi) {
      let randomSuffix = String(Math.floor(Math.random() * 900) + 100);
      idTransaksi = "TRX-" + yyyy + mm + "-" + randomSuffix;
      noInvoice = "INV/HD/" + yyyy + "/" + mm + "/" + randomSuffix;
    }

    const tglBayar = dd + "/" + mm + "/" + yyyy;

    // 7. Simpan Transaksi ke Sheet KEUANGAN (Nominal tersimpan numerik)
    keuSheet.appendRow([
      idTransaksi,
      noInvoice,
      idSiswa,
      studentNama,
      studentKelas || '-',
      studentKategori || '-',
      jenis,
      periodeBulan,
      periodeTahun,
      nominal, // Nilai angka (Number)
      tglBayar,
      payload.metodePembayaran || 'TRANSFER BANK',
      fileUrl,
      'PENDING VERIFIKASI',
      payload.keterangan || 'Pembayaran Mandiri via Portal Siswa',
      '-',
      '-',
      'FALSE'
    ]);

    logActivity(idSiswa, 'UPLOAD_BUKTI_SPP', 'KEUANGAN', 'Upload bukti bayar: ' + idTransaksi + ' (' + jenis + ' ' + periodeBulan + ' ' + periodeTahun + ')');

    return sanitizeForRPC({
      success: true,
      message: "Bukti transfer berhasil diunggah! Status transaksi saat ini PENDING VERIFIKASI.",
      idTransaksi: idTransaksi,
      noInvoice: noInvoice
    });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal menyimpan pembayaran: " + err.toString() });
  }
}

function updateStatusPembayaran(idTransaksi, statusPembayaran, verifikator, waktuVerifikasi, catatanAdmin, authToken) {
  try {
    const auth = requireAuth_(authToken, ['ADMIN', 'SUPER_ADMIN']);
    if (!auth.authorized) return auth.error;

    if (String(auth.session.role || '').toUpperCase() === 'ADMIN' && !canEditModule_(auth.session, 'pembayaran')) {
      return sanitizeForRPC({ success: false, message: "Hak akses Anda untuk modul Pembayaran hanya baca (Readonly)." });
    }

    const allowedStatuses = ['LUNAS', 'DITOLAK', 'PENDING VERIFIKASI'];
    const normStatus = String(statusPembayaran || '').toUpperCase().trim();
    if (allowedStatuses.indexOf(normStatus) === -1) {
      return sanitizeForRPC({ success: false, message: "Status pembayaran tidak valid (hanya LUNAS, DITOLAK, atau PENDING VERIFIKASI)." });
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const keuSheet = ss.getSheetByName('KEUANGAN');
    if (!keuSheet) return sanitizeForRPC({ success: false, message: "Sheet KEUANGAN tidak ditemukan!" });

    const data = keuSheet.getDataRange().getValues();
    if (data.length <= 1) return sanitizeForRPC({ success: false, message: "Data KEUANGAN kosong." });

    const headersClean = data[0].map(h => String(h || '').toUpperCase().trim());
    let idIdx = headersClean.indexOf('ID TRANSAKSI');
    if (idIdx === -1) idIdx = 0;

    let statusIdx = headersClean.indexOf('STATUS PEMBAYARAN');
    if (statusIdx === -1) statusIdx = 13;

    let verifikatorIdx = headersClean.indexOf('VERIFIKATOR');
    if (verifikatorIdx === -1) verifikatorIdx = 15;

    let waktuIdx = headersClean.indexOf('WAKTU VERIFIKASI');
    if (waktuIdx === -1) waktuIdx = 16;

    let isDeletedIdx = headersClean.indexOf('IS DELETED');
    if (isDeletedIdx === -1) isDeletedIdx = headersClean.indexOf('IS_DELETED');

    let catatanIdx = headersClean.indexOf('CATATAN ADMIN');
    if (catatanIdx === -1) {
      catatanIdx = headersClean.length;
      keuSheet.getRange(1, catatanIdx + 1).setValue('CATATAN ADMIN');
    }

    const targetId = String(idTransaksi || '').trim().toLowerCase();
    const actorId = (auth.session && (auth.session.uid || auth.session.email)) || verifikator || 'ADMIN';
    const actorName = (auth.session && auth.session.nama) || verifikator || 'Admin Sanggar';

    const now = new Date();
    const pad = (n) => String(n).padStart(2, '0');
    const waktuVerifStr = `${pad(now.getDate())}/${pad(now.getMonth() + 1)}/${now.getFullYear()} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;

    for (let i = 1; i < data.length; i++) {
      let rowId = String(data[i][idIdx] || '').trim().toLowerCase();
      if (rowId === targetId) {
        if (isDeletedIdx > -1 && String(data[i][isDeletedIdx] || '').trim().toUpperCase() === 'TRUE') {
          return sanitizeForRPC({ success: false, message: "Transaksi " + idTransaksi + " sudah dihapus sebelumnya." });
        }
        keuSheet.getRange(i + 1, statusIdx + 1).setValue(normStatus);
        keuSheet.getRange(i + 1, verifikatorIdx + 1).setValue(actorName);
        keuSheet.getRange(i + 1, waktuIdx + 1).setValue(waktuVerifStr);
        const cleanCatatan = String(catatanAdmin || '').trim().slice(0, 500);
        keuSheet.getRange(i + 1, catatanIdx + 1).setValue(cleanCatatan);

        logActivity(actorId, 'VERIFIKASI_PEMBAYARAN', 'KEUANGAN', 'Ubah status pembayaran ' + idTransaksi + ' menjadi ' + normStatus + (cleanCatatan ? (' dengan catatan: ' + cleanCatatan) : ''));
        return sanitizeForRPC({ success: true, message: "Status pembayaran berhasil diubah menjadi " + normStatus });
      }
    }

    return sanitizeForRPC({ success: false, message: "ID Transaksi " + idTransaksi + " tidak ditemukan." });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal memperbarui status pembayaran: " + err.toString() });
  }
}

function cekDuplikatTagihan_(kData, kHeaders, idSiswa, jenis, periodeBulan, periodeTahun) {
  if (!jenis || jenis.toUpperCase().includes('SUKARELA') || kData.length <= 1) return null;
  let idSiswaCol = kHeaders.indexOf('ID SISWA');
  let jenisCol = kHeaders.indexOf('JENIS TRANSAKSI');
  let bulanCol = kHeaders.indexOf('PERIODE BULAN');
  let tahunCol = kHeaders.indexOf('PERIODE TAHUN');
  let statusCol = kHeaders.indexOf('STATUS PEMBAYARAN');
  let delCol = kHeaders.indexOf('IS DELETED');
  if (delCol === -1) delCol = kHeaders.indexOf('IS_DELETED');

  for (let r = 1; r < kData.length; r++) {
    let rowId = String(kData[r][idSiswaCol] || '').trim().toLowerCase();
    let rowJenis = String(kData[r][jenisCol] || '').trim().toUpperCase();
    let rowBulan = String(kData[r][bulanCol] || '').trim().toLowerCase();
    let rowTahun = String(kData[r][tahunCol] || '').trim();
    let rowStatus = String(kData[r][statusCol] || '').trim().toUpperCase();
    let rowDel = delCol > -1 ? String(kData[r][delCol] || '').trim().toUpperCase() : 'FALSE';

    if (rowDel === 'TRUE') continue;

    if (rowId === idSiswa.toLowerCase() &&
        rowJenis === jenis.toUpperCase() &&
        rowBulan === String(periodeBulan).toLowerCase() &&
        rowTahun === String(periodeTahun)) {
      if (rowStatus === 'LUNAS') {
        return "Pembayaran untuk " + jenis + " periode " + periodeBulan + " " + periodeTahun + " sudah berstatus LUNAS.";
      }
      if (rowStatus.includes('PENDING')) {
        return "Transaksi untuk " + jenis + " periode " + periodeBulan + " " + periodeTahun + " sedang dalam proses PENDING VERIFIKASI.";
      }
    }
  }
  return null;
}

function createTransaksiManual(payload, authToken) {
  try {
    const auth = requireAuth_(authToken, ['ADMIN', 'SUPER_ADMIN']);
    if (!auth.authorized) return auth.error;

    if (String(auth.session.role || '').toUpperCase() === 'ADMIN' && !canEditModule_(auth.session, 'pembayaran')) {
      return sanitizeForRPC({ success: false, message: "Hak akses Anda untuk modul Pembayaran hanya baca (Readonly)." });
    }

    if (!payload || typeof payload !== 'object') {
      return sanitizeForRPC({ success: false, message: "Payload transaksi manual tidak valid." });
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const idSiswa = String(payload.idSiswa || '').trim();
    if (!idSiswa) {
      return sanitizeForRPC({ success: false, message: "ID Siswa wajib dipilih." });
    }

    // 1. Validasi Siswa di BIODATA
    let bioSheet = ss.getSheetByName('BIODATA');
    let studentFound = false;
    let studentNama = String(payload.namaSiswa || '').trim();
    let studentKelas = String(payload.kelas || '').trim();
    let studentKategori = String(payload.kategori || '').trim();

    if (bioSheet && bioSheet.getLastRow() > 1) {
      const bioData = bioSheet.getDataRange().getValues();
      const bioHeaders = bioData[0].map(h => String(h || '').toUpperCase().trim());
      let idCol = bioHeaders.findIndex(h => h === 'ID SISWA' || h === 'IDSISWA' || h === 'ID');
      let delCol = bioHeaders.indexOf('IS DELETED');
      if (delCol === -1) delCol = bioHeaders.indexOf('IS_DELETED');
      let namaCol = bioHeaders.indexOf('NAMA LENGKAP') > -1 ? bioHeaders.indexOf('NAMA LENGKAP') : bioHeaders.indexOf('NAMA');
      let kelasCol = bioHeaders.indexOf('KELAS');
      let katCol = bioHeaders.indexOf('KATEGORI');

      for (let b = 1; b < bioData.length; b++) {
        let rowId = String(bioData[b][idCol] || '').trim().toLowerCase();
        let isDel = delCol > -1 ? String(bioData[b][delCol] || '').trim().toUpperCase() : 'FALSE';
        if (rowId === idSiswa.toLowerCase() && isDel !== 'TRUE') {
          studentFound = true;
          if (!studentNama && namaCol > -1) studentNama = String(bioData[b][namaCol] || '').trim();
          if (!studentKelas && kelasCol > -1) studentKelas = String(bioData[b][kelasCol] || '').trim();
          if (!studentKategori && katCol > -1) studentKategori = String(bioData[b][katCol] || '').trim();
          break;
        }
      }
    }

    if (!studentFound) {
      return sanitizeForRPC({ success: false, message: "ID Siswa tidak ditemukan atau status siswa sudah nonaktif." });
    }

    // 2. Validasi Nominal
    let rawNom = payload.nominal;
    let nominal = typeof rawNom === 'number' ? rawNom : parseInt(String(rawNom || '').replace(/[^0-9]/g, ''), 10);
    if (isNaN(nominal) || nominal <= 0 || nominal > 100000000) {
      return sanitizeForRPC({ success: false, message: "Nominal pembayaran tidak valid (Rp 1 s.d. Rp 100.000.000)." });
    }

    // 3. Validasi Jenis & Periode
    let jenis = String(payload.jenisTransaksi || 'INFAK BULANAN').trim();
    if (!jenis || jenis.length > 60) {
      return sanitizeForRPC({ success: false, message: "Jenis transaksi wajib diisi dan maksimal 60 karakter." });
    }

    const now = new Date();
    const yyyy = now.getFullYear();
    const mm = String(now.getMonth() + 1).padStart(2, '0');
    const dd = String(now.getDate()).padStart(2, '0');

    const monthsList = [
      'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
      'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
    ];
    let periodeBulan = String(payload.periodeBulan || monthsList[now.getMonth()]).trim();
    let periodeTahun = parseInt(payload.periodeTahun, 10);
    if (isNaN(periodeTahun) || periodeTahun < 2020 || periodeTahun > 2099) {
      periodeTahun = yyyy;
    }

    let keuSheet = ss.getSheetByName('KEUANGAN');
    if (!keuSheet) {
      keuSheet = ss.insertSheet('KEUANGAN');
      keuSheet.appendRow([
        'ID TRANSAKSI', 'NO INVOICE', 'ID SISWA', 'NAMA SISWA', 'KELAS', 'KATEGORI',
        'JENIS TRANSAKSI', 'PERIODE BULAN', 'PERIODE TAHUN', 'NOMINAL', 'TANGGAL BAYAR', 
        'METODE PEMBAYARAN', 'BUKTI TRANSFER', 'STATUS PEMBAYARAN', 'KETERANGAN', 'VERIFIKATOR', 
        'WAKTU VERIFIKASI', 'IS DELETED', 'CATATAN ADMIN'
      ]);
    }

    const kData = keuSheet.getDataRange().getValues();
    const kHeaders = kData[0].map(h => String(h || '').toUpperCase().trim());

    // 4. Validasi Anti-Duplikat
    const dupErr = cekDuplikatTagihan_(kData, kHeaders, idSiswa, jenis, periodeBulan, periodeTahun);
    if (dupErr) {
      return sanitizeForRPC({ success: false, message: dupErr });
    }

    // 5. Generate ID Transaksi & Invoice Unik
    let idTransaksi = "";
    let noInvoice = "";
    let tries = 0;
    while (tries < 10) {
      let randomNum = Math.floor(Math.random() * 900) + 100;
      let tempId = "TRX-" + yyyy + mm + "-" + randomNum;
      let isDuplicate = false;
      let idCol = kHeaders.indexOf('ID TRANSAKSI');
      if (idCol > -1) {
        for (let r = 1; r < kData.length; r++) {
          if (String(kData[r][idCol] || '').trim().toLowerCase() === tempId.toLowerCase()) {
            isDuplicate = true;
            break;
          }
        }
      }
      if (!isDuplicate) {
        idTransaksi = tempId;
        noInvoice = "INV/HD/" + yyyy + "/" + mm + "/" + randomNum;
        break;
      }
      tries++;
    }

    if (!idTransaksi) {
      let randomSuffix = String(Math.floor(Math.random() * 900) + 100);
      idTransaksi = "TRX-" + yyyy + mm + "-" + randomSuffix;
      noInvoice = "INV/HD/" + yyyy + "/" + mm + "/" + randomSuffix;
    }

    const tglBayar = payload.tanggalBayar || (dd + "/" + mm + "/" + yyyy);
    const actorId = (auth.session && (auth.session.uid || auth.session.email)) || 'ADMIN';
    const actorName = (auth.session && auth.session.nama) || 'Admin Sanggar';

    const rawStatus = String(payload.statusPembayaran || 'LUNAS').toUpperCase().trim();
    const statusValid = (rawStatus === 'PENDING VERIFIKASI') ? 'PENDING VERIFIKASI' : 'LUNAS';

    const pad = (n) => String(n).padStart(2, '0');
    const waktuVerifStr = statusValid === 'LUNAS' ? `${pad(now.getDate())}/${pad(now.getMonth() + 1)}/${now.getFullYear()} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}` : '-';
    const verifName = statusValid === 'LUNAS' ? actorName : '-';

    // 6. Simpan baris ke Sheet KEUANGAN
    keuSheet.appendRow([
      idTransaksi,
      noInvoice,
      idSiswa,
      studentNama,
      studentKelas || '-',
      studentKategori || '-',
      jenis,
      periodeBulan,
      periodeTahun,
      nominal,
      tglBayar,
      payload.metodePembayaran || 'TUNAI / CASH',
      payload.buktiUrl || '-',
      statusValid,
      payload.keterangan || 'Dicatat manual oleh Admin Sanggar',
      verifName,
      waktuVerifStr,
      'FALSE',
      payload.catatanAdmin || ''
    ]);

    logActivity(actorId, 'CREATE_TRANSAKSI_MANUAL', 'KEUANGAN', 'Admin mencatat transaksi ' + idTransaksi + ' (' + jenis + ' ' + periodeBulan + ' ' + periodeTahun + ') nominal: ' + nominal + ' status: ' + statusValid);

    return sanitizeForRPC({
      success: true,
      message: "Transaksi manual " + idTransaksi + " berhasil dicatat!",
      idTransaksi: idTransaksi,
      noInvoice: noInvoice
    });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal mencatat transaksi manual: " + err.toString() });
  }
}

function softDeleteTransaksi(idTransaksi, authToken) {
  try {
    const auth = requireAuth_(authToken, ['ADMIN', 'SUPER_ADMIN']);
    if (!auth.authorized) return auth.error;

    if (String(auth.session.role || '').toUpperCase() === 'ADMIN' && !canEditModule_(auth.session, 'pembayaran')) {
      return sanitizeForRPC({ success: false, message: "Hak akses Anda untuk modul Pembayaran hanya baca (Readonly)." });
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const keuSheet = ss.getSheetByName('KEUANGAN');
    if (!keuSheet) return sanitizeForRPC({ success: false, message: "Sheet KEUANGAN tidak ditemukan!" });

    const data = keuSheet.getDataRange().getValues();
    if (data.length <= 1) return sanitizeForRPC({ success: false, message: "Data KEUANGAN kosong." });

    const headersClean = data[0].map(h => String(h || '').toUpperCase().trim());
    let idIdx = headersClean.indexOf('ID TRANSAKSI');
    if (idIdx === -1) idIdx = 0;

    let isDeletedIdx = headersClean.indexOf('IS DELETED');
    if (isDeletedIdx === -1) isDeletedIdx = headersClean.indexOf('IS_DELETED');
    if (isDeletedIdx === -1) {
      isDeletedIdx = headersClean.length;
      keuSheet.getRange(1, isDeletedIdx + 1).setValue('IS DELETED');
    }

    const targetId = String(idTransaksi || '').trim().toLowerCase();
    const actorId = (auth.session && (auth.session.uid || auth.session.email)) || 'ADMIN';

    for (let i = 1; i < data.length; i++) {
      let rowId = String(data[i][idIdx] || '').trim().toLowerCase();
      if (rowId === targetId) {
        let currentDel = String(data[i][isDeletedIdx] || '').trim().toUpperCase();
        if (currentDel === 'TRUE') {
          return sanitizeForRPC({ success: true, message: "Transaksi " + idTransaksi + " sudah dihapus sebelumnya." });
        }
        keuSheet.getRange(i + 1, isDeletedIdx + 1).setValue('TRUE');

        logActivity(actorId, 'SOFT_DELETE_TRANSAKSI', 'KEUANGAN', 'Hapus (soft delete) transaksi: ' + idTransaksi);
        return sanitizeForRPC({ success: true, message: "Transaksi " + idTransaksi + " berhasil dihapus." });
      }
    }

    return sanitizeForRPC({ success: false, message: "ID Transaksi " + idTransaksi + " tidak ditemukan." });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal menghapus transaksi: " + err.toString() });
  }
}

function logLaporanAccess(action, authToken) {
  try {
    const auth = requireAuth_(authToken, ['ADMIN', 'SUPER_ADMIN']);
    if (!auth.authorized) return auth.error;

    const allowedActions = ['VIEW_LAPORAN', 'EXPORT_LAPORAN'];
    const act = String(action || 'VIEW_LAPORAN').toUpperCase().trim();
    if (!allowedActions.includes(act)) {
      return sanitizeForRPC({ success: false, message: "Aksi tidak valid." });
    }

    const actorId = (auth.session && (auth.session.uid || auth.session.email)) || 'ADMIN';
    const actorName = (auth.session && auth.session.nama) || 'Admin';
    const detail = (act === 'EXPORT_LAPORAN')
      ? 'Unduh berkas Excel Laporan Keuangan oleh ' + actorName
      : 'Membuka / melihat modul Laporan Keuangan oleh ' + actorName;

    logActivity(actorId, act, 'LAPORAN', detail);
    return sanitizeForRPC({ success: true, message: "Aktivitas dicatat." });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal mencatat log: " + err.toString() });
  }
}

function savePengeluaran(payload, authToken) {
  try {
    const token = authToken || (payload && (payload.__token || payload.token));
    const auth = requireAuth_(token, ['ADMIN', 'SUPER_ADMIN']);
    if (!auth.authorized) return auth.error;

    // Guard permissions matrix untuk role ADMIN
    if (String(auth.session.role || '').toUpperCase() === 'ADMIN' && !canEditModule_(auth.session, 'pengeluaran')) {
      return sanitizeForRPC({ success: false, message: "Akses ditolak: Anda tidak memiliki izin Editor untuk modul Pengeluaran." });
    }

    const actorId = (auth.session && (auth.session.uid || auth.session.email)) || 'ADMIN';
    const actorName = (auth.session && auth.session.nama) || 'Admin';

    if (!payload || typeof payload !== 'object') {
      return sanitizeForRPC({ success: false, message: "Data pengeluaran tidak valid." });
    }

    // 1. Validasi Tanggal
    let tgl = String(payload.tanggal || '').trim();
    if (!tgl) {
      const now = new Date();
      tgl = String(now.getDate()).padStart(2, '0') + '/' + String(now.getMonth() + 1).padStart(2, '0') + '/' + now.getFullYear();
    } else if (/^\d{4}-\d{2}-\d{2}$/.test(tgl)) {
      const p = tgl.split('-');
      tgl = p[2] + '/' + p[1] + '/' + p[0];
    }

    // 2. Validasi Kategori
    const kategori = String(payload.kategori || 'Lain-lain').trim() || 'Lain-lain';

    // 3. Validasi Keterangan
    const keterangan = String(payload.keterangan || '').trim();
    if (!keterangan) {
      return sanitizeForRPC({ success: false, message: "Keterangan pengeluaran wajib diisi!" });
    }
    if (keterangan.length > 1000) {
      return sanitizeForRPC({ success: false, message: "Keterangan pengeluaran maksimal 1.000 karakter!" });
    }

    // 4. Validasi Nominal
    const cleanNominalStr = String(payload.nominal || '').replace(/[^0-9]/g, '');
    const nominal = parseInt(cleanNominalStr, 10);
    if (isNaN(nominal) || nominal <= 0) {
      return sanitizeForRPC({ success: false, message: "Nominal pengeluaran harus berupa angka lebih besar dari 0!" });
    }
    if (nominal > 500000000) {
      return sanitizeForRPC({ success: false, message: "Nominal pengeluaran maksimal Rp 500.000.000." });
    }

    // 5. Validasi Metode
    const metode = String(payload.metode || 'TUNAI / CASH').trim() || 'TUNAI / CASH';

    // 6. Tangani Upload Bukti Nota (Opsional, max 2MB)
    let buktiUrl = String(payload.buktiUrl || '').trim();
    if (payload.buktiBase64) {
      let cleanBase64 = String(payload.buktiBase64);
      if (cleanBase64.indexOf(',') !== -1) {
        cleanBase64 = cleanBase64.split(',')[1];
      }
      cleanBase64 = cleanBase64.replace(/[^A-Za-z0-9+/=]/g, "");
      if (cleanBase64) {
        let approxBytes = Math.ceil(cleanBase64.length * 3 / 4);
        if (approxBytes > 2202009) {
          return sanitizeForRPC({ success: false, message: "Ukuran file bukti nota melebihi batas 2.0 MB." });
        }
        let decodedBytes = Utilities.base64Decode(cleanBase64);
        if (decodedBytes && decodedBytes.length > 0) {
          let isJpeg = decodedBytes.length > 2 && (decodedBytes[0] & 0xFF) === 0xFF && (decodedBytes[1] & 0xFF) === 0xD8;
          let isPng = decodedBytes.length > 4 && (decodedBytes[0] & 0xFF) === 0x89 && (decodedBytes[1] & 0xFF) === 0x50 && (decodedBytes[2] & 0xFF) === 0x4E && (decodedBytes[3] & 0xFF) === 0x47;
          if (isJpeg || isPng) {
            let ext = isPng ? ".png" : ".jpg";
            let mime = isPng ? "image/png" : "image/jpeg";
            let folderConfig = null;
            try {
              folderConfig = DriveApp.getFolderById(CONFIG.FOLDER_BUKTI_ID);
            } catch (fErr) {}
            if (folderConfig) {
              let fileName = "NOTA_" + new Date().getTime() + "_" + Math.floor(Math.random() * 1000) + ext;
              let blob = Utilities.newBlob(decodedBytes, mime, fileName);
              let file = folderConfig.createFile(blob);
              try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
              buktiUrl = "https://lh3.googleusercontent.com/d/" + file.getId();
            }
          }
        }
      }
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let pengSheet = ss.getSheetByName('PENGELUARAN');
    const headers = [
      'ID PENGELUARAN', 'TANGGAL', 'KATEGORI', 'KETERANGAN', 'NOMINAL',
      'METODE', 'BUKTI NOTA', 'DICATAT OLEH', 'WAKTU CATAT', 'IS DELETED'
    ];
    if (!pengSheet) {
      pengSheet = ss.insertSheet('PENGELUARAN');
      pengSheet.appendRow(headers);
    } else if (pengSheet.getLastRow() === 0) {
      pengSheet.appendRow(headers);
    }

    const data = pengSheet.getDataRange().getValues();
    const headersClean = data[0].map(h => String(h || '').toUpperCase().trim());
    const idIdx = headersClean.indexOf('ID PENGELUARAN');
    const isDeletedIdx = headersClean.indexOf('IS DELETED');

    const targetId = String(payload.idPengeluaran || '').trim();
    let existingRow = -1;
    if (targetId && data.length > 1 && idIdx > -1) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][idIdx] || '').trim().toLowerCase() === targetId.toLowerCase()) {
          existingRow = i + 1;
          break;
        }
      }
    }

    const now = new Date();
    const pad = n => String(n).padStart(2, '0');
    const waktuCatatStr = `${pad(now.getDate())}/${pad(now.getMonth() + 1)}/${now.getFullYear()} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;

    if (existingRow > -1) {
      // Update data baris
      const tglIdx = headersClean.indexOf('TANGGAL');
      const katIdx = headersClean.indexOf('KATEGORI');
      const ketIdx = headersClean.indexOf('KETERANGAN');
      const nomIdx = headersClean.indexOf('NOMINAL');
      const metIdx = headersClean.indexOf('METODE');
      const bktIdx = headersClean.indexOf('BUKTI NOTA');
      const dicIdx = headersClean.indexOf('DICATAT OLEH');
      const wktIdx = headersClean.indexOf('WAKTU CATAT');

      if (tglIdx > -1) pengSheet.getRange(existingRow, tglIdx + 1).setValue(tgl);
      if (katIdx > -1) pengSheet.getRange(existingRow, katIdx + 1).setValue(kategori);
      if (ketIdx > -1) pengSheet.getRange(existingRow, ketIdx + 1).setValue(keterangan);
      if (nomIdx > -1) pengSheet.getRange(existingRow, nomIdx + 1).setValue(nominal);
      if (metIdx > -1) pengSheet.getRange(existingRow, metIdx + 1).setValue(metode);
      if (bktIdx > -1) {
        if (buktiUrl) {
          pengSheet.getRange(existingRow, bktIdx + 1).setValue(buktiUrl);
        } else if (payload.hapusBukti === true || String(payload.hapusBukti).toLowerCase() === 'true') {
          pengSheet.getRange(existingRow, bktIdx + 1).setValue('-');
        }
      }
      if (dicIdx > -1) pengSheet.getRange(existingRow, dicIdx + 1).setValue(actorName);
      if (wktIdx > -1) pengSheet.getRange(existingRow, wktIdx + 1).setValue(waktuCatatStr);
      if (isDeletedIdx > -1) pengSheet.getRange(existingRow, isDeletedIdx + 1).setValue('FALSE');

      logActivity(actorId, 'UPDATE_PENGELUARAN', 'PENGELUARAN', 'Ubah pengeluaran ' + targetId + ' (' + kategori + '): Rp ' + nominal.toLocaleString('id-ID'));
      return sanitizeForRPC({ success: true, message: "Pengeluaran " + targetId + " berhasil diperbarui!", idPengeluaran: targetId });
    } else {
      // Generate ID unik EXP-YYYYMM-XXX secara sekuensial & terjamin unik
      const yyyy = String(now.getFullYear());
      const mm = pad(now.getMonth() + 1);
      const prefix = "EXP-" + yyyy + mm + "-";
      const existingIds = {};
      let maxSeq = 0;
      if (idIdx > -1 && data.length > 1) {
        for (let r = 1; r < data.length; r++) {
          const val = String(data[r][idIdx] || '').trim().toUpperCase();
          if (val) existingIds[val] = true;
          if (val.startsWith(prefix)) {
            const numPart = parseInt(val.substring(prefix.length), 10);
            if (!isNaN(numPart) && numPart > maxSeq) {
              maxSeq = numPart;
            }
          }
        }
      }
      let nextSeq = maxSeq + 1;
      let idPengeluaran = prefix + String(nextSeq).padStart(3, '0');
      while (existingIds[idPengeluaran.toUpperCase()]) {
        nextSeq++;
        idPengeluaran = prefix + String(nextSeq).padStart(3, '0');
      }

      pengSheet.appendRow([
        idPengeluaran,
        tgl,
        kategori,
        keterangan,
        nominal,
        metode,
        buktiUrl || '-',
        actorName,
        waktuCatatStr,
        'FALSE'
      ]);

      logActivity(actorId, 'CREATE_PENGELUARAN', 'PENGELUARAN', 'Catat pengeluaran baru ' + idPengeluaran + ' (' + kategori + '): Rp ' + nominal.toLocaleString('id-ID'));
      return sanitizeForRPC({ success: true, message: "Pengeluaran " + idPengeluaran + " berhasil dicatat!", idPengeluaran: idPengeluaran });
    }
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal menyimpan pengeluaran: " + err.toString() });
  }
}

function softDeletePengeluaran(idPengeluaran, authToken) {
  try {
    const auth = requireAuth_(authToken, ['ADMIN', 'SUPER_ADMIN']);
    if (!auth.authorized) return auth.error;

    // Guard permissions matrix untuk role ADMIN
    if (String(auth.session.role || '').toUpperCase() === 'ADMIN' && !canEditModule_(auth.session, 'pengeluaran')) {
      return sanitizeForRPC({ success: false, message: "Akses ditolak: Anda tidak memiliki izin menghapus data pengeluaran." });
    }

    const actorId = (auth.session && (auth.session.uid || auth.session.email)) || 'ADMIN';
    const targetId = String(idPengeluaran || '').trim().toLowerCase();
    if (!targetId) {
      return sanitizeForRPC({ success: false, message: "ID Pengeluaran tidak boleh kosong." });
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const pengSheet = ss.getSheetByName('PENGELUARAN');
    if (!pengSheet) return sanitizeForRPC({ success: false, message: "Sheet PENGELUARAN tidak ditemukan!" });

    const data = pengSheet.getDataRange().getValues();
    if (data.length <= 1) return sanitizeForRPC({ success: false, message: "Data PENGELUARAN kosong." });

    const headers = data[0].map(h => String(h || '').toUpperCase().trim());
    const idIdx = headers.indexOf('ID PENGELUARAN');
    let isDeletedIdx = headers.indexOf('IS DELETED');

    if (idIdx === -1) return sanitizeForRPC({ success: false, message: "Kolom ID PENGELUARAN tidak ditemukan." });

    if (isDeletedIdx === -1) {
      isDeletedIdx = headers.length;
      pengSheet.getRange(1, isDeletedIdx + 1).setValue('IS DELETED');
    }

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idIdx] || '').trim().toLowerCase() === targetId) {
        const currStatus = String(data[i][isDeletedIdx] || '').toUpperCase().trim();
        if (currStatus === 'TRUE') {
          return sanitizeForRPC({ success: true, message: "Pengeluaran " + idPengeluaran + " sudah dihapus sebelumnya." });
        }
        pengSheet.getRange(i + 1, isDeletedIdx + 1).setValue('TRUE');

        logActivity(actorId, 'SOFT_DELETE_PENGELUARAN', 'PENGELUARAN', 'Hapus (soft delete) pengeluaran: ' + idPengeluaran);
        return sanitizeForRPC({ success: true, message: "Pengeluaran " + idPengeluaran + " berhasil dihapus." });
      }
    }

    return sanitizeForRPC({ success: false, message: "ID Pengeluaran " + idPengeluaran + " tidak ditemukan." });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal menghapus pengeluaran: " + err.toString() });
  }
}

function fetchGaleriData(authToken) {
  try {
    if (authToken) {
      const auth = requireAuth_(authToken, null);
      if (!auth.authorized) return auth.error;
      const res = fetchTableData('GALERI', authToken, true);
      const role = String((auth.session && auth.session.role) || '').toUpperCase();
      if (role !== 'ADMIN' && role !== 'SUPER_ADMIN') {
        const uid = String((auth.session && auth.session.uid) || '').trim().toLowerCase();
        if (res && res.success && Array.isArray(res.data)) {
          res.data = res.data.filter(function(r) {
            const st = String((r && (r['STATUS TAYANG'] || r.statusTayang)) || '').toUpperCase().trim();
            if (st === 'PUBLISHED') return true;
            const idu = String((r && (r['ID SISWA PENGUNGGAH'] || r['ID SISWA'] || r.idSiswaPengunggah || r.idSiswa)) || '').trim().toLowerCase();
            return !!(idu && uid && idu === uid);
          });
        }
      }
      return res;
    }
    const res = fetchTableData('GALERI', '', true);
    if (res && res.success && Array.isArray(res.data)) {
      res.data = res.data.filter(function(r) {
        return String((r && (r['STATUS TAYANG'] || r.statusTayang)) || '').toUpperCase().trim() === 'PUBLISHED';
      });
    }
    return res;
  } catch (err) {
    return sanitizeForRPC({ success: false, message: err.toString(), data: [] });
  }
}

function submitUploadGaleriSiswa(payload, authToken) {
  try {
    const token = authToken || (payload && (payload.__token || payload.token));
    const auth = requireAuth_(token, null);
    if (!auth.authorized) return auth.error;

    const targetSiswa = String((payload && (payload.idSiswaPengunggah || payload.idSiswa)) || '').trim();
    if (targetSiswa && !requireOwnership_(auth.session, targetSiswa)) {
      return sanitizeForRPC({ success: false, message: "Tidak diizinkan mengunggah karya untuk siswa lain." });
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let galSheet = ss.getSheetByName('GALERI');

    const galHeaders = [
      'ID GALERI', 'JUDUL MEDIA', 'KATEGORI', 'DESKRIPSI', 'URL FOTO',
      'ID SISWA PENGUNGGAH', 'NAMA SISWA', 'STATUS TAYANG', 'WAKTU UPLOAD',
      'VERIFIKATOR', 'IS DELETED'
    ];

    if (!galSheet) {
      galSheet = ss.insertSheet('GALERI');
      galSheet.appendRow(galHeaders);
    } else if (galSheet.getLastRow() === 0) {
      galSheet.appendRow(galHeaders);
    }

    // Validasi input wajib
    const cleanJudul = String((payload && payload.judulMedia) || '').trim();
    const cleanDesk = String((payload && payload.deskripsi) || '').trim();
    if (!cleanJudul) {
      return sanitizeForRPC({ success: false, message: "Judul media/karya wajib diisi." });
    }
    if (cleanJudul.length > 120) {
      return sanitizeForRPC({ success: false, message: "Judul karya terlalu panjang (maksimal 120 karakter)." });
    }
    if (!cleanDesk) {
      return sanitizeForRPC({ success: false, message: "Deskripsi karya wajib diisi." });
    }
    if (cleanDesk.length > 1000) {
      return sanitizeForRPC({ success: false, message: "Deskripsi karya terlalu panjang (maksimal 1000 karakter)." });
    }

    const validKategori = ['Tugas', 'Ekstrakurikuler', 'Prestasi'];
    const finalKategori = (payload && validKategori.includes(payload.kategori)) ? payload.kategori : 'Tugas';

    if (!payload || !payload.fotoBase64) {
      return sanitizeForRPC({ success: false, message: "File foto karya wajib diunggah." });
    }

    let cleanBase64 = String(payload.fotoBase64);
    if (cleanBase64.indexOf(',') !== -1) {
      cleanBase64 = cleanBase64.split(',')[1];
    }
    cleanBase64 = cleanBase64.replace(/[^A-Za-z0-9+/=]/g, "");

    // Cek batas ukuran maksimal ~2.1 MB
    if (cleanBase64.length * 0.75 > 2.2 * 1024 * 1024) {
      return sanitizeForRPC({ success: false, message: "Ukuran foto terlalu besar (maksimal 2.0 MB). Silakan gunakan foto yang telah terkompresi." });
    }

    let fileUrl = "";
    const now = new Date();
    const tz = Session.getScriptTimeZone() || 'Asia/Jakarta';
    const serverIdGaleri = 'GAL-' + Utilities.formatDate(now, tz, 'yyyyMMddHHmmss') + '-' + Math.floor(1000 + Math.random() * 9000);
    const serverWaktuUpload = Utilities.formatDate(now, tz, 'dd/MM/yyyy HH:mm:ss');
    const idSiswa = String((payload && payload.idSiswaPengunggah) || '').trim();
    const namaSiswa = String((payload && payload.namaSiswa) || '').trim();

    let fileName = "KARYA_" + (idSiswa || 'SISWA') + "_" + now.getTime() + ".jpg";
    let blob = Utilities.newBlob(Utilities.base64Decode(cleanBase64), "image/jpeg", fileName);

    try {
      let folderConfig = DriveApp.getFolderById(CONFIG.FOLDER_GALERI_ID);
      let file = folderConfig.createFile(blob);
      try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
      fileUrl = "https://lh3.googleusercontent.com/d/" + file.getId();
    } catch (err1) {
      try {
        let file = DriveApp.createFile(blob);
        try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e2) {}
        fileUrl = "https://lh3.googleusercontent.com/d/" + file.getId();
      } catch (err2) {
        return sanitizeForRPC({ success: false, message: "Gagal menyimpan file foto ke Google Drive: " + err2.toString() });
      }
    }

    if (!fileUrl && !(payload && payload.urlFoto)) {
      return sanitizeForRPC({ success: false, message: "Gagal memproses URL foto karya." });
    }

    galSheet.appendRow([
      serverIdGaleri,
      cleanJudul,
      finalKategori,
      cleanDesk,
      fileUrl || (payload && payload.urlFoto ? payload.urlFoto : ''),
      idSiswa,
      namaSiswa,
      'PENDING',
      serverWaktuUpload,
      '',
      'FALSE'
    ]);

    logActivity(idSiswa || '-', 'UPLOAD_GALERI', 'GALERI', 'Siswa mengunggah karya galeri: ' + cleanJudul);

    return sanitizeForRPC({
      success: true,
      message: "Foto karya berhasil diunggah! Status karya saat ini PENDING MODERASI.",
      idGaleri: serverIdGaleri,
      urlFoto: fileUrl,
      waktuUpload: serverWaktuUpload
    });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal mengunggah foto karya: " + err.toString() });
  }
}

function updateGaleriSiswa(idGaleri, judulMedia, kategori, deskripsi, idSiswa, authToken) {
  try {
    const auth = requireAuth_(authToken, null);
    if (!auth.authorized) return auth.error;
    if (idSiswa && !requireOwnership_(auth.session, idSiswa)) {
      return sanitizeForRPC({ success: false, message: "Tidak diizinkan mengedit karya siswa lain." });
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const galSheet = ss.getSheetByName('GALERI');
    if (!galSheet) return sanitizeForRPC({ success: false, message: "Sheet GALERI tidak ditemukan!" });

    const cleanJudul = String(judulMedia || '').trim();
    const cleanDesk = String(deskripsi || '').trim();
    if (!cleanJudul) {
      return sanitizeForRPC({ success: false, message: "Judul media/karya wajib diisi." });
    }
    if (cleanJudul.length > 120) {
      return sanitizeForRPC({ success: false, message: "Judul karya terlalu panjang (maksimal 120 karakter)." });
    }
    if (!cleanDesk) {
      return sanitizeForRPC({ success: false, message: "Deskripsi karya wajib diisi." });
    }
    if (cleanDesk.length > 1000) {
      return sanitizeForRPC({ success: false, message: "Deskripsi karya terlalu panjang (maksimal 1000 karakter)." });
    }

    const validKategori = ['Tugas', 'Ekstrakurikuler', 'Prestasi'];
    const finalKategori = validKategori.includes(kategori) ? kategori : 'Tugas';

    const data = galSheet.getDataRange().getValues();
    if (data.length <= 1) return sanitizeForRPC({ success: false, message: "Data GALERI kosong." });

    const headersClean = data[0].map(h => String(h || '').toUpperCase().trim());
    let idIdx = headersClean.indexOf('ID GALERI');
    if (idIdx === -1) idIdx = 0;

    let judulIdx = headersClean.indexOf('JUDUL MEDIA');
    if (judulIdx === -1) judulIdx = 1;

    let katIdx = headersClean.indexOf('KATEGORI');
    if (katIdx === -1) katIdx = 2;

    let deskIdx = headersClean.indexOf('DESKRIPSI');
    if (deskIdx === -1) deskIdx = 3;

    let idSiswaIdx = headersClean.indexOf('ID SISWA PENGUNGGAH');
    if (idSiswaIdx === -1) idSiswaIdx = headersClean.indexOf('ID SISWA');

    let isDelIdx = headersClean.indexOf('IS DELETED');

    const targetId = String(idGaleri || '').trim().toLowerCase();
    const reqSiswa = String(idSiswa || '').trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      let rowId = String(data[i][idIdx] || '').trim().toLowerCase();
      let rowIsDel = isDelIdx > -1 ? String(data[i][isDelIdx] || '').trim().toUpperCase() : 'FALSE';

      if (rowId === targetId && rowIsDel !== 'TRUE') {
        if (reqSiswa && idSiswaIdx > -1) {
          let rowSiswa = String(data[i][idSiswaIdx] || '').trim().toLowerCase();
          if (rowSiswa && rowSiswa !== reqSiswa) {
            return sanitizeForRPC({ success: false, message: "Anda tidak memiliki akses untuk mengedit karya ini." });
          }
        }

        if (judulIdx > -1) galSheet.getRange(i + 1, judulIdx + 1).setValue(cleanJudul);
        if (katIdx > -1) galSheet.getRange(i + 1, katIdx + 1).setValue(finalKategori);
        if (deskIdx > -1) galSheet.getRange(i + 1, deskIdx + 1).setValue(cleanDesk);

        logActivity(reqSiswa || targetId, 'UPDATE_GALERI', 'GALERI', 'Siswa mengedit karya galeri: ' + idGaleri);
        return sanitizeForRPC({ success: true, message: "Informasi karya galeri berhasil diperbarui!" });
      }
    }

    return sanitizeForRPC({ success: false, message: "ID Galeri " + idGaleri + " tidak ditemukan atau sudah dihapus." });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal memperbarui karya galeri: " + err.toString() });
  }
}

function deleteGaleriSiswa(idGaleri, idSiswa, authToken) {
  try {
    const auth = requireAuth_(authToken, null);
    if (!auth.authorized) return auth.error;
    if (idSiswa && !requireOwnership_(auth.session, idSiswa)) {
      return sanitizeForRPC({ success: false, message: "Tidak diizinkan menghapus karya siswa lain." });
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const galSheet = ss.getSheetByName('GALERI');
    if (!galSheet) return sanitizeForRPC({ success: false, message: "Sheet GALERI tidak ditemukan!" });

    const data = galSheet.getDataRange().getValues();
    if (data.length <= 1) return sanitizeForRPC({ success: false, message: "Data GALERI kosong." });

    const headersClean = data[0].map(h => String(h || '').toUpperCase().trim());
    let idIdx = headersClean.indexOf('ID GALERI');
    if (idIdx === -1) idIdx = 0;

    let idSiswaIdx = headersClean.indexOf('ID SISWA PENGUNGGAH');
    if (idSiswaIdx === -1) idSiswaIdx = headersClean.indexOf('ID SISWA');

    let isDelIdx = headersClean.indexOf('IS DELETED');
    if (isDelIdx === -1) {
      isDelIdx = headersClean.length;
      galSheet.getRange(1, isDelIdx + 1).setValue('IS DELETED');
    }

    const targetId = String(idGaleri || '').trim().toLowerCase();
    const reqSiswa = String(idSiswa || '').trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      let rowId = String(data[i][idIdx] || '').trim().toLowerCase();
      let rowIsDel = isDelIdx > -1 ? String(data[i][isDelIdx] || '').trim().toUpperCase() : 'FALSE';

      if (rowId === targetId && rowIsDel !== 'TRUE') {
        if (reqSiswa && idSiswaIdx > -1) {
          let rowSiswa = String(data[i][idSiswaIdx] || '').trim().toLowerCase();
          if (rowSiswa && rowSiswa !== reqSiswa) {
            return sanitizeForRPC({ success: false, message: "Anda tidak memiliki akses untuk menghapus karya ini." });
          }
        }

        galSheet.getRange(i + 1, isDelIdx + 1).setValue('TRUE');
        logActivity(reqSiswa || targetId, 'DELETE_GALERI', 'GALERI', 'Siswa menarik/menghapus (soft delete) karya galeri: ' + idGaleri);
        return sanitizeForRPC({ success: true, message: "Karya galeri berhasil ditarik dan ditandai terhapus di database." });
      }
    }

    return sanitizeForRPC({ success: false, message: "ID Galeri " + idGaleri + " tidak ditemukan atau sudah dihapus." });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal menghapus karya galeri: " + err.toString() });
  }
}

function ajukanUlangGaleri(idGaleri, idSiswa, authToken) {
  try {
    const auth = requireAuth_(authToken, null);
    if (!auth.authorized) return auth.error;
    if (idSiswa && !requireOwnership_(auth.session, idSiswa)) {
      return sanitizeForRPC({ success: false, message: "Tidak diizinkan mengajukan ulang karya siswa lain." });
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const galSheet = ss.getSheetByName('GALERI');
    if (!galSheet) return sanitizeForRPC({ success: false, message: "Sheet GALERI tidak ditemukan!" });

    const data = galSheet.getDataRange().getValues();
    if (data.length <= 1) return sanitizeForRPC({ success: false, message: "Data GALERI kosong." });

    const headersClean = data[0].map(h => String(h || '').toUpperCase().trim());
    let idIdx = headersClean.indexOf('ID GALERI');
    if (idIdx === -1) idIdx = 0;

    let statusIdx = headersClean.indexOf('STATUS TAYANG');
    if (statusIdx === -1) statusIdx = 7;

    let verifikatorIdx = headersClean.indexOf('VERIFIKATOR');
    if (verifikatorIdx === -1) verifikatorIdx = 9;

    let catatanIdx = headersClean.indexOf('CATATAN ADMIN');

    let idSiswaIdx = headersClean.indexOf('ID SISWA PENGUNGGAH');
    if (idSiswaIdx === -1) idSiswaIdx = headersClean.indexOf('ID SISWA');

    let isDelIdx = headersClean.indexOf('IS DELETED');

    const targetId = String(idGaleri || '').trim().toLowerCase();
    const reqSiswa = String(idSiswa || '').trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      let rowId = String(data[i][idIdx] || '').trim().toLowerCase();
      let rowIsDel = isDelIdx > -1 ? String(data[i][isDelIdx] || '').trim().toUpperCase() : 'FALSE';

      if (rowId === targetId && rowIsDel !== 'TRUE') {
        if (reqSiswa && idSiswaIdx > -1) {
          let rowSiswa = String(data[i][idSiswaIdx] || '').trim().toLowerCase();
          if (rowSiswa && rowSiswa !== reqSiswa) {
            return sanitizeForRPC({ success: false, message: "Anda tidak memiliki akses untuk mengajukan ulang karya ini." });
          }
        }

        galSheet.getRange(i + 1, statusIdx + 1).setValue('PENDING');
        if (verifikatorIdx > -1) galSheet.getRange(i + 1, verifikatorIdx + 1).setValue('');
        if (catatanIdx > -1) galSheet.getRange(i + 1, catatanIdx + 1).setValue('');

        logActivity(reqSiswa || targetId, 'AJUKAN_ULANG_GALERI', 'GALERI', 'Siswa mengajukan ulang karya galeri: ' + idGaleri);
        return sanitizeForRPC({ success: true, message: "Karya galeri berhasil diajukan ulang! Status kembali menjadi PENDING MODERASI." });
      }
    }

    return sanitizeForRPC({ success: false, message: "Karya dengan ID " + idGaleri + " tidak ditemukan atau sudah dihapus." });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal mengajukan ulang karya galeri: " + err.toString() });
  }
}

function updateStatusGaleri(idGaleri, statusTayang, verifikator, catatanAdmin, authToken) {
  try {
    const auth = requireAuth_(authToken, ['ADMIN', 'SUPER_ADMIN']);
    if (!auth.authorized) return auth.error;
    if (String(auth.session.role || '').toUpperCase() === 'ADMIN' && !canEditModule_(auth.session, 'galeri_admin')) {
      return sanitizeForRPC({ success: false, message: "Mode baca saja (Readonly) — Anda tidak memiliki izin untuk memoderasi karya galeri." });
    }

    const normStatus = String(statusTayang || '').trim().toUpperCase();
    const allowedStatus = ['PENDING', 'PUBLISHED', 'REJECTED'];
    if (!allowedStatus.includes(normStatus)) {
      return sanitizeForRPC({ success: false, message: "Status moderasi tidak valid. Harus PENDING, PUBLISHED, atau REJECTED." });
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const galSheet = ss.getSheetByName('GALERI');
    if (!galSheet) return sanitizeForRPC({ success: false, message: "Sheet GALERI tidak ditemukan!" });

    const data = galSheet.getDataRange().getValues();
    if (data.length <= 1) return sanitizeForRPC({ success: false, message: "Data GALERI kosong." });

    const headersClean = data[0].map(h => String(h || '').toUpperCase().trim());
    let idIdx = headersClean.indexOf('ID GALERI');
    if (idIdx === -1) idIdx = 0;

    let statusIdx = headersClean.indexOf('STATUS TAYANG');
    if (statusIdx === -1) statusIdx = 7;

    let verifikatorIdx = headersClean.indexOf('VERIFIKATOR');
    if (verifikatorIdx === -1) verifikatorIdx = 9;

    let catatanIdx = headersClean.indexOf('CATATAN ADMIN');
    if (catatanIdx === -1) {
      catatanIdx = headersClean.length;
      galSheet.getRange(1, catatanIdx + 1).setValue('CATATAN ADMIN');
    }

    let isDeletedIdx = headersClean.indexOf('IS DELETED');

    const cleanCatatan = String(catatanAdmin || '').trim().slice(0, 500);
    const actorName = String((auth.session && (auth.session.nama || auth.session.email)) || 'Admin Sanggar').trim();
    const actorId = String((auth.session && (auth.session.uid || auth.session.email)) || 'ADMIN').trim();

    const targetId = String(idGaleri || '').trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      let rowId = String(data[i][idIdx] || '').trim().toLowerCase();
      let rowIsDel = isDeletedIdx > -1 ? String(data[i][isDeletedIdx] || '').trim().toUpperCase() : 'FALSE';
      if (rowId === targetId) {
        if (rowIsDel === 'TRUE') {
          return sanitizeForRPC({ success: false, message: "Karya galeri " + idGaleri + " sudah dihapus sebelumnya." });
        }
        galSheet.getRange(i + 1, statusIdx + 1).setValue(normStatus);
        galSheet.getRange(i + 1, verifikatorIdx + 1).setValue(actorName);
        galSheet.getRange(i + 1, catatanIdx + 1).setValue(cleanCatatan);

        logActivity(actorId, 'MODERASI_GALERI', 'GALERI', 'Ubah status galeri ' + idGaleri + ' menjadi ' + normStatus + (cleanCatatan ? (' dengan catatan: ' + cleanCatatan) : ''));
        return sanitizeForRPC({ success: true, message: "Status karya galeri berhasil diubah menjadi " + normStatus });
      }
    }

    return sanitizeForRPC({ success: false, message: "ID Galeri " + idGaleri + " tidak ditemukan." });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal memperbarui status galeri: " + err.toString() });
  }
}

function softDeleteGaleri(idGaleri, authToken) {
  try {
    const auth = requireAuth_(authToken, ['ADMIN', 'SUPER_ADMIN']);
    if (!auth.authorized) return auth.error;
    if (String(auth.session.role || '').toUpperCase() === 'ADMIN' && !canEditModule_(auth.session, 'galeri_admin')) {
      return sanitizeForRPC({ success: false, message: "Mode baca saja (Readonly) — Anda tidak memiliki izin untuk menghapus karya galeri." });
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const galSheet = ss.getSheetByName('GALERI');
    if (!galSheet) return sanitizeForRPC({ success: false, message: "Sheet GALERI tidak ditemukan!" });

    const data = galSheet.getDataRange().getValues();
    if (data.length <= 1) return sanitizeForRPC({ success: false, message: "Data GALERI kosong." });

    const headersClean = data[0].map(h => String(h || '').toUpperCase().trim());
    let idIdx = headersClean.indexOf('ID GALERI');
    if (idIdx === -1) idIdx = 0;

    let isDeletedIdx = headersClean.indexOf('IS DELETED');
    if (isDeletedIdx === -1) {
      isDeletedIdx = headersClean.length;
      galSheet.getRange(1, isDeletedIdx + 1).setValue('IS DELETED');
    }

    const actorId = String((auth.session && (auth.session.uid || auth.session.email)) || 'ADMIN').trim();
    const targetId = String(idGaleri || '').trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      let rowId = String(data[i][idIdx] || '').trim().toLowerCase();
      let rowIsDel = isDeletedIdx > -1 ? String(data[i][isDeletedIdx] || '').trim().toUpperCase() : 'FALSE';
      if (rowId === targetId) {
        if (rowIsDel === 'TRUE') {
          return sanitizeForRPC({ success: false, message: "Karya galeri " + idGaleri + " sudah dihapus sebelumnya." });
        }
        galSheet.getRange(i + 1, isDeletedIdx + 1).setValue('TRUE');
        logActivity(actorId, 'SOFT_DELETE_GALERI', 'GALERI', 'Admin menghapus (soft delete) karya galeri: ' + idGaleri);
        return sanitizeForRPC({ success: true, message: "Karya galeri berhasil dihapus (IS DELETED = 'TRUE')." });
      }
    }

    return sanitizeForRPC({ success: false, message: "ID Galeri " + idGaleri + " tidak ditemukan." });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal menghapus karya galeri: " + err.toString() });
  }
}

function fetchJadwalLes(authToken) {
  if (authToken) {
    const auth = requireAuth_(authToken, null);
    if (!auth.authorized) return auth.error;
  }
  return fetchTableData('JADWAL', authToken, true);
}

function saveJadwalLes(payload, authToken) {
  try {
    const token = authToken || (payload && (payload.__token || payload.token));
    const auth = requireAuth_(token, ['ADMIN', 'SUPER_ADMIN']);
    if (!auth.authorized) return auth.error;

    if (String(auth.session.role).toUpperCase() === 'ADMIN' && !canEditModule_(auth.session, 'absensi')) {
      return sanitizeForRPC({ success: false, message: "Mode baca saja (Readonly). Anda tidak memiliki izin mengubah modul Absensi." });
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let jadwalSheet = ss.getSheetByName('JADWAL');

    if (!jadwalSheet) {
      jadwalSheet = ss.insertSheet('JADWAL');
      jadwalSheet.appendRow(['ID JADWAL', 'TANGGAL', 'HARI', 'JAM', 'PERTEMUAN', 'KELAS', 'KATEGORI', 'STATUS', 'CREATED_AT', 'IS DELETED']);
    }

    const data = jadwalSheet.getDataRange().getValues();
    let headers = data[0].map(h => String(h).toUpperCase().trim());
    
    // Periksa dan sisipkan kolom KATEGORI jika belum ada di sheet
    let katIdx = headers.indexOf('KATEGORI');
    if (katIdx === -1) {
        katIdx = headers.length;
        jadwalSheet.getRange(1, katIdx + 1).setValue('KATEGORI');
        headers.push('KATEGORI');
    }

    const targetId = String((payload && payload.idJadwal) || '').trim().toLowerCase();
    let existingRow = -1;
    if (data.length > 1) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0] || '').trim().toLowerCase() === targetId) {
          existingRow = i + 1;
          break;
        }
      }
    }

    const getValue = (key, def) => payload[key] !== undefined ? payload[key] : def;
    const colIdx = (name) => headers.indexOf(name) + 1;

    if (existingRow > -1) {
      if(colIdx('TANGGAL')) jadwalSheet.getRange(existingRow, colIdx('TANGGAL')).setValue(getValue('tanggal', ''));
      if(colIdx('HARI')) jadwalSheet.getRange(existingRow, colIdx('HARI')).setValue(getValue('hari', 'Minggu'));
      if(colIdx('JAM')) jadwalSheet.getRange(existingRow, colIdx('JAM')).setValue(getValue('jam', '09:00 - 12:00 WIB'));
      if(colIdx('PERTEMUAN')) jadwalSheet.getRange(existingRow, colIdx('PERTEMUAN')).setValue(getValue('pertemuan', 'Pertemuan 1'));
      if(colIdx('KELAS')) jadwalSheet.getRange(existingRow, colIdx('KELAS')).setValue(getValue('kelas', ''));
      if(colIdx('KATEGORI')) jadwalSheet.getRange(existingRow, colIdx('KATEGORI')).setValue(getValue('kategori', ''));
      if(colIdx('STATUS')) jadwalSheet.getRange(existingRow, colIdx('STATUS')).setValue(getValue('status', 'AKTIF'));
      if(colIdx('IS DELETED')) jadwalSheet.getRange(existingRow, colIdx('IS DELETED')).setValue('FALSE');
    } else {
      const newRow = new Array(headers.length).fill('');
      if(colIdx('ID JADWAL')) newRow[colIdx('ID JADWAL') - 1] = payload.idJadwal || ('JAD-' + new Date().getTime());
      if(colIdx('TANGGAL')) newRow[colIdx('TANGGAL') - 1] = getValue('tanggal', '');
      if(colIdx('HARI')) newRow[colIdx('HARI') - 1] = getValue('hari', 'Minggu');
      if(colIdx('JAM')) newRow[colIdx('JAM') - 1] = getValue('jam', '09:00 - 12:00 WIB');
      if(colIdx('PERTEMUAN')) newRow[colIdx('PERTEMUAN') - 1] = getValue('pertemuan', 'Pertemuan 1');
      if(colIdx('KELAS')) newRow[colIdx('KELAS') - 1] = getValue('kelas', '');
      if(colIdx('KATEGORI')) newRow[colIdx('KATEGORI') - 1] = getValue('kategori', '');
      if(colIdx('STATUS')) newRow[colIdx('STATUS') - 1] = getValue('status', 'AKTIF');
      if(colIdx('CREATED_AT')) newRow[colIdx('CREATED_AT') - 1] = new Date().toISOString();
      if(colIdx('IS DELETED')) newRow[colIdx('IS DELETED') - 1] = 'FALSE';
      
      jadwalSheet.appendRow(newRow);
    }

    const actorId = auth.session ? (auth.session.uid || auth.session.email || 'ADMIN') : 'ADMIN';
    logActivity(actorId, 'SAVE_JADWAL', 'JADWAL', (existingRow > -1 ? 'Ubah' : 'Tambah') + ' jadwal: ' + (payload.idJadwal || ''));

    return sanitizeForRPC({ success: true, message: "Jadwal les berhasil disimpan!" });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal menyimpan jadwal les: " + err.toString() });
  }
}

function saveBatchJadwalLes(schedulesArray, authToken) {
  try {
    const auth = requireAuth_(authToken, ['ADMIN', 'SUPER_ADMIN']);
    if (!auth.authorized) return auth.error;

    if (String(auth.session.role).toUpperCase() === 'ADMIN' && !canEditModule_(auth.session, 'absensi')) {
      return sanitizeForRPC({ success: false, message: "Mode baca saja (Readonly). Anda tidak memiliki izin mengubah modul Absensi." });
    }

    if (!Array.isArray(schedulesArray) || schedulesArray.length === 0) {
      return sanitizeForRPC({ success: false, message: "Daftar jadwal kosong." });
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let jadwalSheet = ss.getSheetByName('JADWAL');
    
    if (!jadwalSheet) {
      jadwalSheet = ss.insertSheet('JADWAL');
      jadwalSheet.appendRow(['ID JADWAL', 'TANGGAL', 'HARI', 'JAM', 'PERTEMUAN', 'KELAS', 'KATEGORI', 'STATUS', 'CREATED_AT', 'IS DELETED']);
    }
    
    const data = jadwalSheet.getDataRange().getValues();
    let headers = data[0].map(h => String(h).toUpperCase().trim());
    
    // Periksa dan sisipkan kolom KATEGORI jika belum ada di sheet
    let katIdx = headers.indexOf('KATEGORI');
    if (katIdx === -1) {
        katIdx = headers.length;
        jadwalSheet.getRange(1, katIdx + 1).setValue('KATEGORI');
        headers.push('KATEGORI');
    }

    const colIdx = (name) => headers.indexOf(name);

    schedulesArray.forEach(function(item) {
      const newRow = new Array(headers.length).fill('');
      if(colIdx('ID JADWAL') > -1) newRow[colIdx('ID JADWAL')] = item.idJadwal || ('JAD-' + Date.now());
      if(colIdx('TANGGAL') > -1) newRow[colIdx('TANGGAL')] = item.tanggal || '';
      if(colIdx('HARI') > -1) newRow[colIdx('HARI')] = item.hari || 'Minggu';
      if(colIdx('JAM') > -1) newRow[colIdx('JAM')] = item.jam || '09:00 - 12:00 WIB';
      if(colIdx('PERTEMUAN') > -1) newRow[colIdx('PERTEMUAN')] = item.pertemuan || 'Pertemuan 1';
      if(colIdx('KELAS') > -1) newRow[colIdx('KELAS')] = item.kelas || '';
      if(colIdx('KATEGORI') > -1) newRow[colIdx('KATEGORI')] = item.kategori || '';
      if(colIdx('STATUS') > -1) newRow[colIdx('STATUS')] = item.status || 'AKTIF';
      if(colIdx('CREATED_AT') > -1) newRow[colIdx('CREATED_AT')] = new Date().toISOString();
      if(colIdx('IS DELETED') > -1) newRow[colIdx('IS DELETED')] = 'FALSE';
      
      jadwalSheet.appendRow(newRow);
    });

    const actorId = auth.session ? (auth.session.uid || auth.session.email || 'ADMIN') : 'ADMIN';
    logActivity(actorId, 'GENERATE_JADWAL', 'JADWAL', 'Generate batch ' + schedulesArray.length + ' jadwal les');

    return sanitizeForRPC({ success: true, message: "Batch jadwal les berhasil dibuat!" });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal generate batch jadwal: " + err.toString() });
  }
}

function deleteJadwalLes(idJadwal, authToken) {
  try {
    const auth = requireAuth_(authToken, ['ADMIN', 'SUPER_ADMIN']);
    if (!auth.authorized) return auth.error;

    if (String(auth.session.role).toUpperCase() === 'ADMIN' && !canEditModule_(auth.session, 'absensi')) {
      return sanitizeForRPC({ success: false, message: "Mode baca saja (Readonly). Anda tidak memiliki izin mengubah modul Absensi." });
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const jadwalSheet = ss.getSheetByName('JADWAL');
    if (!jadwalSheet) return sanitizeForRPC({ success: false, message: "Sheet JADWAL tidak ditemukan!" });

    const data = jadwalSheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h || '').toUpperCase().trim());
    let delIdx = headers.indexOf('IS DELETED');
    if (delIdx === -1) {
      delIdx = headers.length;
      jadwalSheet.getRange(1, delIdx + 1).setValue('IS DELETED');
      headers.push('IS DELETED');
    }
    let statusIdx = headers.indexOf('STATUS');

    const targetId = String(idJadwal || '').trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim().toLowerCase() === targetId) {
        jadwalSheet.getRange(i + 1, delIdx + 1).setValue('TRUE');
        if (statusIdx > -1) {
          jadwalSheet.getRange(i + 1, statusIdx + 1).setValue('NONAKTIF');
        }
        const actorId = auth.session ? (auth.session.uid || auth.session.email || 'ADMIN') : 'ADMIN';
        logActivity(actorId, 'DELETE_JADWAL', 'JADWAL', 'Soft delete jadwal les: ' + idJadwal);
        return sanitizeForRPC({ success: true, message: "Jadwal les berhasil dihapus (nonaktif)." });
      }
    }

    return sanitizeForRPC({ success: false, message: "ID Jadwal " + idJadwal + " tidak ditemukan." });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal menghapus jadwal les: " + err.toString() });
  }
}

function saveBatchAbsensi(batchData, authToken) {
  try {
    const auth = requireAuth_(authToken, ['ADMIN', 'SUPER_ADMIN']);
    if (!auth.authorized) return auth.error;

    if (String(auth.session.role).toUpperCase() === 'ADMIN' && !canEditModule_(auth.session, 'absensi')) {
      return sanitizeForRPC({ success: false, message: "Mode baca saja (Readonly). Anda tidak memiliki izin mengubah modul Absensi." });
    }

    if (!Array.isArray(batchData) || batchData.length === 0) {
      return sanitizeForRPC({ success: false, message: "Data batch presensi kosong." });
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let absensiSheet = ss.getSheetByName('ABSENSI');
    const headers = [
      'ID ABSENSI', 'ID JADWAL', 'TANGGAL', 'ID SISWA', 'NAMA SISWA',
      'KELAS', 'STATUS KEHADIRAN', 'CATATAN', 'VERIFIKATOR', 'WAKTU RECORD', 'IS DELETED'
    ];

    if (!absensiSheet) {
      absensiSheet = ss.insertSheet('ABSENSI');
      absensiSheet.appendRow(headers);
    }

    const data = absensiSheet.getDataRange().getValues();
    const headersClean = data[0].map(h => String(h || '').toUpperCase().replace(/[^A-Z0-9]/g, ''));
    let iJadwal = headersClean.findIndex(h => h === 'IDJADWAL');
    if (iJadwal === -1) iJadwal = 1;
    let iSiswa = headersClean.findIndex(h => h === 'IDSISWA' || h === 'SISWA');
    if (iSiswa === -1) iSiswa = 3;
    let iStatus = headersClean.findIndex(h => h.includes('STATUS'));
    if (iStatus === -1) iStatus = 6;
    let iCatatan = headersClean.findIndex(h => h.includes('CATATAN') || h.includes('KETERANGAN'));
    if (iCatatan === -1) iCatatan = 7;
    let iVerifikator = headersClean.findIndex(h => h.includes('VERIFIKATOR'));
    if (iVerifikator === -1) iVerifikator = 8;
    let iWaktu = headersClean.findIndex(h => h.includes('WAKTU') || h.includes('RECORD'));
    if (iWaktu === -1) iWaktu = 9;
    let iDeleted = headersClean.findIndex(h => h.includes('ISDELETED') || h.includes('HAPUS'));
    if (iDeleted === -1) iDeleted = headersClean.length - 1;

    const existingMap = {};
    if (data.length > 1) {
      for (let i = 1; i < data.length; i++) {
        const isDeleted = String(data[i][iDeleted] || '').toUpperCase() === 'TRUE';
        if (!isDeleted) {
          const key = String(data[i][iJadwal] || '').trim().toLowerCase() + '_' + String(data[i][iSiswa] || '').trim().toLowerCase();
          existingMap[key] = i + 1;
        }
      }
    }

    batchData.forEach(function(item) {
      const key = String(item.idJadwal || '').trim().toLowerCase() + '_' + String(item.idSiswa || '').trim().toLowerCase();
      const existingRow = existingMap[key];

      if (existingRow) {
        absensiSheet.getRange(existingRow, iStatus + 1).setValue(item.statusKehadiran || 'HADIR');
        absensiSheet.getRange(existingRow, iCatatan + 1).setValue(item.catatan || '');
        absensiSheet.getRange(existingRow, iVerifikator + 1).setValue(item.verifikator || 'Admin Sanggar');
        absensiSheet.getRange(existingRow, iWaktu + 1).setValue(item.waktuRecord || new Date().toLocaleString('id-ID'));
      } else {
        absensiSheet.appendRow([
          item.idAbsensi || ('ABS-' + Date.now()),
          item.idJadwal || '',
          item.tanggal || '',
          item.idSiswa || '',
          item.namaSiswa || '',
          item.kelas || '',
          item.statusKehadiran || 'HADIR',
          item.catatan || '',
          item.verifikator || 'Admin Sanggar',
          item.waktuRecord || new Date().toLocaleString('id-ID'),
          'FALSE'
        ]);
      }
    });

    bumpAbsensiCacheVer();
    const actorId = auth.session ? (auth.session.uid || auth.session.email || 'ADMIN') : 'ADMIN';
    const firstJadwal = (batchData[0] && (batchData[0].idJadwal || batchData[0].tanggal)) || '-';
    logActivity(actorId, 'SAVE_PRESENSI', 'ABSENSI', 'Simpan presensi ' + batchData.length + ' siswa untuk sesi: ' + firstJadwal);
    return sanitizeForRPC({ success: true, message: "Data presensi " + batchData.length + " siswa berhasil disimpan ke Sheet ABSENSI!" });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal menyimpan presensi: " + err.toString() });
  }
}

function fetchAbsensiByJadwal(idJadwal, authToken) {
  try {
    const auth = requireAuth_(authToken, ['ADMIN', 'SUPER_ADMIN']);
    if (!auth.authorized) return auth.error;

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('ABSENSI');
    if (!sheet) return sanitizeForRPC({ success: true, data: [] });

    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return sanitizeForRPC({ success: true, data: [] });

    const headers = values[0];
    const headersClean = headers.map(h => String(h || '').toUpperCase().replace(/[^A-Z0-9]/g, ''));
    let iJadwal = headersClean.findIndex(h => h === 'IDJADWAL');
    if (iJadwal === -1) iJadwal = 1;
    let iDeleted = headersClean.findIndex(h => h.includes('ISDELETED') || h.includes('HAPUS'));
    if (iDeleted === -1) iDeleted = headersClean.length - 1;

    const targetJadwal = String(idJadwal || '').trim().toLowerCase();
    const result = [];

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const isDeleted = String(row[iDeleted] || '').toUpperCase() === 'TRUE';
      const rJadwal = String(row[iJadwal] || '').trim().toLowerCase();

      if (!isDeleted && rJadwal === targetJadwal) {
        let item = {};
        headers.forEach((h, idx) => {
          let val = row[idx];
          if (val instanceof Date) {
            val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd/MM/yyyy');
          }
          item[String(h)] = (val !== undefined && val !== null) ? String(val) : '';
        });
        result.push(item);
      }
    }

    return sanitizeForRPC({ success: true, data: result });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: err.toString(), data: [] });
  }
}

function getAbsensiCacheVer() {
  try {
    const v = CacheService.getScriptCache().get('ABSEN_CACHE_VER');
    return v ? parseInt(v, 10) || 1 : 1;
  } catch (e) {
    return 1;
  }
}

function bumpAbsensiCacheVer() {
  try {
    const cur = getAbsensiCacheVer();
    const nextVer = (cur + 1) % 1000000;
    CacheService.getScriptCache().put('ABSEN_CACHE_VER', String(nextVer), 21600);
  } catch (e) {}
}

function fetchAbsensiSiswa(idSiswa, authToken) {
  try {
    const auth = requireAuth_(authToken, null);
    if (!auth.authorized) return auth.error;
    if (!requireOwnership_(auth.session, idSiswa)) {
      return sanitizeForRPC({ success: false, message: "Akses data absensi siswa lain tidak diizinkan." });
    }

    const targetSiswa = String(idSiswa || '').trim().toLowerCase();
    if (!targetSiswa) {
      return sanitizeForRPC({ success: true, data: [] });
    }

    // 1. Cek server-side cache (TTL 60 detik)
    const ver = getAbsensiCacheVer();
    const cacheKey = 'ABS_V' + ver + '_' + targetSiswa.replace(/[^a-z0-9_-]/gi, '_').slice(0, 50);
    try {
      const cached = CacheService.getScriptCache().get(cacheKey);
      if (cached) {
        const parsed = JSON.parse(cached);
        if (Array.isArray(parsed)) {
          return sanitizeForRPC({ success: true, data: parsed, fromCache: true });
        }
      }
    } catch (eCache) {}

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('ABSENSI');
    if (!sheet) return sanitizeForRPC({ success: true, data: [] });

    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return sanitizeForRPC({ success: true, data: [] });

    const headers = values[0];
    const headersClean = headers.map(h => String(h || '').toUpperCase().replace(/[^A-Z0-9]/g, ''));
    let iSiswa = headersClean.findIndex(h => h === 'IDSISWA' || h === 'SISWA');
    if (iSiswa === -1) iSiswa = 3;
    let iDeleted = headersClean.findIndex(h => h.includes('ISDELETED') || h.includes('HAPUS'));
    if (iDeleted === -1) iDeleted = headersClean.length - 1;

    const result = [];

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const isDeleted = String(row[iDeleted] || '').toUpperCase() === 'TRUE';
      const rSiswa = String(row[iSiswa] || '').trim().toLowerCase();

      if (!isDeleted && rSiswa === targetSiswa) {
        let item = {};
        headers.forEach((h, idx) => {
          let val = row[idx];
          if (val instanceof Date) {
            val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd/MM/yyyy');
          }
          item[String(h)] = (val !== undefined && val !== null) ? String(val) : '';
        });
        result.push(item);
      }
    }

    // 2. Simpan ke server-side cache (TTL 60 detik)
    try {
      CacheService.getScriptCache().put(cacheKey, JSON.stringify(result), 60);
    } catch (ePut) {}

    return sanitizeForRPC({ success: true, data: result });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: err.toString(), data: [] });
  }
}

function submitPengajuanIzinSakit(payload, authToken) {
  try {
    const token = authToken || (payload && (payload.__token || payload.token));
    const auth = requireAuth_(token, null);
    if (!auth.authorized) return auth.error;

    if (!payload || !payload.idSiswa) {
      return sanitizeForRPC({ success: false, message: "ID Siswa tidak valid." });
    }

    if (!requireOwnership_(auth.session, payload.idSiswa)) {
      return sanitizeForRPC({ success: false, message: "Tidak diizinkan mengajukan izin untuk siswa lain." });
    }

    const rawStatus = String(payload.statusKehadiran || 'IZIN').trim().toUpperCase();
    if (rawStatus !== 'IZIN' && rawStatus !== 'SAKIT') {
      return sanitizeForRPC({ success: false, message: "Status pengajuan tidak valid. Hanya diperbolehkan IZIN atau SAKIT." });
    }
    const statusKehadiran = rawStatus;

    if (!payload.pesan || !String(payload.pesan).trim()) {
      return sanitizeForRPC({ success: false, message: "Pesan / alasan detail ketidakhadiran wajib diisi." });
    }

    // Batasi ukuran base64 foto surat di server (maksimal 2.0 MB)
    if (payload.fotoBase64) {
      const approxBytes = Math.floor(String(payload.fotoBase64).length * 0.75);
      if (approxBytes > 2 * 1024 * 1024) {
        return sanitizeForRPC({ success: false, message: "Ukuran lampiran foto surat melebihi batas 2.0 MB." });
      }
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let absensiSheet = ss.getSheetByName('ABSENSI');
    const headers = [
      'ID ABSENSI', 'ID JADWAL', 'TANGGAL', 'ID SISWA', 'NAMA SISWA',
      'KELAS', 'STATUS KEHADIRAN', 'CATATAN', 'VERIFIKATOR', 'WAKTU RECORD', 'IS DELETED'
    ];

    if (!absensiSheet) {
      absensiSheet = ss.insertSheet('ABSENSI');
      absensiSheet.appendRow(headers);
    }

    const absData = absensiSheet.getDataRange().getValues();
    const absHeadersClean = absData[0].map(h => String(h || '').toUpperCase().replace(/[^A-Z0-9]/g, ''));
    let iJadwal = absHeadersClean.findIndex(h => h === 'IDJADWAL');
    if (iJadwal === -1) iJadwal = 1;
    let iSiswa = absHeadersClean.findIndex(h => h === 'IDSISWA' || h === 'SISWA');
    if (iSiswa === -1) iSiswa = 3;
    let iDeleted = absHeadersClean.findIndex(h => h.includes('ISDELETED') || h.includes('HAPUS'));
    if (iDeleted === -1) iDeleted = absHeadersClean.length - 1;

    const targetSiswa = String(payload.idSiswa || '').trim().toLowerCase();
    const targetJadwal = String(payload.idJadwal || '').trim().toLowerCase();

    if (absData.length > 1) {
      for (let i = 1; i < absData.length; i++) {
        const isDeleted = String(absData[i][iDeleted] || '').toUpperCase() === 'TRUE';
        const rJadwal = String(absData[i][iJadwal] || '').trim().toLowerCase();
        const rSiswa = String(absData[i][iSiswa] || '').trim().toLowerCase();

        if (!isDeleted && rSiswa === targetSiswa && rJadwal === targetJadwal) {
          return sanitizeForRPC({
            success: false,
            message: "Peringatan: Anda sudah pernah melakukan presensi/pengajuan pada sesi latihan ini. Presensi hanya diperbolehkan 1 kali per sesi."
          });
        }
      }
    }

    // Validasi sesi latihan dari Sheet JADWAL & kecocokan kelas siswa
    if (!targetJadwal) {
      return sanitizeForRPC({ success: false, message: "Pilih sesi latihan yang valid." });
    }

    const jadwalSheet = ss.getSheetByName('JADWAL');
    if (!jadwalSheet) {
      return sanitizeForRPC({ success: false, message: "Sheet JADWAL tidak ditemukan di database." });
    }

    const jValues = jadwalSheet.getDataRange().getValues();
    if (jValues.length <= 1) {
      return sanitizeForRPC({ success: false, message: "Data jadwal sesi latihan kosong." });
    }

    const jHeadersClean = jValues[0].map(h => String(h || '').toUpperCase().replace(/[^A-Z0-9]/g, ''));
    let jIdIdx = jHeadersClean.findIndex(h => h === 'IDJADWAL' || h === 'ID');
    if (jIdIdx === -1) jIdIdx = 0;
    let jTglIdx = jHeadersClean.findIndex(h => h === 'TANGGAL');
    let jKelasIdx = jHeadersClean.findIndex(h => h === 'KELAS' || h === 'PROGRAM');
    let jStatusIdx = jHeadersClean.findIndex(h => h === 'STATUS');
    let jDeletedIdx = jHeadersClean.findIndex(h => h.includes('ISDELETED') || h.includes('HAPUS'));

    let jadwalFound = false;
    let jadwalRow = null;
    let tanggalFinal = payload.tanggal || '-';
    let kelasFinal = payload.kelas || '-';
    let rawTglVal = null;

    for (let j = 1; j < jValues.length; j++) {
      let jId = String(jValues[j][jIdIdx] || '').trim().toLowerCase();
      if (jId === targetJadwal) {
        jadwalFound = true;
        jadwalRow = jValues[j];
        break;
      }
    }

    if (!jadwalFound || !jadwalRow) {
      return sanitizeForRPC({ success: false, message: "Sesi latihan tidak ditemukan di Sheet JADWAL. Silakan muat ulang dan pilih sesi lain." });
    }

    // Validasi Status Aktif & IS DELETED
    const isJadwalDeleted = jDeletedIdx > -1 && String(jadwalRow[jDeletedIdx] || '').trim().toUpperCase() === 'TRUE';
    const jadwalStatus = jStatusIdx > -1 ? String(jadwalRow[jStatusIdx] || '').trim().toUpperCase() : 'AKTIF';
    if (isJadwalDeleted || (jadwalStatus && jadwalStatus !== 'AKTIF')) {
      return sanitizeForRPC({ success: false, message: "Sesi latihan ini sudah tidak aktif atau dibatalkan." });
    }

    if (jTglIdx > -1 && jadwalRow[jTglIdx]) {
      rawTglVal = jadwalRow[jTglIdx];
      if (rawTglVal instanceof Date) {
        tanggalFinal = Utilities.formatDate(rawTglVal, Session.getScriptTimeZone(), 'dd/MM/yyyy');
      } else {
        tanggalFinal = String(rawTglVal).trim();
      }
    }
    if (jKelasIdx > -1 && jadwalRow[jKelasIdx]) {
      kelasFinal = String(jadwalRow[jKelasIdx]).trim();
    }

    // Validasi Tanggal Tidak Lampau (boleh hari ini atau mendatang)
    try {
      let tglJadwalDate = null;
      if (rawTglVal instanceof Date) {
        tglJadwalDate = new Date(rawTglVal.getTime());
      } else if (tanggalFinal && tanggalFinal !== '-') {
        const partsSlash = tanggalFinal.split('/');
        const partsDash = tanggalFinal.split('-');
        if (partsSlash.length === 3) {
          tglJadwalDate = new Date(parseInt(partsSlash[2], 10), parseInt(partsSlash[1], 10) - 1, parseInt(partsSlash[0], 10));
        } else if (partsDash.length === 3) {
          tglJadwalDate = new Date(parseInt(partsDash[0], 10), parseInt(partsDash[1], 10) - 1, parseInt(partsDash[2], 10));
        }
      }

      if (tglJadwalDate && !isNaN(tglJadwalDate.getTime())) {
        const today = new Date();
        const startOfToday = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 0, 0, 0, 0);
        const startOfJadwal = new Date(tglJadwalDate.getFullYear(), tglJadwalDate.getMonth(), tglJadwalDate.getDate(), 0, 0, 0, 0);
        if (startOfJadwal < startOfToday) {
          return sanitizeForRPC({
            success: false,
            message: "Pengajuan izin tidak diperkenankan untuk sesi latihan yang sudah lewat (" + tanggalFinal + ")."
          });
        }
      }
    } catch (eDate) {}

    // Validasi Kecocokan Kelas Siswa
    if (kelasFinal && kelasFinal !== '-' && kelasFinal.toLowerCase() !== 'semua kelas') {
      try {
        const bioRes = fetchBiodataSiswa(payload.idSiswa);
        const studentKelas = (bioRes && bioRes.biodata && bioRes.biodata.kelas) ? String(bioRes.biodata.kelas).trim() : '';
        if (studentKelas && studentKelas !== '-' && studentKelas.toLowerCase() !== 'semua kelas') {
          const kTarget = kelasFinal.toLowerCase();
          const kSiswa = studentKelas.toLowerCase();
          const matchKelas = (kTarget === kSiswa) || kTarget.includes(kSiswa) || kSiswa.includes(kTarget);
          if (!matchKelas) {
            return sanitizeForRPC({
              success: false,
              message: "Sesi latihan ini diperuntukkan bagi kelas '" + kelasFinal + "', sedangkan Anda terdaftar di kelas '" + studentKelas + "'."
            });
          }
        }
      } catch (eBio) {}
    }

    let urlFotoSurat = '';
    if (payload.fotoBase64) {
      try {
        const folderId = CONFIG.FOLDER_BUKTI_ID || CONFIG.FOLDER_GALERI_ID;
        const folder = DriveApp.getFolderById(folderId);
        const bytes = Utilities.base64Decode(payload.fotoBase64);
        const fileName = 'SURAT_' + payload.idSiswa + '_' + Date.now() + '.jpg';
        const blob = Utilities.newBlob(bytes, 'image/jpeg', fileName);
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        urlFotoSurat = 'https://lh3.googleusercontent.com/d/' + file.getId();
      } catch (fErr) {
        urlFotoSurat = '';
      }
    }

    let catatanLengkap = payload.pesan || '';
    if (urlFotoSurat) {
      catatanLengkap += ' [Lampiran Surat: ' + urlFotoSurat + ']';
    }

    const idAbsensi = 'ABS-REQ-' + Date.now();
    const waktuRecord = new Date().toLocaleString('id-ID');

    absensiSheet.appendRow([
      idAbsensi,
      payload.idJadwal || '-',
      tanggalFinal,
      payload.idSiswa,
      payload.namaSiswa || '-',
      kelasFinal,
      statusKehadiran,
      catatanLengkap,
      'Pending Verification (Pengajuan Online)',
      waktuRecord,
      'FALSE'
    ]);

    logActivity(payload.idSiswa, 'PENGAJUAN_IZIN', 'ABSENSI', 'Siswa mengajukan ' + statusKehadiran + ' untuk jadwal: ' + (payload.idJadwal || '-'));

    return sanitizeForRPC({
      success: true,
      message: "Pengajuan " + statusKehadiran + " berhasil dikirim! Status telah tercatat di Sheet ABSENSI.",
      idAbsensi: idAbsensi,
      tanggal: tanggalFinal,
      kelas: kelasFinal
    });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal mengajukan izin/sakit: " + err.toString() });
  }
}

/**
 * Membatalkan pengajuan izin/sakit siswa yang masih berstatus 'Pending Verification'
 * Menggunakan soft delete (IS DELETED = 'TRUE') untuk menjaga integritas data.
 */
function batalkanPengajuanIzin(idAbsensi, idSiswa, authToken) {
  try {
    const auth = requireAuth_(authToken, null);
    if (!auth.authorized) return auth.error;

    if (!requireOwnership_(auth.session, idSiswa)) {
      return sanitizeForRPC({ success: false, message: "Tidak diizinkan membatalkan izin siswa lain." });
    }

    const targetAbs = String(idAbsensi || '').trim().toLowerCase();
    const targetSiswa = String(idSiswa || '').trim().toLowerCase();

    if (!targetAbs || !targetSiswa) {
      return sanitizeForRPC({ success: false, message: "ID Absensi dan ID Siswa wajib diisi." });
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('ABSENSI');
    if (!sheet || sheet.getLastRow() <= 1) {
      return sanitizeForRPC({ success: false, message: "Data presensi tidak ditemukan." });
    }

    const data = sheet.getDataRange().getValues();
    const headersClean = data[0].map(h => String(h || '').toUpperCase().replace(/[^A-Z0-9]/g, ''));

    let idAbsIdx = headersClean.findIndex(h => h === 'IDABSENSI' || h === 'ID');
    if (idAbsIdx === -1) idAbsIdx = 0;

    let idSiswaIdx = headersClean.findIndex(h => h === 'IDSISWA' || h === 'SISWA');
    if (idSiswaIdx === -1) idSiswaIdx = 3;

    let vfIdx = headersClean.findIndex(h => h.includes('VERIFIKATOR'));
    if (vfIdx === -1) vfIdx = 8;

    let delIdx = headersClean.findIndex(h => h.includes('ISDELETED') || h.includes('HAPUS'));
    if (delIdx === -1) delIdx = headersClean.length - 1;

    let matchRow = -1;
    for (let i = 1; i < data.length; i++) {
      let rAbs = String(data[i][idAbsIdx] || '').trim().toLowerCase();
      let rSiswa = String(data[i][idSiswaIdx] || '').trim().toLowerCase();
      let isDeleted = String(data[i][delIdx] || '').toUpperCase() === 'TRUE';

      if (!isDeleted && rAbs === targetAbs && rSiswa === targetSiswa) {
        matchRow = i + 1;
        let rVf = String(data[i][vfIdx] || '').toLowerCase();
        if (!rVf.includes('pending') && !rVf.includes('pengajuan online') && !rVf.includes('online request')) {
          return sanitizeForRPC({
            success: false,
            message: "Pengajuan presensi ini telah diverifikasi oleh Admin dan tidak dapat dibatalkan secara mandiri."
          });
        }
        break;
      }
    }

    if (matchRow === -1) {
      return sanitizeForRPC({ success: false, message: "Catatan pengajuan presensi aktif tidak ditemukan." });
    }

    sheet.getRange(matchRow, delIdx + 1).setValue('TRUE');

    logActivity(idSiswa, 'BATAL_PENGAJUAN', 'ABSENSI', 'Siswa membatalkan pengajuan izin: ' + idAbsensi);

    return sanitizeForRPC({
      success: true,
      message: "Pengajuan izin/sakit berhasil dibatalkan."
    });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal membatalkan pengajuan: " + err.toString() });
  }
}

/**
 * Menolak pengajuan izin/sakit siswa oleh Admin/Super Admin
 */
function rejectPengajuanIzin(payload, authToken) {
  try {
    const token = authToken || (payload && (payload.__token || payload.token));
    const auth = requireAuth_(token, ['ADMIN', 'SUPER_ADMIN']);
    if (!auth.authorized) return auth.error;

    if (String(auth.session.role).toUpperCase() === 'ADMIN' && !canEditModule_(auth.session, 'absensi')) {
      return sanitizeForRPC({ success: false, message: "Mode baca saja (Readonly). Anda tidak memiliki izin mengubah modul Absensi." });
    }

    const idAbsensi = String((payload && (payload.idAbsensi || payload.id)) || '').trim();
    const alasan = String((payload && payload.alasan) || '').trim();
    const targetJadwal = String((payload && (payload.idJadwal || payload.jadwal)) || '').trim().toLowerCase();
    const targetSiswa = String((payload && (payload.idSiswa || payload.siswa)) || '').trim().toLowerCase();

    if (!idAbsensi && !(targetJadwal && targetSiswa)) {
      return sanitizeForRPC({ success: false, message: "ID Pengajuan Presensi tidak valid." });
    }
    if (!alasan) {
      return sanitizeForRPC({ success: false, message: "Alasan penolakan wajib diisi." });
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('ABSENSI');
    if (!sheet || sheet.getLastRow() <= 1) {
      return sanitizeForRPC({ success: false, message: "Sheet ABSENSI kosong atau tidak ditemukan." });
    }

    const data = sheet.getDataRange().getValues();
    const headersClean = data[0].map(h => String(h || '').toUpperCase().replace(/[^A-Z0-9]/g, ''));

    let idAbsIdx = headersClean.findIndex(h => h === 'IDABSENSI' || h === 'ID');
    if (idAbsIdx === -1) idAbsIdx = 0;

    let idJadwalIdx = headersClean.findIndex(h => h === 'IDJADWAL');
    if (idJadwalIdx === -1) idJadwalIdx = 1;

    let idSiswaIdx = headersClean.findIndex(h => h === 'IDSISWA' || h === 'SISWA');
    if (idSiswaIdx === -1) idSiswaIdx = 3;

    let catatanIdx = headersClean.findIndex(h => h.includes('CATATAN') || h.includes('KETERANGAN'));
    if (catatanIdx === -1) catatanIdx = 7;

    let vfIdx = headersClean.findIndex(h => h.includes('VERIFIKATOR'));
    if (vfIdx === -1) vfIdx = 8;

    let waktuIdx = headersClean.findIndex(h => h.includes('WAKTU') || h.includes('RECORD'));
    if (waktuIdx === -1) waktuIdx = 9;

    let delIdx = headersClean.findIndex(h => h.includes('ISDELETED') || h.includes('HAPUS'));
    if (delIdx === -1) delIdx = headersClean.length - 1;

    let targetAbs = idAbsensi.toLowerCase();
    let matchRow = -1;
    let oldCatatan = '';

    for (let i = 1; i < data.length; i++) {
      let rAbs = String(data[i][idAbsIdx] || '').trim().toLowerCase();
      let rJadwal = String(data[i][idJadwalIdx] || '').trim().toLowerCase();
      let rSiswa = String(data[i][idSiswaIdx] || '').trim().toLowerCase();
      let isDeleted = String(data[i][delIdx] || '').toUpperCase() === 'TRUE';

      const matchById = targetAbs && (rAbs === targetAbs);
      const matchByJadwalSiswa = targetJadwal && targetSiswa && (rJadwal === targetJadwal && rSiswa === targetSiswa);

      if (!isDeleted && (matchById || matchByJadwalSiswa)) {
        matchRow = i + 1;
        oldCatatan = String(data[i][catatanIdx] || '');
        let rVf = String(data[i][vfIdx] || '').toLowerCase();
        if (!rVf.includes('pending') && !rVf.includes('pengajuan online') && !rVf.includes('online request')) {
          return sanitizeForRPC({
            success: false,
            message: "Pengajuan ini tidak lagi berstatus pending (sudah diverifikasi/diproses sebelumnya)."
          });
        }
        break;
      }
    }

    if (matchRow === -1) {
      return sanitizeForRPC({ success: false, message: "Pengajuan izin tidak ditemukan di database." });
    }

    const adminName = (auth.session && (auth.session.nama || auth.session.email)) || 'Admin Sanggar';
    const updatedCatatan = oldCatatan ? (oldCatatan + ' [DITOLAK: ' + alasan + ']') : ('[DITOLAK: ' + alasan + ']');
    const nowStr = new Date().toLocaleString('id-ID');

    sheet.getRange(matchRow, vfIdx + 1).setValue('DITOLAK - ' + adminName);
    sheet.getRange(matchRow, catatanIdx + 1).setValue(updatedCatatan);
    sheet.getRange(matchRow, waktuIdx + 1).setValue(nowStr);

    bumpAbsensiCacheVer();

    const actorId = auth.session ? (auth.session.uid || auth.session.email || 'ADMIN') : 'ADMIN';
    logActivity(actorId, 'REJECT_PENGAJUAN_IZIN', 'ABSENSI', 'Tolak pengajuan izin ' + (idAbsensi || (targetJadwal + '/' + targetSiswa)) + ': ' + alasan);

    return sanitizeForRPC({
      success: true,
      message: "Pengajuan izin berhasil ditolak."
    });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal menolak pengajuan: " + err.toString() });
  }
}

function savePengumuman(payload, authToken) {
  try {
    const token = authToken || (payload && (payload.__token || payload.token));
    const auth = requireAuth_(token, ['ADMIN', 'SUPER_ADMIN']);
    if (!auth.authorized) return auth.error;
    const actorId = (auth.session && (auth.session.uid || auth.session.email)) || 'ADMIN';

    // Validasi input backend
    const judul = String((payload && payload.judul) || '').trim();
    if (!judul) {
      return sanitizeForRPC({ success: false, message: "Judul pengumuman wajib diisi!" });
    }
    if (judul.length > 500) {
      return sanitizeForRPC({ success: false, message: "Judul pengumuman maksimal 500 karakter!" });
    }
    const isi = String((payload && payload.isi) || '').trim();
    if (!isi) {
      return sanitizeForRPC({ success: false, message: "Isi pengumuman wajib diisi!" });
    }
    if (isi.length > 10000) {
      return sanitizeForRPC({ success: false, message: "Isi pengumuman maksimal 10.000 karakter!" });
    }
    const kategori = String((payload && payload.kategori) || '').trim() || 'Pengumuman';
    const tanggal = String((payload && payload.tanggal) || '').trim();
    if (!tanggal) {
      return sanitizeForRPC({ success: false, message: "Tanggal pengumuman wajib diisi!" });
    }
    const rawStatus = String((payload && payload.status) || 'AKTIF').trim().toUpperCase();
    const status = (rawStatus === 'NONAKTIF') ? 'NONAKTIF' : 'AKTIF';

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName('PENGUMUMAN');
    const headers = ['ID PENGUMUMAN', 'JUDUL', 'KATEGORI', 'TANGGAL', 'ISI', 'STATUS', 'CREATED_AT', 'IS DELETED'];
    if (!sheet) {
      sheet = ss.insertSheet('PENGUMUMAN');
      sheet.appendRow(headers);
    }

    const data = sheet.getDataRange().getValues();
    const targetId = String((payload && payload.idPengumuman) || '').trim().toLowerCase();

    let existingRow = -1;
    if (data.length > 1) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0] || '').trim().toLowerCase() === targetId) {
          existingRow = i + 1;
          break;
        }
      }
    }

    if (existingRow > -1) {
      sheet.getRange(existingRow, 2).setValue(judul);
      sheet.getRange(existingRow, 3).setValue(kategori);
      sheet.getRange(existingRow, 4).setValue(tanggal);
      sheet.getRange(existingRow, 5).setValue(isi);
      sheet.getRange(existingRow, 6).setValue(status);
      sheet.getRange(existingRow, 8).setValue('FALSE');
    } else {
      sheet.appendRow([
        payload.idPengumuman || ('PNG-' + new Date().getTime()),
        judul,
        kategori,
        tanggal,
        isi,
        status,
        new Date().toISOString(),
        'FALSE'
      ]);
    }

    logActivity(actorId, 'SAVE_PENGUMUMAN', 'PENGUMUMAN', 'Simpan pengumuman: ' + judul);
    return sanitizeForRPC({ success: true, message: "Pengumuman berhasil disimpan ke Sheet PENGUMUMAN!" });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal menyimpan pengumuman: " + err.toString() });
  }
}

function deletePengumuman(idPengumuman, authToken) {
  try {
    const auth = requireAuth_(authToken, ['ADMIN', 'SUPER_ADMIN']);
    if (!auth.authorized) return auth.error;
    const actorId = (auth.session && (auth.session.uid || auth.session.email)) || 'ADMIN';

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('PENGUMUMAN');
    if (!sheet) return sanitizeForRPC({ success: false, message: "Sheet PENGUMUMAN tidak ditemukan!" });

    const data = sheet.getDataRange().getValues();
    const targetId = String(idPengumuman || '').trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim().toLowerCase() === targetId) {
        // Soft delete: tandai IS DELETED = 'TRUE'
        sheet.getRange(i + 1, 8).setValue('TRUE');
        logActivity(actorId, 'DELETE_PENGUMUMAN', 'PENGUMUMAN', 'Hapus pengumuman: ' + idPengumuman);
        return sanitizeForRPC({ success: true, message: "Pengumuman berhasil dihapus dari Sheet PENGUMUMAN." });
      }
    }

    return sanitizeForRPC({ success: false, message: "ID Pengumuman " + idPengumuman + " tidak ditemukan." });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal menghapus pengumuman: " + err.toString() });
  }
}

function fetchFormOptions() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    
    let referensi = [];
    let referensiKeuangan = [];
    const refSheet = ss.getSheetByName('REFERENSI');
    if (refSheet) {
      const refData = refSheet.getDataRange().getValues();
      if (refData.length > 1) {
        const h = refData[0].map(x => String(x).trim().toUpperCase());
        const kIdx = h.indexOf('KELAS');
        const katIdx = h.indexOf('KATEGORI');
        const jenisIdx = h.findIndex(x => x.includes('JENIS'));
        const nomIdx = h.findIndex(x => x.includes('NOMINAL'));

        for (let i = 1; i < refData.length; i++) {
          if (kIdx > -1 && katIdx > -1 && refData[i][kIdx]) {
            referensi.push({ kelas: String(refData[i][kIdx]), kategori: String(refData[i][katIdx]) });
          }
          if (jenisIdx > -1 && nomIdx > -1 && refData[i][jenisIdx]) {
            let jVal = String(refData[i][jenisIdx]).trim();
            let nVal = Number(refData[i][nomIdx]) || 0;
            if (jVal) {
              referensiKeuangan.push({ 'JENIS TRANSAKSI': jVal, 'NOMINAL PEMBAYARAN': nVal });
            }
          }
        }
      }
    }

    if (referensiKeuangan.length === 0) {
      referensiKeuangan = [
        { 'JENIS TRANSAKSI': 'PENDAFTARAN', 'NOMINAL PEMBAYARAN': 85000 },
        { 'JENIS TRANSAKSI': 'INFAK BULANAN', 'NOMINAL PEMBAYARAN': 150000 },
        { 'JENIS TRANSAKSI': 'SUKARELA', 'NOMINAL PEMBAYARAN': 50000 }
      ];
    }

    let alamat = {};
    const alSheet = ss.getSheetByName('ALAMAT');
    if (alSheet) {
      const alData = alSheet.getDataRange().getValues();
      if (alData.length > 1) {
        const h = alData[0].map(x => String(x).trim().toUpperCase());
        const pIdx = h.indexOf('PROVINSI');
        const kbIdx = h.indexOf('KABUPATEN');
        const kcIdx = h.indexOf('KECAMATAN');
        const klIdx = h.indexOf('KELURAHAN');
        
        if (pIdx > -1) {
          for (let i = 1; i < alData.length; i++) {
            let prov = String(alData[i][pIdx] || '');
            let kab = kbIdx > -1 ? String(alData[i][kbIdx] || '') : '';
            let kec = kcIdx > -1 ? String(alData[i][kcIdx] || '') : '';
            let kel = klIdx > -1 ? String(alData[i][klIdx] || '') : '';
            
            if (prov) {
              if (!alamat[prov]) alamat[prov] = {};
              if (kab) {
                if (!alamat[prov][kab]) alamat[prov][kab] = {};
                if (kec) {
                  if (!alamat[prov][kab][kec]) alamat[prov][kab][kec] = [];
                  if (kel && !alamat[prov][kab][kec].includes(kel)) {
                    alamat[prov][kab][kec].push(kel);
                  }
                }
              }
            }
          }
        }
      }
    }

    return sanitizeForRPC({ success: true, referensi: referensi, alamat: alamat, referensiKeuangan: referensiKeuangan });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: err.toString() });
  }
}

function submitBiodataPendaftaran(payload) {
  const lock = LockService.getScriptLock();
  try {
    const lockAcquired = lock.tryLock(15000);
    if (!lockAcquired) {
      return sanitizeForRPC({ success: false, message: "Sistem sedang sibuk. Silakan coba kembali beberapa saat lagi." });
    }

    const targetEmail = String((payload && payload.email) || '').trim().toLowerCase();
    if (!targetEmail) return sanitizeForRPC({ success: false, message: "Email wajib diisi!" });
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(targetEmail)) {
      return sanitizeForRPC({ success: false, message: "Format email tidak valid!" });
    }

    const targetNama = String((payload && payload.nama) || '').trim();
    if (!targetNama) return sanitizeForRPC({ success: false, message: "Nama lengkap siswa wajib diisi!" });

    let cleanHp = String((payload && payload.noHp) || '').replace(/[^0-9]/g, '');
    if (cleanHp.length < 8) {
      return sanitizeForRPC({ success: false, message: "Nomor HP/WhatsApp tidak valid (minimal 8 digit angka)!" });
    }
    if (cleanHp.startsWith('62')) cleanHp = cleanHp.substring(2);
    if (cleanHp.startsWith('0')) cleanHp = cleanHp.substring(1);
    const formattedHp = "+62" + cleanHp;

    const targetKelas = String((payload && (payload.kelas || payload.program)) || '').trim();
    if (!targetKelas) return sanitizeForRPC({ success: false, message: "Pilihan program/kelas wajib diisi!" });

    // A5: Rate limiting via CacheService
    const cache = CacheService.getScriptCache();
    const globalCountKey = 'REG_RATE_GLOBAL';
    const emailHashKey = 'REG_RATE_MAIL_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, targetEmail)).substring(0, 16);

    const currGlobal = parseInt(cache.get(globalCountKey) || '0', 10);
    if (currGlobal >= 60) {
      return sanitizeForRPC({ success: false, message: "Sistem pendaftaran sedang padat. Silakan tunggu beberapa menit lalu coba lagi." });
    }

    const currEmail = parseInt(cache.get(emailHashKey) || '0', 10);
    if (currEmail >= 5) {
      return sanitizeForRPC({ success: false, message: "Terlalu banyak permintaan untuk email ini. Silakan coba lagi dalam 1 jam." });
    }

    cache.put(globalCountKey, (currGlobal + 1).toString(), 3600);
    cache.put(emailHashKey, (currEmail + 1).toString(), 3600);

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

    // A2: Cek duplikasi di Sheet DAFTAR dengan toleransi baris DITOLAK
    let daftarSheet = ss.getSheetByName('DAFTAR');
    let existingDitolakRow = -1;
    let existingDitolakNo = null;

    if (daftarSheet && daftarSheet.getLastRow() > 1) {
      const daftarValues = daftarSheet.getDataRange().getValues();
      const daftarHeaders = daftarValues[0].map(h => String(h || '').toUpperCase().trim());
      const emailColIdx = daftarHeaders.indexOf('EMAIL');
      const statusColIdx = daftarHeaders.indexOf('STATUS');

      if (emailColIdx > -1) {
        for (let i = 1; i < daftarValues.length; i++) {
          const existingEmail = String(daftarValues[i][emailColIdx] || '').trim().toLowerCase();
          if (existingEmail === targetEmail) {
            const st = statusColIdx > -1 ? String(daftarValues[i][statusColIdx] || '').trim().toUpperCase() : '';
            if (st === 'DITOLAK') {
              existingDitolakRow = i + 1;
              existingDitolakNo = daftarValues[i][0];
            } else {
              return sanitizeForRPC({ success: false, message: "Pendaftaran Gagal: Email '" + targetEmail + "' sudah pernah didaftarkan dan sedang ditinjau." });
            }
          }
        }
      }
    }

    // A2: Cek akun di Sheet USER (blokir hanya jika akun AKTIF)
    let userSheet = ss.getSheetByName('USER') || ss.getSheetByName('USERS');
    if (userSheet && userSheet.getLastRow() > 1) {
      const userValues = userSheet.getDataRange().getValues();
      const uHeaders = userValues[0].map(h => String(h || '').toLowerCase().replace(/[^a-z0-9]/g, ''));
      let emailIdx = uHeaders.findIndex(h => h === 'email' || h === 'surel');
      if (emailIdx === -1) emailIdx = 1;
      let statusIdx = uHeaders.findIndex(h => h === 'status');
      if (statusIdx === -1) statusIdx = 5;
      let delIdx = uHeaders.findIndex(h => h.includes('delete') || h.includes('hapus'));

      for (let i = 1; i < userValues.length; i++) {
        const existingUserEmail = String(userValues[i][emailIdx] || '').trim().toLowerCase();
        if (existingUserEmail === targetEmail) {
          const uStatus = String(userValues[i][statusIdx] || '').trim().toUpperCase();
          const uDeleted = delIdx > -1 ? String(userValues[i][delIdx] || '').trim().toUpperCase() : 'FALSE';
          if (uDeleted !== 'TRUE' && (uStatus === 'AKTIF' || uStatus === 'ACTIVE')) {
            return sanitizeForRPC({ success: false, message: "Pendaftaran Gagal: Email '" + targetEmail + "' sudah memiliki akun aktif di sistem. Silakan langsung login." });
          }
        }
      }
    }

    if (!daftarSheet) {
      daftarSheet = ss.insertSheet('DAFTAR');
      daftarSheet.appendRow(['NO', 'NAMA SISWA', 'TEMPAT LAHIR', 'TANGGAL LAHIR', 'USIA', 'ASAL SEKOLAH', 'KELAS', 'KATEGORI', 'NAMA ORANGTUA', 'NOMOR HP', 'EMAIL', 'PASSWORD', 'ALAMAT', 'PROVINSI', 'KABUPATEN', 'KECAMATAN', 'KELURAHAN', 'FOTO', 'WAKTU DAFTAR', 'STATUS']);
    }

    const dValues = daftarSheet.getDataRange().getValues();
    let dHeaders = dValues[0].map(h => String(h || '').toUpperCase().trim());
    let passColIdx = dHeaders.indexOf('PASSWORD');
    if (passColIdx === -1) {
      passColIdx = dHeaders.length;
      daftarSheet.getRange(1, passColIdx + 1).setValue('PASSWORD');
      dHeaders.push('PASSWORD');
    }
    let statusColIdx = dHeaders.indexOf('STATUS');
    if (statusColIdx === -1) {
      statusColIdx = dHeaders.length;
      daftarSheet.getRange(1, statusColIdx + 1).setValue('STATUS');
      dHeaders.push('STATUS');
    }

    let no;
    if (existingDitolakRow > -1 && existingDitolakNo) {
      no = existingDitolakNo;
    } else {
      const props = PropertiesService.getScriptProperties();
      let lastSeq = parseInt(props.getProperty('GLOBAL_SEQ_SISWA') || props.getProperty('GLOBAL_SEQ_DAFTAR') || '0', 10);
      if (isNaN(lastSeq) || lastSeq < 0) lastSeq = 0;

      // 1. Cek nomor urut tertinggi di Sheet DAFTAR
      if (daftarSheet.getLastRow() > 1) {
        const dataIds = daftarSheet.getRange(2, 1, daftarSheet.getLastRow() - 1, 1).getValues();
        for (let r = 0; r < dataIds.length; r++) {
          let currNo = parseInt(dataIds[r][0], 10);
          if (!isNaN(currNo) && currNo > lastSeq) lastSeq = currNo;
        }
      }

      // 2. Cek nomor urut tertinggi & 4 digit terakhir ID di Sheet BIODATA agar nomor urut pendaftaran tidak pernah reset/duplikat
      let bioSheetCheck = ss.getSheetByName('BIODATA');
      if (bioSheetCheck && bioSheetCheck.getLastRow() > 1) {
        const bioDataIds = bioSheetCheck.getDataRange().getValues();
        const bioH = bioDataIds[0].map(h => String(h || '').toUpperCase().trim());
        let bioNoIdx = bioH.indexOf('NO');
        let bioIdIdx = bioH.findIndex(h => h === 'ID SISWA' || h === 'IDSISWA' || h === 'ID');

        for (let r = 1; r < bioDataIds.length; r++) {
          if (bioNoIdx > -1) {
            let currBioNo = parseInt(bioDataIds[r][bioNoIdx], 10);
            if (!isNaN(currBioNo) && currBioNo > lastSeq) lastSeq = currBioNo;
          }
          if (bioIdIdx > -1) {
            let bIdStr = String(bioDataIds[r][bioIdIdx] || '').trim();
            if (bIdStr.length >= 4) {
              let sNum = parseInt(bIdStr.slice(-4), 10);
              if (!isNaN(sNum) && sNum > lastSeq) lastSeq = sNum;
            }
          }
        }
      }

      no = lastSeq + 1;
      props.setProperty('GLOBAL_SEQ_DAFTAR', no.toString());
      props.setProperty('GLOBAL_SEQ_SISWA', no.toString());
    }

    let fileUrl = "";

    // A5: Validasi ukuran foto & upload Drive
    if (payload && payload.fotoBase64) {
      let cleanBase64 = String(payload.fotoBase64);
      if (cleanBase64.indexOf(',') !== -1) {
        cleanBase64 = cleanBase64.split(',')[1];
      }
      cleanBase64 = cleanBase64.replace(/[^A-Za-z0-9+/=]/g, "");

      if (cleanBase64.length > 5500000) {
        return sanitizeForRPC({ success: false, message: "Ukuran foto terlalu besar (maksimal 4 MB). Silakan kompres atau pilih foto lain." });
      }

      let safeName = String(targetNama || 'Siswa').replace(/[^a-zA-Z0-9]/g, '_');
      let fileName = "REG_" + no + "_" + safeName + ".jpg";
      let blob = Utilities.newBlob(Utilities.base64Decode(cleanBase64), "image/jpeg", fileName);
      
      let file = null;
      let uploadSuccess = false;
      let driveError = "";

      try {
        let folderConfig = DriveApp.getFolderById(CONFIG.FOLDER_FOTO_ID);
        file = folderConfig.createFile(blob);
        uploadSuccess = true;
      } catch (err1) {
        try {
          const folders = DriveApp.getFoldersByName('FOTO');
          let folderBackup = folders.hasNext() ? folders.next() : DriveApp.createFolder('FOTO');
          file = folderBackup.createFile(blob);
          uploadSuccess = true;
        } catch (err2) {
          try {
            file = DriveApp.createFile(blob);
            uploadSuccess = true;
          } catch (err3) {
            driveError = err3.message;
          }
        }
      }

      if (uploadSuccess && file) {
        try {
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        } catch (shareErr) {}
        fileUrl = "https://lh3.googleusercontent.com/d/" + file.getId();
      } else {
        return sanitizeForRPC({ success: false, message: "Upload Foto Gagal. Akses Drive ditolak: " + driveError });
      }
    }

    // Amankan password yang dibuat pendaftar dengan hashing V2
    const rawPassInput = String((payload && payload.password) || '').trim();
    const studentPasswordHash = rawPassInput ? hashPasswordV2(rawPassInput) : hashPasswordV2('siswa123');

    // Susun baris pendaftaran sesuai urutan kolom Sheet DAFTAR
    const newDaftarRow = dHeaders.map(h => {
      if (h === 'NO') return no;
      if (h === 'NAMA SISWA' || h === 'NAMA') return targetNama;
      if (h === 'TEMPAT LAHIR') return payload ? payload.tempatLahir || '' : '';
      if (h === 'TANGGAL LAHIR') return payload ? payload.tglLahir || '' : '';
      if (h === 'USIA') return payload ? payload.usia || '' : '';
      if (h === 'ASAL SEKOLAH') return payload ? payload.asalSekolah || '' : '';
      if (h === 'KELAS' || h === 'PROGRAM') return targetKelas;
      if (h === 'KATEGORI' || h === 'LEVEL') return payload ? payload.kategori || '' : '';
      if (h === 'NAMA ORANGTUA' || h === 'NAMA ORTU') return payload ? payload.namaOrtu || '' : '';
      if (h === 'NOMOR HP' || h === 'NO HP') return formattedHp;
      if (h === 'EMAIL') return targetEmail;
      if (h === 'PASSWORD') return studentPasswordHash;
      if (h === 'ALAMAT') return payload ? payload.alamat || '' : '';
      if (h === 'PROVINSI') return payload ? payload.provinsi || '' : '';
      if (h === 'KABUPATEN' || h === 'KOTA') return payload ? payload.kabupaten || '' : '';
      if (h === 'KECAMATAN') return payload ? payload.kecamatan || '' : '';
      if (h === 'KELURAHAN' || h === 'DESA') return payload ? payload.kelurahan || '' : '';
      if (h === 'FOTO') return fileUrl;
      if (h === 'WAKTU DAFTAR' || h === 'TIMESTAMP') return new Date().toISOString();
      if (h === 'STATUS') return 'PENDING';
      return '';
    });

    if (existingDitolakRow > -1) {
      daftarSheet.getRange(existingDitolakRow, 1, 1, newDaftarRow.length).setValues([newDaftarRow]);
    } else {
      daftarSheet.appendRow(newDaftarRow);
    }

    // Daftarkan juga draf akun ke Sheet USER dengan status 'PENDING'
    try {
      if (!userSheet) {
        userSheet = ss.insertSheet('USER');
        userSheet.appendRow(['id_user', 'email', 'password', 'id_role', 'nama', 'status', 'created_at', 'updated_at', 'is_deleted']);
      }
      const uValues = userSheet.getDataRange().getValues();
      const uH = uValues[0].map(h => String(h || '').toLowerCase().replace(/[^a-z0-9]/g, ''));
      let uEIdx = uH.findIndex(h => h === 'email' || h === 'surel');
      if (uEIdx === -1) uEIdx = 1;
      let uPIdx = uH.findIndex(h => h.includes('pass') || h.includes('pwd') || h.includes('sandi'));
      if (uPIdx === -1) uPIdx = 2;
      let uNamaIdx = uH.findIndex(h => h === 'nama' || h === 'namalengkap');
      if (uNamaIdx === -1) uNamaIdx = 4;
      let uStatusIdx = uH.findIndex(h => h === 'status');
      if (uStatusIdx === -1) uStatusIdx = 5;
      let uUpdatedIdx = uH.findIndex(h => h.includes('update'));
      if (uUpdatedIdx === -1) uUpdatedIdx = 7;
      let uDelIdx = uH.findIndex(h => h.includes('delete') || h.includes('hapus'));
      if (uDelIdx === -1 && uValues[0].length >= 9) uDelIdx = 8;

      let uExisting = -1;
      for (let u = 1; u < uValues.length; u++) {
        if (String(uValues[u][uEIdx] || '').trim().toLowerCase() === targetEmail) {
          uExisting = u + 1;
          break;
        }
      }

      const nowIso = new Date().toISOString();
      if (uExisting > -1) {
        if (uPIdx > -1) userSheet.getRange(uExisting, uPIdx + 1).setValue(studentPasswordHash);
        if (uNamaIdx > -1) userSheet.getRange(uExisting, uNamaIdx + 1).setValue(targetNama);
        if (uStatusIdx > -1) userSheet.getRange(uExisting, uStatusIdx + 1).setValue('PENDING');
        if (uUpdatedIdx > -1) userSheet.getRange(uExisting, uUpdatedIdx + 1).setValue(nowIso);
        if (uDelIdx > -1) userSheet.getRange(uExisting, uDelIdx + 1).setValue('FALSE');
      } else {
        userSheet.appendRow([
          'REG-' + no,
          targetEmail,
          studentPasswordHash,
          'SISWA',
          targetNama,
          'PENDING',
          nowIso,
          nowIso,
          'FALSE'
        ]);
      }
    } catch (uErr) {
      Logger.log("Draft user note: " + uErr.message);
    }

    logActivity('SYSTEM', 'PENDAFTARAN', 'DAFTAR', 'Siswa mendaftar: ' + targetNama);

    return sanitizeForRPC({ 
      success: true, 
      message: "Formulir Pendaftaran berhasil disimpan! Silakan tunggu konfirmasi Admin.", 
      noUrut: no,
      photoUrl: fileUrl
    });
  } catch (error) {
    return sanitizeForRPC({ success: false, message: "Gagal menyimpan pendaftaran: " + error.toString() });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function fetchBiodataSiswa(idSiswa, authToken) {
  try {
    const auth = requireAuth_(authToken, null);
    if (!auth.authorized) return auth.error;
    if (!requireOwnership_(auth.session, idSiswa)) {
      return sanitizeForRPC({ success: false, message: "Akses data biodata siswa lain tidak diizinkan." });
    }

    const targetId = String(idSiswa || '').trim().toLowerCase();
    if (!targetId) {
      return sanitizeForRPC({ success: false, message: "ID Siswa wajib diisi." });
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const bioSheet = ss.getSheetByName('BIODATA');
    if (!bioSheet) return sanitizeForRPC({ success: false, message: "Sheet BIODATA tidak ditemukan!" });

    const bioData = bioSheet.getDataRange().getValues();
    if (bioData.length <= 1) return sanitizeForRPC({ success: false, message: "Data BIODATA kosong." });

    const headers = bioData[0];
    const headersClean = headers.map(h => String(h || '').toUpperCase().replace(/[^A-Z0-9]/g, ''));
    const idIdx = headersClean.findIndex(h => h === 'IDSISWA' || h === 'ID' || h === 'NOID');
    const isDelIdx = headersClean.indexOf('ISDELETED');

    let matchRow = -1;
    for (let b = 1; b < bioData.length; b++) {
      let cellId = idIdx > -1 ? String(bioData[b][idIdx] || '').trim().toLowerCase() : '';
      if (cellId === targetId) {
        matchRow = b;
        break;
      }
    }
    if (matchRow === -1) {
      for (let b = 1; b < bioData.length; b++) {
        for (let c = 0; c < bioData[b].length; c++) {
          if (String(bioData[b][c] || '').trim().toLowerCase() === targetId) {
            matchRow = b;
            break;
          }
        }
        if (matchRow > -1) break;
      }
    }

    if (matchRow === -1) {
      return sanitizeForRPC({ success: false, message: "Siswa dengan ID '" + idSiswa + "' tidak ditemukan." });
    }

    if (isDelIdx > -1) {
      const isDel = String(bioData[matchRow][isDelIdx] || '').trim().toUpperCase();
      if (isDel === 'TRUE') {
        return sanitizeForRPC({ success: false, message: "Akun siswa ini telah dinonaktifkan." });
      }
    }

    const row = bioData[matchRow];
    const findCol = (exacts, includesArr) => {
      let idx = headersClean.findIndex(h => exacts.includes(h));
      if (idx === -1 && includesArr) {
        idx = headersClean.findIndex(h => includesArr.some(inc => h.includes(inc)));
      }
      return idx;
    };

    const getVal = (exacts, includesArr) => {
      let idx = findCol(exacts, includesArr);
      if (idx > -1 && row[idx] !== undefined && row[idx] !== null && String(row[idx]).trim() !== '') {
        let val = row[idx];
        if (val instanceof Date) {
          const yyyy = val.getFullYear();
          const mm = String(val.getMonth() + 1).padStart(2, '0');
          const dd = String(val.getDate()).padStart(2, '0');
          return yyyy + '-' + mm + '-' + dd;
        }
        return String(val).trim();
      }
      return '-';
    };

    const biodataObj = {
      id: getVal(['IDSISWA', 'ID SISWA', 'ID', 'NOID'], []),
      nama: getVal(['NAMASISWA', 'NAMA SISWA', 'NAMA', 'NAMALENGKAP'], ['NAMA']),
      tempatLahir: getVal(['TEMPATLAHIR', 'TEMPAT'], ['TEMPAT']),
      tglLahir: getVal(['TANGGALLAHIR', 'TGLLAHIR'], ['TANGGAL', 'LAHIR']),
      usia: getVal(['USIA', 'UMUR'], ['USIA', 'UMUR']),
      asalSekolah: getVal(['ASALSEKOLAH', 'SEKOLAHASAL', 'SEKOLAH'], ['SEKOLAH']),
      kelas: getVal(['KELAS', 'PROGRAM'], ['KELAS']),
      kategori: getVal(['KATEGORI', 'LEVEL'], ['KATEGORI']),
      namaOrtu: getVal(['NAMAORANGTUA', 'NAMAORTU', 'ORANGTUA'], ['ORTU', 'ORANGTUA']),
      noHp: getVal(['NOMORHP', 'NOHP', 'WHATSAPP', 'WA', 'HP'], ['HP', 'TELEPON', 'WA']),
      email: getVal(['EMAIL', 'SUREL'], ['EMAIL']),
      alamat: getVal(['ALAMAT', 'ALAMATLENGKAP'], ['ALAMAT']),
      provinsi: getVal(['PROVINSI'], ['PROVINSI']),
      kabupaten: getVal(['KABUPATEN', 'KOTA'], ['KABUPATEN', 'KOTA']),
      kecamatan: getVal(['KECAMATAN'], ['KECAMATAN']),
      kelurahan: getVal(['KELURAHAN', 'DESA'], ['KELURAHAN', 'DESA']),
      foto: getVal(['FOTO', 'PASFOTO', 'FOTOPROFIL'], ['FOTO'])
    };

    return sanitizeForRPC({
      success: true,
      biodata: biodataObj
    });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal mengambil biodata siswa: " + err.toString() });
  }
}

function updateBiodataSiswa(payload, authToken) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(15000)) {
      return sanitizeForRPC({ success: false, message: "Sistem sedang sibuk memproses data lain. Harap coba lagi beberapa saat lagi." });
    }

    const token = authToken || (payload && (payload.__token || payload.token));
    const auth = requireAuth_(token, null);
    if (!auth.authorized) return auth.error;

    if (!payload || typeof payload !== 'object') {
      return sanitizeForRPC({ success: false, message: "Payload tidak valid." });
    }

    const idVal = String(payload.id || payload.idSiswa || '').trim();
    if (!idVal) {
      return sanitizeForRPC({ success: false, message: "ID Siswa wajib diisi." });
    }

    if (!requireOwnership_(auth.session, idVal)) {
      return sanitizeForRPC({ success: false, message: "Tidak diizinkan mengubah biodata siswa lain." });
    }

    const roleNorm = String(auth.session.role || '').toUpperCase();
    if (roleNorm === 'ADMIN' && !canEditModule_(auth.session, 'master')) {
      return sanitizeForRPC({ success: false, message: "Hak akses Anda untuk modul Master Siswa hanya baca (Readonly)." });
    }

    const namaVal = String(payload.nama || '').trim();
    const tempatLahirVal = String(payload.tempatLahir || '').trim();
    const tglLahirVal = String(payload.tglLahir || '').trim();
    const kelasVal = String(payload.kelas || '').trim();
    const namaOrtuVal = String(payload.namaOrtu || '').trim();
    const rawHp = String(payload.noHp || '').trim();
    const emailVal = String(payload.email || '').trim().toLowerCase();
    const alamatVal = String(payload.alamat || '').trim();
    const provVal = String(payload.provinsi || '').trim();
    const kabVal = String(payload.kabupaten || '').trim();
    const kecVal = String(payload.kecamatan || '').trim();
    const kelVal = String(payload.kelurahan || '').trim();
    const asalSekolahVal = String(payload.asalSekolah || '').trim();

    // 1. Validasi Kolom Wajib
    if (!namaVal || !tempatLahirVal || !tglLahirVal || !kelasVal || !namaOrtuVal || !rawHp || !emailVal || !alamatVal || !provVal || !kabVal || !kecVal || !kelVal) {
      return sanitizeForRPC({ success: false, message: "Semua kolom bertanda * wajib diisi lengkap." });
    }

    // 2. Batas Panjang Isian
    if (namaVal.length > 100 || tempatLahirVal.length > 80 || namaOrtuVal.length > 100 || asalSekolahVal.length > 120) {
      return sanitizeForRPC({ success: false, message: "Panjang isian nama/tempat lahir/sekolah melebihi batas wajar." });
    }
    if (alamatVal.length > 255 || provVal.length > 80 || kabVal.length > 80 || kecVal.length > 80 || kelVal.length > 80) {
      return sanitizeForRPC({ success: false, message: "Panjang isian alamat wilayah melebihi batas wajar." });
    }

    // 3. Format Email
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(emailVal) || emailVal.length > 120) {
      return sanitizeForRPC({ success: false, message: "Format alamat email tidak valid." });
    }

    // 4. Format Nomor HP / WhatsApp
    let cleanHp = rawHp.replace(/[^0-9]/g, '');
    if (cleanHp.startsWith('62')) cleanHp = cleanHp.substring(2);
    if (cleanHp.startsWith('0')) cleanHp = cleanHp.substring(1);
    if (cleanHp.length < 9 || cleanHp.length > 14) {
      return sanitizeForRPC({ success: false, message: "Nomor WhatsApp/HP harus berupa 9 sampai 14 digit angka." });
    }
    const formattedHp = "+62" + cleanHp;

    // 5. Format Tanggal Lahir (YYYY-MM-DD)
    if (!/^\d{4}-\d{2}-\d{2}$/.test(tglLahirVal)) {
      return sanitizeForRPC({ success: false, message: "Format tanggal lahir harus YYYY-MM-DD." });
    }

    // 6. Validasi & Kalkulasi Usia
    let usiaVal = payload.usia !== undefined && payload.usia !== null && payload.usia !== '' ? Number(payload.usia) : NaN;
    if (isNaN(usiaVal) || usiaVal < 1 || usiaVal > 100) {
      try {
        const birthD = new Date(tglLahirVal);
        const nowD = new Date();
        let calculatedAge = nowD.getFullYear() - birthD.getFullYear();
        const m = nowD.getMonth() - birthD.getMonth();
        if (m < 0 || (m === 0 && nowD.getDate() < birthD.getDate())) calculatedAge--;
        usiaVal = Math.max(1, calculatedAge);
      } catch (e) {
        usiaVal = 7;
      }
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const bioSheet = ss.getSheetByName('BIODATA');
    if (!bioSheet) return sanitizeForRPC({ success: false, message: "Sheet BIODATA tidak ditemukan!" });

    const bioData = bioSheet.getDataRange().getValues();
    if (bioData.length <= 1) return sanitizeForRPC({ success: false, message: "Data BIODATA kosong." });

    const bioHeaders = bioData[0].map(h => String(h || '').trim());
    const bioHeadersClean = bioHeaders.map(h => h.toUpperCase().replace(/[^A-Z0-9]/g, ''));
    const targetId = idVal.toLowerCase();

    const idColIdx = bioHeadersClean.findIndex(h => h === 'IDSISWA' || h === 'ID' || h === 'NOID');
    const isDelColIdx = bioHeadersClean.indexOf('ISDELETED');

    let matchRow = -1;
    for (let b = bioData.length - 1; b >= 1; b--) {
      let cellVal = idColIdx > -1 ? String(bioData[b][idColIdx] || '').trim().toLowerCase() : '';
      if (cellVal === targetId) {
        matchRow = b;
        break;
      }
    }
    if (matchRow === -1) {
      for (let b = bioData.length - 1; b >= 1; b--) {
        for (let c = 0; c < bioData[b].length; c++) {
          let cellVal = String(bioData[b][c] || '').trim().toLowerCase();
          if (cellVal !== '' && cellVal === targetId) {
            matchRow = b;
            break;
          }
        }
        if (matchRow > -1) break;
      }
    }

    if (matchRow === -1) {
      return sanitizeForRPC({ success: false, message: "Siswa dengan ID '" + idVal + "' tidak ditemukan di Sheet BIODATA." });
    }

    // Cek IS DELETED
    if (isDelColIdx > -1) {
      const isDelVal = String(bioData[matchRow][isDelColIdx] || '').trim().toUpperCase();
      if (isDelVal === 'TRUE') {
        return sanitizeForRPC({ success: false, message: "Akun siswa ini telah dinonaktifkan (IS DELETED = TRUE)." });
      }
    }

    // 7. Cek Email Unik di Sheet USER
    const userSheet = ss.getSheetByName('USER') || ss.getSheetByName('USERS');
    let userWarning = '';
    if (userSheet && userSheet.getLastRow() > 1) {
      const uData = userSheet.getDataRange().getValues();
      const uHeadersClean = uData[0].map(h => String(h || '').toLowerCase().replace(/[^a-z0-9]/g, ''));
      let idIdxU = uHeadersClean.findIndex(h => h === 'iduser' || h === 'id' || h.includes('id'));
      if (idIdxU === -1) idIdxU = 0;
      let emailIdxU = uHeadersClean.findIndex(h => h === 'email' || h === 'surel');
      if (emailIdxU === -1) emailIdxU = 1;

      // Cari baris USER milik siswa ini
      let myUserRow = -1;
      for (let u = 1; u < uData.length; u++) {
        let uId = String(uData[u][idIdxU] || '').trim().toLowerCase();
        if (uId === targetId) {
          myUserRow = u;
          break;
        }
      }

      // Pastikan email baru belum digunakan akun lain
      if (emailIdxU > -1) {
        for (let u = 1; u < uData.length; u++) {
          if (u !== myUserRow) {
            let otherEmail = String(uData[u][emailIdxU] || '').trim().toLowerCase();
            if (otherEmail === emailVal) {
              return sanitizeForRPC({ success: false, message: "Email '" + emailVal + "' sudah terdaftar pada akun lain. Gunakan email yang berbeda." });
            }
          }
        }
      }

      // Update baris USER
      if (myUserRow > -1) {
        let namaIdxU = uHeadersClean.findIndex(h => h.includes('nama'));
        if (namaIdxU > -1) userSheet.getRange(myUserRow + 1, namaIdxU + 1).setValue(namaVal);
        if (emailIdxU > -1) userSheet.getRange(myUserRow + 1, emailIdxU + 1).setValue(emailVal);
      } else {
        userWarning = ' (Catatan: Baris ID siswa di Sheet USER tidak ditemukan)';
      }
    }

    // 8. Hitung Kategori dari Sheet REFERENSI
    let computedKategori = String(payload.kategori || '').trim();
    try {
      const refSheet = ss.getSheetByName('REFERENSI');
      if (refSheet && refSheet.getLastRow() > 1) {
        const refValues = refSheet.getDataRange().getValues();
        const rHeaders = refValues[0].map(h => String(h || '').toUpperCase().trim());
        const kIdx = rHeaders.findIndex(h => h.includes('KELAS'));
        const katIdx = rHeaders.findIndex(h => h.includes('KATEGORI') || h.includes('LEVEL'));
        if (kIdx > -1 && katIdx > -1) {
          for (let r = 1; r < refValues.length; r++) {
            if (String(refValues[r][kIdx] || '').trim().toLowerCase() === kelasVal.toLowerCase()) {
              const matchedKat = String(refValues[r][katIdx] || '').trim();
              if (matchedKat) {
                computedKategori = matchedKat;
                break;
              }
            }
          }
        }
      }
    } catch (refErr) {}
    if (!computedKategori) computedKategori = '-';

    // 9. Unggah & Validasi Pas Foto
    let fileUrl = (payload.existingFoto && typeof payload.existingFoto === 'string') ? payload.existingFoto.trim() : '';
    if (fileUrl.includes('unsplash.com') || fileUrl.includes('placehold.co')) {
      fileUrl = '';
    }

    if (payload.fotoBase64) {
      let cleanBase64 = String(payload.fotoBase64);
      if (cleanBase64.indexOf(',') !== -1) {
        cleanBase64 = cleanBase64.split(',')[1];
      }
      cleanBase64 = cleanBase64.replace(/[^A-Za-z0-9+/=]/g, "");

      if (cleanBase64.length > 2202009) {
        return sanitizeForRPC({ success: false, message: "Ukuran pas foto terlalu besar. Maksimal 2.0 MB setelah dikompresi." });
      }

      let decodedBytes;
      try {
        decodedBytes = Utilities.base64Decode(cleanBase64);
      } catch (decErr) {
        return sanitizeForRPC({ success: false, message: "Format encoding foto tidak valid (gagal decode base64)." });
      }

      let mimeType = 'image/jpeg';
      let ext = '.jpg';
      if (decodedBytes && decodedBytes.length >= 4) {
        let b0 = decodedBytes[0] & 0xFF;
        let b1 = decodedBytes[1] & 0xFF;
        let b2 = decodedBytes[2] & 0xFF;
        let b3 = decodedBytes[3] & 0xFF;

        if (b0 === 0xFF && b1 === 0xD8) {
          mimeType = 'image/jpeg';
          ext = '.jpg';
        } else if (b0 === 0x89 && b1 === 0x50 && b2 === 0x4E && b3 === 0x47) {
          mimeType = 'image/png';
          ext = '.png';
        } else {
          return sanitizeForRPC({ success: false, message: "Format pas foto tidak didukung. Harap unggah berkas foto asli dalam format JPG atau PNG." });
        }
      } else {
        return sanitizeForRPC({ success: false, message: "Berkas pas foto rusak atau kosong." });
      }

      let safeName = namaVal.replace(/[^a-zA-Z0-9]/g, '_');
      let fileName = idVal + "_" + safeName + "_" + Date.now() + ext;
      let blob = Utilities.newBlob(decodedBytes, mimeType, fileName);

      let targetFolder = null;
      try {
        targetFolder = DriveApp.getFolderById(CONFIG.FOLDER_FOTO_ID);
      } catch (folderErr) {
        return sanitizeForRPC({ success: false, message: "Gagal mengakses folder FOTO di Google Drive: " + folderErr.toString() });
      }

      if (!targetFolder) {
        return sanitizeForRPC({ success: false, message: "Folder FOTO tidak ditemukan di Google Drive." });
      }

      let file = null;
      try {
        file = targetFolder.createFile(blob);
      } catch (createErr) {
        return sanitizeForRPC({ success: false, message: "Gagal mengunggah foto ke Google Drive: " + createErr.toString() });
      }

      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (shareErr) {}
      fileUrl = "https://lh3.googleusercontent.com/d/" + file.getId();
    }

    // 10. Tulis Perubahan ke Sheet BIODATA (ID SISWA TIDAK DITULIS ULANG)
    const updateCol = (exacts, includesArr, value) => {
      let idx = bioHeadersClean.findIndex(h => exacts.includes(h));
      if (idx === -1 && includesArr) {
        idx = bioHeadersClean.findIndex(h => includesArr.some(inc => h.includes(inc)));
      }
      if (idx > -1) {
        bioSheet.getRange(matchRow + 1, idx + 1).setValue(value);
      }
    };

    updateCol(['NAMASISWA', 'NAMA SISWA', 'NAMA', 'NAMALENGKAP'], ['NAMA'], namaVal);
    updateCol(['TEMPATLAHIR', 'TEMPAT'], ['TEMPAT'], tempatLahirVal);
    updateCol(['TANGGALLAHIR', 'TGLLAHIR'], ['TANGGAL', 'LAHIR'], tglLahirVal);
    updateCol(['USIA', 'UMUR'], ['USIA', 'UMUR'], usiaVal);
    updateCol(['ASALSEKOLAH', 'SEKOLAHASAL', 'SEKOLAH'], ['SEKOLAH'], asalSekolahVal);
    updateCol(['KELAS', 'PROGRAM'], ['KELAS'], kelasVal);
    updateCol(['KATEGORI', 'LEVEL'], ['KATEGORI'], computedKategori);
    updateCol(['NAMAORANGTUA', 'NAMAORTU', 'ORANGTUA'], ['ORTU', 'ORANGTUA'], namaOrtuVal);
    updateCol(['NOMORHP', 'NOHP', 'WHATSAPP', 'WA', 'HP'], ['HP', 'TELEPON', 'WA'], formattedHp);
    updateCol(['EMAIL', 'SUREL'], ['EMAIL'], emailVal);
    updateCol(['ALAMAT', 'ALAMATLENGKAP'], ['ALAMAT'], alamatVal);
    updateCol(['PROVINSI'], ['PROVINSI'], provVal);
    updateCol(['KABUPATEN', 'KOTA'], ['KABUPATEN', 'KOTA'], kabVal);
    updateCol(['KECAMATAN'], ['KECAMATAN'], kecVal);
    updateCol(['KELURAHAN', 'DESA'], ['KELURAHAN', 'DESA'], kelVal);
    if (fileUrl) {
      updateCol(['FOTO', 'PASFOTO'], ['FOTO'], fileUrl);
    }

    const updatedBio = {
      id: idVal,
      nama: namaVal,
      tempatLahir: tempatLahirVal,
      tglLahir: tglLahirVal,
      usia: String(usiaVal),
      asalSekolah: asalSekolahVal,
      kelas: kelasVal,
      kategori: computedKategori,
      namaOrtu: namaOrtuVal,
      noHp: formattedHp,
      email: emailVal,
      alamat: alamatVal,
      provinsi: provVal,
      kabupaten: kabVal,
      kecamatan: kecVal,
      kelurahan: kelVal,
      foto: fileUrl
    };

    const actorId = auth.session ? (auth.session.uid || auth.session.email || idVal) : idVal;
    logActivity(actorId, 'UPDATE_BIODATA', 'BIODATA', 'Memperbarui biodata siswa: ' + namaVal + ' (ID: ' + idVal + ')');

    return sanitizeForRPC({
      success: true,
      message: "Biodata siswa berhasil diperbarui!" + userWarning,
      biodata: updatedBio
    });

  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal memperbarui biodata: " + err.toString() });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function updateAkunSiswa(payload, authToken) {
  try {
    const token = authToken || (payload && (payload.__token || payload.token));
    const auth = requireAuth_(token, null);
    if (!auth.authorized) return auth.error;

    const targetId = String((payload && (payload.id || payload.idSiswa)) || '').trim().toLowerCase();
    if (targetId && !requireOwnership_(auth.session, targetId)) {
      return sanitizeForRPC({ success: false, message: "Tidak diizinkan mengubah akun pengguna lain." });
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const userSheet = ss.getSheetByName('USER') || ss.getSheetByName('USERS');
    if (!userSheet) return sanitizeForRPC({ success: false, message: "Sheet USER tidak ditemukan!" });

    const uData = userSheet.getDataRange().getValues();
    if (uData.length <= 1) return sanitizeForRPC({ success: false, message: "Data pengguna kosong." });

    const uHeadersClean = uData[0].map(h => String(h || '').toLowerCase().replace(/[^a-z0-9]/g, ''));
    
    let idIdx = uHeadersClean.findIndex(h => h.includes('iduser') || h === 'id');
    if (idIdx === -1) {
      let cand = uHeadersClean.findIndex(h => h.includes('id'));
      idIdx = cand > -1 ? cand : 0;
    }

    let emailIdx = uHeadersClean.findIndex(h => h === 'email' || h === 'surel');
    if (emailIdx === -1) emailIdx = 1;

    let passIdx = uHeadersClean.findIndex(h => h.includes('pass') || h.includes('pwd') || h.includes('sandi'));
    if (passIdx === -1) passIdx = 2;

    let updIdx = uHeadersClean.findIndex(h => h.includes('updated') || h.includes('terakhir') || h === 'lastupdate');
    let namaIdx = uHeadersClean.findIndex(h => h.includes('nama'));

    const currPass = String((payload && payload.currentPassword) || '').trim();
    const newEmail = String((payload && payload.newEmail) || '').trim().toLowerCase();
    const newPass = String((payload && payload.newPassword) || '').trim();

    if (!newEmail && !newPass) {
      return sanitizeForRPC({ success: false, message: "Tidak ada data perubahan yang dikirim." });
    }

    let userRow = -1;
    for (let u = 1; u < uData.length; u++) {
      let uId = String(uData[u][idIdx] || '').trim().toLowerCase();
      let uEmail = String(uData[u][emailIdx] || '').trim().toLowerCase();
      if (uId === targetId || uEmail === targetId || uEmail === String((payload && payload.currentEmail) || '').trim().toLowerCase()) {
        userRow = u;
        break;
      }
    }

    if (userRow === -1) {
      return sanitizeForRPC({ success: false, message: "Akun pengguna tidak ditemukan di Sheet USER." });
    }

    const storedPass = String(uData[userRow][passIdx] || '').trim();
    if (!verifyPassword_(currPass, storedPass)) {
      return sanitizeForRPC({ success: false, message: "Password Saat Ini tidak sesuai. Silakan periksa kembali." });
    }

    const existingEmail = String(uData[userRow][emailIdx] || '').trim().toLowerCase();

    if (newEmail) {
      const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
      if (!emailRegex.test(newEmail)) {
        return sanitizeForRPC({ success: false, message: "Format alamat Email baru tidak valid." });
      }
      if (newEmail === existingEmail && !newPass) {
        return sanitizeForRPC({ success: false, message: "Alamat Email baru sama dengan Email yang sedang digunakan saat ini." });
      }
      if (newEmail !== existingEmail) {
        for (let u = 1; u < uData.length; u++) {
          if (u !== userRow) {
            let otherEmail = String(uData[u][emailIdx] || '').trim().toLowerCase();
            if (otherEmail === newEmail) {
              return sanitizeForRPC({ success: false, message: "Email '" + newEmail + "' sudah terdaftar pada akun lain." });
            }
          }
        }
      }
    }

    if (newPass) {
      if (newPass.length < 6) {
        return sanitizeForRPC({ success: false, message: "Password baru minimal harus 6 karakter." });
      }
      if (newPass === currPass || verifyPassword_(newPass, storedPass)) {
        return sanitizeForRPC({ success: false, message: "Password baru tidak boleh sama dengan Password saat ini." });
      }
    }

    const nowIso = new Date().toISOString();
    let emailUpdated = false;
    let passUpdated = false;

    if (newEmail && newEmail !== existingEmail) {
      userSheet.getRange(userRow + 1, emailIdx + 1).setValue(newEmail);
      emailUpdated = true;
    }
    if (newPass) {
      const hashedNewPass = hashPasswordV2(newPass);
      userSheet.getRange(userRow + 1, passIdx + 1).setValue(hashedNewPass);
      passUpdated = true;
    }

    if ((emailUpdated || passUpdated) && updIdx > -1) {
      userSheet.getRange(userRow + 1, updIdx + 1).setValue(nowIso);
    }

    if (emailUpdated) {
      const bioSheet = ss.getSheetByName('BIODATA');
      if (bioSheet && bioSheet.getLastRow() > 1) {
        const bioData = bioSheet.getDataRange().getValues();
        const bioHeadersClean = bioData[0].map(h => String(h || '').toUpperCase().replace(/[^A-Z0-9]/g, ''));
        let bioIdIdx = bioHeadersClean.findIndex(h => h === 'IDSISWA' || h === 'ID SISWA' || h === 'ID' || h === 'NOID');
        let bioEmailIdx = bioHeadersClean.findIndex(h => h === 'EMAIL' || h === 'SUREL');
        
        if (bioEmailIdx > -1) {
          for (let b = bioData.length - 1; b >= 1; b--) {
            let cellId = bioIdIdx > -1 ? String(bioData[b][bioIdIdx] || '').trim().toLowerCase() : '';
            let cellEmail = String(bioData[b][bioEmailIdx] || '').trim().toLowerCase();
            
            if ((cellId !== '' && cellId === targetId) || (existingEmail !== '' && cellEmail === existingEmail)) {
              bioSheet.getRange(b + 1, bioEmailIdx + 1).setValue(newEmail);
              break;
            }
          }
        }
      }
    }

    // Sinkronkan pembaruan EMAIL dan PASSWORD ke Sheet DAFTAR jika baris pendaftar ada
    if (emailUpdated || passUpdated) {
      try {
        const daftarSheet = ss.getSheetByName('DAFTAR');
        if (daftarSheet && daftarSheet.getLastRow() > 1) {
          const dData = daftarSheet.getDataRange().getValues();
          const dHeaders = dData[0].map(h => String(h || '').toUpperCase().trim());
          let dEmailIdx = dHeaders.findIndex(h => h === 'EMAIL' || h === 'SUREL');
          let dPassIdx = dHeaders.findIndex(h => h === 'PASSWORD' || h === 'SANDI');
          let dNamaIdx = dHeaders.findIndex(h => h === 'NAMA SISWA' || h === 'NAMA');

          let userNameVal = namaIdx > -1 ? String(uData[userRow][namaIdx] || '').trim().toLowerCase() : '';

          for (let d = dData.length - 1; d >= 1; d--) {
            let rowEmail = dEmailIdx > -1 ? String(dData[d][dEmailIdx] || '').trim().toLowerCase() : '';
            let rowNama = dNamaIdx > -1 ? String(dData[d][dNamaIdx] || '').trim().toLowerCase() : '';

            let isMatchDaftar = (existingEmail && rowEmail === existingEmail) || (userNameVal && rowNama === userNameVal);
            if (isMatchDaftar) {
              if (emailUpdated && dEmailIdx > -1) {
                daftarSheet.getRange(d + 1, dEmailIdx + 1).setValue(newEmail);
              }
              if (passUpdated && dPassIdx > -1) {
                daftarSheet.getRange(d + 1, dPassIdx + 1).setValue(hashPasswordV2(newPass));
              }
              break;
            }
          }
        }
      } catch (dErr) {
        Logger.log("DAFTAR sync note: " + dErr.message);
      }
    }

    logActivity(payload ? payload.id : '-', 'UPDATE_AKUN', 'AUTH', 'Siswa memperbarui kredensial akun');

    return sanitizeForRPC({
      success: true,
      message: "Pengaturan akun berhasil diperbarui!",
      email: emailUpdated ? newEmail : existingEmail,
      updatedAt: nowIso
    });

  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal memperbarui akun: " + err.toString() });
  }
}

/**
 * Mengambil ringkasan informasi akun siswa untuk panel Info Akun
 */
function getInfoAkunSiswa(payload, authToken) {
  try {
    const token = authToken || (payload && (payload.__token || payload.token));
    const auth = requireAuth_(token, null);
    if (!auth.authorized) return auth.error;

    const targetId = String((payload && (payload.id || payload.idSiswa)) || '').trim().toLowerCase();
    if (targetId && !requireOwnership_(auth.session, targetId)) {
      return sanitizeForRPC({ success: false, message: "Tidak diizinkan melihat informasi akun pengguna lain." });
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const userSheet = ss.getSheetByName('USER') || ss.getSheetByName('USERS');
    if (!userSheet) return sanitizeForRPC({ success: false, message: "Sheet USER tidak ditemukan!" });

    const uData = userSheet.getDataRange().getValues();
    if (uData.length <= 1) return sanitizeForRPC({ success: false, message: "Data pengguna kosong." });

    const uHeadersClean = uData[0].map(h => String(h || '').toLowerCase().replace(/[^a-z0-9]/g, ''));
    let idIdx = uHeadersClean.findIndex(h => h.includes('iduser') || h === 'id');
    if (idIdx === -1) {
      let cand = uHeadersClean.findIndex(h => h.includes('id'));
      idIdx = cand > -1 ? cand : 0;
    }
    let emailIdx = uHeadersClean.findIndex(h => h === 'email' || h === 'surel');
    if (emailIdx === -1) emailIdx = 1;
    let namaIdx = uHeadersClean.findIndex(h => h.includes('nama'));
    let roleIdx = uHeadersClean.findIndex(h => h.includes('role') || h.includes('peran'));
    let statusIdx = uHeadersClean.findIndex(h => h.includes('status'));
    let createdIdx = uHeadersClean.findIndex(h => h.includes('created') || h.includes('buat'));
    let updatedIdx = uHeadersClean.findIndex(h => h.includes('updated') || h.includes('terakhir') || h === 'lastupdate');

    const targetEmail = String((payload && payload.email) || '').trim().toLowerCase();

    let userRow = -1;
    for (let u = 1; u < uData.length; u++) {
      let uId = String(uData[u][idIdx] || '').trim().toLowerCase();
      let uEmail = emailIdx > -1 ? String(uData[u][emailIdx] || '').trim().toLowerCase() : '';
      if ((targetId && (uId === targetId || uEmail === targetId)) || (targetEmail && uEmail === targetEmail)) {
        userRow = u;
        break;
      }
    }

    if (userRow === -1) {
      return sanitizeForRPC({ success: false, message: "Pengguna tidak ditemukan." });
    }

    const row = uData[userRow];
    const uIdFound = String(row[idIdx] || '-').trim();
    const uEmailFound = emailIdx > -1 ? String(row[emailIdx] || '-').trim() : '-';
    const uNamaFound = namaIdx > -1 ? String(row[namaIdx] || '-').trim() : '-';
    const uRoleFound = roleIdx > -1 ? String(row[roleIdx] || 'SISWA').trim() : 'SISWA';
    const uStatusFound = statusIdx > -1 ? String(row[statusIdx] || 'AKTIF').trim() : 'AKTIF';
    const uCreatedFound = createdIdx > -1 ? row[createdIdx] : '-';
    const uUpdatedFound = updatedIdx > -1 ? row[updatedIdx] : '-';

    // Cari riwayat login terakhir dari sheet LOG
    let lastLogin = '-';
    try {
      const logSheet = ss.getSheetByName('LOG');
      if (logSheet && logSheet.getLastRow() > 1) {
        const logData = logSheet.getDataRange().getValues();
        const logHeadersClean = logData[0].map(h => String(h || '').toLowerCase().replace(/[^a-z0-9]/g, ''));
        let lUserIdx = logHeadersClean.findIndex(h => h.includes('user'));
        let lActionIdx = logHeadersClean.findIndex(h => h.includes('action') || h.includes('aksi'));
        let lTimeIdx = logHeadersClean.findIndex(h => h.includes('time') || h.includes('waktu') || h.includes('tanggal'));

        if (lTimeIdx > -1) {
          for (let l = logData.length - 1; l >= 1; l--) {
            let logUser = lUserIdx > -1 ? String(logData[l][lUserIdx] || '').trim().toLowerCase() : '';
            let logAction = lActionIdx > -1 ? String(logData[l][lActionIdx] || '').trim().toUpperCase() : '';
            if (logAction === 'LOGIN' && (logUser === targetId || logUser === targetEmail || (uIdFound && logUser === uIdFound.toLowerCase()))) {
              lastLogin = logData[l][lTimeIdx];
              break;
            }
          }
        }
      }
    } catch (logErr) {}

    return sanitizeForRPC({
      success: true,
      info: {
        id: uIdFound,
        email: uEmailFound,
        nama: uNamaFound,
        role: uRoleFound,
        status: uStatusFound,
        createdAt: uCreatedFound,
        updatedAt: uUpdatedFound,
        lastLogin: lastLogin
      }
    });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal memuat info akun: " + err.toString() });
  }
}

function parseNominal_(val) {
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  if (!val) return 0;
  const clean = String(val).replace(/[^0-9]/g, '');
  const n = parseInt(clean, 10);
  return isNaN(n) ? 0 : n;
}

function matchPeriodDate_(val, targetYear, targetMonthIdx) {
  if (!val) return false;
  const anyMonth = (targetMonthIdx === -1 || targetMonthIdx === null || targetMonthIdx === undefined);
  if (val instanceof Date) {
    return val.getFullYear() === targetYear && (anyMonth || val.getMonth() === targetMonthIdx);
  }
  const s = String(val).trim();
  let ddmmyyyy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (ddmmyyyy) {
    const y = parseInt(ddmmyyyy[3], 10);
    const m = parseInt(ddmmyyyy[2], 10) - 1;
    return y === targetYear && (anyMonth || m === targetMonthIdx);
  }
  let yyyymmdd = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (yyyymmdd) {
    const y = parseInt(yyyymmdd[1], 10);
    const m = parseInt(yyyymmdd[2], 10) - 1;
    return y === targetYear && (anyMonth || m === targetMonthIdx);
  }
  let parsed = Date.parse(s);
  if (!isNaN(parsed)) {
    const d = new Date(parsed);
    return d.getFullYear() === targetYear && (anyMonth || d.getMonth() === targetMonthIdx);
  }
  return false;
}

function extractYearFromDate_(val) {
  if (!val) return null;
  if (val instanceof Date) return val.getFullYear();
  const s = String(val).trim();
  let m = s.match(/\b(20\d{2})\b/);
  if (m) {
    const y = parseInt(m[1], 10);
    if (y >= 2020 && y <= 2040) return y;
  }
  let parsed = Date.parse(s);
  if (!isNaN(parsed)) {
    const y = new Date(parsed).getFullYear();
    if (y >= 2020 && y <= 2040) return y;
  }
  return null;
}

function fetchDashboardStats(authToken, filterYear, filterMonth) {
  try {
    const auth = requireAuth_(authToken, ['ADMIN', 'SUPER_ADMIN']);
    if (!auth.authorized) return auth.error;

    const siswaData = fetchTableData('BIODATA', authToken).data || [];
    const kelasData = fetchTableData('REFERENSI', authToken).data || [];
    const keuData = fetchTableData('KEUANGAN', authToken).data || [];
    const galeriData = fetchTableData('GALERI', authToken).data || [];
    const pengData = fetchTableData('PENGELUARAN', authToken).data || [];

    const now = new Date();
    const curYear = now.getFullYear();
    const curMonth = now.getMonth(); // 0-indexed (0 = Januari, 11 = Desember)
    const indonesianMonths = ['januari', 'februari', 'maret', 'april', 'mei', 'juni', 'juli', 'agustus', 'september', 'oktober', 'november', 'desember'];

    const fYear = (filterYear && !isNaN(parseInt(filterYear, 10))) ? parseInt(filterYear, 10) : curYear;
    let fMonthIdx = curMonth;
    const rawFMonth = (filterMonth !== undefined && filterMonth !== null) ? String(filterMonth).trim().toLowerCase() : '';
    if (rawFMonth === '0' || rawFMonth === 'all' || rawFMonth === 'semua') {
      fMonthIdx = -1;
    } else if (rawFMonth !== '') {
      const pm = parseInt(rawFMonth, 10);
      if (!isNaN(pm) && pm >= 1 && pm <= 12) {
        fMonthIdx = pm - 1;
      }
    }

    // Filter Siswa Terdaftar & Kelas Aktif sesuai periode
    let periodKelasSet = new Set();
    let allTimeKelasSet = new Set();
    let periodSiswaCount = 0;
    siswaData.forEach(item => {
      const kAll = String(item['KELAS'] || item.kelas || '').trim();
      if (kAll) allTimeKelasSet.add(kAll);
      const tglDaftar = item['WAKTU DAFTAR'] || item.waktuDaftar || '';
      if (matchPeriodDate_(tglDaftar, fYear, fMonthIdx)) {
        periodSiswaCount++;
        if (kAll) periodKelasSet.add(kAll);
      }
    });

    let totalPendapatan = 0;
    let totalIncomeAll = 0;
    let tagihanPending = 0;
    let pembayaranDitolak = 0;

    keuData.forEach(item => {
      const status = String(item['STATUS PEMBAYARAN'] || item.status || '').toUpperCase().trim();
      const isDel = String(item['IS DELETED'] || item.isDeleted || '').toUpperCase().trim() === 'TRUE';
      if (isDel) return;

      const nom = parseNominal_(item.NOMINAL || item.nominal || 0);
      const tglBayar = String(item['TANGGAL BAYAR'] || item['TANGGAL_BAYAR'] || item.tanggalBayar || '').trim();
      const thn = parseInt(item['PERIODE TAHUN'] || item['PERIODE_TAHUN'] || item.periodeTahun || 0, 10);
      const blnStr = String(item['PERIODE BULAN'] || item['PERIODE_BULAN'] || item.periodeBulan || '').toLowerCase().trim();
      const bIdx = indonesianMonths.indexOf(blnStr);

      const isMonthMatch = (fMonthIdx === -1 || bIdx === fMonthIdx);
      const isPeriodMatch = matchPeriodDate_(tglBayar, fYear, fMonthIdx) || (thn === fYear && isMonthMatch);

      if (status === 'LUNAS') {
        totalIncomeAll += nom;
        if (isPeriodMatch) {
          totalPendapatan += nom;
        }
      } else if (status.includes('PENDING')) {
        if (isPeriodMatch) {
          tagihanPending++;
        }
      } else if (status === 'DITOLAK') {
        const waktuVf = item['WAKTU VERIFIKASI'] || item.waktuVerifikasi || '';
        const isRejectPeriod = matchPeriodDate_(waktuVf, fYear, fMonthIdx) || isPeriodMatch;
        if (isRejectPeriod) {
          pembayaranDitolak++;
        }
      }
    });

    let totalPengeluaranBulanIni = 0;
    let totalExpenseAll = 0;
    pengData.forEach(item => {
      const isDel = String(item['IS DELETED'] || item.isDeleted || '').toUpperCase().trim() === 'TRUE';
      if (isDel) return;

      const nom = parseNominal_(item['NOMINAL'] || item.nominal || 0);
      totalExpenseAll += nom;
      const tgl = String(item['TANGGAL'] || item.tanggal || '').trim();
      if (matchPeriodDate_(tgl, fYear, fMonthIdx)) {
        totalPengeluaranBulanIni += nom;
      }
    });

    const saldoBersih = totalPendapatan - totalPengeluaranBulanIni;

    let galeriPending = 0;
    galeriData.forEach(item => {
      const status = String(item['STATUS TAYANG'] || '').toUpperCase().trim();
      const tglUpload = item['WAKTU UPLOAD'] || item.waktuUpload || '';
      if (status === 'PENDING' && matchPeriodDate_(tglUpload, fYear, fMonthIdx)) {
        galeriPending++;
      }
    });

    let pendaftarPending = 0;
    try {
      const daftarData = fetchTableData('DAFTAR', authToken).data || [];
      daftarData.forEach(item => {
        const st = String(item['STATUS'] || item.status || '').toUpperCase().trim();
        const isDel = String(item['IS DELETED'] || item.isDeleted || '').toUpperCase().trim() === 'TRUE';
        const tglDaftar = item['WAKTU DAFTAR'] || item.waktuDaftar || '';
        if (!isDel && (st === 'PENDING' || st === 'MENUNGGU VERIFIKASI' || st === '') && matchPeriodDate_(tglDaftar, fYear, fMonthIdx)) {
          pendaftarPending++;
        }
      });
    } catch (eDaftar) {}

    return sanitizeForRPC({
      success: true,
      stats: { 
        totalSiswa: periodSiswaCount, 
        totalKelas: periodKelasSet.size, 
        totalPendapatan: totalPendapatan, 
        tagihanPending: tagihanPending,
        galeriPending: galeriPending,
        pendaftarPending: pendaftarPending,
        pembayaranDitolak: pembayaranDitolak,
        totalPengeluaranBulanIni: totalPengeluaranBulanIni,
        saldoBersih: saldoBersih
      },
      totals: {
        totalPendapatan: totalIncomeAll,
        totalPengeluaran: totalExpenseAll,
        saldoBersih: totalIncomeAll - totalExpenseAll,
        totalSiswa: siswaData.length,
        totalKelas: allTimeKelasSet.size
      },
      filter: {
        year: fYear,
        month: fMonthIdx === -1 ? 0 : (fMonthIdx + 1),
        isAllMonths: (fMonthIdx === -1)
      }
    });
  } catch (error) {
    return sanitizeForRPC({ success: false, stats: { totalSiswa: 0, totalKelas: 0, totalPendapatan: 0, tagihanPending: 0, galeriPending: 0, pendaftarPending: 0, pembayaranDitolak: 0, totalPengeluaranBulanIni: 0, saldoBersih: 0 }, totals: { totalPendapatan: 0, totalPengeluaran: 0, saldoBersih: 0, totalSiswa: 0, totalKelas: 0 } });
  }
}

function fetchDashboardOverview(authToken, filterYear, filterMonth) {
  try {
    const auth = requireAuth_(authToken, ['ADMIN', 'SUPER_ADMIN']);
    if (!auth.authorized) return auth.error;

    const now = new Date();
    const curYear = now.getFullYear();
    const curMonth = now.getMonth(); // 0 = Jan, 11 = Des

    const fYear = (filterYear && !isNaN(parseInt(filterYear, 10))) ? parseInt(filterYear, 10) : curYear;
    let fMonthIdx = curMonth;
    const rawFMonth = (filterMonth !== undefined && filterMonth !== null) ? String(filterMonth).trim().toLowerCase() : '';
    if (rawFMonth === '0' || rawFMonth === 'all' || rawFMonth === 'semua') {
      fMonthIdx = -1;
    } else if (rawFMonth !== '') {
      const pm = parseInt(rawFMonth, 10);
      if (!isNaN(pm) && pm >= 1 && pm <= 12) {
        fMonthIdx = pm - 1;
      }
    }

    const prevMonth = (fMonthIdx === 0) ? 11 : (fMonthIdx === -1 ? -1 : fMonthIdx - 1);
    const prevYear = (fMonthIdx === 0 || fMonthIdx === -1) ? fYear - 1 : fYear;
    const indonesianMonths = ['januari', 'februari', 'maret', 'april', 'mei', 'juni', 'juli', 'agustus', 'september', 'oktober', 'november', 'desember'];

    // Cek cache server (TTL 60 detik) berdasarkan periode terpilih
    const ver = getKeuanganCacheVer();
    const cacheMonthKey = (fMonthIdx === -1) ? 'ALL' : String(fMonthIdx + 1);
    const cacheKey = 'DASH_OVERVIEW_V' + ver + '_Y' + fYear + '_M' + cacheMonthKey;
    try {
      const cached = CacheService.getScriptCache().get(cacheKey);
      if (cached) {
        const parsed = JSON.parse(cached);
        if (parsed && parsed.success) {
          return sanitizeForRPC(parsed);
        }
      }
    } catch (eCache) {}

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

    // Helper pembacaan sheet tunggal dengan filter IS DELETED
    const getSheetRows_ = function(name) {
      let sh = ss.getSheetByName(name);
      if (!sh && (name === 'USER' || name === 'USERS')) sh = ss.getSheetByName('USER') || ss.getSheetByName('USERS');
      if (!sh && name === 'PENGELUARAN') sh = ensurePengeluaranSheet_();
      if (!sh || sh.getLastRow() <= 1 || sh.getLastColumn() <= 0) return [];
      const vals = sh.getDataRange().getValues();
      const headers = vals[0].map(h => String(h || '').trim());
      const headersClean = headers.map(h => h.toUpperCase().replace(/[^A-Z0-9]/g, ''));
      const isDelIdx = headersClean.indexOf('ISDELETED');
      const rows = [];
      for (let i = 1; i < vals.length; i++) {
        const r = vals[i];
        if (isDelIdx > -1 && String(r[isDelIdx] || '').toUpperCase() === 'TRUE') continue;
        const obj = {};
        headers.forEach((h, idx) => {
          let v = r[idx];
          if (v instanceof Date) {
            const yyyy = v.getFullYear();
            const mm = String(v.getMonth() + 1).padStart(2, '0');
            const dd = String(v.getDate()).padStart(2, '0');
            v = dd + '/' + mm + '/' + yyyy;
          }
          obj[h] = (v !== undefined && v !== null) ? String(v) : '';
        });
        rows.push(obj);
      }
      return rows;
    };

    const siswaData = getSheetRows_('BIODATA');
    const kelasData = getSheetRows_('REFERENSI');
    const keuData = getSheetRows_('KEUANGAN');
    const galeriData = getSheetRows_('GALERI');
    const pengData = getSheetRows_('PENGELUARAN');
    const daftarData = getSheetRows_('DAFTAR');

    // Kumpulkan seluruh tahun yang tersedia untuk opsi filter (murni data + tahun ini & terpilih)
    const yearsSet = new Set();
    yearsSet.add(curYear);
    yearsSet.add(fYear);
    keuData.forEach(item => {
      const thn = parseInt(item['PERIODE TAHUN'] || item.periodeTahun || 0, 10);
      if (thn >= 2020 && thn <= 2040) yearsSet.add(thn);
      const yTgl = extractYearFromDate_(item['TANGGAL BAYAR'] || item.tanggalBayar);
      if (yTgl) yearsSet.add(yTgl);
    });
    pengData.forEach(item => {
      const yTgl = extractYearFromDate_(item['TANGGAL'] || item.tanggal);
      if (yTgl) yearsSet.add(yTgl);
    });
    siswaData.forEach(item => {
      const yTgl = extractYearFromDate_(item['WAKTU DAFTAR'] || item.waktuDaftar);
      if (yTgl) yearsSet.add(yTgl);
    });
    daftarData.forEach(item => {
      const yTgl = extractYearFromDate_(item['WAKTU DAFTAR'] || item.waktuDaftar);
      if (yTgl) yearsSet.add(yTgl);
    });
    galeriData.forEach(item => {
      const yTgl = extractYearFromDate_(item['WAKTU UPLOAD'] || item.waktuUpload);
      if (yTgl) yearsSet.add(yTgl);
    });
    const availableYears = Array.from(yearsSet).sort((a, b) => b - a);

    // Filter Siswa Terdaftar & Kelas Aktif sesuai periode terpilih
    let periodKelasSet = new Set();
    let allTimeKelasSet = new Set();
    let periodSiswaCount = 0;
    siswaData.forEach(item => {
      const kAll = String(item['KELAS'] || item.kelas || '').trim();
      if (kAll) allTimeKelasSet.add(kAll);
      const tglDaftar = item['WAKTU DAFTAR'] || item.waktuDaftar || '';
      if (matchPeriodDate_(tglDaftar, fYear, fMonthIdx)) {
        periodSiswaCount++;
        if (kAll) periodKelasSet.add(kAll);
      }
    });

    // Series bulanan pendapatan (12 bulan tahun terpilih)
    const incomeSeries = new Array(12).fill(0);
    let totalPendapatan = 0;
    let totalIncomeAll = 0;
    const yearIncMap = {};
    let prevMonthPendapatan = 0;
    let tagihanPending = 0;
    let pembayaranDitolak = 0;

    keuData.forEach(item => {
      const status = String(item['STATUS PEMBAYARAN'] || item.status || '').toUpperCase().trim();
      const isDel = String(item['IS DELETED'] || item.isDeleted || '').toUpperCase().trim() === 'TRUE';
      if (isDel) return;

      const nom = parseNominal_(item.NOMINAL || item.nominal || 0);
      const tglBayar = String(item['TANGGAL BAYAR'] || item['TANGGAL_BAYAR'] || item.tanggalBayar || '').trim();
      const thn = parseInt(item['PERIODE TAHUN'] || item['PERIODE_TAHUN'] || item.periodeTahun || 0, 10);
      const blnStr = String(item['PERIODE BULAN'] || item['PERIODE_BULAN'] || item.periodeBulan || '').toLowerCase().trim();
      const bIdx = indonesianMonths.indexOf(blnStr);

      let mIdx = -1;
      let yVal = fYear;

      let ddmmyyyy = tglBayar.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      if (ddmmyyyy) {
        yVal = parseInt(ddmmyyyy[3], 10);
        mIdx = parseInt(ddmmyyyy[2], 10) - 1;
      } else {
        let yyyymmdd = tglBayar.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
        if (yyyymmdd) {
          yVal = parseInt(yyyymmdd[1], 10);
          mIdx = parseInt(yyyymmdd[2], 10) - 1;
        }
      }

      if (mIdx < 0 || mIdx > 11) {
        if (bIdx >= 0 && thn > 0) {
          yVal = thn;
          mIdx = bIdx;
        }
      }

      const isMonthMatch = (fMonthIdx === -1 || mIdx === fMonthIdx);
      const isPeriodMatch = (yVal === fYear && isMonthMatch);

      if (status === 'LUNAS') {
        totalIncomeAll += nom;
        if (yVal > 0) {
          yearIncMap[yVal] = (yearIncMap[yVal] || 0) + nom;
        }
        if (yVal === fYear && mIdx >= 0 && mIdx < 12) {
          incomeSeries[mIdx] += nom;
          if (isMonthMatch) totalPendapatan += nom;
        }
        if (fMonthIdx === -1) {
          if (yVal === prevYear && mIdx >= 0 && mIdx < 12) {
            prevMonthPendapatan += nom;
          }
        } else {
          if (yVal === prevYear && mIdx === prevMonth) {
            prevMonthPendapatan += nom;
          }
        }
      } else if (status.includes('PENDING')) {
        if (isPeriodMatch) {
          tagihanPending++;
        }
      } else if (status === 'DITOLAK') {
        const waktuVf = item['WAKTU VERIFIKASI'] || item.waktuVerifikasi || '';
        const isRejectPeriod = matchPeriodDate_(waktuVf, fYear, fMonthIdx) || isPeriodMatch;
        if (isRejectPeriod) {
          pembayaranDitolak++;
        }
      }
    });

    // Series bulanan pengeluaran (12 bulan tahun terpilih)
    const expenseSeries = new Array(12).fill(0);
    let totalPengeluaranBulanIni = 0;
    let totalExpenseAll = 0;
    const yearExpMap = {};
    let prevMonthPengeluaran = 0;

    pengData.forEach(item => {
      const isDel = String(item['IS DELETED'] || item.isDeleted || '').toUpperCase().trim() === 'TRUE';
      if (isDel) return;

      const nom = parseNominal_(item['NOMINAL'] || item.nominal || 0);
      totalExpenseAll += nom;
      const tgl = String(item['TANGGAL'] || item.tanggal || '').trim();
      let mIdx = -1;
      let yVal = fYear;

      let ddmmyyyy = tgl.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      if (ddmmyyyy) {
        yVal = parseInt(ddmmyyyy[3], 10);
        mIdx = parseInt(ddmmyyyy[2], 10) - 1;
      } else {
        let yyyymmdd = tgl.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
        if (yyyymmdd) {
          yVal = parseInt(yyyymmdd[1], 10);
          mIdx = parseInt(yyyymmdd[2], 10) - 1;
        }
      }

      if (yVal > 0) {
        yearExpMap[yVal] = (yearExpMap[yVal] || 0) + nom;
      }

      const isExpMonthMatch = (fMonthIdx === -1 || mIdx === fMonthIdx);
      if (yVal === fYear && mIdx >= 0 && mIdx < 12) {
        expenseSeries[mIdx] += nom;
        if (isExpMonthMatch) totalPengeluaranBulanIni += nom;
      }
      if (fMonthIdx === -1) {
        if (yVal === prevYear && mIdx >= 0 && mIdx < 12) {
          prevMonthPengeluaran += nom;
        }
      } else {
        if (yVal === prevYear && mIdx === prevMonth) {
          prevMonthPengeluaran += nom;
        }
      }
    });

    const saldoBersih = totalPendapatan - totalPengeluaranBulanIni;

    // Series perbandingan per tahun untuk grafik Total (urutan naik/kronologis)
    const allChartYears = Array.from(new Set([
      ...availableYears,
      ...Object.keys(yearIncMap).map(Number),
      ...Object.keys(yearExpMap).map(Number)
    ])).filter(y => y >= 2020 && y <= 2040).sort((a, b) => a - b);
    if (allChartYears.length === 0) allChartYears.push(curYear);
    const totalSeriesIncome = allChartYears.map(y => yearIncMap[y] || 0);
    const totalSeriesExpense = allChartYears.map(y => yearExpMap[y] || 0);

    // Pending galeri karya pada periode terpilih
    let galeriPending = 0;
    galeriData.forEach(item => {
      const status = String(item['STATUS TAYANG'] || '').toUpperCase().trim();
      const tglUpload = item['WAKTU UPLOAD'] || item.waktuUpload || '';
      if (status === 'PENDING' && matchPeriodDate_(tglUpload, fYear, fMonthIdx)) {
        galeriPending++;
      }
    });

    // Pendaftar pending pada periode terpilih
    let pendaftarPending = 0;
    daftarData.forEach(item => {
      const st = String(item['STATUS'] || item.status || '').toUpperCase().trim();
      const isDel = String(item['IS DELETED'] || item.isDeleted || '').toUpperCase().trim() === 'TRUE';
      const tglDaftar = item['WAKTU DAFTAR'] || item.waktuDaftar || '';
      if (!isDel && (st === 'PENDING' || st === 'MENUNGGU VERIFIKASI' || st === '') && matchPeriodDate_(tglDaftar, fYear, fMonthIdx)) {
        pendaftarPending++;
      }
    });

    // Hitung MoM
    let incomeMoM = 0;
    if (prevMonthPendapatan > 0) {
      incomeMoM = Math.round(((totalPendapatan - prevMonthPendapatan) / prevMonthPendapatan) * 100);
    }
    let expenseMoM = 0;
    if (prevMonthPengeluaran > 0) {
      expenseMoM = Math.round(((totalPengeluaranBulanIni - prevMonthPengeluaran) / prevMonthPengeluaran) * 100);
    }

    // Top 5 aktivitas terbaru dari sheet LOG (tetap global)
    const recentActivities = [];
    try {
      const logSheet = ss.getSheetByName('LOG');
      if (logSheet && logSheet.getLastRow() > 1) {
        const lastRow = logSheet.getLastRow();
        const startRow = Math.max(2, lastRow - 99);
        const numRows = lastRow - startRow + 1;
        const logVals = logSheet.getRange(startRow, 1, numRows, 6).getValues();
        for (let i = logVals.length - 1; i >= 0 && recentActivities.length < 5; i--) {
          const row = logVals[i];
          const ts = row[1];
          const action = String(row[3] || '');
          const mod = String(row[4] || '');
          const detail = String(row[5] || '');
          if (action || detail) {
            recentActivities.push({
              timestamp: (ts instanceof Date) ? ts.toISOString() : String(ts || ''),
              action: action,
              module: mod,
              detail: detail
            });
          }
        }
      }
    } catch (eLog) {}

    const payload = {
      success: true,
      stats: {
        totalSiswa: periodSiswaCount,
        totalKelas: periodKelasSet.size,
        totalPendapatan: totalPendapatan,
        totalPengeluaranBulanIni: totalPengeluaranBulanIni,
        saldoBersih: saldoBersih,
        tagihanPending: tagihanPending,
        galeriPending: galeriPending,
        pendaftarPending: pendaftarPending,
        pembayaranDitolak: pembayaranDitolak,
        incomeMoM: incomeMoM,
        expenseMoM: expenseMoM,
        prevMonthPendapatan: prevMonthPendapatan,
        prevMonthPengeluaran: prevMonthPengeluaran
      },
      totals: {
        totalPendapatan: totalIncomeAll,
        totalPengeluaran: totalExpenseAll,
        saldoBersih: totalIncomeAll - totalExpenseAll,
        totalSiswa: siswaData.length,
        totalKelas: allTimeKelasSet.size,
        series: {
          years: allChartYears,
          income: totalSeriesIncome,
          expense: totalSeriesExpense
        }
      },
      filter: {
        year: fYear,
        month: fMonthIdx === -1 ? 0 : (fMonthIdx + 1),
        monthName: fMonthIdx === -1 ? 'semua' : indonesianMonths[fMonthIdx],
        isAllMonths: (fMonthIdx === -1)
      },
      availableYears: availableYears,
      series: {
        currentYear: fYear,
        income: incomeSeries,
        expense: expenseSeries
      },
      activities: recentActivities,
      generatedAt: new Date().toISOString()
    };

    // Cache server TTL 60 detik
    try {
      CacheService.getScriptCache().put(cacheKey, JSON.stringify(payload), 60);
    } catch (ePut) {}

    return sanitizeForRPC(payload);
  } catch (error) {
    return sanitizeForRPC({
      success: false,
      message: "Gagal memuat overview dashboard: " + error.toString(),
      stats: { totalSiswa: 0, totalKelas: 0, totalPendapatan: 0, tagihanPending: 0, galeriPending: 0, pendaftarPending: 0, pembayaranDitolak: 0, totalPengeluaranBulanIni: 0, saldoBersih: 0 },
      totals: { totalPendapatan: 0, totalPengeluaran: 0, saldoBersih: 0, totalSiswa: 0, totalKelas: 0, series: { years: [curYear], income: [0], expense: [0] } },
      filter: { year: fYear, month: fMonthIdx + 1 }
    });
  }
}

function logActivity(userId, action, module, detail) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let logSheet = ss.getSheetByName('LOG');
    if (!logSheet) {
      logSheet = ss.insertSheet('LOG');
      logSheet.appendRow(['id_log', 'timestamp', 'user_id', 'action', 'module', 'detail']);
    }
    logSheet.appendRow(['LOG-' + new Date().getTime(), new Date().toISOString(), userId || '-', action || '-', module || '-', detail || '-']);
    invalidateSyncStampCache();
    const modUpper = String(module || '').toUpperCase().trim();
    if (modUpper === 'KEUANGAN') {
      bumpKeuanganCacheVer();
    } else if (modUpper === 'ABSENSI') {
      bumpAbsensiCacheVer();
    }
  } catch (e) {}
}

function saveUserAdmin(payload, authToken) {
  try {
    const token = authToken || (payload && (payload.__token || payload.token));
    const auth = requireAuth_(token, ['SUPER_ADMIN']);
    if (!auth.authorized) return auth.error;

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName('USER') || ss.getSheetByName('USERS');
    if (!sheet) {
      sheet = ss.insertSheet('USER');
      sheet.appendRow(['id_user', 'email', 'password', 'id_role', 'nama', 'status', 'created_at', 'updated_at', 'is_deleted', 'permissions']);
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).toLowerCase().replace(/[^a-z0-9]/g, ''));
    let idIdx = headers.findIndex(h => h.includes('id'));
    if (idIdx === -1) idIdx = 0;

    let emailIdx = headers.findIndex(h => h === 'email' || h === 'surel');
    if (emailIdx === -1) emailIdx = 1;

    let passIdx = headers.findIndex(h => h.includes('pass') || h.includes('pwd') || h.includes('sandi'));
    if (passIdx === -1) passIdx = 2;

    let roleIdx = headers.findIndex(h => h.includes('role') || h.includes('hak') || h.includes('akses'));
    if (roleIdx === -1) roleIdx = 3;

    let namaIdx = headers.findIndex(h => h.includes('nama') || h.includes('name'));
    if (namaIdx === -1) namaIdx = 4;

    let statusIdx = headers.findIndex(h => h.includes('status'));
    if (statusIdx === -1) statusIdx = 5;

    let delIdx = headers.findIndex(h => h.includes('deleted') || h.includes('hapus'));
    if (delIdx === -1) {
      delIdx = headers.length;
      sheet.getRange(1, delIdx + 1).setValue('is_deleted');
    }

    let permIdx = headers.findIndex(h => h.includes('permis') || h.includes('hakakses') || h.includes('otorisasi'));
    if (permIdx === -1) {
      permIdx = sheet.getLastColumn();
      sheet.getRange(1, permIdx + 1).setValue('permissions');
    }

    let targetId = String(payload.id_user || payload.idUser || payload.id || '').trim().toLowerCase();
    const targetEmail = String(payload.email || '').trim().toLowerCase();
    const targetNama = String(payload.nama || '').trim();
    const targetRole = String(payload.id_role || payload.role || 'SISWA').trim();
    const targetStatus = String(payload.status || 'AKTIF').trim();
    const targetPassword = payload.password !== undefined ? String(payload.password).trim() : '';

    if (!targetEmail) {
      return sanitizeForRPC({ success: false, message: "Email wajib diisi." });
    }

    let existingRow = -1;
    let oldEmail = '';
    let oldRole = '';
    let oldStatus = '';
    let oldUid = '';

    if (data.length > 1) {
      for (let i = 1; i < data.length; i++) {
        let uId = String(data[i][idIdx] || '').trim().toLowerCase();
        let uEmail = String(data[i][emailIdx] || '').trim().toLowerCase();
        if (targetId && (uId === targetId || uEmail === targetId)) {
          existingRow = i + 1;
          oldUid = String(data[i][idIdx] || '').trim();
          oldEmail = uEmail;
          oldRole = String(data[i][roleIdx] || '').trim().toUpperCase();
          oldStatus = String(data[i][statusIdx] || 'AKTIF').trim().toUpperCase();
          break;
        }
      }
    }

    // Validasi keunikan email di antara akun aktif lainnya
    if (data.length > 1) {
      for (let i = 1; i < data.length; i++) {
        if (existingRow > -1 && (i + 1) === existingRow) continue;
        let uEmail = String(data[i][emailIdx] || '').trim().toLowerCase();
        let isDel = delIdx > -1 && delIdx < data[i].length ? String(data[i][delIdx] || '').trim().toUpperCase() : 'FALSE';
        if (uEmail === targetEmail && isDel !== 'TRUE') {
          return sanitizeForRPC({ success: false, message: "Email " + targetEmail + " sudah terdaftar untuk pengguna lain." });
        }
      }
    }

    const sessionCallerUid = String(auth.session.uid || '').trim().toLowerCase();
    const sessionCallerEmail = String(auth.session.email || '').trim().toLowerCase();
    const isTargetSelf = (targetId && (targetId === sessionCallerUid || targetId === sessionCallerEmail)) || (oldEmail && oldEmail === sessionCallerEmail);

    if (existingRow > -1) {
      // Cegah Super Admin menonaktifkan akunnya sendiri
      if (isTargetSelf && targetStatus !== 'AKTIF') {
        return sanitizeForRPC({ success: false, message: "Anda tidak dapat menonaktifkan akun Anda sendiri." });
      }

      // Cegah Super Admin mencabut role Super Admin dari akunnya sendiri
      if (isTargetSelf && oldRole.includes('SUPER') && !targetRole.toUpperCase().includes('SUPER')) {
        return sanitizeForRPC({ success: false, message: "Anda tidak dapat mencabut hak akses Super Admin dari akun Anda sendiri." });
      }

      // Jika mendegradasi atau menonaktifkan akun SUPER_ADMIN, periksa apakah ini SUPER_ADMIN terakhir
      if (oldRole.includes('SUPER') && (!targetRole.toUpperCase().includes('SUPER') || targetStatus !== 'AKTIF')) {
        let activeSuperAdminCount = 0;
        for (let i = 1; i < data.length; i++) {
          let r = String(data[i][roleIdx] || '').trim().toUpperCase();
          let s = String(data[i][statusIdx] || 'AKTIF').trim().toUpperCase();
          let isDel = delIdx > -1 && delIdx < data[i].length ? String(data[i][delIdx] || '').trim().toUpperCase() : 'FALSE';
          if (r.includes('SUPER') && s === 'AKTIF' && isDel !== 'TRUE') {
            activeSuperAdminCount++;
          }
        }
        if (activeSuperAdminCount <= 1) {
          return sanitizeForRPC({ success: false, message: "Tidak dapat mengubah role atau menonaktifkan Super Admin terakhir." });
        }
      }

      // Validasi password jika diisi
      if (targetPassword !== '') {
        if (targetPassword.length < 8) {
          return sanitizeForRPC({ success: false, message: "Password minimal 8 karakter." });
        }
      }

      sheet.getRange(existingRow, emailIdx + 1).setValue(targetEmail);
      if (targetPassword !== '') {
        sheet.getRange(existingRow, passIdx + 1).setValue(hashPasswordV2(targetPassword));
      }
      sheet.getRange(existingRow, roleIdx + 1).setValue(targetRole);
      sheet.getRange(existingRow, namaIdx + 1).setValue(targetNama);
      sheet.getRange(existingRow, statusIdx + 1).setValue(targetStatus);
      
      let updIdx = headers.findIndex(h => h.includes('updated'));
      if (updIdx > -1) {
        sheet.getRange(existingRow, updIdx + 1).setValue(new Date().toISOString());
      }

      if (payload.permissions !== undefined) {
        sheet.getRange(existingRow, permIdx + 1).setValue(String(payload.permissions || ''));
      }

      logActivity(oldUid || targetId, 'UPDATE_USER', 'USERS', 'Super Admin memperbarui user: ' + targetEmail);

      // Blok E: Jika password, role, atau status berubah -> cabut semua sesi aktif target seketika
      if (targetPassword !== '' || targetRole !== oldRole || targetStatus !== oldStatus) {
        revokeUserSessions_(oldEmail);
        revokeUserSessions_(targetEmail);
        revokeUserSessions_(oldUid);
        revokeUserSessions_(targetId);
      }
    } else {
      // Akun Baru: Password wajib dan minimal 8 karakter
      if (!targetPassword) {
        return sanitizeForRPC({ success: false, message: "Password wajib diisi untuk pengguna baru." });
      }
      if (targetPassword.length < 8) {
        return sanitizeForRPC({ success: false, message: "Password pengguna baru minimal 8 karakter." });
      }

      const newId = payload.id_user || payload.idUser || ('USR-' + Date.now());
      const hashedPass = hashPasswordV2(targetPassword);
      const nowStr = new Date().toISOString();
      const permsStr = payload.permissions !== undefined ? String(payload.permissions || '') : '';

      const maxCols = Math.max(sheet.getLastColumn(), permIdx + 1, delIdx + 1);
      const newRow = new Array(maxCols).fill('');
      newRow[idIdx] = newId;
      newRow[emailIdx] = targetEmail;
      newRow[passIdx] = hashedPass;
      newRow[roleIdx] = targetRole;
      newRow[namaIdx] = targetNama;
      newRow[statusIdx] = targetStatus;
      
      let crIdx = headers.findIndex(h => h.includes('create'));
      if (crIdx > -1) newRow[crIdx] = nowStr;
      let updIdx = headers.findIndex(h => h.includes('update'));
      if (updIdx > -1) newRow[updIdx] = nowStr;
      if (delIdx > -1) newRow[delIdx] = 'FALSE';
      if (permIdx > -1) newRow[permIdx] = permsStr;

      sheet.appendRow(newRow);
      logActivity(newId, 'ADD_USER', 'USERS', 'Super Admin membuat user baru: ' + targetEmail);
    }

    return sanitizeForRPC({ success: true, message: "Data user berhasil disimpan ke Sheet USERS!" });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal menyimpan user: " + err.toString() });
  }
}

function deleteUserAdmin(idUser, authToken) {
  return softDeleteUserAdmin(idUser, authToken);
}

function softDeleteUserAdmin(idUser, authToken) {
  try {
    const auth = requireAuth_(authToken, ['SUPER_ADMIN']);
    if (!auth.authorized) return auth.error;

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('USER') || ss.getSheetByName('USERS');
    if (!sheet) return sanitizeForRPC({ success: false, message: "Sheet USER tidak ditemukan!" });

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return sanitizeForRPC({ success: false, message: "Data pengguna kosong." });

    const targetId = String(idUser || '').trim().toLowerCase();

    const headers = data[0].map(h => String(h || '').toLowerCase().replace(/[^a-z0-9]/g, ''));
    let idIdx = headers.findIndex(h => h.includes('id'));
    if (idIdx === -1) idIdx = 0;

    let emailIdx = headers.findIndex(h => h === 'email' || h === 'surel');
    if (emailIdx === -1) emailIdx = 1;

    let roleIdx = headers.findIndex(h => h.includes('role') || h.includes('hak') || h.includes('akses'));
    if (roleIdx === -1) roleIdx = 3;

    let statusIdx = headers.findIndex(h => h.includes('status'));
    if (statusIdx === -1) statusIdx = 5;

    let delIdx = headers.findIndex(h => h.includes('deleted') || h.includes('hapus'));
    if (delIdx === -1) {
      delIdx = headers.length;
      sheet.getRange(1, delIdx + 1).setValue('is_deleted');
    }

    const sessionCallerUid = String(auth.session.uid || '').trim().toLowerCase();
    const sessionCallerEmail = String(auth.session.email || '').trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      let uId = String(data[i][idIdx] || '').trim().toLowerCase();
      let uEmail = String(data[i][emailIdx] || '').trim().toLowerCase();
      let uRole = roleIdx > -1 ? String(data[i][roleIdx] || '').trim().toUpperCase() : '';

      if ((uId && uId === targetId) || (uEmail && uEmail === targetId)) {
        // Anti-lockout: Cegah Super Admin menghapus akunnya sendiri
        if (uId === sessionCallerUid || uEmail === sessionCallerEmail) {
          return sanitizeForRPC({ success: false, message: "Anda tidak dapat menghapus akun Anda sendiri." });
        }

        // Anti-lockout: Cegah menghapus Super Admin terakhir
        if (uRole.includes('SUPER')) {
          let activeSuperAdminCount = 0;
          for (let k = 1; k < data.length; k++) {
            let r = roleIdx > -1 ? String(data[k][roleIdx] || '').trim().toUpperCase() : '';
            let s = statusIdx > -1 ? String(data[k][statusIdx] || 'AKTIF').trim().toUpperCase() : 'AKTIF';
            let isDel = delIdx > -1 && delIdx < data[k].length ? String(data[k][delIdx] || '').trim().toUpperCase() : 'FALSE';
            if (r.includes('SUPER') && s === 'AKTIF' && isDel !== 'TRUE') {
              activeSuperAdminCount++;
            }
          }
          if (activeSuperAdminCount <= 1) {
            return sanitizeForRPC({ success: false, message: "Tidak dapat menghapus Super Admin terakhir." });
          }
        }

        if (statusIdx > -1 && statusIdx < data[i].length) {
          sheet.getRange(i + 1, statusIdx + 1).setValue('NONAKTIF');
        }
        sheet.getRange(i + 1, delIdx + 1).setValue('TRUE');
        logActivity('ADMIN', 'DELETE_USER', 'USERS', 'Super Admin menghapus (soft delete) user: ' + idUser);

        // Blok E: Cabut sesi aktif user yang dihapus
        revokeUserSessions_(uEmail);
        revokeUserSessions_(uId);

        return sanitizeForRPC({ success: true, message: "User " + idUser + " berhasil dihapus (Soft Delete)." });
      }
    }

    return sanitizeForRPC({ success: false, message: "ID User " + idUser + " tidak ditemukan." });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal menghapus user: " + err.toString() });
  }
}

function reactivateUserAdmin(idUser, authToken) {
  try {
    const auth = requireAuth_(authToken, ['SUPER_ADMIN']);
    if (!auth.authorized) return auth.error;

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('USER') || ss.getSheetByName('USERS');
    if (!sheet) return sanitizeForRPC({ success: false, message: "Sheet USER tidak ditemukan!" });

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return sanitizeForRPC({ success: false, message: "Data pengguna kosong." });

    const targetId = String(idUser || '').trim().toLowerCase();

    const headers = data[0].map(h => String(h || '').toLowerCase().replace(/[^a-z0-9]/g, ''));
    let idIdx = headers.findIndex(h => h.includes('id'));
    if (idIdx === -1) idIdx = 0;

    let emailIdx = headers.findIndex(h => h === 'email' || h === 'surel');
    if (emailIdx === -1) emailIdx = 1;

    let statusIdx = headers.findIndex(h => h.includes('status'));
    let delIdx = headers.findIndex(h => h.includes('deleted') || h.includes('hapus'));

    for (let i = 1; i < data.length; i++) {
      let uId = String(data[i][idIdx] || '').trim().toLowerCase();
      let uEmail = String(data[i][emailIdx] || '').trim().toLowerCase();

      if ((uId && uId === targetId) || (uEmail && uEmail === targetId)) {
        if (statusIdx > -1 && statusIdx < data[i].length) {
          sheet.getRange(i + 1, statusIdx + 1).setValue('AKTIF');
        }
        if (delIdx > -1 && delIdx < data[i].length) {
          sheet.getRange(i + 1, delIdx + 1).setValue('FALSE');
        }
        
        logActivity('ADMIN', 'REACTIVATE_USER', 'USERS', 'Super Admin mengaktifkan kembali user: ' + idUser);
        return sanitizeForRPC({ success: true, message: "User " + idUser + " berhasil diaktifkan kembali." });
      }
    }

    return sanitizeForRPC({ success: false, message: "ID User " + idUser + " tidak ditemukan." });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal mengaktifkan kembali user: " + err.toString() });
  }
}

/**
 * Mereset Nomor Urut Pendaftaran dimulai dari 1 (ScriptProperties & Sheet DAFTAR),
 * dan mereset ID Siswa (4 digit terakhir) menjadi format 0001 (0001, 0002, dst.)
 * dengan tetap mempertahankan 10 digit pola tanggal di depannya.
 * Memperbarui referensi ID di sheet USER, KEUANGAN, ABSENSI, dan GALERI secara aman.
 */
function resetNomorUrutDanIdSiswa(authToken) {
  try {
    const auth = requireAuth_(authToken, ['SUPER_ADMIN', 'ADMIN']);
    if (!auth.authorized) return auth.error;

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const props = PropertiesService.getScriptProperties();

    let resultLog = [];

    // 1. Reset / Susun Ulang Nomor Urut di Sheet DAFTAR
    let daftarSheet = ss.getSheetByName('DAFTAR');
    let totalDaftar = 0;
    if (daftarSheet && daftarSheet.getLastRow() > 1) {
      let dData = daftarSheet.getDataRange().getValues();
      let dHeaders = dData[0].map(h => String(h || '').toUpperCase().trim());
      let noIdx = dHeaders.indexOf('NO');
      if (noIdx === -1) noIdx = 0;
      let dEmailIdx = dHeaders.findIndex(h => h === 'EMAIL' || h === 'SUREL');

      // Hapus data uji coba jika ada
      if (dEmailIdx > -1) {
        for (let d = dData.length - 1; d >= 1; d--) {
          if (String(dData[d][dEmailIdx] || '').trim().toLowerCase() === 'test_siswa_pertama@example.com') {
            daftarSheet.deleteRow(d + 1);
          }
        }
      }

      if (daftarSheet.getLastRow() > 1) {
        dData = daftarSheet.getDataRange().getValues();
        for (let i = 1; i < dData.length; i++) {
          let newNo = i; // Mulai dari 1
          daftarSheet.getRange(i + 1, noIdx + 1).setValue(newNo);
          totalDaftar++;
        }
        props.setProperty('GLOBAL_SEQ_DAFTAR', totalDaftar.toString());
        resultLog.push("Sheet DAFTAR: " + totalDaftar + " pendaftar diurutkan ulang mulai dari NO 1.");
      } else {
        props.setProperty('GLOBAL_SEQ_DAFTAR', '0');
        resultLog.push("Sheet DAFTAR: Kosong. Counter GLOBAL_SEQ_DAFTAR di-reset ke 0 (pendaftar berikutnya mulai dari 1).");
      }
    } else {
      props.setProperty('GLOBAL_SEQ_DAFTAR', '0');
      resultLog.push("Sheet DAFTAR: Kosong. Counter GLOBAL_SEQ_DAFTAR di-reset ke 0 (pendaftar berikutnya mulai dari 1).");
    }

    // 2. Susun Ulang Nomor Urut & 4 Digit Terakhir ID Siswa di Sheet BIODATA (0001, 0002, dst. secara kronologis)
    let bioSheet = ss.getSheetByName('BIODATA');
    let idMapping = {}; // oldId -> newId
    let emailMapping = {}; // email -> newId
    let totalSiswa = 0;

    if (bioSheet && bioSheet.getLastRow() > 1) {
      let bioData = bioSheet.getDataRange().getValues();
      let bioHeaders = bioData[0].map(h => String(h || '').toUpperCase().trim());
      let bioEmailIdx = bioHeaders.findIndex(h => h === 'EMAIL' || h === 'SUREL');

      // Hapus data uji coba jika ada
      if (bioEmailIdx > -1) {
        for (let b = bioData.length - 1; b >= 1; b--) {
          if (String(bioData[b][bioEmailIdx] || '').trim().toLowerCase() === 'test_siswa_pertama@example.com') {
            bioSheet.deleteRow(b + 1);
          }
        }
      }

      if (bioSheet.getLastRow() > 1) {
        bioData = bioSheet.getDataRange().getValues();
        bioHeaders = bioData[0].map(h => String(h || '').toUpperCase().trim());

        let noIdx = bioHeaders.indexOf('NO');
        if (noIdx === -1) noIdx = 0;

        let idIdx = bioHeaders.findIndex(h => h === 'ID SISWA' || h === 'IDSISWA' || h === 'ID');
        if (idIdx === -1) idIdx = 1;

        let isDelIdx = bioHeaders.indexOf('IS DELETED');
        if (isDelIdx === -1) isDelIdx = bioHeaders.indexOf('IS_DELETED');

        let tglLahirIdx = bioHeaders.findIndex(h => h.includes('LAHIR') && h.includes('TGL'));
        let waktuDaftarIdx = bioHeaders.findIndex(h => h.includes('WAKTU') || h.includes('TIMESTAMP'));

        // Kumpulkan seluruh siswa aktif
        let activeRows = [];
        for (let b = 1; b < bioData.length; b++) {
          let isDel = isDelIdx > -1 ? String(bioData[b][isDelIdx] || '').trim().toUpperCase() : 'FALSE';
          if (isDel === 'TRUE') continue; // Jangan renumber siswa terhapus

          let rowEmail = bioEmailIdx > -1 ? String(bioData[b][bioEmailIdx] || '').trim().toLowerCase() : '';
          let rowWaktu = waktuDaftarIdx > -1 ? bioData[b][waktuDaftarIdx] : '';
          let rowId = idIdx > -1 ? String(bioData[b][idIdx] || '').trim() : '';
          let rowNo = noIdx > -1 ? parseInt(bioData[b][noIdx], 10) : b;

          // Parse waktu daftar secara presisi
          let tDaftar = 0;
          if (rowWaktu instanceof Date && !isNaN(rowWaktu.getTime())) {
            tDaftar = rowWaktu.getTime();
          } else if (rowWaktu) {
            let parsed = Date.parse(String(rowWaktu));
            if (!isNaN(parsed)) tDaftar = parsed;
          }
          // Jika waktu daftar kosong, coba ekstrak tahun-bulan registrasi dari digit 6-10 di oldId (YYMM)
          if (!tDaftar && rowId.length >= 10) {
            let regYY = parseInt(rowId.substring(6, 8), 10);
            let regMM = parseInt(rowId.substring(8, 10), 10);
            if (!isNaN(regYY) && !isNaN(regMM)) {
              tDaftar = new Date(2000 + regYY, regMM - 1, 1).getTime();
            }
          }
          if (!tDaftar && !isNaN(rowNo) && rowNo > 0) {
            tDaftar = rowNo;
          }

          activeRows.push({
            rowIndex: b + 1,
            dataIndex: b,
            email: rowEmail,
            oldId: rowId,
            tDaftar: tDaftar,
            rowNo: rowNo,
            tglLahirVal: tglLahirIdx > -1 ? bioData[b][tglLahirIdx] : '',
            waktuDaftarVal: rowWaktu
          });
        }

        // Urutkan siswa secara kronologis: yang mendaftar lebih awal mendapat nomor urut lebih awal (0001, 0002, dst.)
        activeRows.sort(function(a, b) {
          if (a.tDaftar && b.tDaftar && a.tDaftar !== b.tDaftar) {
            return a.tDaftar - b.tDaftar;
          }
          if (a.rowNo && b.rowNo && a.rowNo !== b.rowNo) {
            return a.rowNo - b.rowNo;
          }
          return a.rowIndex - b.rowIndex;
        });

        totalSiswa = activeRows.length;

        for (let i = 0; i < activeRows.length; i++) {
          let item = activeRows[i];
          let activeCount = i + 1; // 1, 2, 3, dst.
          let newNo = activeCount;
          let oldId = item.oldId;

          let prefix = "";
          if (oldId && oldId.length >= 10) {
            // Ambil 10 digit di depannya: [YY_Lahir][MM_Lahir][DD_Lahir][YY_Daftar][MM_Daftar]
            prefix = oldId.substring(0, 10);
          } else {
            // Fallback: hitung dari tanggal lahir & waktu daftar
            let bDate = new Date(item.tglLahirVal);
            let rDate = new Date(item.waktuDaftarVal);
            let bYY = "00", bMM = "01", bDD = "01", rYY = "26", rMM = "10";
            if (!isNaN(bDate.getTime())) {
              bYY = String(bDate.getFullYear()).slice(-2);
              bMM = String(bDate.getMonth() + 1).padStart(2, '0');
              bDD = String(bDate.getDate()).padStart(2, '0');
            }
            if (!isNaN(rDate.getTime())) {
              rYY = String(rDate.getFullYear()).slice(-2);
              rMM = String(rDate.getMonth() + 1).padStart(2, '0');
            }
            prefix = bYY + bMM + bDD + rYY + rMM;
          }

          // Reset 4 digit terakhir menjadi urut 0001, 0002, dst.
          let newSuffix = String(activeCount).padStart(4, '0');
          let newId = prefix + newSuffix;

          if (oldId && oldId !== newId) {
            idMapping[oldId] = newId;
          }
          if (item.email) {
            emailMapping[item.email] = newId;
          }

          bioSheet.getRange(item.rowIndex, noIdx + 1).setValue(newNo);
          bioSheet.getRange(item.rowIndex, idIdx + 1).setValue(newId);
        }

        props.setProperty('GLOBAL_SEQ_SISWA', totalSiswa.toString());
        props.setProperty('GLOBAL_SEQ_DAFTAR', totalSiswa.toString());
        resultLog.push("Sheet BIODATA: " + totalSiswa + " siswa aktif berhasil diurutkan dengan NO 1 s/d " + totalSiswa + " dan 4 digit terakhir ID Siswa (0001, 0002, dst).");
      } else {
        resultLog.push("Sheet BIODATA: Kosong.");
      }
    } else {
      resultLog.push("Sheet BIODATA: Kosong.");
    }

    // 3. Sinkronkan ID Siswa Baru ke Sheet USER
    let userSheet = ss.getSheetByName('USER') || ss.getSheetByName('USERS');
    let userUpdated = 0;
    if (userSheet && userSheet.getLastRow() > 1 && (Object.keys(idMapping).length > 0 || Object.keys(emailMapping).length > 0)) {
      let uData = userSheet.getDataRange().getValues();
      let uHeaders = uData[0].map(h => String(h || '').toLowerCase().replace(/[^a-z0-9]/g, ''));
      let uEmailIdx = uHeaders.findIndex(h => h === 'email' || h === 'surel');
      let uIdIdx = uHeaders.findIndex(h => h === 'id_user' || h === 'iduser');
      if (uIdIdx === -1) uIdIdx = uHeaders.findIndex(h => h.includes('id') && !h.includes('role'));
      if (uIdIdx === -1) uIdIdx = 0;

      for (let u = 1; u < uData.length; u++) {
        let currentUId = String(uData[u][uIdIdx] || '').trim();
        let currentUEmail = uEmailIdx > -1 ? String(uData[u][uEmailIdx] || '').trim().toLowerCase() : '';
        let targetNewId = (currentUEmail && emailMapping[currentUEmail]) ? emailMapping[currentUEmail] : (idMapping[currentUId] || null);

        if (targetNewId && targetNewId !== currentUId) {
          userSheet.getRange(u + 1, uIdIdx + 1).setValue(targetNewId);
          userUpdated++;
        }
      }
      resultLog.push("Sheet USER: " + userUpdated + " ID akun siswa diperbarui.");
    }

    // 4. Sinkronkan ke Sheet ABSENSI jika ada
    let absSheet = ss.getSheetByName('ABSENSI');
    let absUpdated = 0;
    if (absSheet && absSheet.getLastRow() > 1 && Object.keys(idMapping).length > 0) {
      const aData = absSheet.getDataRange().getValues();
      const aHeaders = aData[0].map(h => String(h || '').toUpperCase().trim());
      let aIdIdx = aHeaders.indexOf('ID SISWA');
      if (aIdIdx > -1) {
        for (let a = 1; a < aData.length; a++) {
          let currId = String(aData[a][aIdIdx] || '').trim();
          if (idMapping[currId]) {
            absSheet.getRange(a + 1, aIdIdx + 1).setValue(idMapping[currId]);
            absUpdated++;
          }
        }
        resultLog.push("Sheet ABSENSI: " + absUpdated + " record disinkronkan.");
      }
    }

    // 5. Sinkronkan ke Sheet KEUANGAN jika ada
    let keuSheet = ss.getSheetByName('KEUANGAN');
    let keuUpdated = 0;
    if (keuSheet && keuSheet.getLastRow() > 1 && Object.keys(idMapping).length > 0) {
      const kData = keuSheet.getDataRange().getValues();
      const kHeaders = kData[0].map(h => String(h || '').toUpperCase().trim());
      let kIdIdx = kHeaders.indexOf('ID SISWA');
      if (kIdIdx > -1) {
        for (let k = 1; k < kData.length; k++) {
          let currId = String(kData[k][kIdIdx] || '').trim();
          if (idMapping[currId]) {
            keuSheet.getRange(k + 1, kIdIdx + 1).setValue(idMapping[currId]);
            keuUpdated++;
          }
        }
        resultLog.push("Sheet KEUANGAN: " + keuUpdated + " record disinkronkan.");
      }
    }

    // 6. Sinkronkan ke Sheet GALERI jika ada
    let galSheet = ss.getSheetByName('GALERI');
    let galUpdated = 0;
    if (galSheet && galSheet.getLastRow() > 1 && Object.keys(idMapping).length > 0) {
      const gData = galSheet.getDataRange().getValues();
      const gHeaders = gData[0].map(h => String(h || '').toUpperCase().trim());
      let gIdIdx = gHeaders.indexOf('ID SISWA PENGUNGGAH');
      if (gIdIdx === -1) gIdIdx = gHeaders.indexOf('ID SISWA');
      if (gIdIdx > -1) {
        for (let g = 1; g < gData.length; g++) {
          let currId = String(gData[g][gIdIdx] || '').trim();
          if (idMapping[currId]) {
            galSheet.getRange(g + 1, gIdIdx + 1).setValue(idMapping[currId]);
            galUpdated++;
          }
        }
        resultLog.push("Sheet GALERI: " + galUpdated + " record disinkronkan.");
      }
    }

    // Invalidate cache server agar client mendapatkan data baru seketika
    invalidateSyncStampCache();
    try { CacheService.getScriptCache().remove('KEU_CACHE_VER'); } catch(e){}
    try { CacheService.getScriptCache().remove('ABSEN_CACHE_VER'); } catch(e){}

    const actorId = (auth.session && (auth.session.uid || auth.session.email)) || 'ADMIN';
    logActivity(actorId, 'RESET_NOMOR_URUT', 'BIODATA', 'Urutkan ulang 4 digit terakhir ID Siswa (0001, 0002, dst) untuk ' + totalSiswa + ' siswa');

    return sanitizeForRPC({
      success: true,
      message: "Nomor urut & 4 digit terakhir ID Siswa berhasil diurutkan ulang!",
      logs: resultLog,
      totalDaftar: totalDaftar,
      totalSiswa: totalSiswa,
      idMapping: idMapping
    });
  } catch (err) {
    return sanitizeForRPC({ success: false, message: "Gagal mereset nomor urut & ID Siswa: " + err.toString() });
  }
}

/**
 * Menghapus cache sync stamp di sisi server saat terjadi mutasi data
 */
function invalidateSyncStampCache() {
  try {
    CacheService.getScriptCache().remove('SYNC_STAMP_V2');
  } catch (e) {
    // Abaikan jika CacheService tidak tersedia
  }
}

/**
 * Mengambil metadata sinkronisasi berupa hash isi data per-sheet dan ringkasan baris.
 * Dilengkapi server-side cache 45 detik (CacheService) agar multi-user hemat kuota Google Apps Script.
 * Mampu mendeteksi penambahan, penghapusan, maupun perubahan in-place (nilai sel) secara presisi.
 */
function getSyncStamp(authToken) {
  try {
    const auth = requireAuth_(authToken, null);
    if (!auth.authorized) return auth.error;
    // 1. Cek cache server (TTL 45 detik)
    try {
      const cached = CacheService.getScriptCache().get('SYNC_STAMP_V2');
      if (cached) {
        const parsed = JSON.parse(cached);
        if (parsed && parsed.success && parsed.stamp) {
          return sanitizeForRPC(parsed);
        }
      }
    } catch (eCache) {
      // Abaikan jika cache miss atau gagal parse
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheets = ss.getSheets();
    const rows = {};
    const sheetMeta = {};
    let totalRows = 0;
    const hashInputs = [];

    // Helper konversi byte array signed dari Utilities.computeDigest ke string hex
    const toHex16 = function(digest) {
      let hex = '';
      const len = Math.min(digest.length, 16);
      for (let j = 0; j < len; j++) {
        const b = digest[j] & 0xFF;
        hex += (b < 16 ? '0' : '') + b.toString(16);
      }
      return hex;
    };

    for (let i = 0; i < sheets.length; i++) {
      const sh = sheets[i];
      const name = sh.getName();
      const r = sh.getLastRow();
      const c = sh.getLastColumn();
      rows[name] = r;
      totalRows += r;

      let sheetHash = '';
      if (r <= 1 || c === 0) {
        sheetHash = '0_0_empty';
      } else {
        try {
          // Baca nilai used range untuk hashing akurat termasuk edit in-place
          const values = sh.getRange(1, 1, r, c).getValues();
          let rawContent = name + ':' + r + 'x' + c + ':';
          for (let vi = 0; vi < values.length; vi++) {
            rawContent += values[vi].join('|') + '\n';
          }
          const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, rawContent, Utilities.Charset.UTF_8);
          sheetHash = toHex16(digest);
        } catch (eHash) {
          // Fallback bila getValues pada sheet gagal
          sheetHash = 'r' + r + 'c' + c;
        }
      }

      sheetMeta[name] = {
        rows: r,
        cols: c,
        hash: sheetHash
      };
      hashInputs.push(name + '=' + sheetHash);
    }

    // Drive timestamp
    let driveTs = 0;
    try {
      driveTs = DriveApp.getFileById(CONFIG.SPREADSHEET_ID).getLastUpdated().getTime();
    } catch (eDrive) {
      driveTs = Date.now();
    }

    // Composite stamp dari seluruh hash sheet yang terurut
    hashInputs.sort();
    const compositeInput = hashInputs.join(';');
    let compositeHash = '';
    try {
      const compDigest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, compositeInput, Utilities.Charset.UTF_8);
      compositeHash = toHex16(compDigest);
    } catch (eComp) {
      compositeHash = String(driveTs) + '_' + String(totalRows);
    }

    const stamp = compositeHash + '|' + String(totalRows);

    const result = {
      success: true,
      stamp: stamp,
      driveUpdatedMs: driveTs,
      totalRows: totalRows,
      rows: rows,
      sheets: sheetMeta,
      serverTime: new Date().toISOString()
    };

    // Simpan ke CacheService (TTL 45 detik)
    try {
      CacheService.getScriptCache().put('SYNC_STAMP_V2', JSON.stringify(result), 45);
    } catch (eSaveCache) {}

    return sanitizeForRPC(result);
  } catch (err) {
    return sanitizeForRPC({
      success: false,
      message: "Gagal mengambil sync stamp: " + err.toString()
    });
  }
}