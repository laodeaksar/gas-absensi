/**
 * =========================================================
 * SETTINGS.GS - Pengaturan aplikasi EduAbsen
 * Tambahkan sebagai file baru di proyek Apps Script
 * (Files > + > Script > beri nama "Settings"). Tidak perlu
 * mengubah isi Code.gs selain 1 baris di submitAttendanceBatch
 * (lihat Index-patch.md, bagian 8).
 *
 * Memakai sheet "Settings" yang sudah ada di DB_SCHEMA:
 * setting_key | setting_value | updated_at
 * =========================================================
 */

const SETTINGS_DEFAULTS = {
  school_name: 'EduAbsen',
  school_address: '',
  principal_name: '',
  principal_nip: '',
  active_academic_year: '',
  edit_window_days: '0'   // 0 = guru boleh mengubah absensi tanggal berapa pun
};

// ---------- Helper internal ----------

function getSettingsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Settings');
  if (!sheet) {
    sheet = ss.insertSheet('Settings');
    const headers = DB_SCHEMA.Settings;
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f4f6');
    sheet.setFrozenRows(1);
  }
  // Simpan nilai sebagai teks agar "2025/2026" atau NIP tidak diubah Sheets
  sheet.getRange(1, 2, sheet.getMaxRows(), 1).setNumberFormat('@');
  return sheet;
}

function getSettingsMap() {
  const map = {};
  Object.keys(SETTINGS_DEFAULTS).forEach(k => { map[k] = SETTINGS_DEFAULTS[k]; });

  getSheetData('Settings').forEach(r => {
    if (r && Object.prototype.hasOwnProperty.call(SETTINGS_DEFAULTS, r.setting_key)) {
      map[r.setting_key] = String(r.setting_value === undefined || r.setting_value === null ? '' : r.setting_value);
    }
  });

  if (!map.school_name) map.school_name = SETTINGS_DEFAULTS.school_name;
  return map;
}

function upsertSettings_(values) {
  const sheet = getSettingsSheet_();
  const data = sheet.getDataRange().getValues();
  const rowByKey = {};
  for (let i = 1; i < data.length; i++) rowByKey[String(data[i][0])] = i + 1;

  const now = new Date().toISOString();
  const newRows = [];
  Object.keys(values).forEach(k => {
    if (rowByKey[k]) sheet.getRange(rowByKey[k], 2, 1, 2).setValues([[values[k], now]]);
    else newRows.push([k, values[k], now]);
  });
  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 3).setValues(newRows);
  }
}

function validateSettings_(d) {
  const out = {};
  const str = (k) => String(d[k] === undefined || d[k] === null ? '' : d[k]).trim();

  if (d.school_name !== undefined) {
    const v = str('school_name');
    if (!v) throw new Error("Nama sekolah wajib diisi.");
    if (v.length > 100) throw new Error("Nama sekolah maksimal 100 karakter.");
    out.school_name = v;
  }
  if (d.school_address !== undefined) {
    const v = str('school_address');
    if (v.length > 200) throw new Error("Alamat sekolah maksimal 200 karakter.");
    out.school_address = v;
  }
  if (d.principal_name !== undefined) {
    const v = str('principal_name');
    if (v.length > 100) throw new Error("Nama kepala sekolah maksimal 100 karakter.");
    out.principal_name = v;
  }
  if (d.principal_nip !== undefined) {
    const v = str('principal_nip');
    if (v && !/^[0-9]{5,30}$/.test(v)) throw new Error("NIP kepala sekolah hanya berisi angka (5-30 digit).");
    out.principal_nip = v;
  }
  if (d.active_academic_year !== undefined) {
    const v = str('active_academic_year');
    if (v) {
      const m = /^(\d{4})\/(\d{4})$/.exec(v);
      if (!m || Number(m[2]) !== Number(m[1]) + 1) throw new Error("Tahun ajaran harus berformat 2025/2026.");
    }
    out.active_academic_year = v;
  }
  if (d.edit_window_days !== undefined) {
    const v = str('edit_window_days') || '0';
    if (!/^[0-9]{1,3}$/.test(v) || Number(v) > 365) throw new Error("Batas edit absensi harus 0-365 hari.");
    out.edit_window_days = String(Number(v));
  }
  return out;
}

// ---------- API untuk frontend ----------

// Tanpa login: hanya data yang aman untuk halaman login
function getPublicSettings() {
  try {
    return { success: true, settings: { school_name: getSettingsMap().school_name } };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// Semua pengguna yang sudah login (dipakai untuk kop & tanda tangan cetak)
function getSettings(token) {
  try {
    requireAuth(token, false, true);
    return { success: true, settings: getSettingsMap() };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// Khusus admin
function saveSettings(token, data) {
  try {
    requireAuth(token, true);
    if (!data) throw new Error("Data pengaturan tidak boleh kosong.");
  } catch (e) {
    return { success: false, message: e.message };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    upsertSettings_(validateSettings_(data));
    return { success: true, message: "Pengaturan berhasil disimpan.", settings: getSettingsMap() };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// Semua pengguna: ganti password sendiri. Sesi lama dibatalkan, harus login ulang.
function changePassword(token, oldPassword, newPassword) {
  let user;
  try {
    user = requireAuth(token, false, true);
    oldPassword = String(oldPassword || '');
    newPassword = String(newPassword || '');
    if (!oldPassword) throw new Error("Password lama wajib diisi.");
    if (newPassword.length < 6) throw new Error("Password baru minimal 6 karakter.");
    if (newPassword === oldPassword) throw new Error("Password baru harus berbeda dari password lama.");
  } catch (e) {
    return { success: false, message: e.message };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const row = getSheetData('Users').find(u => u && u.user_id === user.user_id && isTrue(u.is_active));
    if (!row) throw new Error("Akun tidak ditemukan atau sudah nonaktif.");
    if (hashPassword(oldPassword, row.password_salt) !== row.password_hash) {
      throw new Error("Password lama tidak sesuai.");
    }

    const salt = generateUniqueId();
    updateRecord('Users', 'user_id', user.user_id, {
      password_salt: salt,
      password_hash: hashPassword(newPassword, salt),
      must_change_password: false,
      updated_at: new Date().toISOString()
    });

    const cache = CacheService.getScriptCache();
    cache.put('revoke:' + user.user_id, String(Date.now()), 21600);
    cache.remove(token);
    return { success: true, message: "Password berhasil diubah." };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// Dipanggil dari submitAttendanceBatch: guru dibatasi sesuai edit_window_days
function assertAttendanceEditable(user, date) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(date))) throw new Error("Format tanggal tidak valid.");
  if (!user || user.role === 'admin') return;

  const days = Number(getSettingsMap().edit_window_days) || 0;
  if (days <= 0) return;

  if (String(date) < todayStr(-days)) {
    throw new Error("Absensi tanggal " + date + " sudah dikunci. Guru hanya dapat mengisi atau mengubah absensi maksimal " + days + " hari ke belakang.");
  }
}