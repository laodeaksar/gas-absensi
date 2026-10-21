/**
 * =========================================================
 * GOOGLE APPS SCRIPT - BACKEND ABSENSI SISWA
 * Role: API Backend & Database Manager (V8 V2.2 - Null Safe + Fix Data Siswa)
 * =========================================================
 */

const DB_SCHEMA = {
  Users: ["user_id", "username", "password_hash", "password_salt", "full_name", "role", "assigned_classes", "is_active", "created_at", "updated_at", "nip", "phone", "email", "student_id", "must_change_password"],
  Classes: ["class_id", "class_name", "grade_level", "academic_year", "homeroom_teacher", "description", "is_active", "created_at", "updated_at"],
  Students: ["student_id", "nis", "nisn", "full_name", "gender", "birth_place", "birth_date", "address", "parent_name", "parent_phone", "is_active", "created_at", "updated_at"],
  StudentClassHistory: ["history_id", "student_id", "class_id", "academic_year", "start_date", "end_date", "is_active", "created_at"],
  Attendance: ["attendance_id", "attendance_date", "student_id", "class_id", "status", "notes", "recorded_by", "created_at", "updated_at"],
  AttendanceAudit: ["audit_id", "attendance_id", "action", "old_values", "new_values", "changed_by", "changed_at"],
  Settings: ["setting_key", "setting_value", "updated_at"]
};

// Kolom yang harus disimpan sebagai TEKS (agar angka 0 di depan tidak hilang
// dan tanggal tidak diubah otomatis oleh Sheets menjadi objek Date)
const TEXT_COLUMNS = {
  Students: ["nis", "nisn", "parent_phone", "birth_date"],
  Users: ["nip", "phone"],
  Attendance: ["attendance_date"]
};

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Sistem Absensi Siswa')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==========================================
// A. INISIALISASI & SETUP DATABASE
// ==========================================
function initializeDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let isNew = false;

  for (const [sheetName, headers] of Object.entries(DB_SCHEMA)) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f4f6");
      sheet.setFrozenRows(1);
      sheet.autoResizeColumns(1, headers.length);
      isNew = true;
    }
  }

  migrateSchema();
  applyTextFormats();
  if (isNew) createDummyData();
  return { success: true, message: "Database siap digunakan." };
}

/**
 * Jalankan SEKALI secara manual dari editor jika sheet sudah ada sebelumnya.
 * Mengatur kolom nis, nisn, parent_phone, birth_date, attendance_date menjadi teks.
 */
function applyTextFormats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (const [sheetName, cols] of Object.entries(TEXT_COLUMNS)) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    cols.forEach(col => {
      const idx = headers.indexOf(col) + 1;
      if (idx > 0) sheet.getRange(1, idx, sheet.getMaxRows(), 1).setNumberFormat('@');
    });
  }
}

// Tambahkan kolom schema yang belum ada di sheet (di akhir header), tanpa mengubah data lama
function ensureColumns(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return;
  const textCols = TEXT_COLUMNS[sheetName] || [];
  const current = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0].map(String);
  DB_SCHEMA[sheetName].forEach(h => {
    if (current.indexOf(h) !== -1) return;
    const col = sheet.getLastColumn() + 1;
    if (col > sheet.getMaxColumns()) sheet.insertColumnsAfter(sheet.getMaxColumns(), 1);
    sheet.getRange(1, col).setValue(h).setFontWeight("bold").setBackground("#f3f4f6");
    if (textCols.indexOf(h) !== -1) sheet.getRange(1, col, sheet.getMaxRows(), 1).setNumberFormat('@');
    current.push(h);
  });
}

/** Opsional: jalankan manual untuk migrasi semua sheet. saveTeacher juga memanggil ensureColumns('Users') otomatis. */
function migrateSchema() {
  Object.keys(DB_SCHEMA).forEach(ensureColumns);
}

function createDummyData() {
  const salt = generateUniqueId();
  const passHash = hashPassword("admin123", salt);
  const now = new Date().toISOString();

  insertRecord('Users', {
    user_id: 'USR-001', username: 'admin', password_hash: passHash, password_salt: salt,
    full_name: 'Administrator Utama', role: 'admin', assigned_classes: '[]', is_active: true, created_at: now, updated_at: now
  });

  const saltGuru = generateUniqueId();
  const passHashGuru = hashPassword("guru123", saltGuru);
  insertRecord('Users', {
    user_id: 'USR-002', username: 'guru01', password_hash: passHashGuru, password_salt: saltGuru,
    full_name: 'Budi Santoso, S.Pd', role: 'guru', assigned_classes: '["CLS-001"]', is_active: true, created_at: now, updated_at: now
  });

  insertRecord('Classes', {
    class_id: 'CLS-001', class_name: 'X MIPA 1', grade_level: 'X', academic_year: '2023/2024',
    homeroom_teacher: 'USR-002', description: 'Kelas Unggulan', is_active: true, created_at: now, updated_at: now
  });

  insertRecord('Students', {
    student_id: 'STD-001', nis: '1001', nisn: '0011223344', full_name: 'Andi Wijaya', gender: 'L',
    birth_place: 'Jakarta', birth_date: '2007-05-12', address: 'Jl. Merdeka No.1', parent_name: 'Bapak Andi', parent_phone: '08123456789', is_active: true, created_at: now, updated_at: now
  });
  insertRecord('Students', {
    student_id: 'STD-002', nis: '1002', nisn: '0011223355', full_name: 'Siti Aminah', gender: 'P',
    birth_place: 'Bandung', birth_date: '2007-08-20', address: 'Jl. Sudirman No.2', parent_name: 'Bapak Siti', parent_phone: '08129876543', is_active: true, created_at: now, updated_at: now
  });

  insertRecord('StudentClassHistory', {
    history_id: 'HST-001', student_id: 'STD-001', class_id: 'CLS-001', academic_year: '2023/2024',
    start_date: now, end_date: '', is_active: true, created_at: now
  });
  insertRecord('StudentClassHistory', {
    history_id: 'HST-002', student_id: 'STD-002', class_id: 'CLS-001', academic_year: '2023/2024',
    start_date: now, end_date: '', is_active: true, created_at: now
  });
}

// ==========================================
// B. UTILITAS & CORE DATABASE OPERATIONS
// ==========================================
function generateUniqueId() { return Utilities.getUuid(); }

function isTrue(v) { return v === true || v === 'TRUE' || v === 'true'; }

// Tanggal hari ini (zona waktu script), format yyyy-MM-dd
function todayStr(offsetDays) {
  const d = new Date();
  if (offsetDays) d.setDate(d.getDate() + offsetDays);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

// Ambil bagian tanggal (yyyy-MM-dd) dari nilai apa pun
function dateOnly(v) {
  return v ? String(v).split('T')[0] : '';
}

function hashPassword(password, salt) {
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + salt, Utilities.Charset.UTF_8);
  return rawHash.map(v => ("0" + ((v < 0) ? 256 + v : v).toString(16)).slice(-2)).join("");
}

function getSheetData(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (!data || data.length <= 1) return [];
  const headers = data[0];
  const tz = Session.getScriptTimeZone();
  return data.slice(1).map(row => {
    let obj = {};
    headers.forEach((h, i) => {
      let val = row[i];
      if (val instanceof Date) {
        // Pakai zona waktu lokal agar tanggal tidak mundur 1 hari
        val = Utilities.formatDate(val, tz, "yyyy-MM-dd'T'HH:mm:ss");
      }
      obj[h] = val !== undefined && val !== null ? val : "";
    });
    return obj;
  });
}

function insertRecord(sheetName, recordObj) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(header => recordObj[header] !== undefined && recordObj[header] !== null ? recordObj[header] : "");
  sheet.appendRow(row);
  return recordObj;
}

function updateRecord(sheetName, idColumn, idValue, updateObj) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  if (!data || data.length <= 1) return false;
  const headers = data[0];
  const idIndex = headers.indexOf(idColumn);

  for (let i = 1; i < data.length; i++) {
    if (data[i][idIndex] === idValue) {
      const rowRange = sheet.getRange(i + 1, 1, 1, headers.length);
      const rowValues = rowRange.getValues()[0];
      headers.forEach((header, colIdx) => {
        if (updateObj[header] !== undefined && updateObj[header] !== null) {
          rowValues[colIdx] = updateObj[header];
        }
      });
      rowRange.setValues([rowValues]);
      return true;
    }
  }
  return false;
}

// ==========================================
// C. AUTENTIKASI
// ==========================================
function login(username, password) {
  try {
    const users = getSheetData('Users');
    const user = users.find(u => u && String(u.username) === String(username) && isTrue(u.is_active));
    if (!user) throw new Error("Username atau password salah.");

    const hashedInput = hashPassword(password, user.password_salt);
    if (hashedInput !== user.password_hash) throw new Error("Username atau password salah.");

    const token = generateUniqueId();
    const cache = CacheService.getScriptCache();
    const sessionData = {
      user_id: user.user_id, username: user.username, full_name: user.full_name || 'Pengguna',
      role: user.role, assigned_classes: user.assigned_classes || '[]', issued_at: Date.now(),
      student_id: user.student_id || '', must_change_password: isTrue(user.must_change_password)
    };
    cache.put(token, JSON.stringify(sessionData), 21600); // 6 jam
    return { success: true, token: token, user: sessionData };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove(token);
  return { success: true };
}

function requireAuth(token, requireAdmin = false, allowStudent = false) {
  if (!token) throw new Error("AUTH_FAILED");
  const sessionStr = CacheService.getScriptCache().get(token);
  if (!sessionStr) throw new Error("AUTH_FAILED");
  const user = JSON.parse(sessionStr);
  // Sesi dibatalkan jika akun diubah/dinonaktifkan/direset setelah login
  const revokedAt = CacheService.getScriptCache().get('revoke:' + user.user_id);
  if (revokedAt && Number(revokedAt) > Number(user.issued_at || 0)) {
    CacheService.getScriptCache().remove(token);
    throw new Error("AUTH_FAILED");
  }
  if (user.role === 'siswa' && !allowStudent) throw new Error("FORBIDDEN");
  if (requireAdmin && user.role !== 'admin') throw new Error("FORBIDDEN");
  return user;
}

// ==========================================
// D. MASTER DATA API (CRUD)
// ==========================================
function fetchMasterData(token, type) {
  try {
    const user = requireAuth(token);
    let data = [];
    if (type === 'users' && user.role === 'admin') {
      data = getSheetData('Users').map(u => {
        if (u) { delete u.password_hash; delete u.password_salt; }
        return u;
      }).filter(Boolean);
    } else if (type === 'teachers' && user.role === 'admin') {
      data = getSheetData('Users').filter(u => u && u.role === 'guru').map(u => {
        u.nip = String(u.nip || ''); u.phone = String(u.phone || ''); u.email = String(u.email || '');
        delete u.password_hash; delete u.password_salt;
        return u;
      });
    } else if (type === 'classes') {
      data = getSheetData('Classes').filter(Boolean);
    } else if (type === 'students') {
      // Pastikan semua nilai aman dikirim ke browser (angka/teks/boolean saja)
      data = getSheetData('Students').filter(Boolean).map(s => {
        s.nis = String(s.nis);
        s.nisn = String(s.nisn);
        s.parent_phone = String(s.parent_phone);
        s.birth_date = dateOnly(s.birth_date);
        return s;
      });
    }
    return { success: true, data: data };
  } catch (e) {
    // Kembalikan error sebagai objek agar pesan aslinya terlihat di frontend
    return { success: false, message: e.message };
  }
}

function saveStudent(token, studentData) {
  try {
    requireAuth(token, true);
    if (!studentData) throw new Error("Data siswa tidak boleh kosong.");
  } catch (e) {
    return { success: false, message: e.message };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    // Samakan tipe data (teks) agar perbandingan & leading zero aman
    studentData.nis = String(studentData.nis || '').trim();
    studentData.nisn = String(studentData.nisn || '').trim();
    if (studentData.parent_phone !== undefined) studentData.parent_phone = String(studentData.parent_phone);

    const students = getSheetData('Students');
    const isDuplicate = students.some(s => s &&
      ((studentData.nis && String(s.nis) === studentData.nis) || (studentData.nisn && String(s.nisn) === studentData.nisn)) &&
      s.student_id !== studentData.student_id);
    if (isDuplicate) throw new Error("NIS atau NISN sudah digunakan.");

    studentData.updated_at = new Date().toISOString();
    if (studentData.student_id) {
      updateRecord('Students', 'student_id', studentData.student_id, studentData);
    } else {
      studentData.student_id = 'STD-' + generateUniqueId().substring(0, 8).toUpperCase();
      studentData.created_at = studentData.updated_at;
      studentData.is_active = true;
      insertRecord('Students', studentData);

      // Jika frontend mengirim class_id, langsung masukkan siswa ke kelas tsb
      if (studentData.class_id) {
        const cls = getSheetData('Classes').find(c => c && c.class_id === studentData.class_id);
        insertRecord('StudentClassHistory', {
          history_id: 'HST-' + generateUniqueId().substring(0, 8).toUpperCase(),
          student_id: studentData.student_id, class_id: studentData.class_id,
          academic_year: cls ? cls.academic_year : '', start_date: studentData.updated_at,
          end_date: '', is_active: true, created_at: studentData.updated_at
        });
      }
    }
    return { success: true, message: "Data siswa berhasil disimpan." };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// Manajemen guru (khusus admin). Buat baru jika user_id kosong, ubah jika ada.
// Field yang tidak dikirim dipertahankan dari data lama (mis. hanya is_active saat toggle).
function saveTeacher(token, d) {
  try {
    requireAuth(token, true);
    if (!d) throw new Error("Data guru tidak boleh kosong.");
  } catch (e) {
    return { success: false, message: e.message };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const users = getSheetData('Users');
    const now = new Date().toISOString();
    ensureColumns('Users');
    const isEdit = !!d.user_id;

    let existing = null;
    if (isEdit) {
      existing = users.find(u => u && u.user_id === d.user_id);
      if (!existing) throw new Error("Data guru tidak ditemukan.");
      if (existing.role !== 'guru') throw new Error("Hanya akun guru yang dapat diubah di sini.");
    }

    const username = String(d.username !== undefined ? d.username : (existing ? existing.username : '')).trim();
    const fullName = String(d.full_name !== undefined ? d.full_name : (existing ? existing.full_name : '')).trim();
    if (!fullName) throw new Error("Nama lengkap wajib diisi.");
    if (!/^[A-Za-z0-9._-]{3,30}$/.test(username)) {
      throw new Error("Username 3-30 karakter: huruf, angka, titik, garis bawah, atau strip.");
    }

    const ownId = existing ? existing.user_id : '';
    const isDuplicate = users.some(u => u && String(u.username).toLowerCase() === username.toLowerCase() && u.user_id !== ownId);
    if (isDuplicate) throw new Error("Username sudah digunakan.");

    const pick = (k) => String(d[k] !== undefined && d[k] !== null ? d[k] : (existing ? (existing[k] || '') : '')).trim();
    const nip = pick('nip'), phone = pick('phone'), email = pick('email');
    if (nip && !/^[0-9]{5,30}$/.test(nip)) throw new Error("NIP hanya berisi angka (5-30 digit).");
    if (phone && !/^[0-9+][0-9\s-]{5,19}$/.test(phone)) throw new Error("Nomor HP tidak valid.");
    if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error("Format email tidak valid.");
    if (nip && users.some(u => u && String(u.nip) === nip && u.user_id !== ownId)) throw new Error("NIP sudah digunakan.");

    const pw = d.password ? String(d.password) : '';
    if (!isEdit && pw.length < 6) throw new Error("Password guru baru minimal 6 karakter.");
    if (pw && pw.length < 6) throw new Error("Password minimal 6 karakter.");

    let assigned = existing ? (existing.assigned_classes || '[]') : '[]';
    if (d.assigned_classes !== undefined) {
      if (!Array.isArray(d.assigned_classes)) throw new Error("Data kelas tidak valid.");
      const validIds = getSheetData('Classes').filter(Boolean).map(c => c.class_id);
      assigned = JSON.stringify(d.assigned_classes.filter(id => validIds.indexOf(id) !== -1));
    }

    let isActive = true;
    if (d.is_active !== undefined) isActive = isTrue(d.is_active);
    else if (existing) isActive = isTrue(existing.is_active);

    const record = {
      username: username, full_name: fullName, role: 'guru',
      assigned_classes: assigned, is_active: isActive, updated_at: now,
      nip: nip, phone: phone, email: email
    };
    if (pw) {
      const salt = generateUniqueId();
      record.password_salt = salt;
      record.password_hash = hashPassword(pw, salt);
    }

    if (isEdit) {
      updateRecord('Users', 'user_id', existing.user_id, record);
      // Paksa guru login ulang agar perubahan kelas/status/password berlaku
      CacheService.getScriptCache().put('revoke:' + existing.user_id, String(Date.now()), 21600);
    } else {
      record.user_id = 'USR-' + generateUniqueId().substring(0, 8).toUpperCase();
      record.created_at = now;
      insertRecord('Users', record);
    }
    return { success: true, message: "Data guru berhasil disimpan." };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function saveClass(token, classData) {
  try {
    requireAuth(token, true);
    if (!classData) throw new Error("Data kelas tidak boleh kosong.");
  } catch (e) {
    return { success: false, message: e.message };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const classes = getSheetData('Classes');
    const isDuplicate = classes.some(c => c && c.class_name === classData.class_name && c.academic_year === classData.academic_year && c.class_id !== classData.class_id);
    if (isDuplicate) throw new Error("Nama kelas sudah ada pada tahun ajaran tersebut.");

    if (classData.homeroom_teacher) {
      const wali = getSheetData('Users').find(u => u && u.user_id === classData.homeroom_teacher && u.role === 'guru' && isTrue(u.is_active));
      if (!wali) throw new Error("Wali kelas tidak valid.");
    }

    classData.updated_at = new Date().toISOString();
    if (classData.class_id) {
      updateRecord('Classes', 'class_id', classData.class_id, classData);
    } else {
      classData.class_id = 'CLS-' + generateUniqueId().substring(0, 8).toUpperCase();
      classData.created_at = classData.updated_at;
      classData.is_active = true;
      insertRecord('Classes', classData);
    }
    return { success: true, message: "Data kelas berhasil disimpan." };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// E. ABSENSI (ATTENDANCE)
// ==========================================
function getAttendanceForm(token, date, classId) {
  const user = requireAuth(token);
  if (user.role === 'guru') {
    let assigned = [];
    try { assigned = JSON.parse(user.assigned_classes || "[]"); } catch (e) {}
    if (!assigned.includes(classId)) throw new Error("Anda tidak memiliki akses ke kelas ini.");
  }

  const students = getSheetData('Students').filter(s => s && isTrue(s.is_active));
  const histories = getSheetData('StudentClassHistory').filter(h => h && h.class_id === classId && isTrue(h.is_active));
  const attendances = getSheetData('Attendance').filter(a => a && dateOnly(a.attendance_date) === date && a.class_id === classId);

  const studentMap = {};
  students.forEach(s => { if (s && s.student_id) studentMap[s.student_id] = s; });

  const classStudents = histories.map(h => {
    const s = studentMap[h.student_id];
    if (!s) return null;
    const att = attendances.find(a => a && a.student_id === h.student_id);
    return {
      student_id: s.student_id || '',
      nis: String(s.nis || '-'),
      full_name: s.full_name || 'Tanpa Nama',
      status: att ? (att.status || 'Belum Ada') : 'Belum Ada',
      notes: att ? (att.notes || '') : '',
      attendance_id: att ? att.attendance_id : null
    };
  }).filter(Boolean).sort((a, b) => (a.full_name || '').localeCompare(b.full_name || ''));

  return { success: true, data: classStudents };
}

function submitAttendanceBatch(token, date, classId, attendanceData) {
  const user = requireAuth(token);
  assertAttendanceEditable(user, date);
  if (!Array.isArray(attendanceData)) throw new Error("Data absensi tidak valid.");

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Attendance');
    const existing = getSheetData('Attendance');
    const now = new Date().toISOString();

    const existMap = {};
    existing.filter(a => a && dateOnly(a.attendance_date) === date && a.class_id === classId)
            .forEach(a => { if (a && a.student_id) existMap[a.student_id] = a; });

    const rowsToInsert = [];

    attendanceData.forEach(entry => {
      if (!entry || !entry.student_id) return;
      if (existMap[entry.student_id]) {
        updateRecord('Attendance', 'attendance_id', existMap[entry.student_id].attendance_id, {
          status: entry.status || 'Belum Ada', notes: entry.notes || '', recorded_by: user.user_id, updated_at: now
        });
      } else {
        rowsToInsert.push([
          'ATT-' + generateUniqueId().substring(0, 8).toUpperCase(), date, entry.student_id, classId,
          entry.status || 'Belum Ada', entry.notes || '', user.user_id, now, now
        ]);
      }
    });

    if (rowsToInsert.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, rowsToInsert.length, rowsToInsert[0].length).setValues(rowsToInsert);
    }
    return { success: true, message: "Data absensi berhasil disimpan." };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// F. DASHBOARD & LAPORAN
// ==========================================
function getDashboardStats(token) {
  requireAuth(token);
  const today = todayStr(0);

  const students = getSheetData('Students').filter(s => s && isTrue(s.is_active)).length;
  const classes = getSheetData('Classes').filter(c => c && isTrue(c.is_active)).length;
  const allAtt = getSheetData('Attendance').filter(Boolean);

  let stats = { totalStudents: students, totalClasses: classes, hadir: 0, sakit: 0, izin: 0, alpa: 0, persentase: 0 };

  allAtt.filter(a => a && dateOnly(a.attendance_date) === today).forEach(a => {
    if (a.status === 'Hadir') stats.hadir++;
    else if (a.status === 'Sakit') stats.sakit++;
    else if (a.status === 'Izin') stats.izin++;
    else if (a.status === 'Alpa') stats.alpa++;
  });

  const totalAtt = stats.hadir + stats.sakit + stats.izin + stats.alpa;
  if (totalAtt > 0) stats.persentase = ((stats.hadir / totalAtt) * 100).toFixed(1);

  const chartData = { labels: [], hadir: [], tidakHadir: [] };
  const sevenDaysAgo = todayStr(-7);

  const recentAtt = allAtt.filter(a => a && dateOnly(a.attendance_date) >= sevenDaysAgo);
  const dateMap = {};
  recentAtt.forEach(a => {
    const attDate = dateOnly(a.attendance_date);
    if (!attDate) return;
    if (!dateMap[attDate]) dateMap[attDate] = { h: 0, t: 0 };
    if (a.status === 'Hadir') dateMap[attDate].h++;
    else dateMap[attDate].t++;
  });

  const dates = Object.keys(dateMap).sort();
  dates.forEach(date => {
    chartData.labels.push(date);
    chartData.hadir.push(dateMap[date].h);
    chartData.tidakHadir.push(dateMap[date].t);
  });

  return { success: true, stats: stats, chartData: chartData };
}

function getReportData(token, filters) {
  const user = requireAuth(token);
  filters = filters || {};
  const attendances = getSheetData('Attendance').filter(Boolean);
  const students = getSheetData('Students').filter(Boolean);
  const classes = getSheetData('Classes').filter(Boolean);

  let filtered = attendances.filter(a => {
    if (!a) return false;
    const attDate = dateOnly(a.attendance_date);
    if (filters.startDate && attDate < filters.startDate) return false;
    if (filters.endDate && attDate > filters.endDate) return false;
    if (filters.classId && a.class_id !== filters.classId) return false;
    return true;
  });

  if (user.role === 'guru') {
    let assigned = [];
    try { assigned = JSON.parse(user.assigned_classes || "[]"); } catch (e) {}
    filtered = filtered.filter(a => assigned.includes(a.class_id));
  }

  const studentMap = {}; students.forEach(s => { if (s && s.student_id) studentMap[s.student_id] = s; });
  const classMap = {}; classes.forEach(c => { if (c && c.class_id) classMap[c.class_id] = c; });

  const report = filtered.map(a => {
    const s = studentMap[a.student_id] || {};
    const c = classMap[a.class_id] || {};
    return {
      date: dateOnly(a.attendance_date) || '-',
      student_name: s.full_name || 'Siswa Tidak Ditemukan',
      nis: String(s.nis || '-'),
      class_name: c.class_name || 'Kelas Tidak Ditemukan',
      status: a.status || '-',
      notes: a.notes || '-'
    };
  });

  return { success: true, data: report };
}