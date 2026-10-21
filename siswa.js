/**
 * =========================================================
 * SISWA.GS - Login siswa & lihat kehadiran sendiri
 * Tambahkan sebagai file baru di proyek Apps Script
 * (Files > + > Script > beri nama "Siswa").
 *
 * Aturan akun siswa:
 *   - username      = NIS
 *   - password awal = NISN (siswa diminta menggantinya saat login)
 *   - role          = 'siswa', terhubung ke Students lewat Users.student_id
 * =========================================================
 */

// ---------- Helper internal ----------

// Pastikan kolom baru ada dan kolom username bertipe teks
// (agar NIS seperti "0123" tidak berubah menjadi angka 123)
function prepareUsersSheet_() {
  ensureColumns('Users');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = headers.indexOf('username') + 1;
  if (idx > 0) sheet.getRange(1, idx, sheet.getMaxRows(), 1).setNumberFormat('@');
  return sheet;
}

// Susun record akun dari data siswa. Mengembalikan { record } atau { skip: alasan }
function buildStudentUser_(s, taken, now) {
  const nis = String(s.nis || '').trim();
  const nisn = String(s.nisn || '').trim();
  if (nis.length < 3) return { skip: (s.full_name || s.student_id) + ': NIS kosong/kurang dari 3 karakter' };
  if (nisn.length < 6) return { skip: nis + ': NISN kosong/kurang dari 6 karakter' };
  if (taken[nis.toLowerCase()]) return { skip: nis + ': username sudah dipakai akun lain' };

  const salt = generateUniqueId();
  return {
    record: {
      user_id: 'USR-' + generateUniqueId().substring(0, 8).toUpperCase(),
      username: nis,
      password_hash: hashPassword(nisn, salt),
      password_salt: salt,
      full_name: s.full_name || 'Siswa',
      role: 'siswa',
      assigned_classes: '[]',
      is_active: true,
      created_at: now,
      updated_at: now,
      student_id: s.student_id,
      must_change_password: true
    }
  };
}

// ---------- API admin ----------

// Buat akun untuk semua siswa aktif yang belum punya akun
function generateStudentAccounts(token) {
  try { requireAuth(token, true); } catch (e) { return { success: false, message: e.message }; }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = prepareUsersSheet_();
    const users = getSheetData('Users').filter(Boolean);
    const hasAccount = {}, taken = {};
    users.forEach(u => {
      if (u.student_id) hasAccount[u.student_id] = true;
      taken[String(u.username).toLowerCase()] = true;
    });

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const students = getSheetData('Students').filter(s => s && isTrue(s.is_active));
    const now = new Date().toISOString();
    const rows = [], skipped = [];

    students.forEach(s => {
      if (hasAccount[s.student_id]) return;
      const built = buildStudentUser_(s, taken, now);
      if (built.skip) { skipped.push(built.skip); return; }
      taken[String(built.record.username).toLowerCase()] = true;
      rows.push(headers.map(h => built.record[h] !== undefined ? built.record[h] : ''));
    });

    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
    }

    let msg = rows.length + " akun siswa dibuat.";
    if (skipped.length > 0) {
      msg += " " + skipped.length + " dilewati (" + skipped.slice(0, 3).join("; ") + (skipped.length > 3 ? "; ..." : "") + ").";
    }
    return { success: true, message: msg };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// Buat akun satu siswa, atau reset password-nya ke NISN
function resetStudentAccount(token, studentId) {
  try {
    requireAuth(token, true);
    if (!studentId) throw new Error("Siswa tidak valid.");
  } catch (e) {
    return { success: false, message: e.message };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    prepareUsersSheet_();
    const student = getSheetData('Students').find(s => s && s.student_id === studentId);
    if (!student) throw new Error("Data siswa tidak ditemukan.");

    const users = getSheetData('Users').filter(Boolean);
    const now = new Date().toISOString();
    const existing = users.find(u => u.student_id === studentId);

    if (existing) {
      const nisn = String(student.nisn || '').trim();
      if (nisn.length < 6) throw new Error("NISN siswa kosong atau terlalu pendek untuk dijadikan password awal.");
      const salt = generateUniqueId();
      updateRecord('Users', 'user_id', existing.user_id, {
        password_salt: salt,
        password_hash: hashPassword(nisn, salt),
        must_change_password: true,
        updated_at: now
      });
      CacheService.getScriptCache().put('revoke:' + existing.user_id, String(Date.now()), 21600);
      return { success: true, message: "Password siswa direset ke NISN." };
    }

    const taken = {};
    users.forEach(u => { taken[String(u.username).toLowerCase()] = true; });
    const built = buildStudentUser_(student, taken, now);
    if (built.skip) throw new Error(built.skip);
    insertRecord('Users', built.record);
    return { success: true, message: "Akun siswa dibuat. Username = NIS, password awal = NISN." };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ---------- API siswa ----------

// Hanya data milik siswa yang sedang login (student_id diambil dari sesi, bukan dari parameter)
function getMyAttendance(token, filters) {
  try {
    const user = requireAuth(token, false, true);
    if (user.role !== 'siswa' || !user.student_id) throw new Error("FORBIDDEN");

    filters = filters || {};
    const isDate = (v) => /^\d{4}-\d{2}-\d{2}$/.test(String(v));
    const start = filters.startDate ? String(filters.startDate) : '';
    const end = filters.endDate ? String(filters.endDate) : '';
    if ((start && !isDate(start)) || (end && !isDate(end))) throw new Error("Format tanggal tidak valid.");

    const student = getSheetData('Students').find(s => s && s.student_id === user.student_id);
    if (!student) throw new Error("Data siswa tidak ditemukan.");

    const classMap = {};
    getSheetData('Classes').forEach(c => { if (c) classMap[c.class_id] = c.class_name; });

    const classNames = getSheetData('StudentClassHistory')
      .filter(h => h && h.student_id === student.student_id && isTrue(h.is_active))
      .map(h => classMap[h.class_id]).filter(Boolean);

    const mine = getSheetData('Attendance').filter(a => a && a.student_id === student.student_id);
    const todayRec = mine.find(a => dateOnly(a.attendance_date) === todayStr(0));

    const inRange = mine.filter(a => {
      const d = dateOnly(a.attendance_date);
      return (!start || d >= start) && (!end || d <= end);
    }).sort((a, b) => dateOnly(b.attendance_date).localeCompare(dateOnly(a.attendance_date)));

    const summary = { hadir: 0, sakit: 0, izin: 0, alpa: 0, persentase: 0 };
    inRange.forEach(a => {
      if (a.status === 'Hadir') summary.hadir++;
      else if (a.status === 'Sakit') summary.sakit++;
      else if (a.status === 'Izin') summary.izin++;
      else if (a.status === 'Alpa') summary.alpa++;
    });
    const counted = summary.hadir + summary.sakit + summary.izin + summary.alpa;
    if (counted > 0) summary.persentase = ((summary.hadir / counted) * 100).toFixed(1);

    return {
      success: true,
      profile: { full_name: student.full_name || '', nis: String(student.nis || ''), class_name: classNames.join(', ') || '-' },
      today: todayRec ? { status: todayRec.status || '', notes: todayRec.notes || '' } : null,
      summary: summary,
      records: inRange.slice(0, 500).map(a => ({
        date: dateOnly(a.attendance_date),
        status: a.status || '-',
        notes: a.notes || '',
        class_name: classMap[a.class_id] || '-'
      }))
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}