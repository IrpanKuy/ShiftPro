// ==========================================
// SHIFTPRO - GOOGLE APPS SCRIPT BACKEND
// ==========================================

const SPREADSHEET_ID = null; // null = Gunakan ActiveSpreadsheet()

function getDb() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getOrCreateSheet(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (headers && headers.length > 0) {
      sheet.appendRow(headers);
      sheet
        .getRange(1, 1, 1, headers.length)
        .setFontWeight("bold")
        .setBackground("#e2e8f0");
    }
  }
  return sheet;
}

function hashPassword(password) {
  if (!password) return "";
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    password,
    Utilities.Charset.UTF_8
  );
  return digest
    .map((byte) => (byte < 0 ? byte + 256 : byte).toString(16).padStart(2, "0"))
    .join("");
}

function sanitizeDates(obj) {
  if (obj === null || obj === undefined) return obj;
  if (obj instanceof Date) {
    return obj.toISOString();
  }
  if (Array.isArray(obj)) {
    return obj.map(sanitizeDates);
  }
  if (typeof obj === "object") {
    const sanitized = {};
    for (const key in obj) {
      if (Object.prototype.hasOwnProperty.call(obj, key)) {
        sanitized[key] = sanitizeDates(obj[key]);
      }
    }
    return sanitized;
  }
  return obj;
}

function generateId(prefix) {
  return (
    (prefix || "ID") +
    "" +
    new Date().getTime() +
    "" +
    Math.floor(Math.random() * 1000)
  );
}

// --- SETUP DATABASE & SEEDER ---
function setupDatabase() {
  const ss = getDb();

  // 1. Users
  getOrCreateSheet(ss, "Users", [
    "id",
    "username",
    "password_hash",
    "nama_lengkap",
    "role",
    "status",
    "created_at",
  ]);

  // 2. Teams
  getOrCreateSheet(ss, "Teams", ["id", "nama_regu", "keterangan"]);

  // 3. Shifts (Revisi Schema Input Data Petugas In/Out & Inventaris Manual)
  getOrCreateSheet(ss, "Shifts", [
    "id",
    "nama_serah_terima",
    "regu_list",
    "waktu_shift",
    "jenis_shift",
    "petugas_in_json",
    "petugas_out_json",
    "peralatan_kondisi_json",
    "inventaris_manual_json",
    "transaksi_gangguan_json",
    "created_by",
    "created_at",
  ]);

  // 4. InventoryConfig (Konfigurasi Dinamis Kategori & Item Checklist Inventaris)
  getOrCreateSheet(ss, "InventoryConfig", [
    "id",
    "nama_kategori",
    "items_json",
    "created_at",
  ]);

  // 5. Settings
  getOrCreateSheet(ss, "Settings", ["key", "value"]);

  // Insert Default Admin & Petugas jika Users kosong
  const usersSheet = ss.getSheetByName("Users");
  if (usersSheet.getLastRow() <= 1) {
    const defaultPasswordHash = hashPassword("admin123");
    const now = new Date().toISOString();
    usersSheet.appendRow([
      generateId("USR"),
      "admin",
      defaultPasswordHash,
      "Administrator Utama",
      "Admin",
      "Aktif",
      now,
    ]);
    usersSheet.appendRow([
      generateId("USR"),
      "petugas1",
      hashPassword("petugas123"),
      "Petugas Operasional 1",
      "Petugas",
      "Aktif",
      now,
    ]);
    usersSheet.appendRow([
      generateId("USR"),
      "petugas2",
      hashPassword("petugas123"),
      "Petugas Operasional 2",
      "Petugas",
      "Aktif",
      now,
    ]);
  }

  // Insert Default Master Regu jika kosong
  const teamsSheet = ss.getSheetByName("Teams");
  if (teamsSheet.getLastRow() <= 1) {
    teamsSheet.appendRow([generateId("TM"), "Gasap", "Regu Area Gasap"]);
    teamsSheet.appendRow([generateId("TM"), "Galu", "Regu Area Galu"]);
    teamsSheet.appendRow([generateId("TM"), "Bajo p.", "Regu Area Bajo"]);
  }

  // Insert Default Inventory Config jika kosong
  const cfgSheet = ss.getSheetByName("InventoryConfig");
  if (cfgSheet.getLastRow() <= 1) {
    const now = new Date().toISOString();
    cfgSheet.appendRow([
      generateId("CFG"),
      "Sisa MCB",
      JSON.stringify([
        "2 A",
        "4 A",
        "6 A",
        "10 A",
        "16 A",
        "20 A",
        "25 A",
        "35 A",
      ]),
      now,
    ]);
    cfgSheet.appendRow([
      generateId("CFG"),
      "Sisa Fuse link",
      JSON.stringify([
        "2 A",
        "3 A",
        "4 A",
        "5 A",
        "6 A",
        "8 A",
        "10 A",
        "15 A",
        "20 A",
        "25 A",
        "30 A",
      ]),
      now,
    ]);
  }

  return {
    success: true,
    message: "Database ShiftPro berhasil disiapkan!",
  };
}

// --- ROUTING HTTP / WEB APP ---
function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    return handleApiRequest(e.parameter.action, e.parameter);
  }
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("ShiftPro - Serah Terima Shift Operasional")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    return handleApiRequest(postData.action, postData.payload);
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ success: false, error: err.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleApiRequest(action, payload) {
  let result;
  try {
    switch (action) {
      case "login":
        result = loginUser(payload.username, payload.password);
        break;
      case "getInitialData":
        result = getInitialData(payload.sessionToken);
        break;
      case "getDashboardData":
        result = getDashboardSummary(payload.sessionToken);
        break;
      case "getShifts":
        result = getShifts(payload.sessionToken);
        break;
      case "createShift":
        result = createShift(payload.sessionToken, payload.data);
        break;
      case "updateShift":
        result = updateShift(payload.sessionToken, payload.shiftId, payload.data);
        break;
      case "deleteShift":
        result = deleteShift(payload.sessionToken, payload.shiftId);
        break;
      case "getInventoryConfig":
        result = getInventoryConfig(payload.sessionToken);
        break;
      case "saveInventoryConfig":
        result = saveInventoryConfig(payload.sessionToken, payload.categories);
        break;
      case "getUsers":
        result = getUsers(payload.sessionToken);
        break;
      case "saveUser":
        result = saveUser(payload.sessionToken, payload.data);
        break;
      case "deleteUser":
        result = deleteUser(payload.sessionToken, payload.userId);
        break;
      case "getTeams":
        result = getTeams(payload.sessionToken);
        break;
      case "saveTeams":
        result = saveTeams(payload.sessionToken, payload.teams);
        break;
      case "changePassword":
        result = changePassword(
          payload.sessionToken,
          payload.oldPassword,
          payload.newPassword
        );
        break;
      default:
        result = { success: false, error: "Aksi tidak dikenal: " + action };
    }
  } catch (err) {
    result = { success: false, error: err.toString() };
  }

  return ContentService.createTextOutput(
    JSON.stringify(sanitizeDates(result))
  ).setMimeType(ContentService.MimeType.JSON);
}

// --- SESSION & USER AUTHENTICATION ---
function validateSession(token) {
  if (!token) throw new Error("Sesi tidak valid. Silakan login kembali.");
  const ss = getDb();
  const sheet = ss.getSheetByName("Users");
  if (!sheet) throw new Error("Database belum diinisialisasi.");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] === token && row[5] === "Aktif") {
      return {
        id: row[0],
        username: row[1],
        nama_lengkap: row[3],
        role: row[4],
      };
    }
  }
  throw new Error("Sesi kadaluarsa atau pengguna tidak aktif.");
}

function loginUser(username, password) {
  try {
    setupDatabase();
    const ss = getDb();
    const sheet = ss.getSheetByName("Users");
    const data = sheet.getDataRange().getValues();
    const inputHash = hashPassword(password);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (
        row[1].toLowerCase() === username.toLowerCase() &&
        row[2] === inputHash
      ) {
        if (row[5] !== "Aktif") {
          return { success: false, error: "Akun Anda dinonaktifkan." };
        }
        const userObj = {
          id: row[0],
          username: row[1],
          nama_lengkap: row[3],
          role: row[4],
          token: row[0],
        };
        return sanitizeDates({ success: true, user: userObj });
      }
    }
    return { success: false, error: "Username atau password salah!" };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// --- DATA ACCESS LAYER ---

function getDashboardSummary(sessionToken) {
  validateSession(sessionToken);
  const ss = getDb();

  const shiftsSheet = ss.getSheetByName("Shifts");
  const totalShifts = Math.max(0, shiftsSheet ? shiftsSheet.getLastRow() - 1 : 0);

  let pendingIssues = 0;
  if (shiftsSheet) {
    const shiftRows = shiftsSheet.getDataRange().getValues();
    if (shiftRows.length > 1) {
      const lastShift = shiftRows[shiftRows.length - 1];
      try {
        const txGangguan = JSON.parse(lastShift[9] || lastShift[7] || "{}");
        pendingIssues = Number(txGangguan.gangguan_belum_selesai || 0);
      } catch (e) {}
    }
  }

  const teamsSheet = ss.getSheetByName("Teams");
  const totalTeams = Math.max(0, teamsSheet ? teamsSheet.getLastRow() - 1 : 0);

  return sanitizeDates({
    success: true,
    stats: {
      totalShifts: totalShifts,
      totalTeams: totalTeams,
      pendingIssues: pendingIssues,
    },
  });
}

function getInitialData(sessionToken) {
  const user = validateSession(sessionToken);
  const dashboard = getDashboardSummary(sessionToken);
  const shifts = getShifts(sessionToken);
  const inventoryConfig = getInventoryConfig(sessionToken);
  const teams = getTeams(sessionToken);
  let users = [user];
  if (user.role === "Admin") {
    try {
      const uRes = getUsers(sessionToken);
      users = uRes.data || [user];
    } catch (e) {
      users = [user];
    }
  }

  return sanitizeDates({
    success: true,
    data: {
      dashboard: dashboard.stats,
      shifts: shifts.data,
      inventoryConfig: inventoryConfig.data,
      teams: teams.data,
      users: users,
    },
  });
}

function getShifts(sessionToken) {
  const user = validateSession(sessionToken);
  const ss = getDb();
  const sheet = ss.getSheetByName("Shifts");
  if (!sheet) return sanitizeDates({ success: true, data: [] });

  const rows = sheet.getDataRange().getValues();
  const shifts = [];

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const createdBy = String(r[10] || r[8] || "");

    // Filtering per pengguna jika role Petugas: HANYA TAMPILKAN DATA BUATAN SENDIRI
    if (user.role === "Petugas") {
      const matchName = (user.nama_lengkap || "").toLowerCase();
      const matchUsername = (user.username || "").toLowerCase();
      const createdByLower = createdBy.toLowerCase();

      if (createdByLower !== matchName && createdByLower !== matchUsername) {
        continue;
      }
    }
    
    // Parse Petugas In & Out with fallback for older format
    let petugasIn = [];
    let petugasOut = [];
    if (typeof r[5] === "string" && r[5].startsWith("[")) {
      try { petugasIn = JSON.parse(r[5] || "[]"); } catch(e){}
    } else if (Array.isArray(r[5])) {
      petugasIn = r[5];
    }

    if (typeof r[6] === "string" && r[6].startsWith("[")) {
      try { petugasOut = JSON.parse(r[6] || "[]"); } catch(e){}
    } else if (Array.isArray(r[6])) {
      petugasOut = r[6];
    }

    // Backup parse jika baris lama memakai 1 kolom attendance
    if (petugasIn.length === 0 && petugasOut.length === 0 && r[5]) {
      try {
        const oldAtt = typeof r[5] === "string" ? JSON.parse(r[5]) : r[5];
        if (Array.isArray(oldAtt)) {
          petugasIn = oldAtt.map(p => p.nama_lengkap || p);
        }
      } catch(e){}
    }

    shifts.push({
      id: r[0],
      nama_serah_terima: r[1],
      regu_list: r[2],
      waktu_shift: r[3],
      jenis_shift: r[4],
      petugas_in: petugasIn,
      petugas_out: petugasOut,
      peralatan_kondisi: typeof r[7] === "string" ? JSON.parse(r[7] || "{}") : r[7],
      inventaris_manual: typeof r[8] === "string" ? JSON.parse(r[8] || "[]") : r[8],
      transaksi_gangguan: typeof r[9] === "string" ? JSON.parse(r[9] || "{}") : r[9],
      created_by: createdBy,
      created_at: r[11] || r[9] || "",
    });
  }

  return sanitizeDates({ success: true, data: shifts.reverse() });
}

function createShift(sessionToken, shiftData) {
  const user = validateSession(sessionToken);
  const ss = getDb();
  let sheet = ss.getSheetByName("Shifts");
  if (!sheet) {
    setupDatabase();
    sheet = ss.getSheetByName("Shifts");
  }

  const id = generateId("SFT");
  const now = new Date().toISOString();

  sheet.appendRow([
    id,
    shiftData.nama_serah_terima,
    (shiftData.regu_list || []).join(", "),
    shiftData.waktu_shift || now,
    shiftData.jenis_shift,
    JSON.stringify(shiftData.petugas_in || []),
    JSON.stringify(shiftData.petugas_out || []),
    JSON.stringify(shiftData.peralatan_kondisi || {}),
    JSON.stringify(shiftData.inventaris_manual || []),
    JSON.stringify(shiftData.transaksi_gangguan || {}),
    user.nama_lengkap,
    now,
  ]);

  return sanitizeDates({
    success: true,
    message: "Serah terima shift & data inventaris berhasil disimpan!",
  });
}

function updateShift(sessionToken, shiftId, shiftData) {
  const user = validateSession(sessionToken);
  if (user.role !== "Admin") throw new Error("Akses ditolak: Hanya Admin yang diizinkan merubah data shift.");

  const ss = getDb();
  const sheet = ss.getSheetByName("Shifts");
  if (!sheet) throw new Error("Sheet Shifts tidak ditemukan.");

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === shiftId) {
      const rowIndex = i + 1;
      sheet.getRange(rowIndex, 2).setValue(shiftData.nama_serah_terima);
      sheet.getRange(rowIndex, 3).setValue((shiftData.regu_list || []).join(", "));
      sheet.getRange(rowIndex, 4).setValue(shiftData.waktu_shift);
      sheet.getRange(rowIndex, 5).setValue(shiftData.jenis_shift);
      sheet.getRange(rowIndex, 6).setValue(JSON.stringify(shiftData.petugas_in || []));
      sheet.getRange(rowIndex, 7).setValue(JSON.stringify(shiftData.petugas_out || []));
      sheet.getRange(rowIndex, 8).setValue(JSON.stringify(shiftData.peralatan_kondisi || {}));
      sheet.getRange(rowIndex, 9).setValue(JSON.stringify(shiftData.inventaris_manual || []));
      sheet.getRange(rowIndex, 10).setValue(JSON.stringify(shiftData.transaksi_gangguan || {}));
      return sanitizeDates({
        success: true,
        message: "Data serah terima shift berhasil diperbarui!",
      });
    }
  }
  throw new Error("Data shift tidak ditemukan.");
}

function deleteShift(sessionToken, shiftId) {
  const user = validateSession(sessionToken);
  if (user.role !== "Admin") throw new Error("Akses ditolak: Hanya Admin yang diizinkan menghapus data shift.");

  const ss = getDb();
  const sheet = ss.getSheetByName("Shifts");
  if (!sheet) throw new Error("Sheet Shifts tidak ditemukan.");

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === shiftId) {
      sheet.deleteRow(i + 1);
      return sanitizeDates({
        success: true,
        message: "Data serah terima shift berhasil dihapus!",
      });
    }
  }
  throw new Error("Data shift tidak ditemukan.");
}

// --- CONFIG CHECKLIST INVENTARIS DINAMIS ---

function getInventoryConfig(sessionToken) {
  validateSession(sessionToken);
  const ss = getDb();
  const sheet = ss.getSheetByName("InventoryConfig");
  if (!sheet) return sanitizeDates({ success: true, data: [] });

  const rows = sheet.getDataRange().getValues();
  const config = [];

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    let items = [];
    try {
      items = typeof r[2] === "string" ? JSON.parse(r[2] || "[]") : r[2];
    } catch (e) {}

    config.push({
      id: r[0],
      nama_kategori: r[1],
      items: items,
      created_at: r[3],
    });
  }

  return sanitizeDates({ success: true, data: config });
}

function saveInventoryConfig(sessionToken, categoriesArray) {
  const user = validateSession(sessionToken);
  if (user.role !== "Admin") throw new Error("Akses ditolak: Hanya Admin yang diizinkan.");

  const ss = getDb();
  let sheet = ss.getSheetByName("InventoryConfig");
  if (!sheet) {
    sheet = getOrCreateSheet(ss, "InventoryConfig", ["id", "nama_kategori", "items_json", "created_at"]);
  }

  sheet.clearContents();
  sheet.appendRow(["id", "nama_kategori", "items_json", "created_at"]);
  sheet.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#e2e8f0");

  const now = new Date().toISOString();
  if (Array.isArray(categoriesArray)) {
    categoriesArray.forEach((cat) => {
      if (cat.nama_kategori) {
        const cleanItems = (cat.items || []).filter((i) => i && i.trim() !== "");
        sheet.appendRow([
          cat.id || generateId("CFG"),
          cat.nama_kategori,
          JSON.stringify(cleanItems),
          now,
        ]);
      }
    });
  }

  return sanitizeDates({
    success: true,
    message: "Pengaturan Kategori & Sub-Kategori Checklist Inventaris berhasil disimpan!",
  });
}

// --- MANAGEMENT USER & REGU ---

function getUsers(sessionToken) {
  const user = validateSession(sessionToken);
  if (user.role !== "Admin") throw new Error("Akses ditolak: Hanya Admin yang diizinkan.");

  const ss = getDb();
  const sheet = ss.getSheetByName("Users");
  const rows = sheet.getDataRange().getValues();
  const users = [];

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    users.push({
      id: r[0],
      username: r[1],
      nama_lengkap: r[3],
      role: r[4],
      status: r[5],
      created_at: r[6],
    });
  }

  return sanitizeDates({ success: true, data: users });
}

function saveUser(sessionToken, userData) {
  const currentUser = validateSession(sessionToken);
  if (currentUser.role !== "Admin") throw new Error("Akses ditolak.");

  const ss = getDb();
  const sheet = ss.getSheetByName("Users");
  const rows = sheet.getDataRange().getValues();

  const now = new Date().toISOString();

  if (userData.id) {
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === userData.id) {
        sheet.getRange(i + 1, 2).setValue(userData.username);
        if (userData.password) {
          sheet.getRange(i + 1, 3).setValue(hashPassword(userData.password));
        }
        sheet.getRange(i + 1, 4).setValue(userData.nama_lengkap);
        sheet.getRange(i + 1, 5).setValue(userData.role);
        sheet.getRange(i + 1, 6).setValue(userData.status);
        return sanitizeDates({
          success: true,
          message: "Data petugas/user berhasil diperbarui!",
        });
      }
    }
  } else {
    const newId = generateId("USR");
    const hashed = hashPassword(userData.password || "123456");
    sheet.appendRow([
      newId,
      userData.username,
      hashed,
      userData.nama_lengkap,
      userData.role,
      userData.status,
      now,
    ]);
    return sanitizeDates({
      success: true,
      message: "Petugas/user baru berhasil ditambahkan!",
    });
  }
}

function deleteUser(sessionToken, userId) {
  const currentUser = validateSession(sessionToken);
  if (currentUser.role !== "Admin") throw new Error("Akses ditolak.");

  const ss = getDb();
  const sheet = ss.getSheetByName("Users");
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === userId) {
      sheet.deleteRow(i + 1);
      return sanitizeDates({
        success: true,
        message: "Pengguna berhasil dihapus!",
      });
    }
  }
  throw new Error("Pengguna tidak ditemukan.");
}

function getTeams(sessionToken) {
  validateSession(sessionToken);
  const ss = getDb();
  const sheet = ss.getSheetByName("Teams");
  if (!sheet) return sanitizeDates({ success: true, data: [] });

  const rows = sheet.getDataRange().getValues();
  const teams = [];

  for (let i = 1; i < rows.length; i++) {
    teams.push({
      id: rows[i][0],
      nama_regu: rows[i][1],
      keterangan: rows[i][2],
    });
  }

  return sanitizeDates({ success: true, data: teams });
}

function saveTeams(sessionToken, teamsArray) {
  const user = validateSession(sessionToken);
  if (user.role !== "Admin") throw new Error("Akses ditolak.");

  const ss = getDb();
  const sheet = ss.getSheetByName("Teams");

  sheet.clearContents();
  sheet.appendRow(["id", "nama_regu", "keterangan"]);
  sheet.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#e2e8f0");

  teamsArray.forEach((t) => {
    sheet.appendRow([
      t.id || generateId("TM"),
      t.nama_regu,
      t.keterangan || "",
    ]);
  });

  return sanitizeDates({
    success: true,
    message: "Master Data Regu berhasil disimpan!",
  });
}

function changePassword(sessionToken, oldPassword, newPassword) {
  const user = validateSession(sessionToken);
  const ss = getDb();
  const sheet = ss.getSheetByName("Users");
  const rows = sheet.getDataRange().getValues();

  const oldHash = hashPassword(oldPassword);
  const newHash = hashPassword(newPassword);

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === user.id) {
      if (rows[i][2] !== oldHash) {
        return { success: false, error: "Password lama tidak cocok!" };
      }
      sheet.getRange(i + 1, 3).setValue(newHash);
      return sanitizeDates({
        success: true,
        message: "Password berhasil diganti!",
      });
    }
  }

  throw new Error("Pengguna tidak ditemukan.");
}
