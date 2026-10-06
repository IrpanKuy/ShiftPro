// --- CONFIG & UTILITIES ---
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
    Utilities.Charset.UTF_8,
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

  // 3. Shifts
  getOrCreateSheet(ss, "Shifts", [
    "id",
    "nama_serah_terima",
    "regu_list",
    "waktu_shift",
    "jenis_shift",
    "petugas_attendance_json",
    "peralatan_kondisi_json",
    "transaksi_gangguan_json",
    "created_by",
    "created_at",
  ]);

  // 4. Categories (Parent Inventaris)
  getOrCreateSheet(ss, "Categories", ["id", "nama_kategori", "created_at"]);

  // 5. InventoryItems (Child Inventaris)
  getOrCreateSheet(ss, "InventoryItems", [
    "id",
    "category_id",
    "nama_sub_item",
    "stok",
    "updated_at",
  ]);

  // 6. StockLogs (Audit Log)
  getOrCreateSheet(ss, "StockLogs", [
    "id",
    "item_id",
    "nama_kategori",
    "nama_sub_item",
    "jenis_mutasi",
    "jumlah",
    "stok_sebelum",
    "stok_sesudah",
    "keterangan",
    "pengubah",
    "timestamp",
  ]);

  // 7. Settings
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

  // Insert Default Inventaris Parent-Child Sample jika kosong
  const catSheet = ss.getSheetByName("Categories");
  if (catSheet.getLastRow() <= 1) {
    const cat1Id = generateId("CAT");
    const cat2Id = generateId("CAT");
    const now = new Date().toISOString();

    catSheet.appendRow([cat1Id, "MCB", now]);
    catSheet.appendRow([cat2Id, "Fuse Link", now]);

    const itemSheet = ss.getSheetByName("InventoryItems");
    itemSheet.appendRow([generateId("ITM"), cat1Id, "2 A", 10, now]);
    itemSheet.appendRow([generateId("ITM"), cat1Id, "4 A", 1, now]);
    itemSheet.appendRow([generateId("ITM"), cat1Id, "6 A", 2, now]);
    itemSheet.appendRow([generateId("ITM"), cat1Id, "10 A", 0, now]);
    itemSheet.appendRow([generateId("ITM"), cat1Id, "16 A", 1, now]);
    itemSheet.appendRow([generateId("ITM"), cat1Id, "20 A", 0, now]);

    itemSheet.appendRow([generateId("ITM"), cat2Id, "6 A", 15, now]);
    itemSheet.appendRow([generateId("ITM"), cat2Id, "10 A", 8, now]);
    itemSheet.appendRow([generateId("ITM"), cat2Id, "15 A", 5, now]);
  }

  return {
    success: true,
    message: "Database dan seeder awal berhasil disiapkan!",
  };
}

// --- ROUTING HTTP / WEB APP ---
function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    return handleApiRequest(e.parameter.action, e.parameter);
  }
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Serah Terima Shift & Inventaris Operasional")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    return handleApiRequest(postData.action, postData.payload);
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ success: false, error: err.toString() }),
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
      case "getInventory":
        result = getInventory(payload.sessionToken);
        break;
      case "saveCategoryWithItems":
        result = saveCategoryWithItems(payload.sessionToken, payload.data);
        break;
      case "quickUpdateStock":
        result = quickUpdateStock(
          payload.sessionToken,
          payload.itemId,
          payload.newStock,
          payload.reason,
        );
        break;
      case "mutateStock":
        result = mutateStock(payload.sessionToken, payload.data);
        break;
      case "getStockLogs":
        result = getStockLogs(payload.sessionToken);
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
          payload.newPassword,
        );
        break;
      default:
        result = { success: false, error: "Aksi tidak dikenal: " + action };
    }
  } catch (err) {
    result = { success: false, error: err.toString() };
  }

  return ContentService.createTextOutput(
    JSON.stringify(sanitizeDates(result)),
  ).setMimeType(ContentService.MimeType.JSON);
}

// --- SESSION & USER AUTHENTICATION ---
function validateSession(token) {
  if (!token) throw new Error("Sesi tidak valid. Silakan login kembali.");
  const ss = getDb();
  const sheet = ss.getSheetByName("Users");
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
          token: row[0], // Menggunakan User ID sebagai token sesi sederhana
        };
        return sanitizeDates({ success: true, user: userObj });
      }
    }
    return { success: false, error: "Username atau password salah!" };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// --- DATA ACCESS LAYER (CRUD) ---

function getDashboardSummary(sessionToken) {
  const user = validateSession(sessionToken);
  const ss = getDb();

  // Shifts Count
  const shiftsSheet = ss.getSheetByName("Shifts");
  const totalShifts = Math.max(0, shiftsSheet.getLastRow() - 1);

  // Inventory Items Count & Low Stock Alert
  const itemsSheet = ss.getSheetByName("InventoryItems");
  const itemData = itemsSheet.getDataRange().getValues();
  let totalItems = 0;
  let lowStockCount = 0;
  for (let i = 1; i < itemData.length; i++) {
    totalItems++;
    if (Number(itemData[i][3]) <= 2) {
      lowStockCount++;
    }
  }

  // Pending Issues (Gangguan Belum Selesai dari Shift Terakhir)
  let pendingIssues = 0;
  const shiftRows = shiftsSheet.getDataRange().getValues();
  if (shiftRows.length > 1) {
    const lastShift = shiftRows[shiftRows.length - 1];
    try {
      const txGangguan = JSON.parse(lastShift[7] || "{}");
      pendingIssues = Number(txGangguan.gangguan_belum_selesai || 0);
    } catch (e) {}
  }

  return sanitizeDates({
    success: true,
    stats: {
      totalShifts: totalShifts,
      totalItems: totalItems,
      lowStockCount: lowStockCount,
      pendingIssues: pendingIssues,
    },
  });
}

function getInitialData(sessionToken) {
  const user = validateSession(sessionToken);
  const dashboard = getDashboardSummary(sessionToken);
  const shifts = getShifts(sessionToken);
  const inventory = getInventory(sessionToken);
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
      inventory: inventory.data,
      teams: teams.data,
      users: users,
    },
  });
}

function getShifts(sessionToken) {
  validateSession(sessionToken);
  const ss = getDb();
  const sheet = ss.getSheetByName("Shifts");
  const rows = sheet.getDataRange().getValues();
  const shifts = [];

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    shifts.push({
      id: r[0],
      nama_serah_terima: r[1],
      regu_list: r[2],
      waktu_shift: r[3],
      jenis_shift: r[4],
      petugas_attendance:
        typeof r[5] === "string" ? JSON.parse(r[5] || "[]") : r[5],
      peralatan_kondisi:
        typeof r[6] === "string" ? JSON.parse(r[6] || "{}") : r[6],
      transaksi_gangguan:
        typeof r[7] === "string" ? JSON.parse(r[7] || "{}") : r[7],
      created_by: r[8],
      created_at: r[9],
    });
  }

  return sanitizeDates({ success: true, data: shifts.reverse() });
}

function createShift(sessionToken, shiftData) {
  const user = validateSession(sessionToken);
  const ss = getDb();
  const sheet = ss.getSheetByName("Shifts");

  const id = generateId("SFT");
  const now = new Date().toISOString();

  sheet.appendRow([
    id,
    shiftData.nama_serah_terima,
    (shiftData.regu_list || []).join(", "),
    shiftData.waktu_shift || now,
    shiftData.jenis_shift,
    JSON.stringify(shiftData.petugas_attendance || []),
    JSON.stringify(shiftData.peralatan_kondisi || {}),
    JSON.stringify(shiftData.transaksi_gangguan || {}),
    user.nama_lengkap,
    now,
  ]);

  return sanitizeDates({
    success: true,
    message: "Serah terima shift berhasil disimpan!",
  });
}

function getInventory(sessionToken) {
  validateSession(sessionToken);
  const ss = getDb();
  const catSheet = ss.getSheetByName("Categories");
  const itemSheet = ss.getSheetByName("InventoryItems");

  const catRows = catSheet.getDataRange().getValues();
  const itemRows = itemSheet.getDataRange().getValues();

  const categoriesMap = {};

  for (let i = 1; i < catRows.length; i++) {
    const cid = catRows[i][0];
    categoriesMap[cid] = {
      id: cid,
      nama_kategori: catRows[i][1],
      created_at: catRows[i][2],
      items: [],
    };
  }

  for (let j = 1; j < itemRows.length; j++) {
    const catId = itemRows[j][1];
    if (categoriesMap[catId]) {
      categoriesMap[catId].items.push({
        id: itemRows[j][0],
        category_id: catId,
        nama_sub_item: itemRows[j][2],
        stok: Number(itemRows[j][3]),
        updated_at: itemRows[j][4],
      });
    }
  }

  const result = Object.keys(categoriesMap).map((key) => categoriesMap[key]);
  return sanitizeDates({ success: true, data: result });
}

function saveCategoryWithItems(sessionToken, data) {
  const user = validateSession(sessionToken);
  const ss = getDb();
  const catSheet = ss.getSheetByName("Categories");
  const itemSheet = ss.getSheetByName("InventoryItems");

  const catId = generateId("CAT");
  const now = new Date().toISOString();

  catSheet.appendRow([catId, data.nama_kategori, now]);

  if (Array.isArray(data.items)) {
    data.items.forEach((sub) => {
      if (sub.nama_sub_item) {
        itemSheet.appendRow([
          generateId("ITM"),
          catId,
          sub.nama_sub_item,
          Number(sub.stok || 0),
          now,
        ]);
      }
    });
  }

  return sanitizeDates({
    success: true,
    message: "Kategori & Sub-item berhasil ditambahkan!",
  });
}

function quickUpdateStock(sessionToken, itemId, newStock, reason) {
  const user = validateSession(sessionToken);
  const ss = getDb();
  const itemSheet = ss.getSheetByName("InventoryItems");
  const catSheet = ss.getSheetByName("Categories");
  const logSheet = ss.getSheetByName("StockLogs");

  const itemRows = itemSheet.getDataRange().getValues();
  const catRows = catSheet.getDataRange().getValues();

  let targetRowIndex = -1;
  let targetItem = null;

  for (let i = 1; i < itemRows.length; i++) {
    if (itemRows[i][0] === itemId) {
      targetRowIndex = i + 1;
      targetItem = {
        id: itemRows[i][0],
        category_id: itemRows[i][1],
        nama_sub_item: itemRows[i][2],
        stok_sebelum: Number(itemRows[i][3]),
      };
      break;
    }
  }

  if (targetRowIndex === -1) throw new Error("Item tidak ditemukan.");

  // Find Category Name
  let catName = "Umum";
  for (let j = 1; j < catRows.length; j++) {
    if (catRows[j][0] === targetItem.category_id) {
      catName = catRows[j][1];
      break;
    }
  }

  const stokSesudah = Number(newStock);
  const selisih = stokSesudah - targetItem.stok_sebelum;
  const jenisMutasi = selisih >= 0 ? "Quick Edit (+)" : "Quick Edit (-)";
  const now = new Date().toISOString();

  // Update stok di Sheet
  itemSheet.getRange(targetRowIndex, 4).setValue(stokSesudah);
  itemSheet.getRange(targetRowIndex, 5).setValue(now);

  // Catat Audit Log
  logSheet.appendRow([
    generateId("LOG"),
    itemId,
    catName,
    targetItem.nama_sub_item,
    jenisMutasi,
    Math.abs(selisih),
    targetItem.stok_sebelum,
    stokSesudah,
    reason || "Quick edit stok dari tabel",
    user.username,
    now,
  ]);

  return sanitizeDates({ success: true, message: "Stok berhasil diperbarui!" });
}

function mutateStock(sessionToken, data) {
  const user = validateSession(sessionToken);
  const ss = getDb();
  const itemSheet = ss.getSheetByName("InventoryItems");
  const catSheet = ss.getSheetByName("Categories");
  const logSheet = ss.getSheetByName("StockLogs");

  const itemRows = itemSheet.getDataRange().getValues();
  const catRows = catSheet.getDataRange().getValues();

  let targetRowIndex = -1;
  let targetItem = null;

  for (let i = 1; i < itemRows.length; i++) {
    if (itemRows[i][0] === data.itemId) {
      targetRowIndex = i + 1;
      targetItem = {
        id: itemRows[i][0],
        category_id: itemRows[i][1],
        nama_sub_item: itemRows[i][2],
        stok_sebelum: Number(itemRows[i][3]),
      };
      break;
    }
  }

  if (targetRowIndex === -1) throw new Error("Item tidak ditemukan.");

  let catName = "Umum";
  for (let j = 1; j < catRows.length; j++) {
    if (catRows[j][0] === targetItem.category_id) {
      catName = catRows[j][1];
      break;
    }
  }

  const qty = Number(data.jumlah);
  let stokSesudah = targetItem.stok_sebelum;

  if (data.jenis_mutasi === "Tambah") {
    stokSesudah += qty;
  } else {
    stokSesudah = Math.max(0, stokSesudah - qty);
  }

  const now = new Date().toISOString();

  itemSheet.getRange(targetRowIndex, 4).setValue(stokSesudah);
  itemSheet.getRange(targetRowIndex, 5).setValue(now);

  logSheet.appendRow([
    generateId("LOG"),
    data.itemId,
    catName,
    targetItem.nama_sub_item,
    data.jenis_mutasi,
    qty,
    targetItem.stok_sebelum,
    stokSesudah,
    data.keterangan || "Mutasi Stok",
    user.username,
    now,
  ]);

  return sanitizeDates({
    success: true,
    message: "Mutasi stok berhasil diproses!",
  });
}

function getStockLogs(sessionToken) {
  const user = validateSession(sessionToken);
  const ss = getDb();
  const sheet = ss.getSheetByName("StockLogs");
  const rows = sheet.getDataRange().getValues();
  const logs = [];

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const logUser = r[9];

    // Filter per pengguna untuk role Petugas
    if (user.role === "Petugas" && logUser !== user.username) {
      continue;
    }

    logs.push({
      id: r[0],
      item_id: r[1],
      nama_kategori: r[2],
      nama_sub_item: r[3],
      jenis_mutasi: r[4],
      jumlah: r[5],
      stok_sebelum: r[6],
      stok_sesudah: r[7],
      keterangan: r[8],
      pengubah: r[9],
      timestamp: r[10],
    });
  }

  return sanitizeDates({ success: true, data: logs.reverse() });
}

function getUsers(sessionToken) {
  const user = validateSession(sessionToken);
  if (user.role !== "Admin")
    throw new Error("Akses ditolak: Hanya Admin yang diizinkan.");

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
    // Edit User
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
    // New User
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

  // Reset Sheet
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
