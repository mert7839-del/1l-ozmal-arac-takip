const express = require("express");
const cors = require("cors");
const multer = require("multer");
const nodemailer = require("nodemailer");
const path = require("path");
const fs = require("fs");
const ExcelJS = require("exceljs");
require("dotenv").config();

const app = express();
const PORT = process.env.PORT || 3000;

const DATA_DIR = path.join(__dirname, "data");
const UPLOAD_DIR = path.join(__dirname, "uploads");
const DB_FILE = path.join(DATA_DIR, "db.json");

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR);
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR);

const OPERATIONS = [
  "Domino's Gebze", "Domino's Avrupa", "Domino's İzmir", "Domino's Gaziantep", "Domino's Ankara",
  "Migros Torbalı", "Migros Menemen", "Fasdat Gebze", "Fasdat Hadımköy", "Papa John's", "Barsan"
];

function normalizeOperations(list) {
  if (!Array.isArray(list)) return [];
  return [...new Set(list.map(x => String(x || "").trim()).filter(x => OPERATIONS.includes(x)))];
}

function userCanAccessOperation(user, operation) {
  if (!user) return false;
  if (user.role === "Admin") return true;
  return normalizeOperations(user.operations).includes(String(operation || "").trim());
}

function visibleDbForUser(db, user) {
  if (!user) return { ...db, users: [], vehicles: [], vehicleLogs: [], invoices: [], monthlyKm: [], vehicleDocuments: [], suppliers: [], mailRules: [] };
  
  const sanitizedUsers = (db.users || []).map(u => { const x={...u}; delete x.password; return x; });
  if (user.role === "Admin") {
    return { ...db, users: sanitizedUsers };
  }

  const vehicles = (db.vehicles || []).filter(v => userCanAccessOperation(user, v.operation));
  const plates = new Set(vehicles.map(v => v.plate));

  return {
    ...db,
    users: sanitizedUsers,
    vehicles,
    vehicleLogs: (db.vehicleLogs || []).filter(x => plates.has(x.plate)),
    invoices: (db.invoices || []).filter(x => plates.has(x.plate)),
    monthlyKm: (db.monthlyKm || []).filter(x => plates.has(x.plate)),
    vehicleDocuments: (db.vehicleDocuments || []).filter(x => plates.has(x.plate)),
    suppliers: (db.suppliers || []).filter(x => plates.has(x.plate))
  };
}

const ALL_PERMISSIONS = [
  "dashboard.view",
  "vehicles.view", "vehicles.create", "vehicles.edit", "vehicles.delete", "vehicles.inspection",
  "logs.view", "logs.create", "logs.edit", "logs.delete",
  "monthlyKm.view", "monthlyKm.create", "monthlyKm.delete", "monthlyKm.export",
  "invoices.view", "invoices.create", "invoices.edit", "invoices.approve", "invoices.pay", "invoices.delete",
  "penalties.view", "penalties.create", "penalties.edit", "penalties.delete", "penalties.export",
  "suppliers.view", "suppliers.create", "suppliers.delete",
  "documents.view", "documents.create", "documents.delete",
  "users.manage", "roles.manage", "settings.manage"
];

const DEFAULT_ROLES = [
  { name: "Admin", permissions: ALL_PERMISSIONS.slice(), builtIn: true },
  { name: "Yetkili", permissions: ["dashboard.view","vehicles.view","logs.view","logs.create","logs.edit","monthlyKm.view","monthlyKm.create","invoices.view","invoices.pay","penalties.view","penalties.create","suppliers.view","suppliers.create","documents.view","documents.create"], builtIn: true },
  { name: "Yetkili2", permissions: ["dashboard.view","vehicles.view","vehicles.create","vehicles.edit","vehicles.delete","vehicles.inspection","logs.view","logs.create","logs.edit","logs.delete","monthlyKm.view","monthlyKm.create","monthlyKm.delete","monthlyKm.export","invoices.view","invoices.create","invoices.edit","invoices.approve","invoices.delete","penalties.view","penalties.create","penalties.edit","penalties.delete","penalties.export","suppliers.view","suppliers.create","suppliers.delete","documents.view","documents.create","documents.delete"], builtIn: true },
  { name: "Kullanıcı", permissions: ["dashboard.view","vehicles.view","logs.view","logs.create","monthlyKm.view","monthlyKm.create","invoices.view","invoices.create","penalties.view","penalties.create","suppliers.view","suppliers.create","documents.view","documents.create"], builtIn: true }
];

function normalizeRoles(db) {
  if (!Array.isArray(db.roles)) db.roles = [];
  DEFAULT_ROLES.forEach(def => {
    const existing = db.roles.find(r => r.name === def.name);
    if (!existing) db.roles.push(JSON.parse(JSON.stringify(def)));
    else { existing.permissions = Array.isArray(existing.permissions) ? existing.permissions.filter(p => ALL_PERMISSIONS.includes(p)) : def.permissions.slice(); existing.builtIn = true; }
  });
  db.roles.forEach(r => {
    r.permissions = Array.from(new Set((Array.isArray(r.permissions) ? r.permissions : []).filter(p => ALL_PERMISSIONS.includes(p))));
    if (r.name === "Admin") r.permissions = ALL_PERMISSIONS.slice();
  });
}

function roleByName(db, roleName) { return (db.roles || []).find(r => r.name === roleName); }
function userPermissions(db, user) {
  if (!user) return [];
  if (user.role === "Admin") return ALL_PERMISSIONS.slice();
  const role = roleByName(db, user.role);
  return role ? role.permissions : [];
}
function hasPermission(db, user, permission) { return userPermissions(db, user).includes(permission); }
function canManageUsers(db, user) { return !!user && (user.role === "Admin" || hasPermission(db, user, "users.manage")); }
function canManageRoles(db, user) { return !!user && (user.role === "Admin" || hasPermission(db, user, "roles.manage")); }

const defaultDb = {
  roles: JSON.parse(JSON.stringify(DEFAULT_ROLES)),
  users: [
    { name: "Admin", username: "admin", password: "1234", role: "Admin", mail: "admin@firma.com", operations: [] },
    { name: "Yetkili", username: "yetkili", password: "1234", role: "Yetkili", mail: "yetkili@firma.com", operations: [] },
    { name: "Yetkili 2", username: "yetkili2", password: "1234", role: "Yetkili2", mail: "yetkili2@firma.com", operations: [] },
    { name: "Kullanıcı", username: "kullanici", password: "1234", role: "Kullanıcı", mail: "kullanici@firma.com", operations: [] }
  ],
  vehicles: [],
  vehicleLogs: [],
  invoices: [],
  monthlyKm: [],
  vehicleDocuments: [],
  suppliers: [],
  mailRules: []
};

function readDb() {
  let db;
  try {
    if (!fs.existsSync(DB_FILE)) {
      fs.writeFileSync(DB_FILE, JSON.stringify(defaultDb, null, 2), "utf-8");
      db = JSON.parse(JSON.stringify(defaultDb));
    } else {
      const raw = fs.readFileSync(DB_FILE, "utf-8").trim();
      db = !raw ? JSON.parse(JSON.stringify(defaultDb)) : JSON.parse(raw);
    }
  } catch (err) {
    db = JSON.parse(JSON.stringify(defaultDb));
  }

  db.users ||= [];
  normalizeRoles(db);
  db.users = db.users.map(u => ({ ...u, operations: normalizeOperations(u.operations) }));
  db.vehicles ||= [];
  db.vehicleLogs ||= [];
  db.invoices ||= [];
  db.monthlyKm ||= [];
  db.vehicleDocuments ||= [];
  db.suppliers ||= [];
  db.mailRules ||= [];

  const requiredUsers = [
    { name: "Admin", username: "admin", password: "1234", role: "Admin", mail: "admin@firma.com", operations: [] },
    { name: "Yetkili", username: "yetkili", password: "1234", role: "Yetkili", mail: "yetkili@firma.com", operations: [] },
    { name: "Yetkili 2", username: "yetkili2", password: "1234", role: "Yetkili2", mail: "yetkili2@firma.com", operations: [] },
    { name: "Kullanıcı", username: "kullanici", password: "1234", role: "Kullanıcı", mail: "kullanici@firma.com", operations: [] }
  ];

  requiredUsers.forEach(u => {
    if (!db.users.some(x => x.username === u.username)) db.users.push(u);
  });

  writeDb(db);
  return db;
}

function writeDb(db) {
  fs.writeFileSync(DB_FILE, JSON.stringify(db, null, 2), "utf-8");
}

function recoverDbIfEmpty(db) {
  let changed = false;
  if (!db || typeof db !== "object") {
    db = JSON.parse(JSON.stringify(defaultDb));
    changed = true;
  }
  db.mailRules ||= [];
  if (changed) writeDb(db);
  return db;
}

function today() { return new Date().toLocaleDateString("tr-TR"); }
function oneYearFromTodayISO() {
  const d = new Date();
  d.setFullYear(d.getFullYear() + 1);
  return d.toISOString().slice(0, 10);
}
function userByUsername(db, username) { return db.users.find(u => u.username === username); }

function createTransporter() {
  return nodemailer.createTransport({
    host: process.env.SMTP_HOST,
    port: Number(process.env.SMTP_PORT || 587),
    secure: String(process.env.SMTP_SECURE).toLowerCase() === "true",
    auth: { user: process.env.SMTP_USER, pass: process.env.SMTP_PASS },
    connectionTimeout: 8000,
    greetingTimeout: 8000,
    socketTimeout: 8000
  });
}

// ==================== GELİŞTİRİLMİŞ ÇOKLU ALICI VE OPERASYON/PLAKA MAIL KURALI ====================
function getRecipientsForVehicle(db, plate, operation) {
  let emails = [];
  const rules = Array.isArray(db.mailRules) ? db.mailRules : [];

  const cleanTargetPlate = String(plate || "").trim().toUpperCase();
  const cleanTargetOp = String(operation || "").trim();

  // 1. Spesifik plaka kuralı var mı kontrol et
  const plateRules = rules.filter(r => r.plate && r.plate !== "Tümü" && r.plate.toUpperCase() === cleanTargetPlate);
  if (plateRules.length > 0) {
    plateRules.forEach(r => {
      if (r.email) {
        const splitMails = r.email.split(",").map(e => e.trim()).filter(Boolean);
        emails.push(...splitMails);
      }
    });
  }

  // 2. Operasyon bazlı kurallara bak (Plaka "Tümü" olanlar veya boş olanlar)
  const opRules = rules.filter(r => (!r.plate || r.plate === "Tümü") && r.operation && r.operation === cleanTargetOp);
  opRules.forEach(r => {
    if (r.email) {
      const splitMails = r.email.split(",").map(e => e.trim()).filter(Boolean);
      emails.push(...splitMails);
    }
  });

  // 3. Genel "Tüm Operasyonlar / Genel" kurallarına bak
  const generalRules = rules.filter(r => (!r.operation || r.operation === "Tümü") && (!r.plate || r.plate === "Tümü"));
  generalRules.forEach(r => {
    if (r.email) {
      const splitMails = r.email.split(",").map(e => e.trim()).filter(Boolean);
      emails.push(...splitMails);
    }
  });

  // 4. Eğer hiçbir özel kural eşleşmediyse genel .env veya admin maillerine düş
  if (emails.length === 0) {
    if (process.env.NOTIFY_EMAILS) {
      const envEmails = process.env.NOTIFY_EMAILS.split(",").map(e => e.trim()).filter(Boolean);
      emails.push(...envEmails);
    }
    const adminMails = (db.users || []).filter(u => u.role === "Admin" && u.mail).map(u => u.mail);
    emails.push(...adminMails);
  }

  return Array.from(new Set(emails));
}

async function sendMailForVehicle(db, plate, operation, subject, text, attachment) {
  try {
    if (String(process.env.MAIL_ENABLED || "true").toLowerCase() === "false") {
      return { sent: false, reason: "MAIL_ENABLED=false" };
    }

    const recipients = getRecipientsForVehicle(db, plate, operation);
    if (recipients.length === 0) return { sent: false, reason: "Alıcı e-posta adresi bulunamadı." };
    if (!process.env.SMTP_HOST || !process.env.SMTP_USER || !process.env.SMTP_PASS) {
      return { sent: false, reason: ".env SMTP ayarları eksik." };
    }

    const mail = {
      from: `"1L Özmal Araç Takip" <${process.env.SMTP_USER}>`,
      to: recipients.join(","),
      subject,
      text
    };

    if (attachment) mail.attachments = Array.isArray(attachment) ? attachment : [attachment];

    await createTransporter().sendMail(mail);
    return { sent: true };
  } catch (err) {
    console.error("MAIL_SEND_ERROR:", err.message);
    return { sent: false, reason: err.message };
  }
}

app.use(cors());
app.use(express.json({ limit: "20mb" }));
app.use(express.static(path.join(__dirname, "public")));
app.use("/uploads", express.static(UPLOAD_DIR));

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, UPLOAD_DIR),
  filename: (req, file, cb) => {
    const safe = file.originalname.replace(/[^a-zA-Z0-9ğüşöçıİĞÜŞÖÇ._-]/g, "_");
    cb(null, Date.now() + "_" + safe);
  }
});
const upload = multer({ storage });

app.get("/api/data", (req, res) => {
  const db = recoverDbIfEmpty(readDb());
  const requester = userByUsername(db, req.query.requester);
  if (!requester) return res.status(401).json({ ok:false, message:"Giriş gerekli." });
  const out = visibleDbForUser(JSON.parse(JSON.stringify(db)), requester);
  res.json(out);
});

app.post("/api/login", (req, res) => {
  const db = recoverDbIfEmpty(readDb());
  const { username, password } = req.body;
  const user = db.users.find(u => u.username === username && u.password === password);
  if (!user) return res.status(401).json({ ok: false, message: "Kullanıcı adı veya şifre hatalı." });
  res.json({ ok: true, user: { name: user.name, username: user.username, role: user.role, mail: user.mail, operations: normalizeOperations(user.operations) } });
});

// ==================== ARAÇ İŞLEM / BAKIM KAYDI SİLME ENDPOINTİ ====================
app.delete("/api/logs/:id", (req, res) => {
  try {
    const db = readDb();
    db.vehicleLogs ||= [];
    const requester = userByUsername(db, req.query.requester);
    if (!hasPermission(db, requester, "logs.delete") && requester?.role !== "Admin") {
      return res.status(403).json({ ok: false, message: "İşlem kaydı silme yetkiniz yok." });
    }

    const id = String(req.params.id);
    db.vehicleLogs = db.vehicleLogs.filter(x => String(x.id) !== id);
    writeDb(db);

    const filteredDb = visibleDbForUser(JSON.parse(JSON.stringify(db)), requester);
    return res.json({ ok: true, db: filteredDb });
  } catch (err) {
    res.status(500).json({ ok: false, message: err.message });
  }
});

app.post("/api/vehicle-documents", upload.any(), async (req, res) => {
  try {
    const db = readDb();
    db.vehicleDocuments ||= [];

    const requester = req.body.requester;
    const user = (db.users || []).find(u => u.username === requester);
    if (!user) return res.status(401).json({ ok: false, message: "Giriş gerekli." });

    const plate = req.body.plate || "";
    const docType = req.body.docType || "Diğer";
    const docDate = req.body.docDate || today();
    const note = req.body.note || "";
    const amount = Number(req.body.amount || 0);
    const faultRate = req.body.faultRate || "";
    const otherParty = req.body.otherParty || "";

    const uploadedFiles = req.files || [];
    if (uploadedFiles.length === 0) {
      return res.status(400).json({ ok: false, message: "Yüklenecek dosya bulunamadı." });
    }

    uploadedFiles.forEach(f => {
      const docRecord = {
        id: Date.now() + Math.random(),
        plate,
        docType,
        docDate,
        date: docDate,
        note,
        amount,
        faultRate,
        otherParty,
        fileName: f.originalname,
        fileUrl: "/uploads/" + f.filename,
        mimeType: f.mimetype,
        addedBy: user.name || user.username,
        addedByUsername: user.username,
        createdAt: new Date().toISOString()
      };
      db.vehicleDocuments.unshift(docRecord);
    });

    writeDb(db);
    const filteredDb = visibleDbForUser(JSON.parse(JSON.stringify(db)), user);
    return res.json({ ok: true, db: filteredDb });
  } catch (err) {
    console.error("DOC_UPLOAD_ERROR:", err.message);
    return res.status(500).json({ ok: false, message: err.message });
  }
});

app.delete("/api/vehicle-documents/:id", (req, res) => {
  try {
    const db = readDb();
    db.vehicleDocuments ||= [];
    const requester = userByUsername(db, req.query.requester);
    if (!hasPermission(db, requester, "documents.delete") && requester?.role !== "Admin") {
      return res.status(403).json({ ok: false, message: "Evrak silme yetkiniz yok." });
    }

    db.vehicleDocuments = db.vehicleDocuments.filter(x => x.id !== Number(req.params.id));
    writeDb(db);

    const filteredDb = visibleDbForUser(JSON.parse(JSON.stringify(db)), requester);
    return res.json({ ok: true, db: filteredDb });
  } catch (err) {
    res.status(500).json({ ok: false, message: err.message });
  }
});

app.post("/api/vehicles/:plate/inspection-done", (req, res) => {
  const db = readDb();
  const requester = userByUsername(db, req.body.requester);
  if (!hasPermission(db, requester, "vehicles.inspection") && requester?.role !== "Admin") {
    return res.status(403).json({ ok: false, message: "Muayene yetkiniz yok." });
  }

  const plate = req.params.plate;
  const vehicle = db.vehicles.find(v => v.plate === plate);
  if (!vehicle) return res.status(404).json({ ok: false, message: "Araç bulunamadı." });

  vehicle.inspection = oneYearFromTodayISO();
  db.vehicleLogs.unshift({
    id: Date.now(),
    plate,
    date: today(),
    km: vehicle.km || 0,
    text: "Muayene yapıldı. Yeni muayene tarihi: " + vehicle.inspection,
    cost: 0,
    addedBy: requester.name || requester.username
  });

  writeDb(db);
  const filteredDb = visibleDbForUser(JSON.parse(JSON.stringify(db)), requester);
  res.json({ ok: true, db: filteredDb, message: "Muayene tarihi 1 yıl ileri alındı." });
});

app.post("/api/vehicles", (req, res) => {
  const db = readDb();
  const requester = userByUsername(db, req.body.requester);
  if (!hasPermission(db, requester, "vehicles.create") && requester?.role !== "Admin") {
    return res.status(403).json({ ok: false, message: "Araç ekleme yetkiniz yok." });
  }

  const vehicle = req.body.vehicle || {};
  if (!vehicle.plate) return res.status(400).json({ ok: false, message: "Plaka zorunludur." });
  
  vehicle.plate = vehicle.plate.toUpperCase();
  vehicle.operation = String(vehicle.operation || "").trim();

  if (db.vehicles.some(v => v.plate === vehicle.plate)) {
    return res.status(400).json({ ok: false, message: "Bu plaka zaten kayıtlı." });
  }

  db.vehicles.push(vehicle);
  db.vehicleLogs.unshift({
    id: Date.now(),
    plate: vehicle.plate,
    date: today(),
    km: vehicle.km || 0,
    text: "Araç sisteme eklendi",
    cost: 0,
    addedBy: requester.name || requester.username
  });

  writeDb(db);
  const filteredDb = visibleDbForUser(JSON.parse(JSON.stringify(db)), requester);
  res.json({ ok: true, db: filteredDb });
});

app.post("/api/suppliers", upload.any(), async (req, res) => {
  try {
    const db = readDb();
    db.suppliers ||= [];

    const requester = req.body.requester;
    const user = (db.users || []).find(u => u.username === requester);
    if (!user) return res.status(401).json({ ok: false, message: "Giriş gerekli." });

    const processType = req.body.processType || "Diğer";
    const supplierName = req.body.supplierName || "";
    const plate = req.body.plate || "";
    const amount = Number(req.body.amount || 0);
    const note = req.body.note || "";
    const processDate = req.body.processDate || today();

    const uploadedFiles = req.files || [];
    if (uploadedFiles.length === 0) {
      return res.status(400).json({ ok: false, message: "Fatura veya fotoğraf dosyası zorunludur." });
    }

    const files = uploadedFiles.map(f => ({
      fileName: f.originalname,
      fileUrl: "/uploads/" + f.filename,
      mimeType: f.mimetype,
      path: f.path
    }));

    const record = {
      id: Date.now(),
      processType,
      supplierName,
      plate,
      amount,
      note,
      processDate,
      date: processDate,
      files,
      fileName: files[0]?.fileName || "",
      fileUrl: files[0]?.fileUrl || "",
      addedBy: user.name || user.username,
      addedByUsername: user.username,
      createdAt: new Date().toISOString()
    };

    db.suppliers.unshift(record);
    writeDb(db);

    const filteredDb = visibleDbForUser(JSON.parse(JSON.stringify(db)), user);
    return res.json({ ok: true, db: filteredDb, record });
  } catch (err) {
    console.error("SUPPLIER_UPLOAD_ERROR:", err.message);
    return res.status(500).json({ ok: false, message: err.message });
  }
});

app.delete("/api/suppliers/:id", (req, res) => {
  try {
    const db = readDb();
    db.suppliers ||= [];
    const requester = userByUsername(db, req.query.requester);
    if (!hasPermission(db, requester, "suppliers.delete") && requester?.role !== "Admin") {
      return res.status(403).json({ ok: false, message: "Tedarikçi kaydı silme yetkiniz yok." });
    }

    db.suppliers = db.suppliers.filter(x => x.id !== Number(req.params.id));
    writeDb(db);

    const filteredDb = visibleDbForUser(JSON.parse(JSON.stringify(db)), requester);
    return res.json({ ok: true, db: filteredDb });
  } catch (err) {
    res.status(500).json({ ok: false, message: err.message });
  }
});

app.put("/api/vehicles/:plate", (req, res) => {
  const db = readDb();
  const requester = userByUsername(db, req.body.requester);
  if (!hasPermission(db, requester, "vehicles.edit") && requester?.role !== "Admin") {
    return res.status(403).json({ ok: false, message: "Araç düzenleme yetkiniz yok." });
  }

  const oldPlate = req.params.plate;
  const vehicle = db.vehicles.find(v => v.plate === oldPlate);
  if (!vehicle) return res.status(404).json({ ok: false, message: "Araç bulunamadı." });

  const newPlate = (req.body.plate || oldPlate).toUpperCase();
  if (newPlate !== oldPlate && db.vehicles.some(v => v.plate === newPlate)) {
    return res.status(400).json({ ok: false, message: "Yeni plaka zaten kayıtlı." });
  }

  vehicle.plate = newPlate;
  vehicle.brand = req.body.brand || vehicle.brand;
  vehicle.modelYear = Number(req.body.modelYear || vehicle.modelYear || 0);
  vehicle.operation = String(req.body.operation || "").trim();
  vehicle.km = Number(req.body.km || vehicle.km);
  vehicle.inspection = req.body.inspection || vehicle.inspection;
  vehicle.driverName = req.body.driverName || "";
  vehicle.driverPhone = req.body.driverPhone || "";
  vehicle.salary = Number(req.body.salary || 0);
  vehicle.mealMoney = Number(req.body.mealMoney || 0);
  vehicle.premium = Number(req.body.premium || 0);
  vehicle.srcDate = req.body.srcDate || "";
  vehicle.maintenanceInterval = Number(req.body.maintenanceInterval || 0);
  vehicle.lastMaintenanceKm = Number(req.body.lastMaintenanceKm || vehicle.lastMaintenanceKm || 0);

  if (newPlate !== oldPlate) {
    db.vehicleLogs.forEach(l => { if (l.plate === oldPlate) l.plate = newPlate; });
    db.invoices.forEach(i => { if (i.plate === oldPlate) i.plate = newPlate; });
    db.monthlyKm.forEach(m => { if (m.plate === oldPlate) m.plate = newPlate; });
  }

  writeDb(db);
  const filteredDb = visibleDbForUser(JSON.parse(JSON.stringify(db)), requester);
  res.json({ ok: true, db: filteredDb });
});

// ==================== ARAÇ SİLME ENDPOINTİ ====================
app.delete("/api/vehicles/:plate", (req, res) => {
  const db = readDb();
  const requester = userByUsername(db, req.query.requester);
  if (!hasPermission(db, requester, "vehicles.delete") && requester?.role !== "Admin") {
    return res.status(403).json({ ok: false, message: "Araç silme yetkiniz yok." });
  }

  const plate = decodeURIComponent(req.params.plate);
  db.vehicles = (db.vehicles || []).filter(v => v.plate !== plate);
  db.vehicleLogs = (db.vehicleLogs || []).filter(l => l.plate !== plate);
  db.invoices = (db.invoices || []).filter(i => i.plate !== plate);

  writeDb(db);
  const filteredDb = visibleDbForUser(JSON.parse(JSON.stringify(db)), requester);
  res.json({ ok: true, db: filteredDb });
});

app.put("/api/users/:username/operations", (req, res) => {
  const db = readDb();
  const requester = userByUsername(db, req.body.requester);
  if (!canManageUsers(db, requester)) return res.status(403).json({ ok: false, message: "Sadece admin operasyon yetkisi verebilir." });

  const target = userByUsername(db, req.params.username);
  if (!target) return res.status(404).json({ ok: false, message: "Kullanıcı bulunamadı." });

  target.operations = target.role === "Admin" ? [] : normalizeOperations(req.body.operations);
  writeDb(db);
  const filteredDb = visibleDbForUser(JSON.parse(JSON.stringify(db)), requester);
  res.json({ ok: true, db: filteredDb });
});

app.post("/api/roles", (req, res) => {
  const db = readDb();
  const requester = userByUsername(db, req.body.requester);
  if (!canManageRoles(db, requester)) return res.status(403).json({ ok: false, message: "Sadece admin rol yönetebilir." });

  const role = req.body.role || {};
  const name = String(role.name || "").trim();
  if (!name) return res.status(400).json({ ok: false, message: "Rol adı zorunlu." });
  if (roleByName(db, name)) return res.status(400).json({ ok: false, message: "Bu rol zaten var." });

  const permissions = Array.from(new Set((Array.isArray(role.permissions) ? role.permissions : []).filter(p => ALL_PERMISSIONS.includes(p))));
  db.roles.push({ name, permissions, builtIn: false });
  writeDb(db);
  const filteredDb = visibleDbForUser(JSON.parse(JSON.stringify(db)), requester);
  res.json({ ok: true, db: filteredDb });
});

app.put("/api/roles/:name", (req, res) => {
  const db = readDb();
  const requester = userByUsername(db, req.body.requester);
  if (!canManageRoles(db, requester)) return res.status(403).json({ ok: false, message: "Sadece admin rol yönetebilir." });

  const role = roleByName(db, req.params.name);
  if (!role) return res.status(404).json({ ok: false, message: "Rol bulunamadı." });
  if (role.name === "Admin") return res.status(400).json({ ok: false, message: "Admin rolü değiştirilemez." });

  const patch = req.body.role || {};
  const newName = String(patch.name || role.name).trim();
  if (!newName) return res.status(400).json({ ok: false, message: "Rol adı zorunlu." });

  const oldName = role.name;
  role.name = newName;
  role.permissions = Array.from(new Set((Array.isArray(patch.permissions) ? patch.permissions : role.permissions).filter(p => ALL_PERMISSIONS.includes(p))));
  
  db.users.forEach(u => { if (u.role === oldName) u.role = newName; });
  writeDb(db);
  const filteredDb = visibleDbForUser(JSON.parse(JSON.stringify(db)), requester);
  res.json({ ok: true, db: filteredDb });
});

app.delete("/api/roles/:name", (req, res) => {
  const db = readDb();
  const requester = userByUsername(db, req.query.requester);
  if (!canManageRoles(db, requester)) return res.status(403).json({ ok: false, message: "Sadece admin rol yönetebilir." });

  const role = roleByName(db, req.params.name);
  if (!role) return res.status(404).json({ ok: false, message: "Rol bulunamadı." });
  if (role.builtIn || role.name === "Admin") return res.status(400).json({ ok: false, message: "Sistem rolleri silinemez." });

  if (db.users.some(u => u.role === role.name)) {
    return res.status(400).json({ ok: false, message: "Bu role atanmış kullanıcılar var." });
  }

  db.roles = db.roles.filter(r => r.name !== role.name);
  writeDb(db);
  const filteredDb = visibleDbForUser(JSON.parse(JSON.stringify(db)), requester);
  res.json({ ok: true, db: filteredDb });
});

app.post("/api/users", (req, res) => {
  const db = readDb();
  const requester = userByUsername(db, req.body.requester);
  if (!canManageUsers(db, requester)) return res.status(403).json({ ok: false, message: "Sadece admin üye yönetebilir." });

  const user = req.body.user || {};
  if (!user.name || !user.username || !user.password || !user.role) {
    return res.status(400).json({ ok: false, message: "Üye bilgileri eksik." });
  }
  if (!roleByName(db, user.role)) return res.status(400).json({ ok: false, message: "Geçersiz rol." });
  if (db.users.some(u => u.username === user.username)) {
    return res.status(400).json({ ok: false, message: "Bu kullanıcı adı zaten var." });
  }

  user.operations = user.role === "Admin" ? [] : normalizeOperations(user.operations);
  db.users.push(user);
  writeDb(db);
  const filteredDb = visibleDbForUser(JSON.parse(JSON.stringify(db)), requester);
  res.json({ ok: true, db: filteredDb });
});

app.put("/api/users/:username", (req, res) => {
  const db = readDb();
  const requester = userByUsername(db, req.body.requester);
  if (!canManageUsers(db, requester)) return res.status(403).json({ ok: false, message: "Sadece admin üye yönetebilir." });

  const target = userByUsername(db, req.params.username);
  if (!target) return res.status(404).json({ ok: false, message: "Kullanıcı bulunamadı." });

  const patch = req.body.user || {};
  if (patch.role && !roleByName(db, patch.role)) return res.status(400).json({ ok: false, message: "Geçersiz rol." });
  if (target.username === "admin" && patch.role && patch.role !== "Admin") {
    return res.status(400).json({ ok: false, message: "Ana admin hesabının rolü değiştirilemez." });
  }

  ["name", "username", "password", "mail", "role"].forEach(k => {
    if (patch[k] !== undefined && patch[k] !== "") target[k] = String(patch[k]).trim();
  });
  if (patch.operations !== undefined) {
    target.operations = target.role === "Admin" ? [] : normalizeOperations(patch.operations);
  }

  writeDb(db);
  const filteredDb = visibleDbForUser(JSON.parse(JSON.stringify(db)), requester);
  res.json({ ok: true, db: filteredDb });
});

app.delete("/api/users/:username", (req, res) => {
  const db = readDb();
  const requester = userByUsername(db, req.query.requester);
  if (!canManageUsers(db, requester)) return res.status(403).json({ ok: false, message: "Sadece admin üye yönetebilir." });
  if (req.params.username === "admin") return res.status(400).json({ ok: false, message: "Ana admin silinemez." });

  db.users = db.users.filter(u => u.username !== req.params.username);
  writeDb(db);
  const filteredDb = visibleDbForUser(JSON.parse(JSON.stringify(db)), requester);
  res.json({ ok: true, db: filteredDb });
});

app.post("/api/mail-rules", (req, res) => {
  const db = readDb();
  const requester = userByUsername(db, req.body.requester);
  if (!hasPermission(db, requester, "settings.manage") && requester?.role !== "Admin") {
    return res.status(403).json({ ok: false, message: "Yetkiniz yok." });
  }
  db.mailRules = Array.isArray(req.body.mailRules) ? req.body.mailRules : [];
  writeDb(db);
  const filteredDb = visibleDbForUser(JSON.parse(JSON.stringify(db)), requester);
  res.json({ ok: true, db: filteredDb });
});

app.post("/api/logs", upload.any(), async (req, res) => {
  try {
    const db = readDb();
    const requester = req.body.requester;
    const user = (db.users || []).find(u => u.username === requester);
    if (!user) return res.status(401).json({ ok: false, message: "Giriş gerekli." });

    const plate = req.body.plate || "";
    const maintenanceType = req.body.maintenanceType || "Diğer";
    const text = req.body.text || "";
    const km = Number(req.body.km || 0);
    const cost = Number(req.body.cost || 0);

    if (!plate) return res.status(400).json({ ok: false, message: "Plaka zorunludur." });
    if (!text) return res.status(400).json({ ok: false, message: "Açıklama zorunludur." });
    if (!km || km <= 0) return res.status(400).json({ ok: false, message: "KM zorunludur." });
    if (!cost || cost <= 0) return res.status(400).json({ ok: false, message: "Tutar zorunludur." });

    const uploadedFiles = req.files || [];
    if (uploadedFiles.length === 0) return res.status(400).json({ ok: false, message: "Fatura dosyası zorunludur." });

    const invoiceFiles = uploadedFiles.map(f => ({
      fileName: f.originalname,
      fileUrl: "/uploads/" + f.filename,
      path: f.path,
      mimeType: f.mimetype
    }));

    const log = {
      id: Date.now(),
      plate,
      type: maintenanceType,
      text,
      km,
      cost,
      date: today(),
      user: user.name || user.username,
      username: user.username,
      invoiceName: invoiceFiles[0]?.fileName || "",
      invoiceUrl: invoiceFiles[0]?.fileUrl || "",
      invoiceFiles
    };

    db.vehicleLogs.push(log);

    const invoice = {
      id: Date.now() + 1,
      plate,
      date: today(),
      amount: cost,
      km: km,
      status: "Onay Bekliyor",
      source: "Bakım / İşlem Kaydı",
      description: text,
      fileName: invoiceFiles[0]?.fileName || "",
      fileUrl: invoiceFiles[0]?.fileUrl || "",
      files: invoiceFiles,
      addedBy: user.name || user.username,
      addedByUsername: user.username
    };

    db.invoices.push(invoice);

    const vehicle = (db.vehicles || []).find(v => v.plate === plate);
    const operation = vehicle ? vehicle.operation : "";

    if (vehicle && km >= vehicle.km) {
      vehicle.km = km;
      vehicle.lastAction = text;
      vehicle.lastActionKm = km;
      if (maintenanceType === "Bakım" || text.toLowerCase().includes("bakım")) {
        vehicle.lastMaintenanceKm = km;
      }
    }

    writeDb(db);

    const mailAttachments = invoiceFiles.map(f => ({ filename: f.fileName, path: f.path }));
    await sendMailForVehicle(
      db,
      plate,
      operation,
      `Yeni Fatura / Bakım Kaydı - ${plate}`,
      `${plate} (${operation || "Genel"}) plakalı araç için yeni kayıt eklendi.\nİşlem: ${text}\nKM: ${km}\nTutar: ${cost} ₺`,
      mailAttachments
    );

    const filteredDb = visibleDbForUser(JSON.parse(JSON.stringify(db)), user);
    return res.json({ ok: true, db: filteredDb, log, invoice });
  } catch (err) {
    console.error("LOG_ERROR:", err.message);
    return res.status(500).json({ ok: false, message: err.message });
  }
});

app.delete("/api/invoices/:id", (req, res) => {
  try {
    const db = readDb();
    const requester = userByUsername(db, req.query.requester);
    if (!hasPermission(db, requester, "invoices.delete") && requester?.role !== "Admin") {
      return res.status(403).json({ ok: false, message: "Fatura silme yetkiniz yok." });
    }

    const id = Number(req.params.id);
    db.invoices = (db.invoices || []).filter(i => i.id !== id);
    writeDb(db);
    const filteredDb = visibleDbForUser(JSON.parse(JSON.stringify(db)), requester);
    res.json({ ok: true, db: filteredDb });
  } catch (err) {
    res.status(500).json({ ok: false, message: err.message });
  }
});

app.use((err, req, res, next) => {
  console.error("API_ERROR:", err.message);
  if (res.headersSent) return next(err);
  res.status(500).json({ ok: false, message: err.message || "Sunucu hatası oluştu." });
});

app.listen(PORT, () => console.log(`Sistem çalışıyor: http://localhost:${PORT}`));