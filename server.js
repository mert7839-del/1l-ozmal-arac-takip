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

const defaultDb = {
  users: [
    { name: "Admin", username: "admin", password: "1234", role: "Admin", mail: "admin@firma.com" },
    { name: "Yetkili", username: "yetkili", password: "1234", role: "Yetkili", mail: "yetkili@firma.com" },
    { name: "Yetkili 2", username: "yetkili2", password: "1234", role: "Yetkili2", mail: "yetkili2@firma.com" },
    { name: "Kullanıcı", username: "kullanici", password: "1234", role: "Kullanıcı", mail: "kullanici@firma.com" }
  ],
  vehicles: [
    { plate: "34 ABC 123", brand: "Mercedes Actros", km: 842350, inspection: "2026-06-18", lastAction: "Yağ ve filtre bakımı yapıldı", lastActionKm: 841900, driverName: "Ahmet Yılmaz", driverPhone: "0555 123 45 67", salary: 65000, mealMoney: 8000, premium: 12000, srcDate: "2024-01-15", maintenanceInterval: 20000, lastMaintenanceKm: 824350 },
    { plate: "41 LGS 782", brand: "Ford Trucks F-Max", km: 512120, inspection: "2026-05-14", lastAction: "Ön balata değişimi yapıldı", lastActionKm: 511800 },
    { plate: "16 TRK 456", brand: "Volvo FH", km: 963780, inspection: "2026-07-30", lastAction: "Lastik değişimi ve balans", lastActionKm: 961200 }
  ],
  vehicleLogs: [
    { plate: "34 ABC 123", date: "02.05.2026", km: 841900, text: "Yağ ve filtre bakımı yapıldı", cost: 12500, addedBy: "Admin" }
  ],
  invoices: [],
  monthlyKm: [],
  vehicleDocuments: [],
  suppliers: [],
  notifyMails: []
};

function readDb() {
  let db;

  try {
    if (!fs.existsSync(DB_FILE)) {
      fs.writeFileSync(DB_FILE, JSON.stringify(defaultDb, null, 2), "utf-8");
      db = JSON.parse(JSON.stringify(defaultDb));
    } else {
      const raw = fs.readFileSync(DB_FILE, "utf-8").trim();

      if (!raw) {
        db = JSON.parse(JSON.stringify(defaultDb));
      } else {
        db = JSON.parse(raw);
      }
    }
  } catch (err) {
    console.error("db.json okunamadı, varsayılan veri yükleniyor:", err.message);
    db = JSON.parse(JSON.stringify(defaultDb));
  }

  db.users ||= [];
  db.vehicles ||= [];
  db.vehicleLogs ||= [];
  db.invoices ||= [];
  db.monthlyKm ||= [];
  db.vehicleDocuments ||= [];
  db.suppliers ||= [];
  db.notifyMails ||= [];

  // Eski sistemden gelen Personel rolünü Kullanıcı yap
  db.users = db.users.map(u => {
    if (u.role === "Personel") return { ...u, role: "Kullanıcı" };
    return u;
  });

  // Giriş bozulmasın diye ana kullanıcılar yoksa otomatik ekle
  const requiredUsers = [
    { name: "Admin", username: "admin", password: "1234", role: "Admin", mail: "admin@firma.com" },
    { name: "Yetkili", username: "yetkili", password: "1234", role: "Yetkili", mail: "yetkili@firma.com" },
    { name: "Yetkili 2", username: "yetkili2", password: "1234", role: "Yetkili2", mail: "yetkili2@firma.com" },
    { name: "Kullanıcı", username: "kullanici", password: "1234", role: "Kullanıcı", mail: "kullanici@firma.com" }
  ];

  requiredUsers.forEach(u => {
    if (!db.users.some(x => x.username === u.username)) {
      db.users.push(u);
    }
  });

  // Araç listesi tamamen boşsa demo araçları geri yükle
  if (db.vehicles.length === 0 && defaultDb.vehicles && defaultDb.vehicles.length > 0) {
    db.vehicles = JSON.parse(JSON.stringify(defaultDb.vehicles));
  }

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

  db.users ||= [];
  db.vehicles ||= [];
  db.vehicleLogs ||= [];
  db.invoices ||= [];
  db.monthlyKm ||= [];
  db.vehicleDocuments ||= [];
  db.suppliers ||= [];
  db.notifyMails ||= [];

  const requiredUsers = [
    { name: "Admin", username: "admin", password: "1234", role: "Admin", mail: "admin@firma.com" },
    { name: "Yetkili", username: "yetkili", password: "1234", role: "Yetkili", mail: "yetkili@firma.com" },
    { name: "Yetkili 2", username: "yetkili2", password: "1234", role: "Yetkili2", mail: "yetkili2@firma.com" },
    { name: "Kullanıcı", username: "kullanici", password: "1234", role: "Kullanıcı", mail: "kullanici@firma.com" }
  ];

  requiredUsers.forEach(u => {
    if (!db.users.some(x => x.username === u.username)) {
      db.users.push(u);
      changed = true;
    }
  });

  if (!Array.isArray(db.vehicles) || db.vehicles.length === 0) {
    db.vehicles = JSON.parse(JSON.stringify(defaultDb.vehicles || []));
    changed = true;
  }

  if (changed) writeDb(db);
  return db;
}


function today() {
  return new Date().toLocaleDateString("tr-TR");
}

function oneYearFromTodayISO() {
  const d = new Date();
  d.setFullYear(d.getFullYear() + 1);
  return d.toISOString().slice(0, 10);
}

function userByUsername(db, username) {
  return db.users.find(u => u.username === username);
}

function hasRole(user, roles) {
  return !!user && roles.includes(user.role);
}

function createTransporter() {
  return nodemailer.createTransport({
    host: process.env.SMTP_HOST,
    port: Number(process.env.SMTP_PORT || 587),
    secure: String(process.env.SMTP_SECURE).toLowerCase() === "true",
    auth: { user: process.env.SMTP_USER, pass: process.env.SMTP_PASS },

    // Render free planda SMTP uzun beklerse tüm site HTML hata döndürmesin.
    connectionTimeout: 8000,
    greetingTimeout: 8000,
    socketTimeout: 8000
  });
}


function adminEmails(db) {
  return (db.users || [])
    .filter(u => u.role === "Admin" && u.mail)
    .map(u => u.mail);
}

function roleEmails(db, role) {
  if (role === "Admin") return adminEmails(db);
  return db.users.filter(u => u.role === role && u.mail).map(u => u.mail);
}

async function sendMailTo(db, role, subject, text, attachment) {
  try {
    // Render/Gmail sorunlarında sistemi kilitlememek için mail opsiyoneldir.
    if (String(process.env.MAIL_ENABLED || "true").toLowerCase() === "false") {
      return { sent: false, reason: "MAIL_ENABLED=false olduğu için mail gönderimi kapalı." };
    }

    const recipients = role === "Admin" ? adminEmails(db) : roleEmails(db, role);

    if (recipients.length === 0) {
      return { sent: false, reason: `${role} mail adresi yok.` };
    }

    if (!process.env.SMTP_HOST || !process.env.SMTP_USER || !process.env.SMTP_PASS) {
      return { sent: false, reason: ".env SMTP ayarları eksik." };
    }

    const mail = {
      from: `"1L Özmal Araç Takip" <${process.env.SMTP_USER}>`,
      to: recipients,
      subject,
      text
    };

    if (attachment) mail.attachments = Array.isArray(attachment) ? attachment : [attachment];

    await createTransporter().sendMail(mail);
    return { sent: true };
  } catch (err) {
    console.error("MAIL_SEND_ERROR:", err && err.message ? err.message : err);
    return { sent: false, reason: err && err.message ? err.message : "Mail gönderilemedi." };
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


function safeFileName(originalName) {
  const ext = path.extname(originalName || "").toLowerCase() || "";
  const base = path.basename(originalName || "dosya", ext)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/ı/g, "i")
    .replace(/İ/g, "I")
    .replace(/ğ/g, "g")
    .replace(/Ğ/g, "G")
    .replace(/ü/g, "u")
    .replace(/Ü/g, "U")
    .replace(/ş/g, "s")
    .replace(/Ş/g, "S")
    .replace(/ö/g, "o")
    .replace(/Ö/g, "O")
    .replace(/ç/g, "c")
    .replace(/Ç/g, "C")
    .replace(/[^a-zA-Z0-9_-]+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "")
    .slice(0, 80) || "dosya";

  return `${Date.now()}_${Math.round(Math.random() * 1E9)}_${base}${ext}`;
}




function isAllowedUploadByName(file) {
  if (!file) return false;
  const name = String(file.originalname || file.filename || "").toLowerCase();
  return /\.(pdf|jpg|jpeg|png|webp|gif)$/.test(name);
}

function isAllowedUploadFile(file) {
  if (!file) return false;

  const allowedMimes = [
    "application/pdf",
    "application/x-pdf",
    "application/acrobat",
    "applications/vnd.pdf",
    "text/pdf",
    "text/x-pdf",
    "application/octet-stream",
    "image/jpeg",
    "image/png",
    "image/webp",
    "image/gif"
  ];

  const name = String(file.originalname || "").toLowerCase();
  const extOk = /\.(pdf|jpg|jpeg|png|webp|gif)$/.test(name);

  return allowedMimes.includes(file.mimetype) || extOk;
}


const upload = multer({
  storage,
  fileFilter: (req, file, cb) => {
    return cb(null, true);
    cb(null, true);
  }
});

const vehicleDocUpload = multer({
  storage,
  fileFilter: (req, file, cb) => {
    const allowed = [
      "application/pdf",
      "image/jpeg",
      "image/png",
      "image/webp",
      "image/gif"
    ];

    if (!allowed.includes(file.mimetype)) {
      return cb(new Error("Araç evrakı için PDF veya fotoğraf yükleyebilirsin."));
    }

    cb(null, true);
  }
});


app.post("/api/reset-demo-data", (req, res) => {
  const db = JSON.parse(JSON.stringify(defaultDb));
  writeDb(db);
  res.json({ ok: true, db, message: "Demo veriler sıfırlandı." });
});

app.get("/api/data", (req, res) => res.json(recoverDbIfEmpty(readDb())));

app.post("/api/login", (req, res) => {
  const db = recoverDbIfEmpty(readDb());
  const { username, password } = req.body;
  const user = db.users.find(u => u.username === username && u.password === password);
  if (!user) return res.status(401).json({ ok: false, message: "Kullanıcı adı veya şifre hatalı." });
  res.json({ ok: true, user: { name: user.name, username: user.username, role: user.role, mail: user.mail } });
});

app.post("/api/users", (req, res) => {
  const db = readDb();
  const requester = userByUsername(db, req.body.requester);
  if (!hasRole(requester, ["Admin"])) return res.status(403).json({ ok: false, message: "Sadece admin üye ekleyebilir." });

  const user = req.body.user || {};
  if (!user.name || !user.username || !user.password || !user.role) return res.status(400).json({ ok: false, message: "Üye bilgileri eksik." });
  if (!["Admin", "Yetkili", "Yetkili2", "Kullanıcı"].includes(user.role)) return res.status(400).json({ ok: false, message: "Geçersiz rol." });
  if (db.users.some(u => u.username === user.username)) return res.status(400).json({ ok: false, message: "Bu kullanıcı adı zaten var." });

  db.users.push(user);
  writeDb(db);
  res.json({ ok: true, db });
});

app.delete("/api/users/:username", (req, res) => {
  const db = readDb();
  const requester = userByUsername(db, req.query.requester);
  if (!hasRole(requester, ["Admin"])) return res.status(403).json({ ok: false, message: "Sadece admin üye silebilir." });
  if (req.params.username === "admin") return res.status(400).json({ ok: false, message: "Ana admin silinemez." });
  if (req.params.username === req.query.requester) return res.status(400).json({ ok: false, message: "Kendi hesabını silemezsin." });

  db.users = db.users.filter(u => u.username !== req.params.username);
  writeDb(db);
  res.json({ ok: true, db });
});

app.post("/api/vehicles", (req, res) => {
  const db = readDb();
  const requester = userByUsername(db, req.body.requester);
  if (!hasRole(requester, ["Admin", "Yetkili2"])) return res.status(403).json({ ok: false, message: "Araç ekleme yetkisi admin ve yetkili2 kullanıcılarında." });

  const vehicle = req.body.vehicle;
  if (!vehicle || !vehicle.plate) return res.status(400).json({ ok: false, message: "Araç bilgileri eksik." });
  vehicle.plate = vehicle.plate.toUpperCase();

  if (db.vehicles.some(v => v.plate === vehicle.plate)) return res.status(400).json({ ok: false, message: "Bu plaka kayıtlı." });
  db.vehicles.push(vehicle);
  db.vehicleLogs.unshift({ plate: vehicle.plate, date: today(), km: vehicle.km, text: "Araç sisteme eklendi", cost: 0, addedBy: requester.name });
  writeDb(db);
  res.json({ ok: true, db });
});

app.put("/api/vehicles/:plate", (req, res) => {
  const db = readDb();
  const requester = userByUsername(db, req.body.requester);
  if (!hasRole(requester, ["Admin", "Yetkili2"])) return res.status(403).json({ ok: false, message: "Araç düzenleme yetkisi admin ve yetkili2 kullanıcılarında." });

  const oldPlate = req.params.plate;
  const vehicle = db.vehicles.find(v => v.plate === oldPlate);
  if (!vehicle) return res.status(404).json({ ok: false, message: "Araç bulunamadı." });

  const newPlate = (req.body.plate || oldPlate).toUpperCase();
  if (newPlate !== oldPlate && db.vehicles.some(v => v.plate === newPlate)) return res.status(400).json({ ok: false, message: "Yeni plaka zaten kayıtlı." });

  vehicle.plate = newPlate;
  vehicle.brand = req.body.brand || vehicle.brand;
  vehicle.modelYear = Number(req.body.modelYear || vehicle.modelYear || 0);
  vehicle.operation = req.body.operation || "";
  vehicle.km = Number(req.body.km || vehicle.km);
  vehicle.inspection = req.body.inspection || vehicle.inspection;
  vehicle.driverName = req.body.driverName || "";
  vehicle.driverPhone = req.body.driverPhone || "";
  vehicle.salary = Number(req.body.salary || 0);
  vehicle.mealMoney = Number(req.body.mealMoney || 0);
  vehicle.premium = Number(req.body.premium || 0);
  vehicle.srcDate = req.body.srcDate || "";
  vehicle.maintenanceInterval = Number(req.body.maintenanceInterval || 0);
  vehicle.lastMaintenanceKm = Number(req.body.lastMaintenanceKm || 0);

  if (newPlate !== oldPlate) {
    db.vehicleLogs.forEach(l => { if (l.plate === oldPlate) l.plate = newPlate; });
    db.invoices.forEach(i => { if (i.plate === oldPlate) i.plate = newPlate; });
    db.monthlyKm.forEach(m => { if (m.plate === oldPlate) m.plate = newPlate; });
  }

  db.vehicleLogs.unshift({ plate: newPlate, date: today(), km: vehicle.km, text: "Araç bilgileri düzenlendi", cost: 0, addedBy: requester.name });
  writeDb(db);
  res.json({ ok: true, db });
});

app.delete("/api/vehicles/:plate", (req, res) => {
  const db = readDb();
  const requester = userByUsername(db, req.query.requester);
  if (!hasRole(requester, ["Admin", "Yetkili2"])) return res.status(403).json({ ok: false, message: "Araç silme yetkisi admin ve yetkili2 kullanıcılarında." });

  const plate = req.params.plate;
  db.vehicles = db.vehicles.filter(v => v.plate !== plate);
  db.vehicleLogs = db.vehicleLogs.filter(l => l.plate !== plate);
  db.monthlyKm = db.monthlyKm.filter(m => m.plate !== plate);

  const removedInvoices = db.invoices.filter(i => i.plate === plate);
  removedInvoices.forEach(inv => {
    const p = path.join(UPLOAD_DIR, inv.savedName || "");
    if (fs.existsSync(p)) safeUnlinkUpload((typeof doc!=="undefined" && doc ? (doc.fileUrl || doc.file || doc.url) : (typeof d!=="undefined" && d ? (d.fileUrl || d.file || d.url) : "")));
  });
  db.invoices = db.invoices.filter(i => i.plate !== plate);

  const removedDocs = db.vehicleDocuments.filter(d => d.plate === plate);
  removedDocs.forEach(doc => {
    const p = path.join(UPLOAD_DIR, doc.savedName || "");
    if (fs.existsSync(p)) safeUnlinkUpload((typeof doc!=="undefined" && doc ? (doc.fileUrl || doc.file || doc.url) : (typeof d!=="undefined" && d ? (d.fileUrl || d.file || d.url) : "")));
  });
  db.vehicleDocuments = db.vehicleDocuments.filter(d => d.plate !== plate);

  writeDb(db);
  res.json({ ok: true, db });
});

app.post("/api/vehicles/:plate/inspection-done", (req, res) => {
  const db = readDb();
  const requester = userByUsername(db, req.body.requester);
  if (!hasRole(requester, ["Admin", "Yetkili2"])) return res.status(403).json({ ok: false, message: "Muayene yapma yetkisi admin ve yetkili2 kullanıcılarında." });

  const plate = req.params.plate;
  const vehicle = db.vehicles.find(v => v.plate === plate);
  if (!vehicle) return res.status(404).json({ ok: false, message: "Araç bulunamadı." });

  vehicle.inspection = oneYearFromTodayISO();
  db.vehicleLogs.unshift({ plate, date: today(), km: vehicle.km, text: "Muayene yapıldı. Yeni muayene tarihi: " + vehicle.inspection, cost: 0, addedBy: requester.name });
  writeDb(db);
  res.json({ ok: true, db, message: "Muayene tarihi 1 yıl ileri alındı." });
});

app.post("/api/logs", upload.array("files", 20), async (req, res) => {
  try {
    const db = readDb();

    const requester = req.body.requester;
    const user = (db.users || []).find(u => u.username === requester);
    if (!user) {
      return res.status(401).json({ ok: false, message: "Giriş gerekli." });
    }

    const plate = req.body.plate || "";
    const maintenanceType = req.body.maintenanceType || "Diğer";
    const text = req.body.text || "";
    const km = req.body.km;
    const cost = req.body.cost;

    const kmNumber = Number(km || 0);
    const costNumber = Number(cost || 0);

    if (!plate) {
      return res.status(400).json({ ok: false, message: "Plaka zorunludur." });
    }

    if (!text) {
      return res.status(400).json({ ok: false, message: "Ne yapıldı kısmı zorunludur." });
    }

    if (!kmNumber || kmNumber <= 0) {
      return res.status(400).json({ ok: false, message: "KM zorunludur." });
    }

    if (!costNumber || costNumber <= 0) {
      return res.status(400).json({ ok: false, message: "Tutar zorunludur." });
    }

    const uploadedFiles = req.files || [];
    if (uploadedFiles.length === 0) {
      return res.status(400).json({ ok: false, message: "Fatura dosyası zorunludur." });
    }

    const invoiceFiles = uploadedFiles.map(f => ({
      fileName: f.originalname,
      fileUrl: "/uploads/" + f.filename,
      path: f.path,
      mimeType: f.mimetype
    }));

    db.vehicleLogs = db.vehicleLogs || [];
    db.invoices = db.invoices || [];

    const log = {
      id: Date.now(),
      plate,
      type: maintenanceType,
      text,
      km: kmNumber,
      cost: costNumber,
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
      amount: costNumber,
      km: kmNumber,
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
    if (vehicle) {
      vehicle.km = kmNumber;
      vehicle.lastAction = text;
      vehicle.lastActionKm = kmNumber;

      if (maintenanceType === "Bakım" || text.toLowerCase().includes("bakım")) {
        vehicle.lastMaintenanceKm = kmNumber;
      }
    }

    writeDb(db);

    const mailAttachments = invoiceFiles.map(f => ({
      filename: f.fileName,
      path: f.path
    }));

    try {
      await sendMailTo(
        db,
        "Admin",
        "Yeni Fatura / Bakım Kaydı Eklendi",
        `${plate} plakalı araç için yeni fatura/bakım kaydı eklendi.\nİşlem: ${text}\nKM: ${kmNumber}\nTutar: ${costNumber}`,
        mailAttachments
      );
    } catch (mailErr) {
      console.error("MAIL_LOG_SEND_ERROR:", mailErr && mailErr.message ? mailErr.message : mailErr);
    }

    return res.json({ ok: true, db, log, invoice });
  } catch (err) {
    console.error("LOG_UPLOAD_ERROR:", err && err.stack ? err.stack : err);
    return res.status(500).json({ ok: false, message: err && err.message ? err.message : "Kayıt oluşturulamadı." });
  }
});


app.post("/api/monthly-km", (req, res) => {
  const db = readDb();
  const requester = userByUsername(db, req.body.requester);
  if (!hasRole(requester, ["Admin", "Yetkili", "Yetkili2", "Kullanıcı"])) return res.status(403).json({ ok: false, message: "Giriş gerekli." });

  const record = req.body.record || {};
  const vehicle = db.vehicles.find(x => x.plate === record.plate);
  if (!vehicle) return res.status(404).json({ ok: false, message: "Araç bulunamadı." });

  const idx = db.monthlyKm.findIndex(m => m.plate === record.plate && m.month === record.month);
  const fullRecord = {
    ...record,
    brand: vehicle.brand,
    doneKm: Number(record.endKm) - Number(record.startKm),
    addedBy: requester.name,
    createdAt: today()
  };

  if (fullRecord.doneKm < 0) return res.status(400).json({ ok: false, message: "Ay sonu KM, ay başı KM'den düşük olamaz." });

  if (idx >= 0) db.monthlyKm[idx] = fullRecord;
  else db.monthlyKm.unshift(fullRecord);

  vehicle.km = Math.max(Number(vehicle.km), Number(record.endKm));
  db.vehicleLogs.unshift({
    plate: record.plate,
    date: today(),
    km: Number(record.endKm),
    text: `${record.month} aylık KM kaydı. Yapılan KM: ${fullRecord.doneKm}`,
    cost: 0,
    addedBy: requester.name
  });

  writeDb(db);
  res.json({ ok: true, db });
});

app.delete("/api/monthly-km", (req, res) => {
  const dbCheck = readDb();
  const requesterUser = (dbCheck.users || []).find(u => u.username === req.query.requester);
  if (!requesterUser || requesterUser.role !== "Admin") {
    return res.status(403).json({ ok: false, message: "Aylık KM silme işlemi için admin yetkisi gerekir." });
  }

  const db = readDb();
  const requester = userByUsername(db, req.query.requester);
  if (!hasRole(requester, ["Admin", "Yetkili2"])) return res.status(403).json({ ok: false, message: "Sadece admin ve yetkili2 silebilir." });
  db.monthlyKm = db.monthlyKm.filter(m => !(m.plate === req.query.plate && m.month === req.query.month));
  writeDb(db);
  res.json({ ok: true, db });
});

app.post("/api/invoices", upload.array("files", 10), async (req, res) => {
  try {
    const db = readDb();
    const requester = userByUsername(db, req.body.requester);
    if (!hasRole(requester, ["Admin", "Yetkili", "Yetkili2", "Kullanıcı"])) return res.status(403).json({ ok: false, message: "Giriş gerekli." });

    const plate = (req.body.plate || "").toUpperCase();
    const vehicle = db.vehicles.find(v => v.plate === plate);
    if (!vehicle) return res.status(404).json({ ok: false, message: "Araç bulunamadı." });
    if (!req.file) return res.status(400).json({ ok: false, message: "PDF gerekli." });

    let invoice = {
      id: Date.now(),
      plate,
      fileName: req.file.originalname,
      savedName: req.file.filename,
      fileUrl: "/uploads/" + req.file.filename,
      date: today(),
      addedBy: requester.name,
      addedByUsername: requester.username,
      source: "PDF Fatura",
      status: "YETKILI2_BEKLIYOR",
      yetkili2Approved: false,
      yetkili2ApprovedBy: "",
      yetkili2ApprovedDate: "",
      paid: false,
      paidBy: "",
      paidDate: "",
      mailToYetkili2Sent: false,
      mailToYetkiliSent: false,
      mailError: ""
    };

    try {
      const mailResult = await sendMailTo(
        db,
        "Yetkili2",
        `Yeni Fatura Onayı Bekliyor - ${plate}`,
        `Yeni fatura eklendi.\n\nPlaka: ${plate}\nPDF: ${req.file.originalname}\nEkleyen: ${requester.name}\n\nYetkili2 onayı bekliyor.`,
        { filename: req.file.originalname, path: req.file.path }
      );
      invoice.mailToYetkili2Sent = !!mailResult.sent;
      if (!mailResult.sent) invoice.mailError = mailResult.reason || "";
    } catch (mailErr) {
      invoice.mailError = mailErr.message;
    }

    db.invoices.unshift(invoice);
    db.vehicleLogs.unshift({
      plate,
      date: today(),
      km: vehicle.km || 0,
      text: "PDF fatura yüklendi: " + req.file.originalname + " / Yetkili2 onayı bekliyor",
      cost: 0,
      addedBy: requester.name
    });
    writeDb(db);

    res.json({
      ok: true,
      db,
      message: invoice.mailToYetkili2Sent
        ? "Fatura eklendi ve Yetkili2 kullanıcısına mail gönderildi."
        : "Fatura eklendi fakat Yetkili2 maili gönderilemedi: " + invoice.mailError
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, message: "Hata: " + err.message });
  }
});

app.post("/api/invoices/:id/yetkili2-approve", async (req, res) => {
  try {
    const db = readDb();
    const requester = userByUsername(db, req.body.requester);
    if (!hasRole(requester, ["Admin", "Yetkili2"])) return res.status(403).json({ ok: false, message: "Faturayı Yetkili2 veya Admin onaylayabilir." });

    const inv = db.invoices.find(i => i.id === Number(req.params.id));
    if (!inv) return res.status(404).json({ ok: false, message: "Fatura bulunamadı." });

    inv.yetkili2Approved = true;
    inv.yetkili2ApprovedBy = requester.name;
    inv.yetkili2ApprovedDate = today();
    inv.status = "YETKILI_BEKLIYOR";

    const filePath = path.join(UPLOAD_DIR, inv.savedName || "");
    try {
      const mailResult = await sendMailTo(
        db,
        "Yetkili",
        `Yetkili Onayı Bekleyen Fatura - ${inv.plate}`,
        `Yetkili2 faturayı onayladı.\n\nPlaka: ${inv.plate}\nPDF: ${inv.fileName}\nYetkili2 Onaylayan: ${requester.name}\n\nÖdeme yapılması için kontrol edebilirsiniz.`,
        fs.existsSync(filePath) ? { filename: inv.fileName, path: filePath } : null
      );
      inv.mailToYetkiliSent = !!mailResult.sent;
      if (!mailResult.sent) inv.mailError = mailResult.reason || inv.mailError || "";
    } catch (mailErr) {
      inv.mailError = mailErr.message;
    }

    db.vehicleLogs.unshift({
      plate: inv.plate,
      date: today(),
      km: db.vehicles.find(v => v.plate === inv.plate)?.km || 0,
      text: "Fatura Yetkili2 tarafından onaylandı: " + inv.fileName,
      cost: 0,
      addedBy: requester.name
    });

    writeDb(db);
    res.json({
      ok: true,
      db,
      message: inv.mailToYetkiliSent
        ? "Fatura Yetkili2 tarafından onaylandı ve Yetkili kullanıcısına mail gönderildi."
        : "Fatura onaylandı fakat Yetkili maili gönderilemedi: " + inv.mailError
    });
  } catch (err) {
    res.status(500).json({ ok: false, message: err.message });
  }
});

app.post("/api/invoices/:id/paid", (req, res) => {
  const db = readDb();
  const requester = userByUsername(db, req.body.requester);
  if (!hasRole(requester, ["Admin", "Yetkili"])) return res.status(403).json({ ok: false, message: "Ödeme işaretleme yetkisi Yetkili ve Admin kullanıcılarında." });

  const inv = db.invoices.find(i => i.id === Number(req.params.id));
  if (!inv) return res.status(404).json({ ok: false, message: "Fatura bulunamadı." });
  if (!inv.yetkili2Approved) return res.status(400).json({ ok: false, message: "Fatura önce Yetkili2 tarafından onaylanmalı." });

  inv.paid = true;
  inv.paidBy = requester.name;
  inv.paidDate = today();
  inv.status = "ODEME_YAPILDI";

  db.vehicleLogs.unshift({
    plate: inv.plate,
    date: today(),
    km: db.vehicles.find(v => v.plate === inv.plate)?.km || 0,
    text: "Fatura ödeme yapıldı olarak işaretlendi: " + inv.fileName,
    cost: 0,
    addedBy: requester.name
  });

  writeDb(db);
  res.json({ ok: true, db, message: "Fatura ödeme yapıldı olarak işaretlendi." });
});

app.delete("/api/invoices/:id", (req, res) => {
  const db = readDb();
  const requester = userByUsername(db, req.query.requester);
  if (!hasRole(requester, ["Admin", "Yetkili2"])) return res.status(403).json({ ok: false, message: "Fatura silme yetkisi admin ve yetkili2 kullanıcılarında." });

  const inv = db.invoices.find(i => i.id === Number(req.params.id));
  if (inv) {
    const p = path.join(UPLOAD_DIR, inv.savedName || "");
    if (fs.existsSync(p)) safeUnlinkUpload((typeof doc!=="undefined" && doc ? (doc.fileUrl || doc.file || doc.url) : (typeof d!=="undefined" && d ? (d.fileUrl || d.file || d.url) : "")));
  }
  db.invoices = db.invoices.filter(i => i.id !== Number(req.params.id));
  writeDb(db);
  res.json({ ok: true, db });
});

app.post("/api/settings/notify-mails", (req, res) => {
  const db = readDb();
  const requester = userByUsername(db, req.body.requester);
  if (!hasRole(requester, ["Admin"])) return res.status(403).json({ ok: false, message: "Sadece admin düzenleyebilir." });
  db.notifyMails = req.body.notifyMails || [];
  writeDb(db);
  res.json({ ok: true, db });
});


app.post("/api/suppliers", upload.array("files", 20), async (req, res) => {
  try {
    const db = readDb();
    db.suppliers ||= [];

    const requester = req.body.requester;
    const user = (db.users || []).find(u => u.username === requester);
    if (!user) return res.status(401).json({ ok:false, message:"Giriş gerekli." });

    const processType = req.body.processType || "Diğer";
    const supplierName = req.body.supplierName || "";
    const plate = req.body.plate || "";
    const amount = Number(req.body.amount || 0);
    const note = req.body.note || "";
    const processDate = req.body.processDate || today();

    if (!processType) return res.status(400).json({ ok:false, message:"İşlem türü zorunludur." });

    const uploadedFiles = req.files || [];
    if (uploadedFiles.length === 0) {
      return res.status(400).json({ ok:false, message:"Fatura veya fotoğraf dosyası zorunludur." });
    }

    for (const f of uploadedFiles) {
      if (typeof isAllowedUploadByName === "function" && typeof isAllowedUploadFile === "function") {
        if (!isAllowedUploadByName(f) && !isAllowedUploadFile(f)) {
          return res.status(400).json({ ok:false, message:"Sadece PDF veya fotoğraf yüklenebilir." });
        }
      } else if (typeof isAllowedUploadFile === "function") {
        if (!isAllowedUploadFile(f)) {
          return res.status(400).json({ ok:false, message:"Sadece PDF veya fotoğraf yüklenebilir." });
        }
      }
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
      mimeType: files[0]?.mimeType || "",
      addedBy: user.name || user.username,
      addedByUsername: user.username,
      createdAt: new Date().toISOString()
    };

    db.suppliers.unshift(record);
    writeDb(db);
    return res.json({ ok:true, db, record });
  } catch (err) {
    console.error("SUPPLIER_UPLOAD_ERROR:", err && err.stack ? err.stack : err);
    return res.status(500).json({ ok:false, message: err && err.message ? err.message : "Tedarikçi işlemi eklenemedi." });
  }
});

app.delete("/api/suppliers/:id", (req, res) => {
  const db = readDb();
  db.suppliers ||= [];
  const requester = userByUsername(db, req.query.requester);
  if (!hasRole(requester, ["Admin", "Yetkili2"])) {
    return res.status(403).json({ ok: false, message: "Tedarikçi kaydı silme yetkisi admin ve yetkili2 kullanıcılarında." });
  }

  const rec = db.suppliers.find(x => x.id === Number(req.params.id));
  if (rec) {
    const files = Array.isArray(rec.files) ? rec.files : [];
    files.forEach(f => safeUnlinkUpload(f.fileUrl || f.url || ""));
    if ((!files.length) && rec.fileUrl) safeUnlinkUpload(rec.fileUrl);
  }

  db.suppliers = db.suppliers.filter(x => x.id !== Number(req.params.id));
  writeDb(db);
  return res.json({ ok:true, db });
});


app.post("/api/vehicle-documents", upload.any(), async (req, res) => {
  try {
    const db = readDb();

    const requester = req.body.requester;
    const user = (db.users || []).find(u => u.username === requester);
    if (!user) return res.status(401).json({ ok:false, message:"Giriş gerekli." });

    const uploadedFile = (req.files && req.files[0]) ? req.files[0] : req.file;
    if (!uploadedFile) {
      return res.status(400).json({ ok:false, message:"Dosya seçilmedi." });
    }

    if (typeof isAllowedUploadByName === "function" && typeof isAllowedUploadFile === "function") {
      if (!isAllowedUploadByName(uploadedFile) && !isAllowedUploadFile(uploadedFile)) {
        return res.status(400).json({ ok:false, message:"Sadece PDF veya fotoğraf yüklenebilir." });
      }
    } else if (typeof isAllowedUploadFile === "function") {
      if (!isAllowedUploadFile(uploadedFile)) {
        return res.status(400).json({ ok:false, message:"Sadece PDF veya fotoğraf yüklenebilir." });
      }
    }

    console.log("V77_VEHICLE_DOCUMENT_UPLOAD:", uploadedFile.originalname, uploadedFile.mimetype);

    const doc = {
      id: Date.now(),
      plate: req.body.plate || "",
      docType: req.body.docType || "Diğer",
      docDate: req.body.docDate || today(),
      date: req.body.docDate || today(),
      amount: Number(req.body.amount || 0),
      note: req.body.note || "",
      fileName: uploadedFile.originalname,
      fileUrl: "/uploads/" + uploadedFile.filename,
      mimeType: uploadedFile.mimetype,
      addedBy: user.name || user.username,
      addedByUsername: user.username,
      createdAt: new Date().toISOString()
    };

    db.vehicleDocuments = db.vehicleDocuments || [];
    db.vehicleDocuments.push(doc);

    writeDb(db);
    return res.json({ ok:true, db, doc });
  } catch (err) {
    console.error("VEHICLE_DOCUMENT_UPLOAD_ERROR:", err && err.stack ? err.stack : err);
    return res.status(500).json({ ok:false, message: err && err.message ? err.message : "Evrak yüklenemedi." });
  }
});


app.delete("/api/vehicle-documents/:id", (req, res) => {
  const db = readDb();
  const requester = userByUsername(db, req.query.requester);

  if (!hasRole(requester, ["Admin", "Yetkili2"])) {
    return res.status(403).json({ ok: false, message: "Evrak silme yetkisi admin ve yetkili2 kullanıcılarında." });
  }

  const doc = db.vehicleDocuments.find(d => d.id === Number(req.params.id));

  if (doc) {
    const p = path.join(UPLOAD_DIR, doc.savedName || "");
    if (fs.existsSync(p)) safeUnlinkUpload((typeof doc!=="undefined" && doc ? (doc.fileUrl || doc.file || doc.url) : (typeof d!=="undefined" && d ? (d.fileUrl || d.file || d.url) : "")));
  }

  db.vehicleDocuments = db.vehicleDocuments.filter(d => d.id !== Number(req.params.id));
  writeDb(db);
  res.json({ ok: true, db });
});



app.get("/api/monthly-km/export-xlsx", async (req, res) => {
  try {
    const db = readDb();
    const workbook = new ExcelJS.Workbook();

    workbook.creator = "Lojistik Araç Takip Sistemi";
    workbook.created = new Date();

    const vehicles = db.vehicles || [];
    const monthly = db.monthlyKm || [];

    function cleanSheetName(name) {
      return String(name || "ARAC")
        .replace(/[\\\/\?\*\[\]\:]/g, "-")
        .substring(0, 31);
    }

    function styleSheet(ws) {
      ws.columns = [
        { header: "Ay", key: "month", width: 15 },
        { header: "Plaka", key: "plate", width: 16 },
        { header: "Araç", key: "brand", width: 28 },
        { header: "Ay Başı KM", key: "startKm", width: 16 },
        { header: "Ay Sonu KM", key: "endKm", width: 16 },
        { header: "Yapılan KM", key: "doneKm", width: 16 },
        { header: "Ekleyen", key: "addedBy", width: 20 },
        { header: "Kayıt Tarihi", key: "createdAt", width: 16 }
      ];

      ws.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
      ws.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F766E" } };
      ws.getRow(1).alignment = { vertical: "middle", horizontal: "center" };

      ws.eachRow((row) => {
        row.eachCell((cell) => {
          cell.border = {
            top: { style: "thin", color: { argb: "FFCBD5E1" } },
            left: { style: "thin", color: { argb: "FFCBD5E1" } },
            bottom: { style: "thin", color: { argb: "FFCBD5E1" } },
            right: { style: "thin", color: { argb: "FFCBD5E1" } }
          };
        });
      });

      ws.views = [{ state: "frozen", ySplit: 1 }];
      ws.autoFilter = "A1:H1";
    }

    const summary = workbook.addWorksheet("Özet");
    summary.columns = [
      { header: "Plaka", key: "plate", width: 16 },
      { header: "Araç", key: "brand", width: 28 },
      { header: "Toplam Yapılan KM", key: "totalKm", width: 20 },
      { header: "Kayıt Sayısı", key: "count", width: 14 }
    ];

    vehicles.forEach(vehicle => {
      const records = monthly.filter(m => m.plate === vehicle.plate);
      const totalKm = records.reduce((sum, r) => sum + Number(r.doneKm || 0), 0);
      summary.addRow({
        plate: vehicle.plate,
        brand: vehicle.brand,
        totalKm,
        count: records.length
      });
    });

    summary.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
    summary.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F766E" } };
    summary.views = [{ state: "frozen", ySplit: 1 }];
    summary.autoFilter = "A1:D1";

    vehicles.forEach(vehicle => {
      let sheetName = cleanSheetName(vehicle.plate);
      let baseName = sheetName;
      let counter = 2;

      while (workbook.getWorksheet(sheetName)) {
        sheetName = cleanSheetName(baseName.substring(0, 27) + "-" + counter);
        counter++;
      }

      const ws = workbook.addWorksheet(sheetName);
      styleSheet(ws);

      const records = monthly
        .filter(m => m.plate === vehicle.plate)
        .sort((a, b) => String(a.month).localeCompare(String(b.month)));

      if (records.length === 0) {
        ws.addRow({
          month: "-",
          plate: vehicle.plate,
          brand: vehicle.brand,
          startKm: "",
          endKm: "",
          doneKm: "",
          addedBy: "",
          createdAt: "Kayıt yok"
        });
      } else {
        records.forEach(r => {
          ws.addRow({
            month: r.month,
            plate: r.plate,
            brand: r.brand || vehicle.brand,
            startKm: Number(r.startKm || 0),
            endKm: Number(r.endKm || 0),
            doneKm: Number(r.doneKm || 0),
            addedBy: r.addedBy || "",
            createdAt: r.createdAt || ""
          });
        });
      }
    });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", 'attachment; filename="tum_araclar_aylik_km.xlsx"');

    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, message: "Excel oluşturulamadı: " + err.message });
  }
});



app.use((err, req, res, next) => {
  console.error("Genel API hatası:", err && err.message ? err.message : err);

  if (res.headersSent) {
    return next(err);
  }

  res.status(500).json({
    ok: false,
    message: err && err.message ? err.message : "Sunucu hatası oluştu."
  });
});



app.get("/api/penalties/export-xlsx", async (req, res) => {
  try {
    const db = readDb();
    const docs = Array.isArray(db.vehicleDocuments) ? db.vehicleDocuments : [];
    const penalties = docs.filter(d => d.docType === "Ceza Evrakı");

    const workbook = new ExcelJS.Workbook();
    workbook.creator = "1L Özmal Araç Takip Sistemi";
    workbook.created = new Date();

    const sheet = workbook.addWorksheet("Ceza Takibi");

    sheet.columns = [
      { header: "Tarih", key: "date", width: 16 },
      { header: "Plaka", key: "plate", width: 18 },
      { header: "Evrak Türü", key: "docType", width: 20 },
      { header: "Tutar", key: "amount", width: 15 },
      { header: "Not", key: "note", width: 35 },
      { header: "Dosya Adı", key: "fileName", width: 35 },
      { header: "Dosya Linki", key: "fileUrl", width: 45 },
      { header: "Ekleyen", key: "addedBy", width: 20 }
    ];

    penalties.forEach(p => {
      sheet.addRow({
        date: p.docDate || p.date || "",
        plate: p.plate || "",
        docType: p.docType || "Ceza Evrakı",
        amount: Number(p.amount || 0),
        note: p.note || "",
        fileName: p.fileName || "",
        fileUrl: p.fileUrl || "",
        addedBy: p.addedBy || ""
      });
    });

    sheet.getRow(1).font = { bold: true };
    sheet.getRow(1).alignment = { vertical: "middle", horizontal: "center" };

    sheet.eachRow((row) => {
      row.eachCell((cell) => {
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" }
        };
      });
    });

    const totalRow = sheet.addRow({
      date: "",
      plate: "",
      docType: "TOPLAM",
      amount: penalties.reduce((sum, p) => sum + Number(p.amount || 0), 0),
      note: "",
      fileName: "",
      fileUrl: "",
      addedBy: ""
    });
    totalRow.font = { bold: true };

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=ceza_takibi.xlsx");

    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error("Ceza Excel export hatası:", err.message);
    res.status(500).json({ ok: false, message: err.message });
  }
});



// GLOBAL_JSON_ERROR_HANDLER_V67
app.use((err, req, res, next) => {
  console.error("API_ERROR:", err && err.stack ? err.stack : err);
  if (res.headersSent) return next(err);
  res.status(500).json({ ok:false, message: err && err.message ? err.message : "Sunucu hatası oluştu." });
});



function safeUnlinkUpload(fileUrl) {
  try {
    if (!fileUrl || typeof fileUrl !== "string") return;

    // Sadece /uploads/dosyaadi formatındaki gerçek dosyaları sil.
    if (!fileUrl.startsWith("/uploads/")) return;

    const fileName = path.basename(fileUrl);
    if (!fileName || fileName === "." || fileName === "..") return;

    const fullPath = path.join(__dirname, "uploads", fileName);

    if (!fullPath.startsWith(path.join(__dirname, "uploads"))) return;
    if (!fs.existsSync(fullPath)) return;

    const st = fs.statSync(fullPath);
    if (!st.isFile()) return;

    safeUnlinkUpload((typeof doc!=="undefined" && doc ? (doc.fileUrl || doc.file || doc.url) : (typeof d!=="undefined" && d ? (d.fileUrl || d.file || d.url) : "")));
  } catch (err) {
    console.error("SAFE_UNLINK_UPLOAD_ERROR:", err.message);
  }
}



// V68_GLOBAL_JSON_ERROR
app.use((err, req, res, next) => {
  console.error("API_ERROR:", err && err.stack ? err.stack : err);
  if (res.headersSent) return next(err);
  res.status(500).json({ ok:false, message: err && err.message ? err.message : "Sunucu hatası oluştu." });
});



function canEditRecords(user) {
  return user && ["Admin","Yetkili","Yetkili2"].includes(user.role);
}
function findUserByUsername(db, username) {
  return (db.users || []).find(u => u.username === username);
}


app.put("/api/logs/:id", (req, res) => {
  try {
    const db = readDb();
    const user = findUserByUsername(db, req.body.requester);
    if (!canEditRecords(user)) return res.status(403).json({ ok:false, message:"Yetkin yok." });

    const id = String(req.params.id);
    const log = (db.vehicleLogs || []).find(x => String(x.id) === id);
    if (!log) return res.status(404).json({ ok:false, message:"İşlem bulunamadı." });

    const oldText = log.text;
    log.text = req.body.text || log.text;
    log.km = Number(req.body.km || log.km || 0);
    log.cost = Number(req.body.cost || log.cost || 0);
    log.updatedAt = new Date().toISOString();
    log.updatedBy = user.username;

    (db.invoices || []).forEach(inv => {
      if (inv.plate === log.plate && (inv.description === oldText || inv.description === log.text || inv.id === log.invoiceId)) {
        inv.amount = log.cost;
        inv.km = log.km;
        inv.description = log.text;
      }
    });

    writeDb(db);
    res.json({ ok:true, db, log });
  } catch (err) {
    res.status(500).json({ ok:false, message:err.message });
  }
});

app.delete("/api/logs/:id", (req, res) => {
  try {
    const db = readDb();
    const user = findUserByUsername(db, req.query.requester);
    if (!canEditRecords(user)) return res.status(403).json({ ok:false, message:"Yetkin yok." });

    const id = String(req.params.id);
    const before = (db.vehicleLogs || []).length;
    db.vehicleLogs = (db.vehicleLogs || []).filter(x => String(x.id) !== id);

    if (db.vehicleLogs.length === before) return res.status(404).json({ ok:false, message:"İşlem bulunamadı." });

    writeDb(db);
    res.json({ ok:true, db });
  } catch (err) {
    res.status(500).json({ ok:false, message:err.message });
  }
});

app.put("/api/invoices/:id/plate", (req, res) => {
  try {
    const db = readDb();
    const user = findUserByUsername(db, req.body.requester);
    if (!canEditRecords(user)) return res.status(403).json({ ok:false, message:"Yetkin yok." });

    const id = String(req.params.id);
    const inv = (db.invoices || []).find(x => String(x.id) === id);
    if (!inv) return res.status(404).json({ ok:false, message:"Fatura bulunamadı." });

    inv.plate = String(req.body.plate || "").trim().toUpperCase();
    inv.updatedAt = new Date().toISOString();
    inv.updatedBy = user.username;

    writeDb(db);
    res.json({ ok:true, db, invoice:inv });
  } catch (err) {
    res.status(500).json({ ok:false, message:err.message });
  }
});


app.listen(PORT, () => console.log(`Sistem çalışıyor: http://localhost:${PORT}`));
