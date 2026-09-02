"use strict";
require("dotenv").config();

const express = require("express");
const multer = require("multer");
const cors = require("cors");
const helmet = require("helmet");
const rateLimit = require("express-rate-limit");
const path = require("path");
const fs = require("fs");
const fsp = fs.promises;
const os = require("os");
const dns = require("dns").promises;
const net = require("net");
const { URL } = require("url");
const { execFile } = require("child_process");
const { pipeline } = require("stream/promises");
const { Readable, Transform } = require("stream");
const archiver = require("archiver");
const ExcelJS = require("exceljs");
const { Document, Packer, Paragraph, TextRun } = require("docx");
const { createWorker } = require("tesseract.js");
const { PDFDocument, degrees } = require("pdf-lib");
const sharp = require("sharp");
const QRCode = require("qrcode");
const { GoogleGenerativeAI } = require("@google/generative-ai");

const app = express();
const PORT = Number(process.env.PORT || 3000);
const HOST = "0.0.0.0";
const MAX_FILE_SIZE = Math.max(1, Number(process.env.MAX_FILE_SIZE_MB || 20)) * 1024 * 1024;
const MAX_REMOTE_SIZE = Math.max(1, Number(process.env.MAX_REMOTE_SIZE_MB || 100)) * 1024 * 1024;
const TEMP_ROOT = process.env.TEMP_DIR ? path.resolve(process.env.TEMP_DIR) : path.join(os.tmpdir(), "khmertools");
const PUBLIC_DIR = path.join(__dirname, "public");
const PDFTOPPM = process.env.PDFTOPPM_PATH || "pdftoppm";
const PDFTOPDF = process.env.PDFTOPDF_PATH || "pdftocairo";
const ALLOWED_ORIGIN = process.env.ALLOWED_ORIGIN || "";

fs.mkdirSync(TEMP_ROOT, { recursive: true });

app.disable("x-powered-by");
app.set("trust proxy", 1);
app.use(helmet({ contentSecurityPolicy: false, crossOriginResourcePolicy: { policy: "cross-origin" } }));
app.use(cors({ origin: ALLOWED_ORIGIN ? ALLOWED_ORIGIN.split(",").map(v => v.trim()) : true, methods: ["GET", "POST"], maxAge: 86400 }));
app.use(express.json({ limit: "128kb" }));
app.use(express.urlencoded({ extended: false, limit: "128kb" }));
app.use("/api", rateLimit({ windowMs: 15 * 60 * 1000, limit: 120, standardHeaders: "draft-8", legacyHeaders: false }));
app.use("/api/ai", rateLimit({ windowMs: 15 * 60 * 1000, limit: 30, standardHeaders: "draft-8", legacyHeaders: false }));
app.use(express.static(PUBLIC_DIR, { extensions: ["html"], maxAge: process.env.NODE_ENV === "production" ? "1h" : 0 }));

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: MAX_FILE_SIZE, files: 20, fields: 12 } });
const asyncRoute = fn => (req, res, next) => Promise.resolve(fn(req, res, next)).catch(next);
const execFileAsync = (command, args, options = {}) => new Promise((resolve, reject) => execFile(command, args, { windowsHide: true, maxBuffer: 10 * 1024 * 1024, ...options }, (error, stdout, stderr) => error ? reject(new Error((stderr || error.message).trim())) : resolve({ stdout, stderr })));

function safeName(value, fallback = "file") {
  const name = path.basename(String(value || fallback)).normalize("NFKC").replace(/[<>:"/\\|?*\x00-\x1f]/g, "_").replace(/\s+/g, " ").trim().slice(0, 140);
  return name || fallback;
}
function baseName(name) { return safeName(name).replace(/\.[^.]+$/, ""); }
function contentDisposition(filename) { return `attachment; filename="${safeName(filename).replace(/["\\]/g, "_")}"; filename*=UTF-8''${encodeURIComponent(safeName(filename))}`; }
function tempFolder(prefix) { return fs.mkdtempSync(path.join(TEMP_ROOT, `${prefix}-`)); }
async function cleanup(folder) { if (folder) await fsp.rm(folder, { recursive: true, force: true }).catch(() => {}); }
function sendBuffer(res, buffer, type, filename) { res.set({ "Content-Type": type, "Content-Disposition": contentDisposition(filename), "Content-Length": buffer.length }); res.send(buffer); }
function languageCode(value) { return ({ khmer: "khm", english: "eng", both: "khm+eng", khm: "khm", eng: "eng", "khm+eng": "khm+eng" })[String(value || "both").toLowerCase()] || "khm+eng"; }
function outputFormat(value) { return String(value || "png").toLowerCase() === "jpg" ? "jpg" : "png"; }
function splitText(text) { return String(text || "").split(/\r?\n/).map(v => v.trim()).filter(Boolean); }
function assertText(value, max = 50000) { const text = String(value || "").trim(); if (!text) throw Object.assign(new Error("Text is required."), { status: 400 }); if (text.length > max) throw Object.assign(new Error(`Text is limited to ${max.toLocaleString()} characters.`), { status: 413 }); return text; }

async function detectedType(file) {
  if (!file?.buffer?.length) return null;
  const { fileTypeFromBuffer } = await import("file-type");
  return fileTypeFromBuffer(file.buffer);
}
async function assertFile(file, allowed) {
  if (!file) throw Object.assign(new Error("Please choose a file."), { status: 400 });
  const detected = await detectedType(file);
  const mime = detected?.mime || file.mimetype;
  if (!allowed.includes(mime)) throw Object.assign(new Error(`Unsupported file type (${mime || "unknown"}).`), { status: 415 });
  return mime;
}

async function renderPdf(buffer, folder, format = "png", dpi = 160) {
  const pdfPath = path.join(folder, "input.pdf");
  const prefix = path.join(folder, "page");
  await fsp.writeFile(pdfPath, buffer);
  const flag = format === "jpg" ? "-jpeg" : "-png";
  await execFileAsync(PDFTOPPM, [flag, "-r", String(dpi), pdfPath, prefix]);
  const ext = format === "jpg" ? /\.jpg$/i : /\.png$/i;
  return (await fsp.readdir(folder)).filter(v => ext.test(v)).sort((a, b) => a.localeCompare(b, undefined, { numeric: true })).map(v => path.join(folder, v));
}

async function recognize(input, language = "khm+eng") {
  const worker = await createWorker(language);
  try { const result = await worker.recognize(input); return String(result.data.text || "").trim(); }
  finally { await worker.terminate(); }
}
async function ocrPdf(buffer, language) {
  const folder = tempFolder("ocr");
  try {
    const pages = await renderPdf(buffer, folder, "png", 180);
    if (!pages.length) throw new Error("No PDF pages could be rendered.");
    const worker = await createWorker(language);
    try {
      const chunks = [];
      for (let i = 0; i < pages.length; i++) {
        const result = await worker.recognize(pages[i]);
        chunks.push(`--- Page ${i + 1} ---\n${String(result.data.text || "").trim()}`);
      }
      return { text: chunks.join("\n\n"), pages: pages.length };
    } finally { await worker.terminate(); }
  } finally { await cleanup(folder); }
}
async function extractOrOcr(file, language) {
  const mime = await assertFile(file, ["application/pdf", "image/jpeg", "image/png", "image/webp"]);
  if (mime === "application/pdf") return ocrPdf(file.buffer, language);
  return { text: await recognize(file.buffer, language), pages: 1 };
}

async function wordBuffer(text) {
  const lines = splitText(text);
  const children = (lines.length ? lines : [""]).map(line => new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: line, font: "Noto Sans Khmer", size: 23 })] }));
  return Packer.toBuffer(new Document({ creator: "KhmerTools", sections: [{ properties: {}, children }] }));
}
function columns(line) { if (line.includes("\t")) return line.split("\t").map(v => v.trim()); if (line.includes("|")) return line.split("|").map(v => v.trim()); const values = line.split(/\s{2,}/).map(v => v.trim()).filter(Boolean); return values.length ? values : [line]; }
async function excelBuffer(text) {
  const book = new ExcelJS.Workbook(); book.creator = "KhmerTools";
  const sheet = book.addWorksheet("OCR Result", { views: [{ state: "frozen", ySplit: 1 }] });
  const rows = splitText(text).map(columns); const width = Math.max(1, ...rows.map(v => v.length));
  sheet.addRow(Array.from({ length: width }, (_, i) => `Column ${i + 1}`));
  rows.forEach(row => sheet.addRow(row));
  sheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" }, name: "Noto Sans Khmer" };
  sheet.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF2563EB" } };
  sheet.eachRow(row => row.eachCell(cell => { cell.font = { ...cell.font, name: "Noto Sans Khmer", size: 11 }; cell.alignment = { vertical: "top", wrapText: true }; }));
  sheet.columns.forEach(col => { let max = 10; col.eachCell({ includeEmpty: false }, c => { max = Math.max(max, Math.min(50, String(c.value || "").length + 2)); }); col.width = max; });
  return book.xlsx.writeBuffer();
}

function isPrivateAddress(address) {
  const ip = String(address).toLowerCase();
  if (net.isIPv4(ip)) { const [a, b] = ip.split(".").map(Number); return a === 0 || a === 10 || a === 127 || a >= 224 || (a === 169 && b === 254) || (a === 172 && b >= 16 && b <= 31) || (a === 192 && b === 168); }
  return ip === "::" || ip === "::1" || ip.startsWith("fc") || ip.startsWith("fd") || ip.startsWith("fe80:") || ip.startsWith("::ffff:127.") || ip.startsWith("::ffff:10.") || ip.startsWith("::ffff:192.168.");
}
async function validatePublicUrl(raw) {
  let url; try { url = new URL(String(raw)); } catch { throw Object.assign(new Error("Enter a valid public URL."), { status: 400 }); }
  if (!["http:", "https:"].includes(url.protocol) || url.username || url.password) throw Object.assign(new Error("Only public HTTP/HTTPS URLs without credentials are allowed."), { status: 400 });
  const addresses = await dns.lookup(url.hostname, { all: true });
  if (!addresses.length || addresses.some(entry => isPrivateAddress(entry.address))) throw Object.assign(new Error("Private, local, or reserved network addresses are blocked."), { status: 403 });
  return url;
}
async function fetchPublicMedia(raw) {
  let url = await validatePublicUrl(raw);
  for (let redirects = 0; redirects <= 4; redirects++) {
    const controller = new AbortController(); const timeout = setTimeout(() => controller.abort(), 30000);
    let response;
    try { response = await fetch(url, { redirect: "manual", signal: controller.signal, headers: { "User-Agent": "KhmerTools/2.0 (+public media fetcher)", Accept: "image/*,video/*,audio/*" } }); }
    finally { clearTimeout(timeout); }
    if ([301, 302, 303, 307, 308].includes(response.status)) { const location = response.headers.get("location"); if (!location || redirects === 4) throw Object.assign(new Error("Too many or invalid redirects."), { status: 400 }); url = await validatePublicUrl(new URL(location, url).toString()); continue; }
    if (!response.ok || !response.body) throw Object.assign(new Error(`Remote server returned ${response.status}.`), { status: 400 });
    const type = String(response.headers.get("content-type") || "").split(";")[0].toLowerCase();
    if (!/^(image|video|audio)\//.test(type) || type === "image/svg+xml") throw Object.assign(new Error("The URL must point directly to a supported public image, audio, or video file."), { status: 415 });
    const declared = Number(response.headers.get("content-length") || 0); if (declared > MAX_REMOTE_SIZE) throw Object.assign(new Error("Remote media exceeds the size limit."), { status: 413 });
    return { response, url, type };
  }
  throw new Error("Unable to fetch media.");
}

app.get("/api/status", (req, res) => res.json({ success: true, status: "online", service: "KhmerTools", aiConfigured: Boolean(process.env.GEMINI_API_KEY), maxFileSizeMb: MAX_FILE_SIZE / 1024 / 1024 }));

app.post("/api/ocr", upload.single("file"), asyncRoute(async (req, res) => {
  const result = await extractOrOcr(req.file, languageCode(req.body.language));
  res.json({ success: true, filename: safeName(req.file.originalname), ...result });
}));
app.post("/api/pdf-ocr", upload.single("pdf"), asyncRoute(async (req, res) => {
  await assertFile(req.file, ["application/pdf"]); const result = await ocrPdf(req.file.buffer, languageCode(req.body.language)); res.json({ success: true, filename: safeName(req.file.originalname), ...result });
}));
app.post("/api/pdf-to-word", upload.single("pdf"), asyncRoute(async (req, res) => {
  await assertFile(req.file, ["application/pdf"]); const pdfParse = require("pdf-parse"); const parsed = await pdfParse(req.file.buffer); const text = String(parsed.text || "").trim(); if (!text) throw Object.assign(new Error("No selectable text found. Use Scanned PDF → Word instead."), { status: 422 }); sendBuffer(res, await wordBuffer(text), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", `${baseName(req.file.originalname)}.docx`);
}));
app.post("/api/pdf-ocr-to-word", upload.single("pdf"), asyncRoute(async (req, res) => {
  await assertFile(req.file, ["application/pdf"]); const result = await ocrPdf(req.file.buffer, languageCode(req.body.language)); if (!result.text.trim()) throw Object.assign(new Error("No text was detected."), { status: 422 }); sendBuffer(res, await wordBuffer(result.text), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", `${baseName(req.file.originalname)}-OCR.docx`);
}));
app.post("/api/ocr-to-excel", upload.single("file"), asyncRoute(async (req, res) => {
  const text = req.file ? (await extractOrOcr(req.file, languageCode(req.body.language))).text : assertText(req.body.text); if (!text.trim()) throw Object.assign(new Error("No text was detected."), { status: 422 }); sendBuffer(res, Buffer.from(await excelBuffer(text)), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", `${req.file ? baseName(req.file.originalname) : "OCR-result"}.xlsx`);
}));
app.post("/api/text-to-word", asyncRoute(async (req, res) => sendBuffer(res, await wordBuffer(assertText(req.body.text)), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "khmertools-text.docx")));
app.post("/api/text-to-excel", asyncRoute(async (req, res) => sendBuffer(res, Buffer.from(await excelBuffer(assertText(req.body.text))), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "khmertools-text.xlsx")));

app.post("/api/pdf-to-images", upload.single("pdf"), asyncRoute(async (req, res) => {
  await assertFile(req.file, ["application/pdf"]); const format = outputFormat(req.body.format); const folder = tempFolder("pages");
  try {
    const pages = await renderPdf(req.file.buffer, folder, format, 160); if (!pages.length) throw new Error("No pages were generated.");
    res.set({ "Content-Type": "application/zip", "Content-Disposition": contentDisposition(`${baseName(req.file.originalname)}-${format}-pages.zip`) });
    const archive = archiver("zip", { zlib: { level: 8 } }); archive.on("error", err => res.destroy(err)); archive.pipe(res); pages.forEach((file, i) => archive.file(file, { name: `page-${i + 1}.${format}` })); await archive.finalize();
  } finally { setTimeout(() => cleanup(folder), 1500); }
}));
app.post("/api/pdf-merge", upload.array("files", 12), asyncRoute(async (req, res) => {
  if (!req.files || req.files.length < 2) throw Object.assign(new Error("Choose at least two PDF files."), { status: 400 }); const output = await PDFDocument.create();
  for (const file of req.files) { await assertFile(file, ["application/pdf"]); const source = await PDFDocument.load(file.buffer, { ignoreEncryption: false }); const copied = await output.copyPages(source, source.getPageIndices()); copied.forEach(page => output.addPage(page)); }
  sendBuffer(res, Buffer.from(await output.save()), "application/pdf", "merged.pdf");
}));
app.post("/api/pdf-split", upload.single("pdf"), asyncRoute(async (req, res) => {
  await assertFile(req.file, ["application/pdf"]); const source = await PDFDocument.load(req.file.buffer); const folder = tempFolder("split");
  try { for (let i = 0; i < source.getPageCount(); i++) { const out = await PDFDocument.create(); const [page] = await out.copyPages(source, [i]); out.addPage(page); await fsp.writeFile(path.join(folder, `page-${i + 1}.pdf`), await out.save()); }
    res.set({ "Content-Type": "application/zip", "Content-Disposition": contentDisposition(`${baseName(req.file.originalname)}-split.zip`) }); const archive = archiver("zip"); archive.on("error", err => res.destroy(err)); archive.directory(folder, false).pipe(res); await archive.finalize();
  } finally { setTimeout(() => cleanup(folder), 1500); }
}));
app.post("/api/pdf-rotate", upload.single("pdf"), asyncRoute(async (req, res) => {
  await assertFile(req.file, ["application/pdf"]); const angle = Number(req.body.angle); if (![90, 180, 270].includes(angle)) throw Object.assign(new Error("Rotation must be 90, 180, or 270 degrees."), { status: 400 }); const pdf = await PDFDocument.load(req.file.buffer); pdf.getPages().forEach(page => page.setRotation(degrees((page.getRotation().angle + angle) % 360))); sendBuffer(res, Buffer.from(await pdf.save()), "application/pdf", `${baseName(req.file.originalname)}-rotated.pdf`);
}));
app.post("/api/pdf-protect", upload.single("pdf"), asyncRoute(async (req, res) => {
  await assertFile(req.file, ["application/pdf"]); throw Object.assign(new Error("Strong PDF encryption requires an external qpdf integration and is not enabled on this server. Your file was not stored."), { status: 501 });
}));
app.post("/api/pdf-compress", upload.single("pdf"), asyncRoute(async (req, res) => {
  await assertFile(req.file, ["application/pdf"]); const folder = tempFolder("compress"); const input = path.join(folder, "input.pdf"); const output = path.join(folder, "output.pdf");
  try {
    await fsp.writeFile(input, req.file.buffer);
    await execFileAsync(PDFTOPDF, ["-pdf", "-singlefile", "-jpeg", "-jpegopt", "quality=70", input, path.join(folder, "output")]);
    const data = await fsp.readFile(output);
    sendBuffer(res, data, "application/pdf", `${baseName(req.file.originalname)}-compressed.pdf`);
  } catch (error) {
    throw Object.assign(new Error("PDF compression needs Poppler pdftocairo. Configure PDFTOPDF_PATH or install Poppler."), { status: 503 });
  } finally { await cleanup(folder); }
}));

app.post("/api/image-convert", upload.single("image"), asyncRoute(async (req, res) => {
  await assertFile(req.file, ["image/jpeg", "image/png", "image/webp"]); const format = ["png", "jpeg", "webp"].includes(req.body.format) ? req.body.format : "png"; const image = sharp(req.file.buffer, { limitInputPixels: 40_000_000 }).rotate(); const buffer = await image[format]({ quality: 88 }).toBuffer(); const ext = format === "jpeg" ? "jpg" : format; sendBuffer(res, buffer, `image/${format}`, `${baseName(req.file.originalname)}.${ext}`);
}));
app.post("/api/image-compress", upload.single("image"), asyncRoute(async (req, res) => {
  const mime = await assertFile(req.file, ["image/jpeg", "image/png", "image/webp"]); const quality = Math.min(95, Math.max(20, Number(req.body.quality || 72))); const image = sharp(req.file.buffer, { limitInputPixels: 40_000_000 }).rotate(); let buffer; let ext; if (mime === "image/png") { buffer = await image.png({ compressionLevel: 9, quality }).toBuffer(); ext = "png"; } else if (mime === "image/webp") { buffer = await image.webp({ quality }).toBuffer(); ext = "webp"; } else { buffer = await image.jpeg({ quality, mozjpeg: true }).toBuffer(); ext = "jpg"; } sendBuffer(res, buffer, mime, `${baseName(req.file.originalname)}-compressed.${ext}`);
}));
app.post("/api/image-resize", upload.single("image"), asyncRoute(async (req, res) => {
  const mime = await assertFile(req.file, ["image/jpeg", "image/png", "image/webp"]); const width = Number(req.body.width); const height = Number(req.body.height); if ((!width && !height) || width > 8000 || height > 8000) throw Object.assign(new Error("Provide a width or height up to 8000 pixels."), { status: 400 }); const buffer = await sharp(req.file.buffer, { limitInputPixels: 40_000_000 }).rotate().resize(width || null, height || null, { fit: "inside", withoutEnlargement: false }).toBuffer(); sendBuffer(res, buffer, mime, `${baseName(req.file.originalname)}-resized${path.extname(req.file.originalname) || ".png"}`);
}));
app.post("/api/image-to-pdf", upload.array("images", 20), asyncRoute(async (req, res) => {
  if (!req.files?.length) throw Object.assign(new Error("Choose one or more images."), { status: 400 }); const pdf = await PDFDocument.create();
  for (const file of req.files) { const mime = await assertFile(file, ["image/jpeg", "image/png", "image/webp"]); let buffer = file.buffer; let type = mime; if (mime === "image/webp") { buffer = await sharp(buffer).png().toBuffer(); type = "image/png"; } const image = type === "image/png" ? await pdf.embedPng(buffer) : await pdf.embedJpg(buffer); const page = pdf.addPage([image.width, image.height]); page.drawImage(image, { x: 0, y: 0, width: image.width, height: image.height }); }
  sendBuffer(res, Buffer.from(await pdf.save()), "application/pdf", "images.pdf");
}));
app.post("/api/qr", asyncRoute(async (req, res) => { const text = assertText(req.body.text, 2000); const buffer = await QRCode.toBuffer(text, { type: "png", width: 900, margin: 2, errorCorrectionLevel: "M", color: { dark: "#07111f", light: "#ffffff" } }); sendBuffer(res, buffer, "image/png", "khmertools-qr.png"); }));

app.post("/api/ai/chat", asyncRoute(async (req, res) => {
  const key = process.env.GEMINI_API_KEY; if (!key) throw Object.assign(new Error("AI is not configured. Add GEMINI_API_KEY to the server environment."), { status: 503 });
  const messages = Array.isArray(req.body.messages) ? req.body.messages.slice(-20).map(item => ({ role: item.role === "assistant" ? "model" : "user", parts: [{ text: String(item.content || "").slice(0, 12000) }] })).filter(item => item.parts[0].text.trim()) : [];
  if (!messages.length) throw Object.assign(new Error("Send a message first."), { status: 400 });
  const genAI = new GoogleGenerativeAI(key); const model = genAI.getGenerativeModel({ model: process.env.GEMINI_MODEL || "gemini-3.6-flash", systemInstruction: "You are KhmerTools AI. Be helpful, accurate and concise. Reply in Khmer when the user writes Khmer; otherwise follow the user's language. Never claim web access unless supplied current sources." });
  const latest = messages.pop().parts[0].text; const chat = model.startChat({ history: messages }); const result = await chat.sendMessage(latest); res.json({ success: true, text: result.response.text() });
}));

app.post("/api/media/download", asyncRoute(async (req, res) => {
  const { response, url, type } = await fetchPublicMedia(req.body.url); const extMap = { "image/jpeg": ".jpg", "image/png": ".png", "image/webp": ".webp", "image/gif": ".gif", "video/mp4": ".mp4", "video/webm": ".webm", "audio/mpeg": ".mp3", "audio/mp4": ".m4a" }; const fromUrl = path.basename(decodeURIComponent(url.pathname)); const filename = safeName(fromUrl.includes(".") ? fromUrl : `media${extMap[type] || ".bin"}`);
  res.set({ "Content-Type": type, "Content-Disposition": contentDisposition(filename) }); let total = 0; const limiter = new Transform({ transform(chunk, enc, cb) { total += chunk.length; if (total > MAX_REMOTE_SIZE) cb(Object.assign(new Error("Remote media exceeded the size limit."), { status: 413 })); else cb(null, chunk); } }); await pipeline(Readable.fromWeb(response.body), limiter, res);
}));

app.get("/", (req, res) => res.sendFile(path.join(PUBLIC_DIR, "index.html")));
app.use("/api", (req, res) => res.status(404).json({ success: false, error: "API route not found." }));
app.use((error, req, res, next) => {
  console.error(`[${new Date().toISOString()}]`, error.message);
  if (res.headersSent) return next(error);
  const status = error.status || (error.code === "LIMIT_FILE_SIZE" ? 413 : error instanceof multer.MulterError ? 400 : 500);
  const message = status >= 500 && process.env.NODE_ENV === "production" ? "The request could not be completed." : error.message;
  res.status(status).json({ success: false, error: message || "Unexpected server error." });
});

if (require.main === module) app.listen(PORT, HOST, () => console.log(`KhmerTools running at http://localhost:${PORT}`));
module.exports = app;
