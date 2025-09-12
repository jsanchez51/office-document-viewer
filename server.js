// Servidor en memoria para exponer archivos temporalmente a Office Viewer
// No persiste en disco. Mantiene los binarios en RAM con expiración.

const express = require("express");
const multer = require("multer");
const cors = require("cors");
const crypto = require("crypto");
const path = require("path");

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

// Configuración
const PORT = process.env.PORT ? Number(process.env.PORT) : 3000;
const TTL_MINUTES = process.env.TTL_MINUTES ? Number(process.env.TTL_MINUTES) : 10; // minutos
const MAX_FILE_MB = process.env.MAX_FILE_MB ? Number(process.env.MAX_FILE_MB) : 50;
const PUBLIC_BASE_URL = process.env.PUBLIC_BASE_URL || `http://localhost:${PORT}`;

const MAX_BYTES = MAX_FILE_MB * 1024 * 1024;

// Almacenamiento en memoria: id -> { buffer, mime, filename, expireAt, size }
/** @type {Map<string, {buffer: Buffer, mime: string, filename: string, expireAt: number, size: number}>} */
const inMemoryFiles = new Map();

// Estadísticas de transferencia para progreso del visor: id -> { size, bytesSent, startedAt?, completedAt? }
/** @type {Map<string, { size: number, bytesSent: number, startedAt?: number, completedAt?: number }>} */
const inMemoryStats = new Map();

// CORS para poder subir desde file:// o desde distintos orígenes
app.use(cors({ origin: true }));

// Salud y config
app.get("/health", (_req, res) => res.json({ ok: true }));
app.get("/config", (_req, res) => res.json({ PUBLIC_BASE_URL, TTL_MINUTES, MAX_FILE_MB }));
// Progreso de entrega al visor de Office
app.get("/progress/:id", (req, res) => {
  const stat = inMemoryStats.get(req.params.id);
  if (!stat) return res.status(404).json({ error: "not_found" });
  const now = Date.now();
  const startedAt = stat.startedAt || now;
  const elapsed = (now - startedAt) / 1000;
  const remainingBytes = Math.max((stat.size || 0) - (stat.bytesSent || 0), 0);
  const rate = elapsed > 0 ? (stat.bytesSent || 0) / elapsed : 0; // bytes/s
  const etaSec = rate > 0 ? Math.ceil(remainingBytes / rate) : null;
  res.json({ size: stat.size, bytesSent: stat.bytesSent, etaSec, elapsed });
});

// Subida de archivo a RAM
app.post("/upload", upload.single("file"), (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "file_required" });
    if (req.file.size > MAX_BYTES) return res.status(413).json({ error: "file_too_large", maxMb: MAX_FILE_MB });

    const id = crypto.randomUUID();
    const filename = req.file.originalname || "archivo";
    const mime = req.file.mimetype || guessMimeByExt(filename);
    const expireAt = Date.now() + TTL_MINUTES * 60 * 1000;

    inMemoryFiles.set(id, {
      buffer: req.file.buffer,
      mime,
      filename,
      expireAt,
      size: req.file.buffer.length,
    });
    inMemoryStats.set(id, { size: req.file.buffer.length, bytesSent: 0 });

    // Autoeliminación programada
    setTimeout(() => inMemoryFiles.delete(id), TTL_MINUTES * 60 * 1000).unref?.();

    // Detectar base pública desde el túnel/proxy si existe
    const xfProto = (req.headers["x-forwarded-proto"] || "").toString();
    const host = (req.headers.host || "").toString();
    const detectedBase = xfProto && host ? `${xfProto}://${host}` : undefined;
    const base = detectedBase || PUBLIC_BASE_URL;
    const fileUrl = `${base}/f/${id}`;
    const viewerUrl = `https://view.officeapps.live.com/op/embed.aspx?src=${encodeURIComponent(fileUrl)}`;
    return res.json({ id, fileUrl, viewerUrl, filename, expiresAt: expireAt });
  } catch (e) {
    return res.status(500).json({ error: "internal_error" });
  }
});

// Descarga/lectura del archivo desde memoria
// Soporta Range parcial (206) por si el visor lo solicita
app.get("/f/:id", (req, res) => {
  const file = inMemoryFiles.get(req.params.id);
  if (!file || Date.now() > file.expireAt) {
    inMemoryFiles.delete(req.params.id);
    inMemoryStats.delete(req.params.id);
    return res.status(404).send("Not found or expired");
  }

  const total = file.buffer.length;
  res.setHeader("Accept-Ranges", "bytes");
  res.setHeader("Cache-Control", "no-store");
  res.setHeader("Content-Type", file.mime);
  res.setHeader("Content-Disposition", `inline; filename="${safeFilename(file.filename)}"`);

  const range = req.headers.range;
  if (!range) {
    res.setHeader("Content-Length", String(total));
    const bytes = total;
    const stat = inMemoryStats.get(req.params.id) || { size: total, bytesSent: 0 };
    if (!stat.startedAt) stat.startedAt = Date.now();
    stat.bytesSent += bytes;
    if (stat.bytesSent >= stat.size) stat.completedAt = Date.now();
    inMemoryStats.set(req.params.id, stat);
    return res.status(200).end(file.buffer);
  }

  const match = /bytes=(\d+)-(\d+)?/.exec(range);
  if (!match) {
    res.setHeader("Content-Length", String(total));
    return res.status(200).end(file.buffer);
  }
  const start = Number(match[1]);
  const end = match[2] ? Number(match[2]) : total - 1;
  if (isNaN(start) || isNaN(end) || start > end || end >= total) {
    return res.status(416).send("Requested Range Not Satisfiable");
  }

  const chunk = file.buffer.subarray(start, end + 1);
  res.status(206);
  res.setHeader("Content-Range", `bytes ${start}-${end}/${total}`);
  res.setHeader("Content-Length", String(chunk.length));
  const stat = inMemoryStats.get(req.params.id) || { size: total, bytesSent: 0 };
  if (!stat.startedAt) stat.startedAt = Date.now();
  stat.bytesSent += chunk.length;
  if (stat.bytesSent >= stat.size) stat.completedAt = Date.now();
  inMemoryStats.set(req.params.id, stat);
  return res.end(chunk);
});

// Servir estático opcionalmente (para abrir http://localhost:3000/)
app.use(express.static(process.cwd(), { fallthrough: true }));

// Limpieza periódica de expirados (cada 2 minutos)
setInterval(() => {
  const now = Date.now();
  for (const [id, file] of inMemoryFiles.entries()) {
    if (now > file.expireAt) inMemoryFiles.delete(id);
  }
  for (const [id, stat] of inMemoryStats.entries()) {
    if (!inMemoryFiles.has(id)) inMemoryStats.delete(id);
  }
}, 2 * 60 * 1000).unref?.();

app.listen(PORT, () => {
  console.log(`Servidor en memoria listo en http://localhost:${PORT}`);
  console.log(`PUBLIC_BASE_URL=${PUBLIC_BASE_URL}`);
});

function guessMimeByExt(filename) {
  const ext = path.extname(filename).toLowerCase();
  if (ext === ".docx") return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  if (ext === ".xlsx") return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  if (ext === ".doc") return "application/msword";
  if (ext === ".xls") return "application/vnd.ms-excel";
  if (ext === ".pptx") return "application/vnd.openxmlformats-officedocument.presentationml.presentation";
  if (ext === ".ppt") return "application/vnd.ms-powerpoint";
  return "application/octet-stream";
}

function safeFilename(name) {
  return String(name || "archivo").replace(/[\r\n"<>\\]/g, "_");
}


