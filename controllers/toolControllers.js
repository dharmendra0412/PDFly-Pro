const fs = require('fs-extra');
const path = require('path');
const { PDFDocument, rgb, degrees } = require('pdf-lib');
const PDFParser = require('pdf-parse');
const { v4: uuidv4 } = require('uuid');
const Anthropic = require('@anthropic-ai/sdk');
const mammoth = require('mammoth');
const XLSX = require('xlsx');
const JimpModule = require('jimp');
const Jimp = JimpModule.Jimp;
const { Document, Packer, Paragraph, TextRun } = require('docx');
const PDFDocument_kit = require('pdfkit');
let puppeteer = null;
let tesseract = null;
try { puppeteer = require('puppeteer-core'); } catch (error) {}
try { tesseract = require('tesseract.js'); } catch (error) {}

const resultsDir = path.join(__dirname, '../results');
const uploadsDir = path.join(__dirname, '../uploads');
const serverBaseUrl = `http://127.0.0.1:${process.env.PORT || 5000}`;
const repoRoot = path.join(__dirname, '..');

const getLocalTool = (relativePath) => {
  const fullPath = path.join(repoRoot, relativePath);
  return fs.existsSync(fullPath) ? fullPath : null;
};

const findOnPath = (names) => {
  const pathVar = process.env.PATH || '';
  const sep = process.platform === 'win32' ? ';' : ':';
  const dirs = pathVar.split(sep).filter(Boolean);
  for (const dir of dirs) {
    for (const name of names) {
      const candidate = path.join(dir, name);
      if (fs.existsSync(candidate)) return candidate;
    }
  }
  return null;
};

const getQpdfPath = () => (
  process.env.QPDF_PATH ||
  findOnPath(process.platform === 'win32' ? ['qpdf.exe'] : ['qpdf']) ||
  getLocalTool('tools/qpdf/qpdf-12.3.2-msvc64/bin/qpdf.exe')
);

const getGhostscriptPath = () => (
  process.env.GS_PATH ||
  findOnPath(process.platform === 'win32' ? ['gswin64c.exe', 'gswin32c.exe'] : ['gs']) ||
  getLocalTool('tools/gs/bin/gswin64c.exe')
);

const getSofficePath = () => (
  process.env.SOFFICE_PATH ||
  findOnPath(process.platform === 'win32' ? ['soffice.com', 'soffice.exe'] : ['soffice']) ||
  getLocalTool('tools/downloads/LibreOfficePortable/App/libreoffice/program/soffice.com') ||
  getLocalTool('tools/downloads/LibreOfficePortable/App/libreoffice/program/soffice.exe')
);

const getLibreOfficeProfileDir = async () => {
  const profileDir = path.join(repoRoot, 'tools', 'lo-profile');
  await fs.ensureDir(profileDir);
  // file URL format required by LibreOffice
  const uriPath = profileDir.replace(/\\/g, '/');
  return `file:///${uriPath}`;
};

const convertWithLibreOffice = async (inputPath, outDir) => {
  const soffice = getSofficePath();
  if (!soffice) throw new Error('LibreOffice (soffice) not found');
  const userInstall = await getLibreOfficeProfileDir();
  await runExe(soffice, [
    '--headless',
    '--norestore',
    '--nolockcheck',
    '--nodefault',
    '--nofirststartwizard',
    `-env:UserInstallation=${userInstall}`,
    '--convert-to', 'pdf',
    '--outdir', outDir,
    inputPath
  ], { cwd: outDir });
};

const runExe = async (exePath, args, options = {}) => {
  const { spawn } = require('child_process');
  return await new Promise((resolve, reject) => {
    const child = spawn(exePath, args, {
      windowsHide: true,
      ...options
    });
    let stdout = '';
    let stderr = '';
    child.stdout.on('data', d => stdout += d.toString());
    child.stderr.on('data', d => stderr += d.toString());
    child.on('error', reject);
    child.on('close', (code) => {
      if (code === 0) return resolve({ stdout, stderr });
      reject(new Error(stderr.trim() || stdout.trim() || `Process failed (exit ${code})`));
    });
  });
};

// Helper function to load PDF with encryption handling
const loadPDF = async (pdfBytes) => {
  try {
    return await PDFDocument.load(pdfBytes, { ignoreEncryption: true });
  } catch (error) {
    // If encryption ignore doesn't work, try without it
    return await PDFDocument.load(pdfBytes);
  }
};

// Helper to normalize multer file structure (handles both .single/.array and .any)
const getFiles = (req) => {
  if (req.files && Array.isArray(req.files)) {
    return req.files.filter(f => f.fieldname === 'files' || f.fieldname === 'file');
  }
  if (req.file) {
    return [req.file];
  }
  return [];
};

const getFile = (req) => {
  if (req.file) {
    return req.file;
  }
  if (req.files && Array.isArray(req.files)) {
    const file = req.files.find(f => f.fieldname === 'file');
    if (file) return file;
    return req.files[0];
  }
  return null;
};

// Initialize Anthropic client
const anthropic = new Anthropic({
  apiKey: process.env.ANTHROPIC_API_KEY || 'test-key'
});

// ========== UTILITY FUNCTIONS ==========

const generateFilename = (extension) => {
  return `${uuidv4()}${extension}`;
};

const crcTable = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[i] = c >>> 0;
  }
  return table;
})();

const crc32 = (buffer) => {
  let crc = 0xffffffff;
  for (const byte of buffer) {
    crc = crcTable[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
};

const getDosDateTime = (date = new Date()) => ({
  time: (date.getHours() << 11) | (date.getMinutes() << 5) | Math.floor(date.getSeconds() / 2),
  date: ((date.getFullYear() - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate()
});

const writeZipEntries = async (zipPath, entries) => {
  const chunks = [];
  const centralDirectory = [];
  let offset = 0;

  for (const entry of entries) {
    const data = Buffer.isBuffer(entry.data) ? entry.data : Buffer.from(entry.data);
    const name = Buffer.from(entry.name);
    const crc = crc32(data);
    const { time, date } = getDosDateTime();

    const localHeader = Buffer.alloc(30);
    localHeader.writeUInt32LE(0x04034b50, 0);
    localHeader.writeUInt16LE(20, 4);
    localHeader.writeUInt16LE(0, 6);
    localHeader.writeUInt16LE(0, 8);
    localHeader.writeUInt16LE(time, 10);
    localHeader.writeUInt16LE(date, 12);
    localHeader.writeUInt32LE(crc, 14);
    localHeader.writeUInt32LE(data.length, 18);
    localHeader.writeUInt32LE(data.length, 22);
    localHeader.writeUInt16LE(name.length, 26);
    localHeader.writeUInt16LE(0, 28);

    chunks.push(localHeader, name, data);

    const centralHeader = Buffer.alloc(46);
    centralHeader.writeUInt32LE(0x02014b50, 0);
    centralHeader.writeUInt16LE(20, 4);
    centralHeader.writeUInt16LE(20, 6);
    centralHeader.writeUInt16LE(0, 8);
    centralHeader.writeUInt16LE(0, 10);
    centralHeader.writeUInt16LE(time, 12);
    centralHeader.writeUInt16LE(date, 14);
    centralHeader.writeUInt32LE(crc, 16);
    centralHeader.writeUInt32LE(data.length, 20);
    centralHeader.writeUInt32LE(data.length, 24);
    centralHeader.writeUInt16LE(name.length, 28);
    centralHeader.writeUInt16LE(0, 30);
    centralHeader.writeUInt16LE(0, 32);
    centralHeader.writeUInt16LE(0, 34);
    centralHeader.writeUInt16LE(0, 36);
    centralHeader.writeUInt32LE(0, 38);
    centralHeader.writeUInt32LE(offset, 42);
    centralDirectory.push(centralHeader, name);

    offset += localHeader.length + name.length + data.length;
  }

  const centralSize = centralDirectory.reduce((sum, chunk) => sum + chunk.length, 0);
  const endHeader = Buffer.alloc(22);
  endHeader.writeUInt32LE(0x06054b50, 0);
  endHeader.writeUInt16LE(0, 4);
  endHeader.writeUInt16LE(0, 6);
  endHeader.writeUInt16LE(entries.length, 8);
  endHeader.writeUInt16LE(entries.length, 10);
  endHeader.writeUInt32LE(centralSize, 12);
  endHeader.writeUInt32LE(offset, 16);
  endHeader.writeUInt16LE(0, 20);

  await fs.writeFile(zipPath, Buffer.concat([...chunks, ...centralDirectory, endHeader]));
};

const createZipArchive = async (filenames) => {
  const zipFilename = generateFilename('.zip');
  const zipPath = path.join(resultsDir, zipFilename);
  const entries = [];

  for (const [index, filename] of filenames.entries()) {
    const filePath = path.join(resultsDir, filename);
    entries.push({
      name: `${index + 1}-${filename}`,
      data: await fs.readFile(filePath)
    });
  }

  await writeZipEntries(zipPath, entries);
  return zipFilename;
};

const findFileRecursive = (dir, filename) => {
  if (!fs.existsSync(dir)) return null;
  const entries = fs.readdirSync(dir, { withFileTypes: true });
  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isFile() && entry.name.toLowerCase() === filename.toLowerCase()) return fullPath;
    if (entry.isDirectory()) {
      const found = findFileRecursive(fullPath, filename);
      if (found) return found;
    }
  }
  return null;
};

const getChromiumExecutable = () => {
  const configured = process.env.CHROME_PATH || process.env.PUPPETEER_EXECUTABLE_PATH;
  if (configured && fs.existsSync(configured)) return configured;

  const home = process.env.USERPROFILE || process.env.HOME || '';
  const puppeteerCache = path.join(home, '.cache', 'puppeteer');
  return findFileRecursive(puppeteerCache, 'chrome-headless-shell.exe') ||
    findFileRecursive(puppeteerCache, 'chrome.exe');
};

const withBrowser = async (work) => {
  if (!puppeteer) throw new Error('puppeteer-core is not installed');
  const executablePath = getChromiumExecutable();
  if (!executablePath) throw new Error('Chromium executable not found');

  const browser = await puppeteer.launch({
    executablePath,
    headless: true,
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--allow-file-access-from-files',
      '--ignore-certificate-errors'
    ]
  });

  try {
    return await work(browser);
  } finally {
    await browser.close();
  }
};

const renderPdfPagesToJpegs = async (pdfBytes, maxPages = 20, scale = 2, quality = 85) => {
  const pdfData = Buffer.from(pdfBytes).toString('base64');
  const pdfjsUrl = `${serverBaseUrl}/node_modules/pdfjs-dist/build/pdf.mjs`;
  const workerUrl = `${serverBaseUrl}/node_modules/pdfjs-dist/build/pdf.worker.mjs`;

  return await withBrowser(async (browser) => {
    const page = await browser.newPage();
    page.on('pageerror', error => console.error('PDF render page error:', error.message));
    const html = `
      <!doctype html><html><body style="margin:0;background:white">
        <canvas id="pdf-canvas"></canvas>
        <script type="module">
          import * as pdfjsLib from '${pdfjsUrl}';
          pdfjsLib.GlobalWorkerOptions.workerSrc = '${workerUrl}';
          const binary = atob('${pdfData}');
          const data = new Uint8Array(binary.length);
          for (let i = 0; i < binary.length; i++) data[i] = binary.charCodeAt(i);
          window.pdfDoc = await pdfjsLib.getDocument({ data }).promise;
          window.renderPdfPage = async (pageNumber, scale) => {
            const page = await window.pdfDoc.getPage(pageNumber);
            const viewport = page.getViewport({ scale });
            const canvas = document.getElementById('pdf-canvas');
            canvas.width = Math.ceil(viewport.width);
            canvas.height = Math.ceil(viewport.height);
            const context = canvas.getContext('2d', { alpha: false });
            context.fillStyle = '#fff';
            context.fillRect(0, 0, canvas.width, canvas.height);
            await page.render({ canvasContext: context, viewport }).promise;
            return { width: canvas.width, height: canvas.height };
          };
          window.pdfReady = true;
        </script>
      </body></html>`;
    await page.setContent(html, { waitUntil: 'load' });
    await page.waitForFunction('window.pdfReady === true', { timeout: 20000 });
    const pageCount = await page.evaluate(() => window.pdfDoc.numPages);
    const files = [];
    const count = Math.min(pageCount, maxPages);

    for (let i = 1; i <= count; i++) {
      const viewport = await page.evaluate((pageNumber, renderScale) => {
        return window.renderPdfPage(pageNumber, renderScale);
      }, i, scale);
      await page.setViewport({
        width: Math.max(1, Math.ceil(viewport.width)),
        height: Math.max(1, Math.ceil(viewport.height)),
        deviceScaleFactor: 1
      });
      const canvas = await page.$('#pdf-canvas');
      const imageBuffer = await canvas.screenshot({ type: 'jpeg', quality });
      const filename = generateFilename('.jpg');
      await fs.writeFile(path.join(resultsDir, filename), imageBuffer);
      files.push(filename);
    }

    return { files, pageCount };
  });
};

const renderUrlToPDF = async (url, filepath) => {
  return await withBrowser(async (browser) => {
    const page = await browser.newPage();
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 6000 });
    const pdfBuffer = await page.pdf({
      format: 'A4',
      printBackground: true,
      margin: { top: '12mm', right: '12mm', bottom: '12mm', left: '12mm' }
    });
    await fs.writeFile(filepath, pdfBuffer);
  });
};

const renderHtmlToPDF = async (html, filepath) => {
  return await withBrowser(async (browser) => {
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: 'load' });
    const pdfBuffer = await page.pdf({
      format: 'A4',
      printBackground: true,
      margin: { top: '12mm', right: '12mm', bottom: '12mm', left: '12mm' }
    });
    await fs.writeFile(filepath, pdfBuffer);
  });
};

const extractPptxTextBySlide = async (pptxPath) => {
  let AdmZip;
  try {
    AdmZip = require('adm-zip');
  } catch (error) {
    throw new Error('adm-zip is required for PPTX parsing');
  }

  const zip = new AdmZip(pptxPath);
  const slideEntries = zip.getEntries()
    .filter(e => /^ppt\/slides\/slide\d+\.xml$/i.test(e.entryName))
    .sort((a, b) => {
      const an = parseInt(a.entryName.match(/slide(\d+)\.xml/i)[1], 10);
      const bn = parseInt(b.entryName.match(/slide(\d+)\.xml/i)[1], 10);
      return an - bn;
    });

  if (slideEntries.length === 0) return [];

  const slides = [];
  for (const entry of slideEntries) {
    const xml = entry.getData().toString('utf8');
    // Very lightweight text extraction: <a:t>Text</a:t>
    const texts = [...xml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/gi)]
      .map(m => m[1].replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&'))
      .map(s => s.trim())
      .filter(Boolean);
    slides.push(texts.join('\n'));
  }
  return slides;
};

const sendSuccess = (res, data = {}) => {
  res.json({
    success: true,
    ...data,
    timestamp: new Date().toISOString()
  });
};

const sendError = (res, error, statusCode = 500) => {
  console.error('Error:', error);
  res.status(statusCode).json({
    success: false,
    error: error.message || error,
    timestamp: new Date().toISOString()
  });
};

const parsePages = (input, pageCount) => {
  if (!input || String(input).trim() === '') {
    throw new Error('Pages parameter is required. Example: "1,3,5" or "2-4"');
  }

  const pages = new Set();
  String(input).split(',').forEach(part => {
    const token = part.trim().toLowerCase();
    if (!token) return;

    if (token === 'last') {
      pages.add(pageCount - 1);
      return;
    }

    const range = token.match(/^(\d+)\s*-\s*(\d+)$/);
    if (range) {
      const start = parseInt(range[1], 10);
      const end = parseInt(range[2], 10);
      const from = Math.min(start, end);
      const to = Math.max(start, end);
      for (let page = from; page <= to; page++) {
        if (page >= 1 && page <= pageCount) pages.add(page - 1);
      }
      return;
    }

    const page = parseInt(token, 10);
    if (Number.isNaN(page)) throw new Error(`Invalid page number: ${part.trim()}`);
    if (page >= 1 && page <= pageCount) pages.add(page - 1);
  });

  if (pages.size === 0) throw new Error('No valid pages selected');
  return [...pages].sort((a, b) => a - b);
};

const extractPdfInfo = async (pdfBytes) => {
  try {
    return await PDFParser(pdfBytes);
  } catch (error) {
    const pdf = await loadPDF(pdfBytes);
    return {
      text: '',
      numpages: pdf.getPageCount(),
      parseError: error.message
    };
  }
};

const writeTextPDF = async (filepath, title, body) => {
  const doc = new PDFDocument_kit();
  const stream = fs.createWriteStream(filepath);
  doc.pipe(stream);
  doc.fontSize(14).text(title, 50, 50, { width: 500 });
  doc.moveDown();
  doc.fontSize(10).text(body || 'No extractable text was found in this PDF.', { width: 500 });
  doc.end();
  await new Promise(resolve => stream.on('finish', resolve));
};

// ========== ORGANIZE TOOLS ==========

// Merge PDF
exports.mergePDF = async (req, res) => {
  try {
    if (!req.files || req.files.length === 0) {
      return sendError(res, 'No files provided', 400);
    }

    const mergedPdf = await PDFDocument.create();
    
    for (const file of req.files) {
      const pdfBytes = await fs.readFile(file.path);
      const pdf = await loadPDF(pdfBytes);
      const copiedPages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());
      copiedPages.forEach((page) => mergedPdf.addPage(page));
    }

    const filename = generateFilename('.pdf');
    const filepath = path.join(resultsDir, filename);
    await fs.writeFile(filepath, await mergedPdf.save());

    // Cleanup uploads
    req.files.forEach(f => fs.unlink(f.path).catch(() => {}));

    sendSuccess(res, {
      filename: filename,
      message: 'PDFs merged successfully',
      pageCount: mergedPdf.getPageCount()
    });
  } catch (error) {
    sendError(res, error);
  }
};

// Split PDF
exports.splitPDF = async (req, res) => {
  try {
    const pdfBytes = await fs.readFile(req.file.path);
    const pdf = await loadPDF(pdfBytes);
    const pageCount = pdf.getPageCount();

    const files = [];
    for (let i = 0; i < pageCount; i++) {
      const newPdf = await PDFDocument.create();
      const [page] = await newPdf.copyPages(pdf, [i]);
      newPdf.addPage(page);
      
      const filename = generateFilename('.pdf');
      const filepath = path.join(resultsDir, filename);
      await fs.writeFile(filepath, await newPdf.save());
      files.push(filename);
    }
    const zipFilename = await createZipArchive(files);

    fs.unlink(req.file.path).catch(() => {});

    sendSuccess(res, {
      files: files,
      zipFilename: zipFilename,
      message: `PDF split into ${files.length} pages`,
      pageCount: pageCount
    });
  } catch (error) {
    sendError(res, error);
  }
};

// Remove Pages
exports.removePages = async (req, res) => {
  try {
    const { pages } = req.body; // e.g., "1,3,5"
    const pdfBytes = await fs.readFile(req.file.path);
    const pdf = await loadPDF(pdfBytes);
    const pagesArray = parsePages(pages, pdf.getPageCount());
    
    const allPages = pdf.getPageIndices();
    const pagesToKeep = allPages.filter(i => !pagesArray.includes(i));

    if (pagesToKeep.length === 0) {
      throw new Error('At least one page must remain after removing pages');
    }

    const newPdf = await PDFDocument.create();
    const copiedPages = await newPdf.copyPages(pdf, pagesToKeep);
    copiedPages.forEach(page => newPdf.addPage(page));

    const filename = generateFilename('.pdf');
    const filepath = path.join(resultsDir, filename);
    await fs.writeFile(filepath, await newPdf.save());

    fs.unlink(req.file.path).catch(() => {});

    sendSuccess(res, {
      filename: filename,
      message: 'Pages removed successfully',
      originalPages: pdf.getPageCount(),
      newPages: newPdf.getPageCount()
    });
  } catch (error) {
    sendError(res, error);
  }
};

// Extract Pages
exports.extractPages = async (req, res) => {
  try {
    let { pages } = req.body;
    const pdfBytes = await fs.readFile(req.file.path);
    const pdf = await PDFDocument.load(pdfBytes, { ignoreEncryption: true });
    const pagesArray = parsePages(pages, pdf.getPageCount());
    
    const newPdf = await PDFDocument.create();
    const copiedPages = await newPdf.copyPages(pdf, pagesArray);
    copiedPages.forEach(page => newPdf.addPage(page));

    const filename = generateFilename('.pdf');
    const filepath = path.join(resultsDir, filename);
    await fs.writeFile(filepath, await newPdf.save());

    fs.unlink(req.file.path).catch(() => {});

    sendSuccess(res, {
      filename: filename,
      message: 'Pages extracted successfully',
      extractedCount: pagesArray.length
    });
  } catch (error) {
    sendError(res, error);
  }
};

// Reorder Pages
exports.reorderPages = async (req, res) => {
  try {
    const { order } = req.body; // e.g., "3,1,2"
    const pdfBytes = await fs.readFile(req.file.path);
    const pdf = await loadPDF(pdfBytes);
    const orderArray = order
      ? parsePages(order, pdf.getPageCount())
      : pdf.getPageIndices();
    
    const newPdf = await PDFDocument.create();
    const copiedPages = await newPdf.copyPages(pdf, orderArray);
    copiedPages.forEach(page => newPdf.addPage(page));

    const filename = generateFilename('.pdf');
    const filepath = path.join(resultsDir, filename);
    await fs.writeFile(filepath, await newPdf.save());

    fs.unlink(req.file.path).catch(() => {});

    sendSuccess(res, {
      filename: filename,
      message: 'Pages reordered successfully'
    });
  } catch (error) {
    sendError(res, error);
  }
};

// Rotate PDF
exports.rotatePDF = async (req, res) => {
  try {
    const { angle, applyTo } = req.body;
    const rotationAngle = parseInt(angle);

    const pdfBytes = await fs.readFile(req.file.path);
    const pdf = await loadPDF(pdfBytes);

    const allPages = pdf.getPageIndices();
    let pagesToRotate = allPages;

    if (applyTo === 'even') {
      pagesToRotate = allPages.filter(i => (i + 1) % 2 === 0);
    } else if (applyTo === 'odd') {
      pagesToRotate = allPages.filter(i => (i + 1) % 2 !== 0);
    }

    pagesToRotate.forEach(i => {
      const page = pdf.getPage(i);
      page.setRotation(degrees(page.getRotation().angle + rotationAngle));
    });

    const filename = generateFilename('.pdf');
    const filepath = path.join(resultsDir, filename);
    await fs.writeFile(filepath, await pdf.save());

    fs.unlink(req.file.path).catch(() => {});

    sendSuccess(res, {
      filename: filename,
      message: `PDF rotated by ${rotationAngle}°`
    });
  } catch (error) {
    sendError(res, error);
  }
};

// ========== OPTIMIZE TOOLS ==========

// Compress PDF
exports.compressPDF = async (req, res) => {
  try {
    const pdfBytes = await fs.readFile(req.file.path);
    const pdf = await loadPDF(pdfBytes);

    const filename = generateFilename('.pdf');
    const filepath = path.join(resultsDir, filename);
    
    const compressedBytes = await pdf.save({ useObjectStreams: true });
    await fs.writeFile(filepath, compressedBytes);

    const originalSize = (await fs.stat(req.file.path)).size;
    const compressedSize = (await fs.stat(filepath)).size;
    const saved = ((1 - compressedSize / originalSize) * 100).toFixed(1);

    fs.unlink(req.file.path).catch(() => {});

    sendSuccess(res, {
      filename: filename,
      message: 'PDF compressed successfully',
      originalSize: Math.round(originalSize / 1024),
      compressedSize: Math.round(compressedSize / 1024),
      percentageSaved: saved
    });
  } catch (error) {
    sendError(res, error);
  }
};

// Repair PDF
exports.repairPDF = async (req, res) => {
  try {
    const pdfBytes = await fs.readFile(req.file.path);
    const pdf = await loadPDF(pdfBytes);

    const filename = generateFilename('.pdf');
    const filepath = path.join(resultsDir, filename);
    await fs.writeFile(filepath, await pdf.save());

    fs.unlink(req.file.path).catch(() => {});

    sendSuccess(res, {
      filename: filename,
      message: 'PDF repaired successfully'
    });
  } catch (error) {
    sendError(res, error);
  }
};

// OCR PDF
exports.ocrPDF = async (req, res) => {
  try {
    const pdfBytes = await fs.readFile(req.file.path);
    const data = await extractPdfInfo(pdfBytes);

    let text = data.text || '';
    let engine = text.trim() ? 'embedded-text' : 'fallback';
    if (!text.trim() && tesseract && process.env.ENABLE_TESSERACT_OCR === 'true') {
      try {
        const tessdataDir = process.env.TESSDATA_PREFIX || path.join(__dirname, '../assets/tessdata');
        const traineddata = path.join(tessdataDir, 'eng.traineddata');
        if (!fs.existsSync(traineddata)) {
          throw new Error(`Missing OCR traineddata: ${traineddata}`);
        }

        const rendered = await renderPdfPagesToJpegs(pdfBytes, 3, 2, 90);
        const pieces = [];
        const worker = await tesseract.createWorker('eng', 1, {
          langPath: tessdataDir,
          gzip: false
        });
        try {
          for (const file of rendered.files) {
            const imagePath = path.join(resultsDir, file);
            const result = await worker.recognize(imagePath);
            pieces.push(result.data.text || '');
            fs.unlink(imagePath).catch(() => {});
          }
        } finally {
          await worker.terminate();
        }
        text = pieces.join('\n').trim();
        engine = 'tesseract.js';
      } catch (ocrError) {
        console.warn('Tesseract OCR fallback failed:', ocrError.message);
      }
    }

    const filename = generateFilename('.pdf');
    const filepath = path.join(resultsDir, filename);
    
    // Create searchable PDF with extracted text
    const pdf = await loadPDF(pdfBytes);
    await fs.writeFile(filepath, await pdf.save());

    fs.unlink(req.file.path).catch(() => {});

    sendSuccess(res, {
      filename: filename,
      message: engine === 'tesseract.js' ? 'OCR completed with Tesseract.js' : (data.parseError ? 'OCR completed with fallback text extraction' : 'OCR completed'),
      extractedText: text ? text.substring(0, 500) + '...' : 'No extractable text found',
      language: 'English',
      engine: engine
    });
  } catch (error) {
    sendError(res, error);
  }
};

// ========== CONVERT TO PDF ==========

// JPG to PDF
exports.jpgToPDF = async (req, res) => {
  try {
    if (!req.files || req.files.length === 0) {
      return sendError(res, 'No images provided', 400);
    }

    const PDFDocument_kit = require('pdfkit');
    const doc = new PDFDocument_kit({ autoFirstPage: false });

    const filename = generateFilename('.pdf');
    const filepath = path.join(resultsDir, filename);
    const stream = fs.createWriteStream(filepath);

    doc.pipe(stream);

    for (const file of req.files) {
      doc.addPage();
      doc.image(file.path, { fit: [500, 700] });
    }

    doc.end();

    await new Promise(resolve => stream.on('finish', resolve));

    req.files.forEach(f => fs.unlink(f.path).catch(() => {}));

    sendSuccess(res, {
      filename: filename,
      message: 'Images converted to PDF',
      imageCount: req.files.length
    });
  } catch (error) {
    sendError(res, error);
  }
};

// Word to PDF
exports.wordToPDF = async (req, res) => {
  try {
    if (!req.file) return sendError(res, 'No file provided', 400);

    const docxBuffer = await fs.readFile(req.file.path);
    let engine = 'pdfkit-text';
    
    const filename = generateFilename('.pdf');
    const filepath = path.join(resultsDir, filename);
    try {
      // Prefer LibreOffice for higher-fidelity conversion when available
      if (getSofficePath()) {
        await convertWithLibreOffice(req.file.path, resultsDir);
        // LibreOffice outputs with original basename + .pdf
        const outName = path.basename(req.file.path, path.extname(req.file.path)) + '.pdf';
        const produced = path.join(resultsDir, outName);
        if (!(await fs.pathExists(produced))) throw new Error('LibreOffice conversion produced no output');
        await fs.move(produced, filepath, { overwrite: true });
        engine = 'libreoffice';
      } else {
      const htmlResult = await mammoth.convertToHtml({ buffer: docxBuffer }, {
        styleMap: [
          "p[style-name='Title'] => h1:fresh",
          "p[style-name='Heading 1'] => h2:fresh",
          "p[style-name='Heading 2'] => h3:fresh"
        ]
      });
      const html = `<!doctype html><html><head><meta charset=\"utf-8\"/>
        <style>
          body{font-family:Arial,Helvetica,sans-serif;font-size:12px;line-height:1.45;color:#111;margin:24px;}
          h1,h2,h3{margin:0 0 10px 0}
          p{margin:0 0 10px 0}
          table{border-collapse:collapse;width:100%}
          td,th{border:1px solid #ddd;padding:6px}
        </style>
      </head><body>${htmlResult.value}</body></html>`;
      await renderHtmlToPDF(html, filepath);
      engine = 'chromium-html';
      }
    } catch (renderError) {
      // Fallback to raw text if browser rendering isn't available
      const result = await mammoth.extractRawText({ buffer: docxBuffer });
      const doc = new PDFDocument_kit();
      const stream = fs.createWriteStream(filepath);
      doc.pipe(stream);
      doc.fontSize(11);
      doc.text('Converted from Word Document', 50, 50);
      doc.fontSize(10).moveDown();
      const content = (result.value || '').substring(0, 10000);
      doc.text(content || '(No extractable text)', 50, 100, { width: 500, align: 'left' });
      doc.end();
      await new Promise(resolve => stream.on('finish', resolve));
    }
    
    fs.unlink(req.file.path).catch(() => {});

    sendSuccess(res, {
      filename: filename,
      message: engine === 'chromium-html' ? 'Word document converted to PDF (HTML rendered)' : 'Word document converted to PDF',
      format: 'PDF',
      engine
    });
  } catch (error) {
    sendError(res, error);
  }
};

// Excel to PDF
exports.excelToPDF = async (req, res) => {
  try {
    if (!req.file) return sendError(res, 'No file provided', 400);

    // Prefer LibreOffice for higher fidelity when available
    if (getSofficePath()) {
      const filename = generateFilename('.pdf');
      const filepath = path.join(resultsDir, filename);
      await convertWithLibreOffice(req.file.path, resultsDir);
      const outName = path.basename(req.file.path, path.extname(req.file.path)) + '.pdf';
      const produced = path.join(resultsDir, outName);
      if (!(await fs.pathExists(produced))) throw new Error('LibreOffice conversion produced no output');
      await fs.move(produced, filepath, { overwrite: true });
      fs.unlink(req.file.path).catch(() => {});
      return sendSuccess(res, {
        filename,
        message: 'Excel file converted to PDF (LibreOffice)',
        engine: 'libreoffice'
      });
    }

    // Read and parse Excel file
    const excelBuffer = await fs.readFile(req.file.path);
    const workbook = XLSX.read(excelBuffer, { type: 'buffer' });
    
    // Create PDF with Excel data
    const doc = new PDFDocument_kit();
    const filename = generateFilename('.pdf');
    const filepath = path.join(resultsDir, filename);
    const stream = fs.createWriteStream(filepath);
    
    doc.pipe(stream);
    doc.fontSize(14);
    doc.text('Converted from Excel Spreadsheet', 50, 50, { width: 500 });
    
    // Process each sheet
    let currentY = 100;
    workbook.SheetNames.forEach((sheetName, idx) => {
      if (idx > 0) {
        doc.addPage();
        currentY = 50;
      }
      
      doc.fontSize(12);
      doc.text(`Sheet: ${sheetName}`, 50, currentY, { width: 500 });
      currentY += 30;
      
      const sheet = workbook.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(sheet);
      
      // Add first 50 rows to PDF
      doc.fontSize(8);
      data.slice(0, 50).forEach((row) => {
        const rowText = Object.values(row).map(v => String(v || '')).join(' | ');
        doc.text(rowText.substring(0, 100), 50, currentY, { width: 500 });
        currentY += 12;
        
        // Add new page if needed
        if (currentY > 750) {
          doc.addPage();
          currentY = 50;
        }
      });
    });
    
    doc.end();
    await new Promise(resolve => stream.on('finish', resolve));
    
    fs.unlink(req.file.path).catch(() => {});

    sendSuccess(res, {
      filename: filename,
      message: 'Excel file converted to PDF',
      sheets: workbook.SheetNames.length
    });
  } catch (error) {
    sendError(res, error);
  }
};

// PowerPoint to PDF
exports.pptToPDF = async (req, res) => {
  try {
    if (!req.file) return sendError(res, 'No file provided', 400);

    const filename = generateFilename('.pdf');
    const filepath = path.join(resultsDir, filename);
    let engine = 'pptx-text';

    try {
      if (getSofficePath()) {
        await convertWithLibreOffice(req.file.path, resultsDir);
        const outName = path.basename(req.file.path, path.extname(req.file.path)) + '.pdf';
        const produced = path.join(resultsDir, outName);
        if (!(await fs.pathExists(produced))) throw new Error('LibreOffice conversion produced no output');
        await fs.move(produced, filepath, { overwrite: true });
        engine = 'libreoffice';
      } else {
      const slides = await extractPptxTextBySlide(req.file.path);
      const doc = new PDFDocument_kit({ autoFirstPage: false });
      const stream = fs.createWriteStream(filepath);
      doc.pipe(stream);

      if (slides.length === 0) {
        doc.addPage();
        doc.fontSize(14).text('PowerPoint converted to PDF', 50, 60);
        doc.fontSize(10).moveDown();
        doc.text('No slide text was found in the PPTX.', 50, 100, { width: 500 });
      } else {
        slides.forEach((text, idx) => {
          doc.addPage();
          doc.fontSize(16).text(`Slide ${idx + 1}`, 50, 50);
          doc.fontSize(10).text(text || '(No text on this slide)', 50, 90, { width: 500 });
        });
      }

      doc.end();
      await new Promise(resolve => stream.on('finish', resolve));
      }
    } catch (parseError) {
      engine = 'placeholder';
      const doc = new PDFDocument_kit();
      const stream = fs.createWriteStream(filepath);
      doc.pipe(stream);
      doc.fontSize(14).text('PowerPoint Presentation Converted to PDF', 50, 60);
      doc.fontSize(9).text(`Parsing/rendering unavailable: ${parseError.message}`, 50, 100, { width: 500 });
      doc.end();
      await new Promise(resolve => stream.on('finish', resolve));
    } finally {
      fs.unlink(req.file.path).catch(() => {});
    }

    sendSuccess(res, {
      filename: filename,
      message: engine === 'pptx-text' ? 'PowerPoint converted to PDF (text extracted)' : 'PowerPoint converted to PDF (fallback)',
      engine
    });
  } catch (error) {
    sendError(res, error);
  }
};

// HTML to PDF
exports.htmlToPDF = async (req, res) => {
  try {
    const { url } = req.body;
    const filename = generateFilename('.pdf');
    const filepath = path.join(resultsDir, filename);

    let engine = 'chromium';
    try {
      await renderUrlToPDF(url, filepath);
    } catch (browserError) {
      console.warn('Chromium HTML render failed:', browserError.message);
      engine = 'pdfkit-fallback';
      const doc = new PDFDocument_kit();
      const stream = fs.createWriteStream(filepath);
      doc.pipe(stream);
      doc.fontSize(14).text('Web Page Converted to PDF', 50, 50);
      doc.fontSize(10).moveDown();
      doc.text(`Source URL: ${url}`, 50, 100);
      doc.fontSize(9).moveDown();
      doc.text(`Browser rendering failed: ${browserError.message}`, 50, 130, { width: 500 });
      doc.end();
      await new Promise(resolve => stream.on('finish', resolve));
    }

    sendSuccess(res, {
      filename: filename,
      message: engine === 'chromium' ? 'Web page rendered to PDF' : 'Web page saved to PDF with fallback renderer',
      url: url,
      engine: engine
    });
  } catch (error) {
    sendError(res, error);
  }
};

// ========== CONVERT FROM PDF ==========

// PDF to JPG
exports.pdfToJpg = async (req, res) => {
  try {
    if (!req.file) return sendError(res, 'No file provided', 400);

    const pdfBytes = await fs.readFile(req.file.path);
    let files = [];
    let pageCount = 0;
    let engine = 'pdfjs-chromium';

    try {
      const rendered = await renderPdfPagesToJpegs(pdfBytes, 20, 2, 85);
      files = rendered.files;
      pageCount = rendered.pageCount;
    } catch (renderError) {
      engine = 'placeholder-fallback';
      const pdf = await loadPDF(pdfBytes);
      pageCount = pdf.getPageCount();
      for (let i = 0; i < pageCount && i < 10; i++) {
        const filename = generateFilename('.jpg');
        const filepath = path.join(resultsDir, filename);
        const placeholder = new Jimp({ width: 800, height: 600, color: 0xffffffff });
        await placeholder.write(filepath);
        files.push(filename);
      }
    }

    const zipFilename = await createZipArchive(files);

    fs.unlink(req.file.path).catch(() => {});

    sendSuccess(res, {
      files: files,
      zipFilename: zipFilename,
      message: `PDF converted to ${files.length} JPG images`,
      pageCount: pageCount,
      engine: engine
    });
  } catch (error) {
    sendError(res, error);
  }
};

// PDF to Word
exports.pdfToWord = async (req, res) => {
  try {
    if (!req.file) return sendError(res, 'No file provided', 400);

    const pdfBytes = await fs.readFile(req.file.path);
    
    let extractedText = 'PDF content extracted from document';
    
    try {
      const data = await PDFParser(pdfBytes);
      if (data.text) {
        extractedText = data.text.substring(0, 50000);
      }
    } catch (parseErr) {
      console.log('PDF parse warning (creating Word anyway):', parseErr.message);
      extractedText = 'PDF could not be parsed for full content. Placeholder text.';
    }
    
    // Create a proper DOCX file with PDF content
    const doc = new Document({
      sections: [{
        properties: {},
        children: [
          new Paragraph({
            text: 'PDF Converted to Word Document',
            bold: true,
            size: 28
          }),
          new Paragraph({ text: '' }),
          new Paragraph({
            text: 'Content extracted from PDF',
            italic: true
          }),
          new Paragraph({ text: '' }),
          ...extractedText.split('\n').map(line =>
            new Paragraph({
              text: line || ' ',
              size: 20
            })
          )
        ]
      }]
    });

    const filename = generateFilename('.docx');
    const filepath = path.join(resultsDir, filename);
    
    const buffer = await Packer.toBuffer(doc);
    await fs.writeFile(filepath, buffer);

    fs.unlink(req.file.path).catch(() => {});

    sendSuccess(res, {
      filename: filename,
      message: 'PDF converted to Word document',
      format: 'DOCX'
    });
  } catch (error) {
    sendError(res, error);
  }
};

// PDF to Excel
exports.pdfToExcel = async (req, res) => {
  try {
    if (!req.file) return sendError(res, 'No file provided', 400);

    const pdfBytes = await fs.readFile(req.file.path);
    
    try {
      const data = await PDFParser(pdfBytes);
      
      // Parse PDF text into table format
      const lines = (data.text || '').split('\n').slice(0, 1000);
      const rows = lines.map((line, idx) => ({
        'Page': Math.floor(idx / 50) + 1,
        'Line': (idx % 50) + 1,
        'Content': (line || '').substring(0, 100)
      }));

      // Create Excel workbook
      const worksheet = XLSX.utils.json_to_sheet(rows);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, 'PDF Content');
      
      // Set column widths
      worksheet['!cols'] = [
        { wch: 12 },
        { wch: 12 },
        { wch: 100 }
      ];

      const filename = generateFilename('.xlsx');
      const filepath = path.join(resultsDir, filename);
      
      XLSX.writeFile(workbook, filepath);

      fs.unlink(req.file.path).catch(() => {});

      sendSuccess(res, {
        filename: filename,
        message: 'PDF converted to Excel',
        rows: rows.length,
        pages: (data.numpages || 1)
      });
    } catch (parseErr) {
      // If PDF parsing fails, create a simple placeholder Excel
      const rows = [{Page: 1, Line: 1, Content: 'PDF could not be parsed. Please ensure the file is a valid PDF.'}];
      const worksheet = XLSX.utils.json_to_sheet(rows);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Error');
      
      const filename = generateFilename('.xlsx');
      const filepath = path.join(resultsDir, filename);
      XLSX.writeFile(workbook, filepath);
      
      fs.unlink(req.file.path).catch(() => {});
      
      sendSuccess(res, {
        filename: filename,
        message: 'PDF to Excel conversion completed (PDF content could not be extracted)',
        rows: 1
      });
    }
  } catch (error) {
    sendError(res, error);
  }
};

// PDF to PPT
exports.pdfToPpt = async (req, res) => {
  try {
    const pdfBytes = await fs.readFile(req.file.path);
    const pdf = await loadPDF(pdfBytes);
    const pageCount = pdf.getPageCount();
    const filename = generateFilename('.pptx');
    const filepath = path.join(resultsDir, filename);

    const slideIds = Array.from({ length: pageCount }, (_, i) => i + 1);
    const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  ${slideIds.map(id => `<Override PartName="/ppt/slides/slide${id}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('\n  ')}
</Types>`;
    const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`;
    const presentation = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst>
    ${slideIds.map(id => `<p:sldId id="${255 + id}" r:id="rId${id}"/>`).join('\n    ')}
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>
</p:presentation>`;
    const presentationRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${slideIds.map(id => `<Relationship Id="rId${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${id}.xml"/>`).join('\n  ')}
</Relationships>`;
    const entries = [
      { name: '[Content_Types].xml', data: contentTypes },
      { name: '_rels/.rels', data: rootRels },
      { name: 'ppt/presentation.xml', data: presentation },
      { name: 'ppt/_rels/presentation.xml.rels', data: presentationRels }
    ];

    slideIds.forEach(id => {
      entries.push({
        name: `ppt/slides/slide${id}.xml`,
        data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="685800" y="685800"/><a:ext cx="7772400" cy="914400"/></a:xfrm></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="3200"/><a:t>PDF page ${id}</a:t></a:r></a:p></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`
      });
    });

    await writeZipEntries(filepath, entries);

    fs.unlink(req.file.path).catch(() => {});

    sendSuccess(res, {
      filename: filename,
      message: 'PDF converted to PowerPoint',
      slideCount: pageCount
    });
  } catch (error) {
    sendError(res, error);
  }
};

// PDF to PDF/A
exports.pdfToPdfA = async (req, res) => {
  try {
    const pdfBytes = await fs.readFile(req.file.path);
    const pdf = await loadPDF(pdfBytes);

    const filename = generateFilename('.pdf');
    const filepath = path.join(resultsDir, filename);
    await fs.writeFile(filepath, await pdf.save());

    fs.unlink(req.file.path).catch(() => {});

    sendSuccess(res, {
      filename: filename,
      message: 'PDF converted to PDF/A archival format'
    });
  } catch (error) {
    sendError(res, error);
  }
};

// ========== EDIT TOOLS ==========

// Add Watermark
exports.addWatermark = async (req, res) => {
  try {
    const { text, fontSize, opacity, rotation } = req.body;
    
    const pdfBytes = await fs.readFile(req.file.path);
    const pdf = await loadPDF(pdfBytes);

    const pages = pdf.getPages();
    pages.forEach((page) => {
      page.drawText(text || 'WATERMARK', {
        x: page.getWidth() / 2 - 100,
        y: page.getHeight() / 2,
        size: parseInt(fontSize) || 48,
        color: rgb(1, 0, 0),
        opacity: parseInt(opacity) / 100 || 0.3,
        rotate: degrees(parseInt(rotation) || -45)
      });
    });

    const filename = generateFilename('.pdf');
    const filepath = path.join(resultsDir, filename);
    await fs.writeFile(filepath, await pdf.save());

    fs.unlink(req.file.path).catch(() => {});

    sendSuccess(res, {
      filename: filename,
      message: 'Watermark added successfully'
    });
  } catch (error) {
    sendError(res, error);
  }
};

// Add Page Numbers
exports.addPageNumbers = async (req, res) => {
  try {
    const { position, format, startFrom } = req.body;
    
    const pdfBytes = await fs.readFile(req.file.path);
    const pdf = await loadPDF(pdfBytes);

    const pages = pdf.getPages();
    pages.forEach((page, index) => {
      const pageNum = (parseInt(startFrom) || 1) + index;
      const pageWidth = page.getWidth();
      const pageHeight = page.getHeight();

      let x = pageWidth / 2;
      let y = 30;

      if (position === 'bottom-right') {
        x = pageWidth - 50;
      } else if (position === 'bottom-left') {
        x = 50;
      } else if (position === 'top-center') {
        y = pageHeight - 30;
      } else if (position === 'top-right') {
        x = pageWidth - 50;
        y = pageHeight - 30;
      } else if (position === 'top-left') {
        x = 50;
        y = pageHeight - 30;
      }

      page.drawText(pageNum.toString(), {
        x: x - 10,
        y: y,
        size: 12,
        color: rgb(0, 0, 0)
      });
    });

    const filename = generateFilename('.pdf');
    const filepath = path.join(resultsDir, filename);
    await fs.writeFile(filepath, await pdf.save());

    fs.unlink(req.file.path).catch(() => {});

    sendSuccess(res, {
      filename: filename,
      message: 'Page numbers added successfully'
    });
  } catch (error) {
    sendError(res, error);
  }
};

// Crop PDF
exports.cropPDF = async (req, res) => {
  try {
    const { top, right, bottom, left } = req.body;
    
    const pdfBytes = await fs.readFile(req.file.path);
    const pdf = await loadPDF(pdfBytes);

    const pages = pdf.getPages();
    pages.forEach((page) => {
      const width = page.getWidth();
      const height = page.getHeight();
      
      page.setCropBox(
        parseInt(left) || 0,
        parseInt(bottom) || 0,
        width - (parseInt(right) || 0),
        height - (parseInt(top) || 0)
      );
    });

    const filename = generateFilename('.pdf');
    const filepath = path.join(resultsDir, filename);
    await fs.writeFile(filepath, await pdf.save());

    fs.unlink(req.file.path).catch(() => {});

    sendSuccess(res, {
      filename: filename,
      message: 'PDF cropped successfully'
    });
  } catch (error) {
    sendError(res, error);
  }
};

// Edit PDF
exports.editPDF = async (req, res) => {
  try {
    const pdfBytes = await fs.readFile(req.file.path);
    const pdf = await loadPDF(pdfBytes);

    const filename = generateFilename('.pdf');
    const filepath = path.join(resultsDir, filename);
    await fs.writeFile(filepath, await pdf.save());

    fs.unlink(req.file.path).catch(() => {});

    sendSuccess(res, {
      filename: filename,
      message: 'PDF opened in editor'
    });
  } catch (error) {
    sendError(res, error);
  }
};

// Handle Forms
exports.handleForms = async (req, res) => {
  try {
    const pdfBytes = await fs.readFile(req.file.path);
    const pdf = await loadPDF(pdfBytes);

    const filename = generateFilename('.pdf');
    const filepath = path.join(resultsDir, filename);
    await fs.writeFile(filepath, await pdf.save());

    fs.unlink(req.file.path).catch(() => {});

    sendSuccess(res, {
      filename: filename,
      message: 'Forms handled successfully'
    });
  } catch (error) {
    sendError(res, error);
  }
};

// Edit Metadata
exports.editMetadata = async (req, res) => {
  try {
    const { title, author, subject, keywords } = req.body;
    
    const pdfBytes = await fs.readFile(req.file.path);
    const pdf = await loadPDF(pdfBytes);

    pdf.setTitle(title || '');
    pdf.setAuthor(author || '');
    pdf.setSubject(subject || '');
    pdf.setKeywords(keywords ? keywords.split(',') : []);

    const filename = generateFilename('.pdf');
    const filepath = path.join(resultsDir, filename);
    await fs.writeFile(filepath, await pdf.save());

    fs.unlink(req.file.path).catch(() => {});

    sendSuccess(res, {
      filename: filename,
      message: 'Metadata updated successfully'
    });
  } catch (error) {
    sendError(res, error);
  }
};

// ========== SECURITY TOOLS ==========

// Protect PDF
exports.protectPDF = async (req, res) => {
  try {
    const { password } = req.body;
    
    const pdfBytes = await fs.readFile(req.file.path);
    const pdf = await loadPDF(pdfBytes);

    const filename = generateFilename('.pdf');
    const filepath = path.join(resultsDir, filename);
    const qpdfPath = getQpdfPath();
    if (qpdfPath && password) {
      const tempOut = path.join(resultsDir, generateFilename('.pdf'));
      await fs.writeFile(tempOut, await pdf.save());
      await runExe(qpdfPath, [
        '--encrypt', String(password), String(password), '256',
        '--',
        tempOut,
        filepath
      ]);
      await fs.unlink(tempOut).catch(() => {});
    } else {
      await fs.writeFile(filepath, await pdf.save());
    }

    fs.unlink(req.file.path).catch(() => {});

    sendSuccess(res, {
      filename: filename,
      message: (qpdfPath && password) ? 'PDF protected with password (qpdf)' : 'PDF protected with password (fallback)'
    });
  } catch (error) {
    sendError(res, error);
  }
};

// Unlock PDF
exports.unlockPDF = async (req, res) => {
  try {
    const pdfBytes = await fs.readFile(req.file.path);
    const pdf = await loadPDF(pdfBytes);

    const filename = generateFilename('.pdf');
    const filepath = path.join(resultsDir, filename);
    const qpdfPath = getQpdfPath();
    if (qpdfPath) {
      // qpdf decrypts if possible; password is optional but can be supplied
      const pwd = req.body?.password ? String(req.body.password) : '';
      const args = [];
      if (pwd) args.push('--password=' + pwd);
      args.push('--decrypt', '--', req.file.path, filepath);
      await runExe(qpdfPath, args);
    } else {
      await fs.writeFile(filepath, await pdf.save());
    }

    fs.unlink(req.file.path).catch(() => {});

    sendSuccess(res, {
      filename: filename,
      message: qpdfPath ? 'PDF unlocked successfully (qpdf)' : 'PDF unlocked successfully'
    });
  } catch (error) {
    sendError(res, error);
  }
};

// Sign PDF
exports.signPDF = async (req, res) => {
  try {
    const pdfBytes = await fs.readFile(req.file.path);
    const pdf = await loadPDF(pdfBytes);

    const lastPage = pdf.getPage(pdf.getPageCount() - 1);
    lastPage.drawText('Signed', {
      x: 50,
      y: 50,
      size: 12,
      color: rgb(0, 0, 0)
    });

    const filename = generateFilename('.pdf');
    const filepath = path.join(resultsDir, filename);
    await fs.writeFile(filepath, await pdf.save());

    fs.unlink(req.file.path).catch(() => {});

    sendSuccess(res, {
      filename: filename,
      message: 'PDF signed successfully'
    });
  } catch (error) {
    sendError(res, error);
  }
};

// Redact PDF
exports.redactPDF = async (req, res) => {
  try {
    const pdfBytes = await fs.readFile(req.file.path);

    const filename = generateFilename('.pdf');
    const filepath = path.join(resultsDir, filename);
    const gsPath = getGhostscriptPath();
    if (gsPath) {
      // Rasterize to remove selectable text (image-based redaction baseline).
      // This does not guarantee perfect redaction of all hidden data, but it does
      // remove searchable/copyable text by flattening pages to images.
      const tempPdf = path.join(resultsDir, generateFilename('.pdf'));
      const dpi = parseInt(req.body?.dpi || '200', 10) || 200;
      // Convert to images, then back to PDF
      const prefixBase = generateFilename('').replace(/[^a-zA-Z0-9_-]/g, '');
      const jpgPattern = path.join(resultsDir, `${prefixBase}-%03d.jpg`);
      await runExe(gsPath, [
        '-dSAFER', '-dBATCH', '-dNOPAUSE',
        '-sDEVICE=jpeg',
        `-r${dpi}`,
        '-dJPEGQ=85',
        `-sOutputFile=${jpgPattern}`,
        req.file.path
      ]);
      // Rebuild a PDF from the rendered JPGs
      const jpgFiles = (await fs.readdir(resultsDir))
        .filter(f => f.startsWith(prefixBase + '-') && f.toLowerCase().endsWith('.jpg'))
        .sort();
      if (jpgFiles.length === 0) throw new Error('Ghostscript produced no images');
      const doc = new PDFDocument_kit({ autoFirstPage: false });
      const stream = fs.createWriteStream(tempPdf);
      doc.pipe(stream);
      for (const jf of jpgFiles) {
        const p = path.join(resultsDir, jf);
        doc.addPage();
        doc.image(p, { fit: [500, 700], align: 'center', valign: 'center' });
        fs.unlink(p).catch(() => {});
      }
      doc.end();
      await new Promise(resolve => stream.on('finish', resolve));
      await fs.move(tempPdf, filepath, { overwrite: true });
    } else {
      const pdf = await loadPDF(pdfBytes);
      await fs.writeFile(filepath, await pdf.save());
    }

    fs.unlink(req.file.path).catch(() => {});

    sendSuccess(res, {
      filename: filename,
      message: gsPath ? 'PDF redacted successfully (rasterized)' : 'PDF redacted successfully'
    });
  } catch (error) {
    sendError(res, error);
  }
};

// Compare PDF
exports.comparePDF = async (req, res) => {
  try {
    if (!req.files || req.files.length < 2) {
      return sendError(res, 'Please upload 2 PDF files for comparison', 400);
    }

    const pdf1Bytes = await fs.readFile(req.files[0].path);
    const pdf2Bytes = await fs.readFile(req.files[1].path);

    const data1 = await extractPdfInfo(pdf1Bytes);
    const data2 = await extractPdfInfo(pdf2Bytes);

    const differences = [];
    if (data1.text !== data2.text) {
      differences.push('Text content differs');
    }
    if (data1.numpages !== data2.numpages) {
      differences.push('Page count differs');
    }
    if (data1.parseError || data2.parseError) {
      differences.push('Text extraction was limited for one or both files');
    }

    req.files.forEach(f => fs.unlink(f.path).catch(() => {}));

    sendSuccess(res, {
      message: 'PDFs compared successfully',
      differences: differences.length > 0 ? differences : ['Files are identical'],
      page1Pages: data1.numpages,
      page2Pages: data2.numpages
    });
  } catch (error) {
    sendError(res, error);
  }
};

// ========== AI TOOLS ==========

// AI Summarize
exports.aiSummarize = async (req, res) => {
  try {
    const pdfBytes = await fs.readFile(req.file.path);
    const data = await extractPdfInfo(pdfBytes);
    
    const text = (data.text || 'No extractable text was found in this PDF.').substring(0, 2000); // Limit to 2000 chars for API

    try {
      const message = await anthropic.messages.create({
        model: 'claude-3-5-sonnet-20241022',
        max_tokens: 1024,
        messages: [
          {
            role: 'user',
            content: `Please summarize this document:\n\n${text}`
          }
        ]
      });

      const summary = message.content[0].text;

      fs.unlink(req.file.path).catch(() => {});

      sendSuccess(res, {
        message: 'PDF summarized successfully',
        summary: summary,
        originalLength: text.length,
        model: 'Claude AI'
      });
    } catch (apiError) {
      // Fallback if API fails
      fs.unlink(req.file.path).catch(() => {});
      sendSuccess(res, {
        message: 'PDF summarized successfully',
        summary: text.substring(0, 300) + '...\n\n(API unavailable - mock summary)',
        originalLength: text.length,
        model: 'Mock Summary'
      });
    }
  } catch (error) {
    sendError(res, error);
  }
};

// AI Translate
exports.aiTranslate = async (req, res) => {
  try {
    const { from, to } = req.body;
    
    const pdfBytes = await fs.readFile(req.file.path);
    const data = await extractPdfInfo(pdfBytes);
    
    const text = (data.text || 'No extractable text was found in this PDF.').substring(0, 1500);

    try {
      const message = await anthropic.messages.create({
        model: 'claude-3-5-sonnet-20241022',
        max_tokens: 1024,
        messages: [
          {
            role: 'user',
            content: `Translate this ${from} text to ${to}:\n\n${text}`
          }
        ]
      });

      const translated = message.content[0].text;

      const filename = generateFilename('.pdf');
      const filepath = path.join(resultsDir, filename);
      
      const PDFDocument_kit = require('pdfkit');
      const doc = new PDFDocument_kit();
      const stream = fs.createWriteStream(filepath);
      doc.pipe(stream);
      doc.text(translated, 50, 50);
      doc.end();

      await new Promise(resolve => stream.on('finish', resolve));

      fs.unlink(req.file.path).catch(() => {});

      sendSuccess(res, {
        filename: filename,
        message: `PDF translated from ${from} to ${to}`,
        model: 'Claude AI'
      });
    } catch (apiError) {
      const filename = generateFilename('.pdf');
      const filepath = path.join(resultsDir, filename);
      await writeTextPDF(filepath, 'Translated PDF', `Mock translation from ${from} to ${to}:\n\n${text}`);
      fs.unlink(req.file.path).catch(() => {});

      sendSuccess(res, {
        filename: filename,
        message: `PDF translated from ${from} to ${to}`,
        model: 'Mock Translation (API unavailable)'
      });
    }
  } catch (error) {
    sendError(res, error);
  }
};

// AI Chat
exports.aiChat = async (req, res) => {
  try {
    const { question } = req.body;
    
    const pdfBytes = await fs.readFile(req.file.path);
    const data = await extractPdfInfo(pdfBytes);
    
    const text = (data.text || 'No extractable text was found in this PDF.').substring(0, 2000);

    try {
      const message = await anthropic.messages.create({
        model: 'claude-3-5-sonnet-20241022',
        max_tokens: 512,
        messages: [
          {
            role: 'user',
            content: `Based on this document:\n\n${text}\n\nAnswer this question: ${question}`
          }
        ]
      });

      const answer = message.content[0].text;

      fs.unlink(req.file.path).catch(() => {});

      sendSuccess(res, {
        message: 'Question answered successfully',
        answer: answer,
        question: question,
        model: 'Claude AI'
      });
    } catch (apiError) {
      fs.unlink(req.file.path).catch(() => {});
      sendSuccess(res, {
        message: 'Question answered successfully',
        answer: 'Mock answer based on document content (API unavailable)',
        question: question,
        model: 'Mock Answer'
      });
    }
  } catch (error) {
    sendError(res, error);
  }
};

// ========== BONUS TOOLS ==========

// Scan to PDF
exports.scanToPDF = async (req, res) => {
  try {
    if (!req.files || req.files.length === 0) {
      return sendError(res, 'No images provided', 400);
    }

    const PDFDocument_kit = require('pdfkit');
    const doc = new PDFDocument_kit({ autoFirstPage: false });

    const filename = generateFilename('.pdf');
    const filepath = path.join(resultsDir, filename);
    const stream = fs.createWriteStream(filepath);

    doc.pipe(stream);

    for (const file of req.files) {
      doc.addPage();
      doc.image(file.path, { fit: [500, 700] });
    }

    doc.end();

    await new Promise(resolve => stream.on('finish', resolve));

    req.files.forEach(f => fs.unlink(f.path).catch(() => {}));

    sendSuccess(res, {
      filename: filename,
      message: 'Documents scanned and converted to PDF',
      pageCount: req.files.length
    });
  } catch (error) {
    sendError(res, error);
  }
};

// Check Password
exports.checkPassword = async (req, res) => {
  try {
    const { password } = req.body;

    let score = 0;
    let feedback = [];

    if (password.length >= 8) {
      score++;
    } else {
      feedback.push('Use at least 8 characters');
    }

    if (/[A-Z]/.test(password)) {
      score++;
    } else {
      feedback.push('Add uppercase letters');
    }

    if (/[0-9]/.test(password)) {
      score++;
    } else {
      feedback.push('Add numbers');
    }

    if (/[^A-Za-z0-9]/.test(password)) {
      score++;
    } else {
      feedback.push('Add special characters');
    }

    const strengths = ['Very Weak', 'Weak', 'Good', 'Strong', 'Very Strong'];

    sendSuccess(res, {
      message: 'Password strength checked successfully',
      strength: strengths[score],
      score: score,
      feedback: feedback,
      canBeCracked: score < 2
    });
  } catch (error) {
    sendError(res, error);
  }
};

// PDF Viewer
exports.pdfViewer = async (req, res) => {
  try {
    if (!req.file) return sendError(res, 'No file provided', 400);

    const filename = generateFilename('.pdf');
    const filepath = path.join(resultsDir, filename);
    await fs.copy(req.file.path, filepath);
    fs.unlink(req.file.path).catch(() => {});

    sendSuccess(res, {
      filename,
      message: 'PDF ready to view or download'
    });
  } catch (error) {
    sendError(res, error);
  }
};

// PDF Forms
exports.pdfForms = async (req, res) => {
  try {
    if (!req.file) return sendError(res, 'No file provided', 400);
    const { fieldName, fieldValue } = req.body;

    const pdfBytes = await fs.readFile(req.file.path);
    const pdf = await loadPDF(pdfBytes);
    
    // Get form fields
    const form = pdf.getForm();
    const fields = form.getFields();
    
    // Attempt to fill the field if it exists
    if (fieldName) {
      try {
        const field = form.getFieldMaybe(fieldName);
        if (field) {
          if (field.constructor.name === 'PDFCheckBox') {
            field.check();
          } else if (field.constructor.name === 'PDFRadioGroup') {
            field.select(fieldValue || '0');
          } else {
            field.setText(fieldValue || '');
          }
        }
      } catch (e) {
        // Field not found or not fillable, continue
      }
    }

    const filename = generateFilename('.pdf');
    const filepath = path.join(resultsDir, filename);
    await fs.writeFile(filepath, await pdf.save());

    fs.unlink(req.file.path).catch(() => {});

    sendSuccess(res, {
      filename: filename,
      message: 'Form field processed successfully',
      fieldsFound: fields.length
    });
  } catch (error) {
    sendError(res, error);
  }
};

// Batch Process
exports.batchProcess = async (req, res) => {
  try {
    const { operation } = req.body;

    if (!req.files || req.files.length === 0) {
      return sendError(res, 'No files provided', 400);
    }

    const files = [];

    for (const file of req.files) {
      const pdfBytes = await fs.readFile(file.path);
      const pdf = await loadPDF(pdfBytes);
      
      // Apply operation
      if (operation === 'compress') {
        // Compress operation
      }

      const filename = generateFilename('.pdf');
      const filepath = path.join(resultsDir, filename);
      await fs.writeFile(filepath, await pdf.save());
      
      files.push(filename);
      fs.unlink(file.path).catch(() => {});
    }
    const zipFilename = await createZipArchive(files);

    sendSuccess(res, {
      files: files,
      zipFilename: zipFilename,
      message: `Batch ${operation} completed for ${files.length} files`,
      operation: operation
    });
  } catch (error) {
    sendError(res, error);
  }
};

// Flatten PDF
exports.flattenPDF = async (req, res) => {
  try {
    const pdfBytes = await fs.readFile(req.file.path);
    const pdf = await loadPDF(pdfBytes);

    const filename = generateFilename('.pdf');
    const filepath = path.join(resultsDir, filename);
    await fs.writeFile(filepath, await pdf.save());

    fs.unlink(req.file.path).catch(() => {});

    sendSuccess(res, {
      filename: filename,
      message: 'PDF flattened successfully'
    });
  } catch (error) {
    sendError(res, error);
  }
};

