const express = require('express');
const multer = require('multer');
const cors = require('cors');
const path = require('path');
const fs = require('fs-extra');
const { v4: uuidv4 } = require('uuid');
const dotenv = require('dotenv');

dotenv.config();

const app = express();
const PORT = process.env.PORT || 5000;

// ========== MIDDLEWARE ==========
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname)));

// Create uploads and results directories
const uploadsDir = path.join(__dirname, 'uploads');
const resultsDir = path.join(__dirname, 'results');
fs.ensureDirSync(uploadsDir);
fs.ensureDirSync(resultsDir);

// File upload configuration
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, uploadsDir);
  },
  filename: (req, file, cb) => {
    const uuid = uuidv4();
    cb(null, uuid + path.extname(file.originalname));
  }
});

const upload = multer({
  storage: storage,
  limits: { fileSize: 100 * 1024 * 1024 } // 100MB
});

const formParser = multer();

// ========== IMPORT TOOL CONTROLLERS ==========
const toolControllers = require('./controllers/toolControllers');

// ========== UNIVERSAL FORM DATA PARSER ==========
// This middleware uses multer to parse both files and fields, then normalizes
// single-file uploads so controllers can safely use req.file.
const normalizeUpload = (req, res, next) => {
  if (!req.file && Array.isArray(req.files) && req.files.length > 0) {
    req.file = req.files.find(file => file.fieldname === 'file') || req.files[0];
  }
  next();
};
const parseAllFormData = [upload.any(), normalizeUpload];

// ========== MIDDLEWARE FOR MERGING MULTER FIELDS ==========
// This middleware ensures form fields are properly accessible even with file uploads
app.use((req, res, next) => {
  // For multipart/form-data, multer should have already parsed it
  // Just ensure body is an object
  if (!req.body) req.body = {};
  next();
});

// ========== API ROUTES ==========

// Health check
app.get('/api/health', (req, res) => {
  res.json({ status: 'OK', message: 'PDFly Pro Backend Running' });
});

// Organize Tools
app.post('/api/tools/merge', upload.array('files', 20), toolControllers.mergePDF);
app.post('/api/tools/split', parseAllFormData, toolControllers.splitPDF);
app.post('/api/tools/remove-pages', parseAllFormData, toolControllers.removePages);
app.post('/api/tools/extract', parseAllFormData, toolControllers.extractPages);
app.post('/api/tools/reorder', parseAllFormData, toolControllers.reorderPages);
app.post('/api/tools/rotate', parseAllFormData, toolControllers.rotatePDF);

// Optimize Tools
app.post('/api/tools/compress', parseAllFormData, toolControllers.compressPDF);
app.post('/api/tools/repair', parseAllFormData, toolControllers.repairPDF);
app.post('/api/tools/ocr', parseAllFormData, toolControllers.ocrPDF);

// Convert To PDF
app.post('/api/tools/jpg2pdf', upload.array('files', 50), toolControllers.jpgToPDF);
app.post('/api/tools/word2pdf', parseAllFormData, toolControllers.wordToPDF);
app.post('/api/tools/excel2pdf', parseAllFormData, toolControllers.excelToPDF);
app.post('/api/tools/ppt2pdf', parseAllFormData, toolControllers.pptToPDF);
app.post('/api/tools/html2pdf', parseAllFormData, toolControllers.htmlToPDF);

// Convert From PDF
app.post('/api/tools/pdf2jpg', parseAllFormData, toolControllers.pdfToJpg);
app.post('/api/tools/pdf2word', parseAllFormData, toolControllers.pdfToWord);
app.post('/api/tools/pdf2excel', parseAllFormData, toolControllers.pdfToExcel);
app.post('/api/tools/pdf2ppt', parseAllFormData, toolControllers.pdfToPpt);
app.post('/api/tools/pdf2pdfa', parseAllFormData, toolControllers.pdfToPdfA);

// Edit Tools
app.post('/api/tools/watermark', parseAllFormData, toolControllers.addWatermark);
app.post('/api/tools/pagenumbers', parseAllFormData, toolControllers.addPageNumbers);
app.post('/api/tools/crop', parseAllFormData, toolControllers.cropPDF);
app.post('/api/tools/edit', parseAllFormData, toolControllers.editPDF);
app.post('/api/tools/forms', parseAllFormData, toolControllers.handleForms);
app.post('/api/tools/pdfforms', parseAllFormData, toolControllers.pdfForms);
app.post('/api/tools/metadata', parseAllFormData, toolControllers.editMetadata);

// Security Tools
app.post('/api/tools/protect', parseAllFormData, toolControllers.protectPDF);
app.post('/api/tools/unlock', parseAllFormData, toolControllers.unlockPDF);
app.post('/api/tools/sign', parseAllFormData, toolControllers.signPDF);
app.post('/api/tools/redact', parseAllFormData, toolControllers.redactPDF);
app.post('/api/tools/compare', upload.array('files', 2), toolControllers.comparePDF);

// AI Tools
app.post('/api/tools/ai-summarize', parseAllFormData, toolControllers.aiSummarize);
app.post('/api/tools/ai-translate', parseAllFormData, toolControllers.aiTranslate);
app.post('/api/tools/ai-chat', parseAllFormData, toolControllers.aiChat);

// Bonus Tools
app.post('/api/tools/scan2pdf', upload.array('files', 50), toolControllers.scanToPDF);
app.post('/api/tools/pdf-viewer', parseAllFormData, toolControllers.pdfViewer);
app.post('/api/tools/pwd-checker', formParser.none(), toolControllers.checkPassword);
app.post('/api/tools/batch', upload.array('files', 100), toolControllers.batchProcess);
app.post('/api/tools/flatten', parseAllFormData, toolControllers.flattenPDF);

// Download endpoint
app.get('/api/download/:filename', (req, res) => {
  const filepath = path.join(resultsDir, req.params.filename);
  if (fs.existsSync(filepath)) {
    res.download(filepath, () => {
      // Delete after download (2 hour cleanup with scheduled job is better)
      setTimeout(() => {
        fs.unlink(filepath, () => {});
      }, 5000);
    });
  } else {
    res.status(404).json({ error: 'File not found' });
  }
});

// Serve index.html
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

// ========== ERROR HANDLING ==========
app.use((err, req, res, next) => {
  if (err instanceof multer.MulterError) {
    // Suppress multer field name errors
    if (err.code === 'LIMIT_PART_COUNT' || err.message.includes('Field name')) {
      return res.status(400).json({ error: 'Invalid file upload format' });
    }
  }
  // Log other errors
  console.error(err);
  res.status(500).json({ error: 'Server error' });
});

// ========== START SERVER ==========
app.listen(PORT, '0.0.0.0', () => {
  console.log(`✅ PDFly Pro Backend running on http://0.0.0.0:${PORT}`);
  console.log(`📁 Uploads directory: ${uploadsDir}`);
  console.log(`📁 Results directory: ${resultsDir}`);
});

// Cleanup old files every hour
setInterval(() => {
  const now = Date.now();
  const maxAge = 2 * 60 * 60 * 1000; // 2 hours
  
  fs.readdirSync(resultsDir).forEach(file => {
    const filepath = path.join(resultsDir, file);
    const stats = fs.statSync(filepath);
    if (now - stats.mtimeMs > maxAge) {
      fs.unlinkSync(filepath);
    }
  });

  fs.readdirSync(uploadsDir).forEach(file => {
    const filepath = path.join(uploadsDir, file);
    const stats = fs.statSync(filepath);
    if (now - stats.mtimeMs > maxAge) {
      fs.unlinkSync(filepath);
    }
  });
}, 60 * 60 * 1000);

module.exports = app;



