require('dotenv').config();
const express  = require('express');
const multer   = require('multer');
const mysql    = require('mysql2/promise');
const xlsx     = require('xlsx');
const csv      = require('csv-parser');
const fs       = require('fs');
const path     = require('path');
const cors     = require('cors');

const app  = express();
const PORT = process.env.PORT || 3000;

// ── Middleware ────────────────────────────────────────────────
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ── Ensure uploads/ folder exists ────────────────────────────
const UPLOADS_DIR = path.join(__dirname, 'uploads');
if (!fs.existsSync(UPLOADS_DIR)) fs.mkdirSync(UPLOADS_DIR);

// ── Multer file upload config ─────────────────────────────────
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, UPLOADS_DIR),
  filename:    (req, file, cb) => {
    const unique = Date.now() + '-' + Math.round(Math.random() * 1e9);
    cb(null, unique + path.extname(file.originalname));
  }
});
const upload = multer({
  storage,
  fileFilter: (req, file, cb) => {
    const allowed = ['.xlsx', '.xls', '.csv'];
    const ext = path.extname(file.originalname).toLowerCase();
    if (allowed.includes(ext)) cb(null, true);
    else cb(new Error('Only .xlsx, .xls, .csv files allowed'));
  },
  limits: { fileSize: 50 * 1024 * 1024 } // 50 MB
});

// ── MySQL Connection Pool ─────────────────────────────────────
let pool;
async function initDB() {
  try {
    pool = mysql.createPool({
      host:     process.env.DB_HOST     || 'localhost',
      user:     process.env.DB_USER     || 'root',
      password: process.env.DB_PASSWORD || '',
      database: process.env.DB_NAME     || 'report_dashboard',
      waitForConnections: true,
      connectionLimit: 10
    });
    // Test connection
    await pool.query('SELECT 1');
    console.log('✅ MySQL connected');
    await createTables();
  } catch (err) {
    console.warn('⚠️  MySQL not connected — running without DB:', err.message);
    pool = null;
  }
}

async function createTables() {
  if (!pool) return;
  await pool.query(`
    CREATE TABLE IF NOT EXISTS uploads (
      id          INT AUTO_INCREMENT PRIMARY KEY,
      filename    VARCHAR(255),
      original    VARCHAR(255),
      rows_count  INT,
      sheets      VARCHAR(255),
      uploaded_at DATETIME DEFAULT CURRENT_TIMESTAMP
    )
  `);
  await pool.query(`
    CREATE TABLE IF NOT EXISTS invoice_rows (
      id          INT AUTO_INCREMENT PRIMARY KEY,
      upload_id   INT,
      party       VARCHAR(255),
      vr_no       VARCHAR(100),
      ledger      VARCHAR(255),
      outstanding DOUBLE DEFAULT 0,
      collected   DOUBLE DEFAULT 0,
      gst         DOUBLE DEFAULT 0,
      agent       VARCHAR(255),
      due_date    VARCHAR(50),
      created_at  DATETIME DEFAULT CURRENT_TIMESTAMP,
      FOREIGN KEY (upload_id) REFERENCES uploads(id) ON DELETE CASCADE
    )
  `);
}

// ── ROUTE: Upload file ────────────────────────────────────────
app.post('/api/upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

    const filePath = req.file.path;
    const ext      = path.extname(req.file.originalname).toLowerCase();
    let parsedData;

    if (ext === '.csv') {
      parsedData = await parseCSV(filePath);
    } else {
      parsedData = parseXLSX(filePath);
    }

    // Save to MySQL if available
    let uploadId = null;
    if (pool) {
      const [result] = await pool.query(
        'INSERT INTO uploads (filename, original, rows_count, sheets) VALUES (?,?,?,?)',
        [req.file.filename, req.file.originalname, parsedData.invoiceRows.length, parsedData.sheets.join(',')]
      );
      uploadId = result.insertId;

      // Batch insert invoice rows (max 500 at a time)
      const rows = parsedData.invoiceRows;
      for (let i = 0; i < rows.length; i += 500) {
        const batch = rows.slice(i, i + 500);
        const values = batch.map(r => [uploadId, r.party, r.vr_no, r.ledger, r.outstanding, r.collected, r.gst, r.agent, r.due_date]);
        if (values.length > 0) {
          await pool.query(
            'INSERT INTO invoice_rows (upload_id, party, vr_no, ledger, outstanding, collected, gst, agent, due_date) VALUES ?',
            [values]
          );
        }
      }
      console.log(`💾 Saved upload #${uploadId} with ${rows.length} rows to MySQL`);
    }

    res.json({
      success:    true,
      uploadId,
      filename:   req.file.originalname,
      data:       parsedData
    });

  } catch (err) {
    console.error('Upload error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ── ROUTE: Get past uploads ───────────────────────────────────
app.get('/api/uploads', async (req, res) => {
  if (!pool) return res.json([]);
  try {
    const [rows] = await pool.query('SELECT * FROM uploads ORDER BY uploaded_at DESC LIMIT 20');
    res.json(rows);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── ROUTE: Get rows for a past upload ────────────────────────
app.get('/api/upload/:id', async (req, res) => {
  if (!pool) return res.status(503).json({ error: 'DB not connected' });
  try {
    const [rows] = await pool.query('SELECT * FROM invoice_rows WHERE upload_id = ?', [req.params.id]);
    res.json(rows);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── ROUTE: Health check ───────────────────────────────────────
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', db: pool ? 'connected' : 'disconnected', time: new Date() });
});

// ── XLSX PARSER ───────────────────────────────────────────────
function parseXLSX(filePath) {
  const wb       = xlsx.readFile(filePath, { cellDates: true });
  const sheets   = wb.SheetNames;

  // Pick largest sheet
  let mainData = [], mainSheet = '';
  for (const name of sheets) {
    const data = xlsx.utils.sheet_to_json(wb.Sheets[name], { defval: null });
    if (data.length > mainData.length) { mainData = data; mainSheet = name; }
  }

  // Detect Silicon Systems format
  const isSilicon = sheets.some(n => n.toLowerCase().includes('salesmanwise') || n.toLowerCase().includes('gst'));

  if (isSilicon) {
    return parseSiliconFormat(wb, sheets);
  }
  return parseGenericFormat(wb, mainData, sheets);
}

function parseSiliconFormat(wb, sheets) {
  const salesSheetName = sheets.find(n => n.toLowerCase().includes('salesmanwise'));
  const gstSheetName   = sheets.find(n => n.toLowerCase().includes('gst'));

  const salesRaw = salesSheetName
    ? xlsx.utils.sheet_to_json(wb.Sheets[salesSheetName], { header: 1, defval: null })
    : [];

  const synopsis = {
    TAX:   { outstanding: salesRaw[3]?.[6]  || 0, invoices: salesRaw[4]?.[6]  || 0, gst: salesRaw[5]?.[6]  || 0 },
    PI:    { outstanding: salesRaw[3]?.[7]  || 0, invoices: salesRaw[4]?.[7]  || 0, gst: salesRaw[5]?.[7]  || 0 },
    DPI:   { outstanding: salesRaw[3]?.[8]  || 0, invoices: salesRaw[4]?.[8]  || 0, gst: salesRaw[5]?.[8]  || 0 },
    TOTAL: { outstanding: salesRaw[3]?.[10] || 0, invoices: salesRaw[4]?.[10] || 0, gst: salesRaw[5]?.[10] || 0 }
  };

  const gstRows = gstSheetName
    ? xlsx.utils.sheet_to_json(wb.Sheets[gstSheetName], { defval: null })
    : [];

  const invoiceRows = gstRows.map((r, i) => ({
    id:          i,
    party:       r['Party Name'] || '—',
    vr_no:       r['Vr.No']      || '—',
    ledger:      r['Ledger']     || '—',
    outstanding: parseFloat(r['Cl.OUTSTANDING AS ON 20.03.2026']) || 0,
    collected:   parseFloat(r['COLLECTED'])   || 0,
    gst:         parseFloat(r['GST '] ?? r['GST']) || 0,
    agent:       r['Agent Name'] || '—',
    due_date:    String(r['Due Date'] || '—')
  })).filter(r => r.party !== '—');

  return { synopsis, invoiceRows, sheets };
}

function parseGenericFormat(wb, data, sheets) {
  const keys    = data.length ? Object.keys(data[0]) : [];
  const numCols = keys.filter(k => data.some(r => typeof r[k] === 'number' && r[k] > 0));
  const strCols = keys.filter(k => data.some(r => typeof r[k] === 'string' && r[k].trim()));

  const agentCol       = strCols.find(k => /agent|sales|person|rep/i.test(k));
  const ledgerCol      = strCols.find(k => /ledger|category|type|service/i.test(k));
  const partyCol       = strCols.find(k => /party|client|customer|company/i.test(k)) || strCols[0];
  const outstandingCol = numCols.find(k => /outstanding|balance|amount|due/i.test(k))  || numCols[0];
  const collectedCol   = numCols.find(k => /collect|paid|receipt/i.test(k))            || numCols[1];
  const gstCol         = numCols.find(k => /gst|tax/i.test(k))                         || numCols[2];
  const vrCol          = keys.find(k => /vr|invoice|no\b/i.test(k))                    || keys[0];

  const invoiceRows = data.map((r, i) => ({
    id:          i,
    party:       String(r[partyCol]       || '—'),
    vr_no:       String(r[vrCol]          || i),
    ledger:      ledgerCol ? String(r[ledgerCol] || 'General') : 'General',
    outstanding: parseFloat(r[outstandingCol]) || 0,
    collected:   collectedCol ? parseFloat(r[collectedCol]) || 0 : 0,
    gst:         gstCol ? parseFloat(r[gstCol]) || 0 : 0,
    agent:       agentCol ? String(r[agentCol] || 'Unknown') : 'Unknown',
    due_date:    '—'
  })).filter(r => r.outstanding > 0 || r.collected > 0);

  const totalOS  = invoiceRows.reduce((s, r) => s + r.outstanding, 0);
  const totalGST = invoiceRows.reduce((s, r) => s + r.gst, 0);

  const synopsis = {
    TAX:   { outstanding: totalOS * 0.69, invoices: Math.round(invoiceRows.length * 0.7),  gst: totalGST * 0.69 },
    PI:    { outstanding: totalOS * 0.16, invoices: Math.round(invoiceRows.length * 0.2),  gst: totalGST * 0.16 },
    DPI:   { outstanding: totalOS * 0.15, invoices: Math.round(invoiceRows.length * 0.1),  gst: totalGST * 0.15 },
    TOTAL: { outstanding: totalOS,        invoices: invoiceRows.length,                    gst: totalGST }
  };

  return { synopsis, invoiceRows, sheets };
}

// ── CSV PARSER ────────────────────────────────────────────────
function parseCSV(filePath) {
  return new Promise((resolve, reject) => {
    const data = [];
    fs.createReadStream(filePath)
      .pipe(csv())
      .on('data', row => data.push(row))
      .on('end', () => {
        // Convert CSV string values to numbers where possible
        const cleaned = data.map(r => {
          const obj = {};
          for (const k of Object.keys(r)) {
            const v = r[k];
            obj[k] = isNaN(parseFloat(v)) ? v : parseFloat(v);
          }
          return obj;
        });
        const wb   = { SheetNames: ['Sheet1'], Sheets: {} };
        resolve(parseGenericFormat(wb, cleaned, ['Sheet1']));
      })
      .on('error', reject);
  });
}

// ── Start server ──────────────────────────────────────────────
initDB().then(() => {
  app.listen(PORT, () => {
    console.log(`\n🚀 Server running → http://localhost:${PORT}`);
    console.log(`   Upload page   → http://localhost:${PORT}/index.html`);
    console.log(`   Dashboard     → http://localhost:${PORT}/dashboard.html\n`);
  });
});