/* server.js */
const express = require('express');
const cors = require('cors');
const helmet = require('helmet');
const path = require('path');
const multer = require('multer');
const ExcelJS = require('exceljs');

const {
  toNumber,
  parseWeight,
  weightToKg,
  parseDimensionsToLBH,
  normPartNo,
  mapMaterialClassificationToExcel,
  normalizeNCode
} = require('./utils');
const { SiemensProductScraper } = require('./scraper');
const { checkCompleteness } = require('./completeness-checker');

const app = express();
const PORT = process.env.PORT || 3000;
const SCRAPE_CONCURRENCY = Number(process.env.SCRAPE_CONCURRENCY || 4);

// Ursprüngliche Spalten-Definition (für die Input-Erkennung)
const ORIGINAL_COLS = { Z:'Z', E:'E', C:'C', S:'S', T:'T', U:'U', V:'V', W:'W', P:'P', N:'N' };

// DB/Web-Spaltenpaare
const DB_WEB_PAIRS = [
  { original: 'C', dbCol: null, webCol: null, label: 'Material-Kurztext' },
  { original: 'E', dbCol: null, webCol: null, label: 'Herstellartikelnummer' },
  { original: 'N', dbCol: null, webCol: null, label: 'Fert./Prüfhinweis' },
  { original: 'P', dbCol: null, webCol: null, label: 'Werkstoff' },
  { original: 'S', dbCol: null, webCol: null, label: 'Nettogewicht' },
  { original: 'U', dbCol: null, webCol: null, label: 'Länge' },
  { original: 'V', dbCol: null, webCol: null, label: 'Breite' },
  { original: 'W', dbCol: null, webCol: null, label: 'Höhe' }
];

const HEADER_ROW = 3;
const LABEL_ROW = 4;
const FIRST_DATA_ROW = 5;

// Middleware
app.use(helmet({ 
  contentSecurityPolicy: {
    directives: {
      defaultSrc: ["'self'"],
      scriptSrc: ["'self'", "'unsafe-inline'", "https://cdnjs.cloudflare.com"],
      styleSrc: ["'self'", "'unsafe-inline'"],
      connectSrc: ["'self'"]
    }
  }
}));
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname)));

const scraper = new SiemensProductScraper();

// Helper functions
function getColumnLetter(index) {
  let result = '';
  while (index > 0) {
    index--;
    result = String.fromCharCode(65 + (index % 26)) + result;
    index = Math.floor(index / 26);
  }
  return result;
}

function getColumnIndex(letter) {
  let index = 0;
  for (let i = 0; i < letter.length; i++) {
    index = index * 26 + (letter.charCodeAt(i) - 64);
  }
  return index;
}

function calculateNewColumnStructure(ws) {
  const newStructure = { pairs: [], otherCols: new Map(), totalInsertedCols: 0 };
  let insertedCols = 0;

  for (const pair of DB_WEB_PAIRS) {
    const originalIndex = getColumnIndex(pair.original);
    const adjustedOriginalIndex = originalIndex + insertedCols;
    pair.dbCol = getColumnLetter(adjustedOriginalIndex);
    pair.webCol = getColumnLetter(adjustedOriginalIndex + 1);
    newStructure.pairs.push({ ...pair });
    insertedCols++;
  }
  newStructure.totalInsertedCols = insertedCols;

  const lastCol = ws.lastColumn?.number || ws.columnCount || 26;
  for (let colIndex = 1; colIndex <= lastCol; colIndex++) {
    const originalLetter = getColumnLetter(colIndex);
    const isPairColumn = DB_WEB_PAIRS.some(p => p.original === originalLetter);
    if (!isPairColumn) {
      let insertedBefore = 0;
      for (const p of DB_WEB_PAIRS) {
        if (getColumnIndex(p.original) < colIndex) insertedBefore++;
      }
      const newLetter = getColumnLetter(colIndex + insertedBefore);
      newStructure.otherCols.set(originalLetter, newLetter);
    }
  }
  return newStructure;
}

function fillColor(ws, addr, color) {
  if (!color) return;
  const colorMap = {
    green: 'FFD5F4E6',
    red: 'FFFDEAEA', 
    orange: 'FFFFEAA7',
    dbBlue: 'FFE6F3FF',
    webBlue: 'FFCCE7FF'
  };
  try {
    ws.getCell(addr).fill = { 
      type: 'pattern', 
      pattern: 'solid', 
      fgColor: { argb: colorMap[color] || colorMap.green } 
    };
  } catch (error) {
    console.warn(`Could not set color for cell ${addr}:`, error.message);
  }
}

function hasValue(v) { 
  return v !== null && v !== undefined && v !== '' && String(v).trim() !== ''; 
}

function eqText(a, b) {
  if (a == null || b == null) return false;
  const A = String(a).trim().toLowerCase().replace(/\s+/g, ' ');
  const B = String(b).trim().toLowerCase().replace(/\s+/g, ' ');
  return A === B;
}

function eqPart(a, b) { 
  return normPartNo(a) === normPartNo(b); 
}

function eqN(a, b) { 
  return normalizeNCode(a) === normalizeNCode(b); 
}

function eqWeight(exS, webVal) {
  const { value: wv } = parseWeight(webVal);
  if (wv == null) return false;
  const exNum = toNumber(exS); 
  if (exNum == null) return false;
  return Math.abs(exNum - wv) < 1e-9;
}

function eqDimension(exVal, webDimText, dimType) {
  const exNum = toNumber(exVal); 
  if (exNum == null) return false;
  const d = parseDimensionsToLBH(webDimText);
  const webVal = (dimType === 'L') ? d.L : (dimType === 'B') ? d.B : d.H;
  if (webVal == null) return false;
  return exNum === webVal;
}

// Routes
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

app.get('/api/health', (req, res) => {
  res.json({ ok: true, time: new Date().toISOString() });
});

const upload = multer({ 
  storage: multer.memoryStorage(), 
  limits: { fileSize: 50 * 1024 * 1024 }
});

// Web Search Endpoint
app.post('/api/process-excel', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'Bitte Excel-Datei hochladen (file).' });
    }

    console.log('Processing file:', req.file.originalname, 'Size:', req.file.size);

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(req.file.buffer);

    if (!wb.worksheets || wb.worksheets.length === 0) {
      return res.status(400).json({ error: 'Keine Arbeitsblätter in der Excel-Datei gefunden.' });
    }

    // Collect A2V numbers from column Z
    const tasks = [];
    const rowsPerSheet = new Map();
    
    for (const ws of wb.worksheets) {
      const indices = [];
      const lastRow = ws.lastRow?.number || 10; // Default fallback
      
      for (let r = FIRST_DATA_ROW - 1; r <= lastRow; r++) {
        try {
          const cellValue = ws.getCell(`${ORIGINAL_COLS.Z}${r}`).value;
          const a2v = (cellValue || '').toString().trim().toUpperCase();
          if (a2v.startsWith('A2V')) { 
            indices.push(r); 
            tasks.push(a2v); 
          }
        } catch (error) {
          // Skip invalid cells
          continue;
        }
      }
      rowsPerSheet.set(ws, indices);
    }

    console.log(`Found ${tasks.length} A2V numbers to process`);

    // Scrape data
    const resultsMap = await scraper.scrapeMany(tasks, SCRAPE_CONCURRENCY);

    // Process each worksheet
    for (const ws of wb.worksheets) {
      try {
        const structure = calculateNewColumnStructure(ws);

        // Insert columns (from right to left)
        for (const pair of [...structure.pairs].reverse()) {
          const insertPos = getColumnIndex(pair.original) + 1;
          ws.spliceColumns(insertPos, 0, [null]);
        }

        // Insert label row
        ws.spliceRows(LABEL_ROW, 0, [null]);

        // Set up headers and labels
        for (const pair of structure.pairs) {
          // Copy headers
          try {
            const dbTech = ws.getCell(`${pair.dbCol}2`).value;
            const dbName = ws.getCell(`${pair.dbCol}3`).value;
            ws.getCell(`${pair.webCol}2`).value = dbTech;
            ws.getCell(`${pair.webCol}3`).value = dbName;
          } catch (error) {
            console.warn(`Could not copy headers for ${pair.original}:`, error.message);
          }

          // Set labels
          try {
            ws.getCell(`${pair.dbCol}${LABEL_ROW}`).value = 'DB-Wert';
            ws.getCell(`${pair.webCol}${LABEL_ROW}`).value = 'Web-Wert';
          } catch (error) {
            console.warn(`Could not set labels for ${pair.original}:`, error.message);
          }
        }

        // Process data rows
        const prodRows = rowsPerSheet.get(ws) || [];
        for (const originalRow of prodRows) {
          const currentRow = originalRow + 1; // Account for inserted label row

          try {
            // Get A2V number
            let zCol = ORIGINAL_COLS.Z;
            if (structure.otherCols.has(ORIGINAL_COLS.Z)) {
              zCol = structure.otherCols.get(ORIGINAL_COLS.Z);
            }
            
            const cellValue = ws.getCell(`${zCol}${currentRow}`).value;
            const a2v = (cellValue || '').toString().trim().toUpperCase();
            const web = resultsMap.get(a2v) || {};

            // Process each pair
            for (const pair of structure.pairs) {
              try {
                const dbValue = ws.getCell(`${pair.dbCol}${currentRow}`).value;
                let webValue = null;
                let isEqual = false;

                switch (pair.original) {
                  case 'C': // Material-Kurztext
                    webValue = (web.Produkttitel && web.Produkttitel !== 'Nicht gefunden') ? web.Produkttitel : null;
                    isEqual = webValue ? eqText(dbValue || '', webValue) : false;
                    break;
                  case 'E': // Herstellartikelnummer
                    webValue = (web['Weitere Artikelnummer'] && web['Weitere Artikelnummer'] !== 'Nicht gefunden')
                              ? web['Weitere Artikelnummer']
                              : a2v;
                    isEqual = eqPart(dbValue || a2v, webValue);
                    break;
                  case 'N': // Fert./Prüfhinweis
                    if (web.Materialklassifizierung && web.Materialklassifizierung !== 'Nicht gefunden') {
                      const code = normalizeNCode(mapMaterialClassificationToExcel(web.Materialklassifizierung));
                      if (code) { 
                        webValue = code; 
                        isEqual = eqN(dbValue || '', code); 
                      }
                    }
                    break;
                  case 'P': // Werkstoff
                    webValue = (web.Werkstoff && web.Werkstoff !== 'Nicht gefunden') ? web.Werkstoff : null;
                    isEqual = webValue ? eqText(dbValue || '', webValue) : false;
                    break;
                  case 'S': // Nettogewicht
                    if (web.Gewicht && web.Gewicht !== 'Nicht gefunden') {
                      const { value } = parseWeight(web.Gewicht);
                      if (value != null) { 
                        webValue = value; 
                        isEqual = eqWeight(dbValue, web.Gewicht); 
                      }
                    }
                    break;
                  case 'U': // Länge
                    if (web.Abmessung && web.Abmessung !== 'Nicht gefunden') {
                      const d = parseDimensionsToLBH(web.Abmessung);
                      if (d.L != null) { 
                        webValue = d.L; 
                        isEqual = eqDimension(dbValue, web.Abmessung, 'L'); 
                      }
                    }
                    break;
                  case 'V': // Breite
                    if (web.Abmessung && web.Abmessung !== 'Nicht gefunden') {
                      const d = parseDimensionsToLBH(web.Abmessung);
                      if (d.B != null) { 
                        webValue = d.B; 
                        isEqual = eqDimension(dbValue, web.Abmessung, 'B'); 
                      }
                    }
                    break;
                  case 'W': // Höhe
                    if (web.Abmessung && web.Abmessung !== 'Nicht gefunden') {
                      const d = parseDimensionsToLBH(web.Abmessung);
                      if (d.H != null) { 
                        webValue = d.H; 
                        isEqual = eqDimension(dbValue, web.Abmessung, 'H'); 
                      }
                    }
                    break;
                }

                const hasDb = hasValue(dbValue);
                const hasWeb = webValue !== null;

                if (hasWeb) {
                  ws.getCell(`${pair.webCol}${currentRow}`).value = webValue;
                  fillColor(ws, `${pair.webCol}${currentRow}`, hasDb ? (isEqual ? 'green' : 'red') : 'orange');
                } else {
                  if (hasDb) {
                    fillColor(ws, `${pair.webCol}${currentRow}`, 'orange');
                  }
                }
              } catch (pairError) {
                console.warn(`Error processing pair ${pair.original} in row ${currentRow}:`, pairError.message);
              }
            }
          } catch (rowError) {
            console.warn(`Error processing row ${currentRow}:`, rowError.message);
          }
        }
      } catch (wsError) {
        console.error(`Error processing worksheet ${ws.name}:`, wsError.message);
      }
    }

    const out = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="DB_Produktvergleich_verarbeitet.xlsx"');
    res.send(Buffer.from(out));

  } catch (err) {
    console.error('Error in /api/process-excel:', err);
    res.status(500).json({ error: err.message || 'Unbekannter Fehler beim Verarbeiten der Excel-Datei' });
  }
});

// Completeness Check Endpoint
app.post('/api/check-completeness', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'Bitte Excel-Datei hochladen (file).' });
    }

    console.log('Processing completeness check for:', req.file.originalname);

    const resultBuffer = await checkCompleteness(req.file.buffer);
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="Vollstaendigkeitspruefung.xlsx"');
    return res.send(Buffer.from(resultBuffer));

  } catch (err) {
    console.error('Error in /api/check-completeness:', err);
    res.status(500).json({ error: err.message || 'Fehler bei der Vollständigkeitsprüfung' });
  }
});

// Error handling middleware
app.use((error, req, res, next) => {
  console.error('Unhandled error:', error);
  res.status(500).json({ error: 'Interner Serverfehler' });
});

// 404 handler
app.use((req, res) => {
  res.status(404).json({ error: 'Endpoint nicht gefunden' });
});

// Graceful shutdown
process.on('SIGINT', async () => {
  console.log('\nGraceful shutdown...');
  try {
    await scraper.close();
  } catch (error) {
    console.error('Error closing scraper:', error);
  }
  process.exit(0);
});

process.on('uncaughtException', (error) => {
  console.error('Uncaught Exception:', error);
  process.exit(1);
});

process.on('unhandledRejection', (reason, promise) => {
  console.error('Unhandled Rejection at:', promise, 'reason:', reason);
});

app.listen(PORT, () => {
  console.log(`Server listening on http://localhost:${PORT}`);
  console.log(`Websuche mit ${SCRAPE_CONCURRENCY} parallelen Anfragen`);
});

module.exports = app;
