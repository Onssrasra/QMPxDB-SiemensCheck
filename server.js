const express = require('express');
const cors = require('cors');
const helmet = require('helmet');
const path = require('path');
const multer = require('multer');
const ExcelJS = require('exceljs');
const { spawn } = require('child_process');
const fs = require('fs').promises;
const os = require('os');

const {
  toNumber,
  parseWeight,
  weightToKg,
  parseDimensionsToLBH,
  normPartNo,
  mapMaterialClassificationToExcel,
  normalizeNCode
} = require('./utils');
const { SiemensProductScraper, a2vUrl } = require('./scraper');

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

const HEADER_ROW = 3;      // Spaltennamen
const LABEL_ROW = 4;       // "DB-Wert" / "Web-Wert"
const FIRST_DATA_ROW = 5;  // erste Datenzeile

app.use(helmet({ contentSecurityPolicy: false }));
app.use(cors());
app.use(express.json({ limit: '2mb' }));
app.use(express.static(__dirname));

const scraper = new SiemensProductScraper();

// -------- Helper Functions ----------
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
    pair.dbCol  = getColumnLetter(adjustedOriginalIndex);
    pair.webCol = getColumnLetter(adjustedOriginalIndex + 1);
    newStructure.pairs.push({ ...pair });
    insertedCols++;
  }
  newStructure.totalInsertedCols = insertedCols;

  const lastCol = ws.lastColumn?.number || ws.columnCount || ws.getRow(HEADER_ROW).cellCount || 0;
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
  const map = {
    green:  'FFD5F4E6', // hellgrün
    red:    'FFFDEAEA', // hellrot
    orange: 'FFFFEAA7', // hellorange
    dbBlue: 'FFE6F3FF', // hellblau (Label DB)
    webBlue:'FFCCE7FF'  // noch helleres Blau (Label Web)
  };
  ws.getCell(addr).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: map[color] || map.green } };
}

function copyColumnFormatting(ws, fromCol, toCol, rowStart, rowEnd) {
  for (let row = rowStart; row <= rowEnd; row++) {
    const fromCell = ws.getCell(`${fromCol}${row}`);
    const toCell   = ws.getCell(`${toCol}${row}`);
    if (fromCell.fill)      toCell.fill = fromCell.fill;
    if (fromCell.font)      toCell.font = fromCell.font;
    if (fromCell.border)    toCell.border = fromCell.border;
    if (fromCell.alignment) toCell.alignment = fromCell.alignment;
    if (fromCell.style)     Object.assign(toCell.style, fromCell.style);
  }
}

function applyLabelCellFormatting(ws, addr, isWebCell = false) {
  const cell = ws.getCell(addr);
  fillColor(ws, addr, isWebCell ? 'webBlue' : 'dbBlue');
  cell.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
  cell.font = { bold: true, size: 10 };
  cell.alignment = { horizontal: 'center', vertical: 'middle' };
}

// -------- Vergleichslogik ----------
function hasValue(v){ return v!==null && v!==undefined && v!=='' && String(v).trim()!==''; }
function eqText(a,b){
  if (a==null||b==null) return false;
  const A=String(a).trim().toLowerCase().replace(/\s+/g,' ');
  const B=String(b).trim().toLowerCase().replace(/\s+/g,' ');
  return A===B;
}
function eqPart(a,b){ return normPartNo(a)===normPartNo(b); }
function eqN(a,b){ return normalizeNCode(a)===normalizeNCode(b); }
function eqWeight(exS, webVal){
  const { value: wv } = parseWeight(webVal);
  if (wv==null) return false;
  const exNum = toNumber(exS); if (exNum==null) return false;
  return Math.abs(exNum - wv) < 1e-9;
}
function eqDimension(exVal, webDimText, dimType){
  const exNum = toNumber(exVal); if (exNum==null) return false;
  const d = parseDimensionsToLBH(webDimText);
  const webVal = (dimType==='L')?d.L:(dimType==='B')?d.B:d.H;
  if (webVal==null) return false;
  return exNum===webVal;
}

function applyTopHeader(ws) {
  const b1Fill  = ws.getCell('B1').fill;
  const ag1Fill = ws.getCell('AG1').fill;
  const ah1Fill = ws.getCell('AH1').fill;

  try { ws.unMergeCells('B1:AF1'); } catch {}
  try { ws.unMergeCells('AH1:AJ1'); } catch {}

  ws.mergeCells('B1:AF1');
  const b1 = ws.getCell('B1');
  b1.value = 'DB AG SAP R/3 K MARA Stammdaten Stand 20.Mai 2025';
  if (b1Fill) b1.fill = b1Fill;
  b1.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

  const ag1 = ws.getCell('AG1');
  ag1.value = 'SAP Klassifizierung aus Okt24';
  if (ag1Fill) ag1.fill = ag1Fill;
  ag1.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

  ws.mergeCells('AH1:AJ1');
  const ah1 = ws.getCell('AH1');
  ah1.value = 'Zusatz Herstellerdaten aus Abfragen in 2024';
  if (ah1Fill) ah1.fill = ah1Fill;
  ah1.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
}

function mergePairHeaders(ws, pairs) {
  for (const pair of pairs) {
    const dbCol  = pair.dbCol;
    const webCol = pair.webCol;
    if (!dbCol || !webCol) continue;

    try { ws.unMergeCells(`${dbCol}2:${webCol}2`); } catch {}
    try { ws.unMergeCells(`${dbCol}3:${webCol}3`); } catch {}

    const v2 = ws.getCell(`${dbCol}2`).value;
    const v3 = ws.getCell(`${dbCol}3`).value;

    ws.mergeCells(`${dbCol}2:${webCol}2`);
    ws.mergeCells(`${dbCol}3:${webCol}3`);

    const top2 = ws.getCell(`${dbCol}2`);
    const top3 = ws.getCell(`${dbCol}3`);
    top2.value = v2;
    top3.value = v3;

    top2.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    top3.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

    const src2 = ws.getCell(`${dbCol}2`);
    const src3 = ws.getCell(`${dbCol}3`);
    if (src2.fill)  top2.fill  = src2.fill;
    if (src2.font)  top2.font  = src2.font;
    if (src2.border)top2.border= src2.border;

    if (src3.fill)  top3.fill  = src3.fill;
    if (src3.font)  top3.font  = src3.font;
    if (src3.border)top3.border= src3.border;
  }
}

// -------- Quality Check Functions (Python Integration) ----------
async function runQualityCheck(inputBuffer) {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'siemens-quality-'));
  const inputFile = path.join(tempDir, 'input.xlsx');
  const outputFile = path.join(tempDir, 'output.xlsx');
  
  // Look for Python script in multiple possible locations
  const possiblePaths = [
    path.join(__dirname, 'Bewertung_Ende.py'),
    path.join(process.cwd(), 'Bewertung_Ende.py'),
    './Bewertung_Ende.py',
    'Bewertung_Ende.py'
  ];

  let pythonScript = null;
  for (const scriptPath of possiblePaths) {
    try {
      await fs.access(scriptPath);
      pythonScript = path.resolve(scriptPath);
      console.log(`Found Python script at: ${pythonScript}`);
      break;
    } catch (error) {
      console.log(`Python script not found at: ${scriptPath}`);
    }
  }

  if (!pythonScript) {
    console.error('Python script not found in any of these locations:', possiblePaths);
    throw new Error('Python script "Bewertung_Ende.py" nicht gefunden. Bitte stellen Sie sicher, dass die Datei im Server-Verzeichnis vorhanden ist.');
  }

  try {
    // Write input file
    await fs.writeFile(inputFile, inputBuffer);
    
    // Read and modify Python script
    let pythonCode = await fs.readFile(pythonScript, 'utf8');
    
    // Replace file paths in Python script
    pythonCode = pythonCode.replace(
      /input_path\s*=\s*.*$/m, 
      `input_path = r"${inputFile}"`
    );
    pythonCode = pythonCode.replace(
      /output_path\s*=\s*.*$/m, 
      `output_path = r"${outputFile}"`
    );
    
    const tempPythonScript = path.join(tempDir, 'quality_check.py');
    await fs.writeFile(tempPythonScript, pythonCode);

    // Try different Python executables
    const pythonExecutables = ['python3', 'python', 'py'];
    let pythonExe = null;

    for (const exe of pythonExecutables) {
      try {
        const testProcess = spawn(exe, ['--version'], { stdio: 'pipe' });
        await new Promise((resolve, reject) => {
          testProcess.on('close', (code) => {
            if (code === 0) {
              pythonExe = exe;
              resolve();
            } else {
              reject();
            }
          });
          testProcess.on('error', reject);
        });
        console.log(`Using Python executable: ${exe}`);
        break;
      } catch (error) {
        console.log(`${exe} not available`);
      }
    }

    if (!pythonExe) {
      throw new Error('Kein Python-Interpreter gefunden. Bitte installieren Sie Python 3.');
    }

    // Run Python script
    return new Promise((resolve, reject) => {
      console.log(`Running Python script: ${pythonExe} ${tempPythonScript}`);
      const python = spawn(pythonExe, [tempPythonScript], {
        cwd: tempDir,
        stdio: ['pipe', 'pipe', 'pipe'],
        timeout: 120000 // 2 minutes timeout
      });

      let stdout = '';
      let stderr = '';

      python.stdout.on('data', (data) => {
        stdout += data.toString();
        console.log('Python stdout:', data.toString());
      });

      python.stderr.on('data', (data) => {
        stderr += data.toString();
        console.error('Python stderr:', data.toString());
      });

      python.on('close', async (code) => {
        try {
          console.log(`Python process exited with code ${code}`);
          console.log('Stdout:', stdout);
          console.log('Stderr:', stderr);

          if (code !== 0) {
            reject(new Error(`Python script failed with code ${code}. Error: ${stderr || 'Unknown error'}`));
            return;
          }

          // Check if output file exists
          try {
            await fs.access(outputFile);
          } catch {
            reject(new Error('Python script completed but output file was not created'));
            return;
          }

          // Read results
          const qualityBuffer = await fs.readFile(outputFile);
          
          // Create release list by reading the Excel file
          const wb = new ExcelJS.Workbook();
          await wb.xlsx.load(qualityBuffer);
          
          const releaseWb = new ExcelJS.Workbook();
          const releaseWs = releaseWb.addWorksheet('Freigabeliste');
          
          // Look for sheet with error-free data
          let sourceSheet = wb.getWorksheet('Ohne_Fehler');
          if (!sourceSheet) {
            // If no "Ohne_Fehler" sheet, try first sheet and filter
            sourceSheet = wb.worksheets[0];
          }

          let totalRows = 0;
          let validRows = 0;

          if (sourceSheet) {
            // Copy headers first
            const headerRow = sourceSheet.getRow(1);
            const newHeaderRow = releaseWs.addRow(headerRow.values);
            
            // Copy formatting for header
            headerRow.eachCell((cell, colNumber) => {
              const newCell = newHeaderRow.getCell(colNumber);
              if (cell.fill) newCell.fill = cell.fill;
              if (cell.font) newCell.font = cell.font;
              if (cell.border) newCell.border = cell.border;
              if (cell.alignment) newCell.alignment = cell.alignment;
            });

            // Copy data rows (skip header)
            sourceSheet.eachRow((row, rowNumber) => {
              if (rowNumber > 1) { // Skip header
                totalRows++;
                
                // Check if this row has errors (if there's an error column)
                let hasError = false;
                const errorCol = row.values.findIndex(cell => 
                  String(cell).includes('Fehler') || String(cell) === '1'
                );
                
                if (errorCol === -1 || !row.getCell(errorCol).value) {
                  validRows++;
                  const newRow = releaseWs.addRow(row.values);
                  
                  // Copy formatting
                  row.eachCell((cell, colNumber) => {
                    const newCell = newRow.getCell(colNumber);
                    if (cell.fill) newCell.fill = cell.fill;
                    if (cell.font) newCell.font = cell.font;
                    if (cell.border) newCell.border = cell.border;
                    if (cell.alignment) newCell.alignment = cell.alignment;
                  });
                }
              }
            });
          }
          
          const releaseBuffer = await releaseWb.xlsx.writeBuffer();

          // Cleanup temp directory
          await fs.rm(tempDir, { recursive: true, force: true });

          resolve({
            qualityReport: Array.from(qualityBuffer),
            releaseList: Array.from(releaseBuffer),
            stats: {
              total: totalRows,
              valid: validRows,
              invalid: totalRows - validRows
            }
          });

        } catch (error) {
          console.error('Error processing Python results:', error);
          await fs.rm(tempDir, { recursive: true, force: true }).catch(() => {});
          reject(error);
        }
      });

      python.on('error', async (error) => {
        console.error('Python process error:', error);
        await fs.rm(tempDir, { recursive: true, force: true }).catch(() => {});
        reject(new Error(`Failed to start Python process: ${error.message}`));
      });

      // Handle timeout
      setTimeout(() => {
        python.kill('SIGTERM');
        reject(new Error('Python script timed out after 2 minutes'));
      }, 120000);
    });

  } catch (error) {
    console.error('Quality check setup error:', error);
    await fs.rm(tempDir, { recursive: true, force: true }).catch(() => {});
    throw error;
  }
}

// -------- Routes ----------
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'index.html')));
app.get('/api/health', (req, res) => res.json({ ok: true, time: new Date().toISOString() }));

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

// Web Search Route
app.post('/api/web-search', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Bitte Excel-Datei hochladen (file).' });

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(req.file.buffer);

    // 1) A2V-Nummern sammeln
    const tasks = [];
    const rowsPerSheet = new Map();
    for (const ws of wb.worksheets) {
      const indices = [];
      const last = ws.lastRow?.number || 0;
      for (let r = FIRST_DATA_ROW - 1; r <= last; r++) {
        const a2v = (ws.getCell(`${ORIGINAL_COLS.Z}${r}`).value || '').toString().trim().toUpperCase();
        if (a2v.startsWith('A2V')) { 
          indices.push(r); 
          tasks.push(a2v); 
        }
      }
      rowsPerSheet.set(ws, indices);
    }

    // 2) Scrapen
    const resultsMap = await scraper.scrapeMany(tasks, SCRAPE_CONCURRENCY);

    // 3) Statistics
    let foundCount = 0;
    let missingCount = 0;
    let diffCount = 0;

    // 4) Umbau pro Worksheet
    for (const ws of wb.worksheets) {
      const structure = calculateNewColumnStructure(ws);

      // Spalten einfügen
      for (const pair of [...structure.pairs].reverse()) {
        const insertPos = getColumnIndex(pair.original) + 1;
        ws.spliceColumns(insertPos, 0, [null]);
      }

      // Zeile 4 (Labels) einfügen
      ws.spliceRows(LABEL_ROW, 0, [null]);

      // Zeilen 2 & 3 Inhalte spiegeln + Labels
      for (const pair of structure.pairs) {
        const dbTech = ws.getCell(`${pair.dbCol}2`).value;
        const dbName = ws.getCell(`${pair.dbCol}3`).value;
        ws.getCell(`${pair.webCol}2`).value = dbTech;
        ws.getCell(`${pair.webCol}3`).value = dbName;
        copyColumnFormatting(ws, pair.dbCol, pair.webCol, 1, 3);

        ws.getCell(`${pair.dbCol}${LABEL_ROW}`).value  = 'DB-Wert';
        ws.getCell(`${pair.webCol}${LABEL_ROW}`).value = 'Web-Wert';
        applyLabelCellFormatting(ws, `${pair.dbCol}${LABEL_ROW}`, false);
        applyLabelCellFormatting(ws, `${pair.webCol}${LABEL_ROW}`, true);
      }

      applyTopHeader(ws);
      mergePairHeaders(ws, structure.pairs);

      // Web-Daten eintragen
      const prodRows = rowsPerSheet.get(ws) || [];
      for (const originalRow of prodRows) {
        const currentRow = originalRow + 1;

        let zCol = ORIGINAL_COLS.Z;
        if (structure.otherCols.has(ORIGINAL_COLS.Z)) zCol = structure.otherCols.get(ORIGINAL_COLS.Z);
        const a2v = (ws.getCell(`${zCol}${currentRow}`).value || '').toString().trim().toUpperCase();
        const web = resultsMap.get(a2v) || {};

        for (const pair of structure.pairs) {
          const dbValue = ws.getCell(`${pair.dbCol}${currentRow}`).value;
          let webValue = null;
          let isEqual = false;

          // Web-Wert extrahieren basierend auf Spalte
          switch (pair.original) {
            case 'C': // Material-Kurztext
              webValue = (web.Produkttitel && web.Produkttitel !== 'Nicht gefunden') ? web.Produkttitel : null;
              isEqual  = webValue ? eqText(dbValue || '', webValue) : false;
              break;
            case 'E': // Herstellartikelnummer
              webValue = (web['Weitere Artikelnummer'] && web['Weitere Artikelnummer'] !== 'Nicht gefunden')
                        ? web['Weitere Artikelnummer']
                        : a2v;
              isEqual  = eqPart(dbValue || a2v, webValue);
              break;
            case 'N': // Fert./Prüfhinweis
              if (web.Materialklassifizierung && web.Materialklassifizierung !== 'Nicht gefunden') {
                const code = normalizeNCode(mapMaterialClassificationToExcel(web.Materialklassifizierung));
                if (code) { webValue = code; isEqual = eqN(dbValue || '', code); }
              }
              break;
            case 'P': // Werkstoff
              webValue = (web.Werkstoff && web.Werkstoff !== 'Nicht gefunden') ? web.Werkstoff : null;
              isEqual  = webValue ? eqText(dbValue || '', webValue) : false;
              break;
            case 'S': // Nettogewicht
              if (web.Gewicht && web.Gewicht !== 'Nicht gefunden') {
                const { value } = parseWeight(web.Gewicht);
                if (value != null) { webValue = value; isEqual = eqWeight(dbValue, web.Gewicht); }
              }
              break;
            case 'U': // Länge
              if (web.Abmessung && web.Abmessung !== 'Nicht gefunden') {
                const d = parseDimensionsToLBH(web.Abmessung);
                if (d.L != null) { webValue = d.L; isEqual = eqDimension(dbValue, web.Abmessung, 'L'); }
              }
              break;
            case 'V': // Breite
              if (web.Abmessung && web.Abmessung !== 'Nicht gefunden') {
                const d = parseDimensionsToLBH(web.Abmessung);
                if (d.B != null) { webValue = d.B; isEqual = eqDimension(dbValue, web.Abmessung, 'B'); }
              }
              break;
            case 'W': // Höhe
              if (web.Abmessung && web.Abmessung !== 'Nicht gefunden') {
                const d = parseDimensionsToLBH(web.Abmessung);
                if (d.H != null) { webValue = d.H; isEqual = eqDimension(dbValue, web.Abmessung, 'H'); }
              }
              break;
          }

          const hasDb = hasValue(dbValue);
          const hasWeb = webValue !== null;

          // Statistiken aktualisieren
          if (hasWeb) {
            foundCount++;
            if (hasDb) {
              if (isEqual) {
                fillColor(ws, `${pair.webCol}${currentRow}`, 'green');
              } else {
                fillColor(ws, `${pair.webCol}${currentRow}`, 'red');
                diffCount++;
              }
            } else {
              fillColor(ws, `${pair.webCol}${currentRow}`, 'orange');
            }
            ws.getCell(`${pair.webCol}${currentRow}`).value = webValue;
          } else {
            missingCount++;
            if (hasDb) fillColor(ws, `${pair.webCol}${currentRow}`, 'orange');
          }
        }
      }
    }

    const out = await wb.xlsx.writeBuffer();
    
    res.json({
      fileBuffer: Array.from(out),
      stats: {
        found: foundCount,
        missing: missingCount,
        differences: diffCount
      }
    });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// Quality Check Route
app.post('/api/quality-check', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Bitte Excel-Datei hochladen (file).' });

    const result = await runQualityCheck(req.file.buffer);
    res.json(result);

  } catch (err) {
    console.error('Quality check error:', err);
    res.status(500).json({ error: err.message });
  }
});

// Graceful shutdown
process.on('SIGTERM', async () => {
  console.log('SIGTERM received, closing scraper...');
  await scraper.close();
  process.exit(0);
});

process.on('SIGINT', async () => {
  console.log('SIGINT received, closing scraper...');
  await scraper.close();
  process.exit(0);
});

app.listen(PORT, () => console.log(`Server running at http://0.0.0.0:${PORT}`));
