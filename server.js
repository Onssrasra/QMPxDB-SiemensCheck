const express = require('express');
const cors = require('cors');
const helmet = require('helmet');
const path = require('path');
const multer = require('multer');
const ExcelJS = require('exceljs');
const fs = require('fs').promises;

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

// -------- Quality Check Functions (JavaScript Implementation) --------
async function runQualityCheck(inputBuffer) {
  try {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(inputBuffer);
    
    // Get the first worksheet
    const ws = wb.worksheets[0];
    if (!ws) {
      throw new Error('Kein Arbeitsblatt in der Excel-Datei gefunden');
    }

    // Find header row (skip first 2 rows, use row 3 as header)
    const headerRow = 3;
    const firstDataRow = 4;
    
    // Get column headers
    const headers = [];
    const headerRowObj = ws.getRow(headerRow);
    headerRowObj.eachCell((cell, colNumber) => {
      headers[colNumber - 1] = cell.value;
    });

    // Find important columns by name
    const colMap = {};
    headers.forEach((header, index) => {
      if (header) {
        const headerStr = String(header).toLowerCase();
        if (headerStr.includes('fert') || headerStr.includes('prüfhinweis')) {
          colMap.fertPruefhinweis = index;
        } else if (headerStr.includes('länge')) {
          colMap.laenge = index;
        } else if (headerStr.includes('breite')) {
          colMap.breite = index;
        } else if (headerStr.includes('höhe')) {
          colMap.hoehe = index;
        } else if (headerStr.includes('materialkurztext') || headerStr.includes('material-kurztext')) {
          colMap.materialkurztext = index;
        }
      }
    });

    // Define valid values for Fert./Prüfhinweis
    const pos1Values = new Set(["OHNE", "1", "2", "3"]);
    const pos2Values = new Set(["N", "3.2", "3.1", "2.2", "2.1"]);
    const pos3Values = new Set(["N", "CL1", "CL2", "CL3"]);
    const pos4Values = new Set(["N", "J"]);
    const pos5Values = new Set(["N", "A1", "A2", "A3", "A5", "A+"]);

    // Quality check functions
    function checkVollstaendig(hinweis) {
      if (!hinweis || hinweis === '') return 1;
      
      const teile = String(hinweis).split("/").map(teil => teil.trim());
      if (teile.length !== 5) return 1;
      
      return (pos1Values.has(teile[0]) && 
              pos2Values.has(teile[1]) && 
              pos3Values.has(teile[2]) && 
              pos4Values.has(teile[3]) && 
              pos5Values.has(teile[4])) ? 0 : 1;
    }

    function checkPflichtfelder(row) {
      // Check columns B-J, N, R-W (indices 1-9, 13, 17-22)
      const relevantIndices = [1, 2, 3, 4, 5, 6, 7, 8, 9, 13, 17, 18, 19, 20, 21, 22];
      
      for (const index of relevantIndices) {
        if (index < headers.length) {
          const cell = row.getCell(index + 1);
          if (!cell.value || String(cell.value).trim() === '') {
            return 1;
          }
        }
      }
      return 0;
    }

    function checkMasspruefung(row) {
      let l = 0, b = 0, h = 0;
      let txt = '';
      
      try {
        if (colMap.laenge !== undefined) {
          const laengeCell = row.getCell(colMap.laenge + 1);
          l = parseFloat(laengeCell.value) || 0;
        }
        if (colMap.breite !== undefined) {
          const breiteCell = row.getCell(colMap.breite + 1);
          b = parseFloat(breiteCell.value) || 0;
        }
        if (colMap.hoehe !== undefined) {
          const hoeheCell = row.getCell(colMap.hoehe + 1);
          h = parseFloat(hoeheCell.value) || 0;
        }
        if (colMap.materialkurztext !== undefined) {
          const txtCell = row.getCell(colMap.materialkurztext + 1);
          txt = String(txtCell.value || '');
        }
      } catch (e) {
        return 1;
      }

      // Check for text measurements like 12×34, 12x34, 12 X 34, 12*34, 12/34
      const pattern = /\d{1,4}[\s×xX*/]{1,3}\d{1,4}/;
      const hasTextMeasurement = pattern.test(txt);

      // If L, B, H are all 0, there must be at least one text measurement
      if (l === 0 && b === 0 && h === 0) {
        return hasTextMeasurement ? 0 : 1;
      }
      return 0;
    }

    // Process all data rows
    const results = [];
    let totalRows = 0;
    let validRows = 0;

    for (let rowNum = firstDataRow; rowNum <= ws.lastRow.number; rowNum++) {
      const row = ws.getRow(rowNum);
      if (!row.hasValues) continue;

      totalRows++;
      
      // Check Fert./Prüfhinweis
      let fertPruefhinweisError = 0;
      if (colMap.fertPruefhinweis !== undefined) {
        const cell = row.getCell(colMap.fertPruefhinweis + 1);
        fertPruefhinweisError = checkVollstaendig(cell.value);
      }

      // Check mandatory fields
      const pflichtfelderError = checkPflichtfelder(row);
      
      // Check measurements
      const massError = checkMasspruefung(row);
      
      // Overall error
      const gesamtError = Math.max(fertPruefhinweisError, pflichtfelderError, massError);
      
      if (gesamtError === 0) {
        validRows++;
      }

      // Add error columns to the row
      row.splice(ws.columnCount + 1, 0, fertPruefhinweisError, pflichtfelderError, massError, gesamtError);
      
      results.push({
        row: rowNum,
        fertPruefhinweisError,
        pflichtfelderError,
        massError,
        gesamtError
      });
    }

    // Add error column headers
    const headerRowObj2 = ws.getRow(headerRow);
    headerRowObj2.splice(ws.columnCount + 1, 0, 
      "Fehler_Vollständigkeit_Fert./Prüfhinweis",
      "Fehler_Vollständigkeit_B-J+N+R-W", 
      "Fehler_Maßprüfung",
      "Fehler"
    );

    // Create quality report workbook
    const qualityWb = new ExcelJS.Workbook();
    
    // Copy original data with error columns
    const qualityWs = qualityWb.addWorksheet('Qualitätsbericht');
    ws.eachRow((row, rowNumber) => {
      const newRow = qualityWs.addRow(row.values);
      
      // Color code based on errors
      if (rowNumber > headerRow) {
        const errorCol = row.values.length - 1; // Last column is overall error
        const errorValue = row.values[errorCol];
        
        if (errorValue === 1) {
          // Red for errors
          newRow.eachCell((cell, colNumber) => {
            if (colNumber > ws.columnCount) {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFCCCC' } };
            }
          });
        } else if (errorValue === 0) {
          // Green for valid
          newRow.eachCell((cell, colNumber) => {
            if (colNumber > ws.columnCount) {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD5F4E6' } };
            }
          });
        }
      }
    });

    // Create release list (only valid rows)
    const releaseWb = new ExcelJS.Workbook();
    const releaseWs = releaseWb.addWorksheet('Freigabeliste');
    
    // Copy headers
    const releaseHeaderRow = releaseWs.addRow(headers);
    releaseHeaderRow.eachCell((cell, colNumber) => {
      cell.font = { bold: true };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF99' } };
    });

    // Copy only valid rows
    for (let rowNum = firstDataRow; rowNum <= ws.lastRow.number; rowNum++) {
      const row = ws.getRow(rowNum);
      if (!row.hasValues) continue;
      
      const errorCol = row.values.length - 1;
      const errorValue = row.values[errorCol];
      
      if (errorValue === 0) {
        // Only copy the original columns (not error columns)
        const originalValues = row.values.slice(0, ws.columnCount);
        releaseWs.addRow(originalValues);
      }
    }

    // Create summary sheet
    const summaryWs = qualityWb.addWorksheet('Zusammenfassung');
    
    const summaryData = [
      ['', 'Fehleranzahl', 'Fehlerquote'],
      [`Gesamt(${totalRows})`, totalRows - validRows, `${((totalRows - validRows) / totalRows * 100).toFixed(2)}%`],
      ['Vollständigkeit_Pflichtfeld', results.filter(r => r.pflichtfelderError === 1).length, 
       `${(results.filter(r => r.pflichtfelderError === 1).length / totalRows * 100).toFixed(2)}%`],
      ['Vollständigkeit_Fert./Prüfhinweis', results.filter(r => r.fertPruefhinweisError === 1).length,
       `${(results.filter(r => r.fertPruefhinweisError === 1).length / totalRows * 100).toFixed(2)}%`],
      ['Gültigkeit_Maß', results.filter(r => r.massError === 1).length,
       `${(results.filter(r => r.massError === 1).length / totalRows * 100).toFixed(2)}%`]
    ];

    summaryData.forEach((row, index) => {
      const newRow = summaryWs.addRow(row);
      newRow.eachCell((cell, colNumber) => {
        cell.alignment = { horizontal: 'center' };
        if (index === 0) {
          cell.font = { bold: true };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF99' } };
        }
      });
    });

    // Set column widths
    summaryWs.getColumn(1).width = 35;
    summaryWs.getColumn(2).width = 14;
    summaryWs.getColumn(3).width = 14;

    // Generate buffers
    const qualityBuffer = await qualityWb.xlsx.writeBuffer();
    const releaseBuffer = await releaseWb.xlsx.writeBuffer();

    return {
      qualityReport: Array.from(qualityBuffer),
      releaseList: Array.from(releaseBuffer),
      stats: {
        total: totalRows,
        valid: validRows,
        invalid: totalRows - validRows
      }
    };

  } catch (error) {
    console.error('Quality check error:', error);
    throw new Error(`Qualitätsprüfung fehlgeschlagen: ${error.message}`);
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
