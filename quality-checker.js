const ExcelJS = require('exceljs');

class QualityChecker {
  constructor() {
    // Gültige Werte für Fert./Prüfhinweis
    this.pos1Values = new Set(["OHNE", "1", "2", "3"]);
    this.pos2Values = new Set(["N", "3.2", "3.1", "2.2", "2.1"]);
    this.pos3Values = new Set(["N", "CL1", "CL2", "CL3"]);
    this.pos4Values = new Set(["N", "J"]);
    this.pos5Values = new Set(["N", "A1", "A2", "A3", "A5", "A+"]);
  }

  // Prüft Fert./Prüfhinweis auf Vollständigkeit und Gültigkeit
  checkFertPruefhinweis(hinweis) {
    if (!hinweis || hinweis === '') return { valid: false, error: 'Fehlender Wert' };
    
    const teile = String(hinweis).split("/").map(teil => teil.trim());
    if (teile.length !== 5) return { valid: false, error: 'Nicht genau 5 Teile' };
    
    const valid = (this.pos1Values.has(teile[0]) && 
                   this.pos2Values.has(teile[1]) && 
                   this.pos3Values.has(teile[2]) && 
                   this.pos4Values.has(teile[3]) && 
                   this.pos5Values.has(teile[4]));
    
    return { 
      valid, 
      error: valid ? null : 'Ungültige Werte in einem oder mehreren Teilen'
    };
  }

  // Prüft Pflichtfelder B-J, N, R-W auf Vollständigkeit
  checkPflichtfelder(row, headers) {
    // Spaltenindizes: B=1, J=9, N=13, R=17, W=22
    const relevantIndices = [1, 2, 3, 4, 5, 6, 7, 8, 9, 13, 17, 18, 19, 20, 21, 22];
    const missingFields = [];
    
    for (const index of relevantIndices) {
      if (index < headers.length) {
        const cell = row.getCell(index + 1);
        if (!cell.value || String(cell.value).trim() === '') {
          const colLetter = this.getColumnLetter(index);
          missingFields.push(`${colLetter} (${headers[index] || 'Unbekannt'})`);
        }
      }
    }
    
    return {
      valid: missingFields.length === 0,
      error: missingFields.length > 0 ? `Fehlende Pflichtfelder: ${missingFields.join(', ')}` : null,
      missingCount: missingFields.length
    };
  }

  // Prüft Maßwerte auf Plausibilität
  checkMasswerte(row, headers) {
    let l = 0, b = 0, h = 0;
    let txt = '';
    let hasTextMeasurement = false;
    
    // Finde Spalten für Länge, Breite, Höhe und Materialkurztext
    const laengeCol = this.findColumnByHeader(headers, ['länge', 'length']);
    const breiteCol = this.findColumnByHeader(headers, ['breite', 'width']);
    const hoeheCol = this.findColumnByHeader(headers, ['höhe', 'height']);
    const materialCol = this.findColumnByHeader(headers, ['materialkurztext', 'material-kurztext', 'material']);
    
    try {
      if (laengeCol !== -1) {
        const cell = row.getCell(laengeCol + 1);
        l = parseFloat(cell.value) || 0;
      }
      if (breiteCol !== -1) {
        const cell = row.getCell(breiteCol + 1);
        b = parseFloat(cell.value) || 0;
      }
      if (hoeheCol !== -1) {
        const cell = row.getCell(hoeheCol + 1);
        h = parseFloat(cell.value) || 0;
      }
      if (materialCol !== -1) {
        const cell = row.getCell(materialCol + 1);
        txt = String(cell.value || '');
      }
    } catch (e) {
      return { valid: false, error: 'Fehler beim Parsen der Maßwerte' };
    }

    // Prüfe auf Textmaße wie 12×34, 12x34, 12 X 34, 12*34, 12/34
    const pattern = /\d{1,4}[\s×xX*/]{1,3}\d{1,4}/;
    hasTextMeasurement = pattern.test(txt);

    // Wenn alle numerischen Maße 0 sind, muss es ein Textmaß geben
    if (l === 0 && b === 0 && h === 0) {
      if (!hasTextMeasurement) {
        return { 
          valid: false, 
          error: 'Alle Maße sind 0, aber kein Textmaß gefunden',
          details: { l, b, h, hasTextMeasurement }
        };
      }
    }

    // Prüfe auf negative oder unplausible Werte
    if (l < 0 || b < 0 || h < 0) {
      return { 
        valid: false, 
        error: 'Negative Maßwerte gefunden',
        details: { l, b, h }
      };
    }

    // Prüfe auf extrem große Werte (über 10000)
    if (l > 10000 || b > 10000 || h > 10000) {
      return { 
        valid: false, 
        error: 'Extrem große Maßwerte (über 10000)',
        details: { l, b, h }
      };
    }

    return { 
      valid: true, 
      error: null,
      details: { l, b, h, hasTextMeasurement }
    };
  }

  // Hilfsfunktion: Spaltenindex nach Header-Text finden
  findColumnByHeader(headers, searchTerms) {
    for (let i = 0; i < headers.length; i++) {
      if (headers[i]) {
        const headerStr = String(headers[i]).toLowerCase();
        for (const term of searchTerms) {
          if (headerStr.includes(term)) {
            return i;
          }
        }
      }
    }
    return -1;
  }

  // Hilfsfunktion: Spaltenbuchstabe aus Index
  getColumnLetter(index) {
    let result = '';
    while (index > 0) {
      index--;
      result = String.fromCharCode(65 + (index % 26)) + result;
      index = Math.floor(index / 26);
    }
    return result;
  }

  // Hauptfunktion: Qualitätsprüfung durchführen
  async checkQuality(inputBuffer) {
    try {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(inputBuffer);
      
      // Erste Arbeitsblatt verwenden
      const ws = wb.worksheets[0];
      if (!ws) {
        throw new Error('Kein Arbeitsblatt in der Excel-Datei gefunden');
      }

      // Header in Zeile 3, Daten ab Zeile 4
      const headerRow = 3;
      const firstDataRow = 4;
      
      // Spaltenüberschriften extrahieren
      const headers = [];
      const headerRowObj = ws.getRow(headerRow);
      headerRowObj.eachCell((cell, colNumber) => {
        headers[colNumber - 1] = cell.value;
      });

      // Fert./Prüfhinweis Spalte finden
      const fertPruefhinweisCol = this.findColumnByHeader(headers, ['fert', 'prüfhinweis', 'fertigungs', 'pruefhinweis']);

      // Qualitätsprüfung für jede Datenzeile
      const results = [];
      let totalRows = 0;
      let validRows = 0;
      let errorRows = 0;

      for (let rowNum = firstDataRow; rowNum <= ws.lastRow.number; rowNum++) {
        const row = ws.getRow(rowNum);
        if (!row.hasValues) continue;

        totalRows++;
        
        // Alle Prüfungen durchführen
        const fertPruefhinweisResult = fertPruefhinweisCol !== -1 ? 
          this.checkFertPruefhinweis(row.getCell(fertPruefhinweisCol + 1).value) : 
          { valid: true, error: null };
        
        const pflichtfelderResult = this.checkPflichtfelder(row, headers);
        const masswerteResult = this.checkMasswerte(row, headers);
        
        // Gesamtbewertung
        const isValid = fertPruefhinweisResult.valid && pflichtfelderResult.valid && masswerteResult.valid;
        
        if (isValid) {
          validRows++;
        } else {
          errorRows++;
        }

        // Ergebnisse speichern
        results.push({
          rowNum,
          fertPruefhinweis: fertPruefhinweisResult,
          pflichtfelder: pflichtfelderResult,
          masswerte: masswerteResult,
          isValid,
          errors: [
            fertPruefhinweisResult.error,
            pflichtfelderResult.error,
            masswerteResult.error
          ].filter(Boolean)
        });

        // Zeile entsprechend einfärben
        this.colorRow(row, isValid);
      }

      // Qualitätsbericht erstellen (gleiche Datei mit Farben)
      const qualityWb = new ExcelJS.Workbook();
      const qualityWs = qualityWb.addWorksheet('Qualitätsbericht');
      
      // Alle Zeilen kopieren (mit Farben)
      ws.eachRow((row, rowNumber) => {
        const newRow = qualityWs.addRow(row.values);
        
        // Formatierung kopieren
        row.eachCell((cell, colNumber) => {
          const newCell = newRow.getCell(colNumber);
          if (cell.fill) newCell.fill = cell.fill;
          if (cell.font) newCell.font = cell.font;
          if (cell.border) newCell.border = cell.border;
          if (cell.alignment) newCell.alignment = cell.alignment;
        });
      });

      // Freigabeliste erstellen (nur fehlerfreie Zeilen im Original-Layout)
      const releaseWb = new ExcelJS.Workbook();
      const releaseWs = releaseWb.addWorksheet('Freigabeliste');
      
      // Header kopieren
      const releaseHeaderRow = releaseWs.addRow(headers);
      releaseHeaderRow.eachCell((cell, colNumber) => {
        cell.font = { bold: true };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF99' } };
      });

      // Nur fehlerfreie Datenzeilen kopieren
      for (let rowNum = firstDataRow; rowNum <= ws.lastRow.number; rowNum++) {
        const row = ws.getRow(rowNum);
        if (!row.hasValues) continue;
        
        const result = results.find(r => r.rowNum === rowNum);
        if (result && result.isValid) {
          // Nur die ursprünglichen Spalten kopieren (keine Fehler-Spalten)
          const originalValues = row.values.slice(0, headers.length);
          releaseWs.addRow(originalValues);
        }
      }

      // Zusammenfassung erstellen
      const summaryWs = qualityWb.addWorksheet('Zusammenfassung');
      
      const summaryData = [
        ['', 'Anzahl', 'Prozent'],
        [`Gesamt (${totalRows})`, totalRows, '100%'],
        ['Fehlerfrei', validRows, `${((validRows / totalRows) * 100).toFixed(1)}%`],
        ['Mit Fehlern', errorRows, `${((errorRows / totalRows) * 100).toFixed(1)}%`],
        ['', '', ''],
        ['Fehlertypen:', '', ''],
        ['Fert./Prüfhinweis', results.filter(r => !r.fertPruefhinweis.valid).length, 
         `${((results.filter(r => !r.fertPruefhinweis.valid).length / totalRows) * 100).toFixed(1)}%`],
        ['Pflichtfelder', results.filter(r => !r.pflichtfelder.valid).length,
         `${((results.filter(r => !r.pflichtfelder.valid).length / totalRows) * 100).toFixed(1)}%`],
        ['Maßwerte', results.filter(r => !r.masswerte.valid).length,
         `${((results.filter(r => !r.masswerte.valid).length / totalRows) * 100).toFixed(1)}%`]
      ];

      summaryData.forEach((row, index) => {
        const newRow = summaryWs.addRow(row);
        newRow.eachCell((cell, colNumber) => {
          cell.alignment = { horizontal: 'center' };
          if (index === 0 || index === 5) {
            cell.font = { bold: true };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF99' } };
          }
        });
      });

      // Spaltenbreiten anpassen
      summaryWs.getColumn(1).width = 35;
      summaryWs.getColumn(2).width = 14;
      summaryWs.getColumn(3).width = 14;

      // Buffer generieren
      const qualityBuffer = await qualityWb.xlsx.writeBuffer();
      const releaseBuffer = await releaseWb.xlsx.writeBuffer();

      return {
        qualityReport: Array.from(qualityBuffer),
        releaseList: Array.from(releaseBuffer),
        stats: {
          total: totalRows,
          valid: validRows,
          invalid: errorRows
        }
      };

    } catch (error) {
      console.error('Qualitätsprüfung fehlgeschlagen:', error);
      throw new Error(`Qualitätsprüfung fehlgeschlagen: ${error.message}`);
    }
  }

  // Zeile einfärben basierend auf Qualität
  colorRow(row, isValid) {
    if (isValid) {
      // Grün für fehlerfreie Zeilen
      row.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD5F4E6' } };
      });
    } else {
      // Rot für Zeilen mit Fehlern
      row.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFCCCC' } };
      });
    }
  }
}

module.exports = QualityChecker; 