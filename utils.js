/* utils.js */

// Konvertiert einen Wert zu einer Zahl (ohne mathjs)
function toNumber(v) {
  if (v == null || v === '') return null;
  const str = String(v).replace(',', '.');
  const n = parseFloat(str);
  return isFinite(n) ? n : null;
}

// Parst Gewichtswerte aus verschiedenen Formaten
function parseWeight(weightText) {
  if (!weightText || typeof weightText !== 'string') {
    return { value: null, unit: null };
  }
  
  const text = weightText.trim();
  
  // Verschiedene Gewichtsmuster
  const patterns = [
    /^(\d+[,.]?\d*)\s*(kg|g|lb|oz|t)$/i,
    /^(\d+[,.]?\d*)\s*(kilo|gramm|pound|ounce|tonne)$/i,
    /^(\d+[,.]?\d*)\s*(\w+)$/i
  ];
  
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) {
      const value = parseFloat(match[1].replace(',', '.'));
      const unit = match[2].toLowerCase();
      
      if (isFinite(value)) {
        // Einheit zu kg konvertieren
        const kgValue = weightToKg(value, unit);
        return { value: kgValue, unit: 'kg' };
      }
    }
  }
  
  return { value: null, unit: null };
}

// Konvertiert verschiedene Gewichtseinheiten zu kg
function weightToKg(value, unit) {
  const unitMap = {
    'kg': 1,
    'kilo': 1,
    'g': 0.001,
    'gramm': 0.001,
    'lb': 0.453592,
    'pound': 0.453592,
    'oz': 0.0283495,
    'ounce': 0.0283495,
    't': 1000,
    'tonne': 1000
  };
  
  const multiplier = unitMap[unit.toLowerCase()] || 1;
  return value * multiplier;
}

// Parst Dimensionen (Länge x Breite x Höhe) aus Text
function parseDimensionsToLBH(dimText) {
  if (!dimText || typeof dimText !== 'string') {
    return { L: null, B: null, H: null };
  }
  
  const text = dimText.trim();
  
  // Verschiedene Dimensionsmuster
  const patterns = [
    /(\d+[,.]?\d*)\s*[xX×]\s*(\d+[,.]?\d*)\s*[xX×]\s*(\d+[,.]?\d*)/,
    /(\d+[,.]?\d*)\s*[xX×]\s*(\d+[,.]?\d*)/,
    /(\d+[,.]?\d*)\s*mm/,
    /(\d+[,.]?\d*)\s*cm/,
    /(\d+[,.]?\d*)\s*m/
  ];
  
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) {
      const values = match.slice(1).map(v => {
        const num = parseFloat(v.replace(',', '.'));
        return isFinite(num) ? num : null;
      }).filter(v => v !== null);
      
      if (values.length === 3) {
        return { L: values[0], B: values[1], H: values[2] };
      } else if (values.length === 2) {
        return { L: values[0], B: values[1], H: null };
      } else if (values.length === 1) {
        return { L: values[0], B: null, H: null };
      }
    }
  }
  
  return { L: null, B: null, H: null };
}

// Normalisiert Artikelnummern für den Vergleich
function normPartNo(partNo) {
  if (!partNo) return '';
  return String(partNo).trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
}

// Normalisiert N-Codes für den Vergleich
function normalizeNCode(code) {
  if (!code) return '';
  return String(code).trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
}

// Mappt Materialklassifizierung zu Excel-Format
function mapMaterialClassificationToExcel(classification) {
  if (!classification) return null;
  
  const mapping = {
    'Kunststoff': 'K',
    'Metall': 'M', 
    'Holz': 'H',
    'Glas': 'G',
    'Keramik': 'C',
    'Textil': 'T',
    'Papier': 'P'
  };
  
  const text = String(classification).trim();
  for (const [key, value] of Object.entries(mapping)) {
    if (text.toLowerCase().includes(key.toLowerCase())) {
      return value;
    }
  }
  
  return classification; // Fallback: Original zurückgeben
}

module.exports = {
  toNumber,
  parseWeight,
  weightToKg,
  parseDimensionsToLBH,
  normPartNo,
  mapMaterialClassificationToExcel,
  normalizeNCode
};
