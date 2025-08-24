/* scraper.js */
const puppeteer = require('puppeteer');

// URL für Siemens Produktsuche
const a2vUrl = 'https://www.siemens.com/global/de/produkte/produktkatalog.html';

class SiemensProductScraper {
  constructor() {
    this.browser = null;
    this.page = null;
  }

  async init() {
    if (!this.browser) {
      this.browser = await puppeteer.launch({
        headless: true,
        args: ['--no-sandbox', '--disable-setuid-sandbox']
      });
      this.page = await this.browser.newPage();
      
      // User-Agent setzen
      await this.page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36');
      
      // Timeout setzen
      await this.page.setDefaultTimeout(30000);
    }
  }

  async close() {
    if (this.browser) {
      await this.browser.close();
      this.browser = null;
      this.page = null;
    }
  }

  async scrapeMany(a2vNumbers, concurrency = 4) {
    if (!a2vNumbers || a2vNumbers.length === 0) {
      return new Map();
    }

    await this.init();
    
    const results = new Map();
    const chunks = this.chunkArray(a2vNumbers, concurrency);
    
    for (const chunk of chunks) {
      const promises = chunk.map(a2v => this.scrapeOne(a2v));
      const chunkResults = await Promise.allSettled(promises);
      
      chunkResults.forEach((result, index) => {
        if (result.status === 'fulfilled') {
          const a2v = chunk[index];
          results.set(a2v, result.value);
        } else {
          console.error(`Fehler beim Scrapen von ${chunk[index]}:`, result.reason);
          results.set(chunk[index], { error: result.reason.message });
        }
      });
      
      // Kurze Pause zwischen Chunks
      await this.sleep(1000);
    }
    
    return results;
  }

  async scrapeOne(a2vNumber) {
    try {
      // Vereinfachte Mock-Implementierung
      // In der echten Implementierung würden Sie hier die Siemens-Website scrapen
      
      // Simuliere Web-Suche
      await this.sleep(Math.random() * 1000 + 500);
      
      // Mock-Daten basierend auf A2V-Nummer
      const mockData = this.generateMockData(a2vNumber);
      
      return mockData;
      
    } catch (error) {
      console.error(`Fehler beim Scrapen von ${a2vNumber}:`, error);
      return { error: error.message };
    }
  }

  generateMockData(a2vNumber) {
    // Generiere realistische Mock-Daten basierend auf der A2V-Nummer
    const lastChar = a2vNumber.slice(-1);
    const isEven = parseInt(lastChar) % 2 === 0;
    
    if (isEven) {
      return {
        Produkttitel: `Siemens ${a2vNumber} - Automatisierungskomponente`,
        'Weitere Artikelnummer': a2vNumber,
        Materialklassifizierung: 'Metall',
        Werkstoff: 'Edelstahl V2A',
        Gewicht: '2.5 kg',
        Abmessung: '120 x 80 x 45 mm'
      };
    } else {
      return {
        Produkttitel: `Siemens ${a2vNumber} - Sensorik Modul`,
        'Weitere Artikelnummer': a2vNumber,
        Materialklassifizierung: 'Kunststoff',
        Werkstoff: 'Polycarbonat',
        Gewicht: '150 g',
        Abmessung: '60 x 40 x 25 mm'
      };
    }
  }

  chunkArray(array, size) {
    const chunks = [];
    for (let i = 0; i < array.length; i += size) {
      chunks.push(array.slice(i, i + size));
    }
    return chunks;
  }

  sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
}

module.exports = {
  SiemensProductScraper,
  a2vUrl
};