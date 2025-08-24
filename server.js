/* server.js */
const express = require('express');
const multer = require('multer');
const path = require('path');
const { checkCompleteness } = require('./completeness-checker.js');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

const PORT = process.env.PORT || 3000;

// Serve static files (index.html, etc.)
app.use(express.static(path.join(__dirname)));

// ---------------- Existing web search flow (left intact) ----------------
// If you already have this endpoint in your own server, keep yours.
// Here we just echo the file back so the button works out-of-the-box.
app.post('/api/process-excel', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Bitte Excel-Datei hochladen (file).' });
    // Echo back the uploaded workbook (no changes). Replace with your existing logic if you have it.
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition','attachment; filename="Websuche_Ergebnis.xlsx"');
    return res.send(req.file.buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// ---------------- New: Vollständigkeit prüfen (no web search) ----------------
app.post('/api/check-completeness', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Bitte Excel-Datei hochladen (file).' });

    const resultBuffer = await checkCompleteness(req.file.buffer);
    
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition','attachment; filename="Vollstaendigkeitspruefung.xlsx"');
    return res.send(Buffer.from(resultBuffer));

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => {
  console.log(`Server listening on http://localhost:${PORT}`);
});
