# Clickbot - Vollständigkeitsprüfung & Websuche

Eine Web-Anwendung zur Excel-Datei-Verarbeitung mit Vollständigkeitsprüfung und Websuche-Funktionalität.

## Features

### 1. Vollständigkeitsprüfung
- **Pflichtfelder prüfen** (Spalten B-J, N, R-W): Werden rot markiert wenn leer
- **Fert./Prüfhinweis validieren**: Ungültige Werte werden orange markiert
- **Maße validieren** (Länge, Breite, Höhe): Negative Werte oder ungültige Kombinationen werden orange markiert
- **Gewicht prüfen**: Werte ≤ 0 werden orange markiert
- **Farbkodierung**: 
  - 🔴 Rot = Pflichtfeld fehlt
  - 🟠 Orange = Ungültiger/unplausibler Wert
  - 🟢 Grün = Zeile ist vollständig und korrekt

### 2. Websuche (Echo-Funktion)
- Lädt Excel-Dateien hoch und gibt sie unverändert zurück
- Für bestehende Websuche-Logik gedacht

## Ausgabe

Die Vollständigkeitsprüfung erstellt eine neue Excel-Datei mit einem Arbeitsblatt:
1. **"Qualitätsbericht"**: Ursprüngliche Daten mit Farbkodierung
   - 🔴 Rot = Pflichtfeld fehlt
   - 🟠 Orange = Ungültiger/unplausibler Wert  
   - 🟢 Grün = Zeile ist vollständig und korrekt

## Installation

1. **Repository klonen:**
   ```bash
   git clone <your-repo-url>
   cd clickbot-completeness-app
   ```

2. **Abhängigkeiten installieren:**
   ```bash
   npm install
   ```

3. **Server starten:**
   ```bash
   npm start
   ```

4. **Im Browser öffnen:**
   ```
   http://localhost:3000
   ```

## Verwendung

1. **Excel-Datei hochladen:**
   - Datei per Drag & Drop oder Klick auswählen
   - **Wichtig**: Header müssen in Zeile 3 stehen, Daten ab Zeile 4

2. **Aktion wählen:**
   - **"Vollständigkeit prüfen"**: Führt die Qualitätsprüfung durch
   - **"Websuche"**: Echo-Funktion (gibt Datei unverändert zurück)

3. **Ergebnis herunterladen:**
   - Nach der Verarbeitung erscheint der Download-Button
   - Excel-Datei mit den Ergebnissen wird heruntergeladen

## Projektstruktur

```
clickbot-completeness-app/
├── server.js              # Hauptserver mit Express-Routen
├── completeness-checker.js # Vollständigkeitsprüfung-Logik
├── index.html             # Web-Oberfläche
├── package.json           # Abhängigkeiten und Skripte
├── .gitignore            # Git-Ignore-Datei
└── README.md             # Diese Datei
```

## Technische Details

- **Backend**: Node.js mit Express
- **Excel-Verarbeitung**: ExcelJS-Bibliothek
- **Datei-Upload**: Multer
- **Frontend**: Vanilla HTML/JavaScript mit Drag & Drop
- **Port**: Standardmäßig 3000 (konfigurierbar über Umgebungsvariable PORT)

## Anforderungen an Excel-Dateien

- **Format**: .xlsx (Excel 2007+)
- **Header**: Muss in Zeile 3 stehen
- **Daten**: Beginnen ab Zeile 4
- **Erwartete Spalten** (für Vollständigkeitsprüfung):
  - Fert./Prüfhinweis
  - Länge, Breite, Höhe
  - Materialkurztext
  - Gewicht (optional)

## Deployment

Für GitHub Pages oder andere statische Hosting-Dienste:
1. `npm run build` (falls verfügbar)
2. Nur die statischen Dateien hochladen
3. Backend-Funktionalität benötigt einen Node.js-Server

## Lizenz

Private Verwendung - nicht für kommerzielle Zwecke bestimmt.
