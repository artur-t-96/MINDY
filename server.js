require('dotenv').config();
const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const Database = require('better-sqlite3');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

// Config
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || 'inframinds2025';
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;

// Ensure directories
if (!fs.existsSync('./data')) fs.mkdirSync('./data');
if (!fs.existsSync('./uploads')) fs.mkdirSync('./uploads');

// Initialize SQLite
const db = new Database('./data/inframinds.db');

// Create tables
db.exec(`
    CREATE TABLE IF NOT EXISTS osoby (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        imie TEXT NOT NULL,
        stanowisko TEXT NOT NULL,
        dzial TEXT NOT NULL CHECK(dzial IN ('rekrutacja', 'sprzedaz')),
        aktywny INTEGER DEFAULT 1,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(imie, stanowisko)
    );

    CREATE TABLE IF NOT EXISTS tygodnie (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        rok INTEGER NOT NULL,
        tydzien INTEGER NOT NULL,
        UNIQUE(rok, tydzien)
    );

    CREATE TABLE IF NOT EXISTS kpi_rekrutacja (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        osoba_id INTEGER NOT NULL,
        tydzien_id INTEGER NOT NULL,
        dni_pracy REAL DEFAULT 5,
        weryfikacje INTEGER DEFAULT 0,
        rekomendacje INTEGER DEFAULT 0,
        cv_dodane INTEGER DEFAULT 0,
        placements INTEGER DEFAULT 0,
        FOREIGN KEY (osoba_id) REFERENCES osoby(id),
        FOREIGN KEY (tydzien_id) REFERENCES tygodnie(id),
        UNIQUE(osoba_id, tydzien_id)
    );

    CREATE TABLE IF NOT EXISTS kpi_sprzedaz (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        osoba_id INTEGER NOT NULL,
        tydzien_id INTEGER NOT NULL,
        dni_pracy REAL DEFAULT 5,
        leady INTEGER DEFAULT 0,
        oferty INTEGER DEFAULT 0,
        mrr REAL DEFAULT 0,
        FOREIGN KEY (osoba_id) REFERENCES osoby(id),
        FOREIGN KEY (tydzien_id) REFERENCES tygodnie(id),
        UNIQUE(osoba_id, tydzien_id)
    );

    CREATE TABLE IF NOT EXISTS hit_ratio (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        osoba_id INTEGER NOT NULL,
        rok INTEGER NOT NULL,
        miesiac INTEGER NOT NULL,
        zamkniete_requesty INTEGER DEFAULT 0,
        placements INTEGER DEFAULT 0,
        hit_ratio REAL DEFAULT 0,
        FOREIGN KEY (osoba_id) REFERENCES osoby(id),
        UNIQUE(osoba_id, rok, miesiac)
    );

    CREATE TABLE IF NOT EXISTS prep_calls (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        osoba_id INTEGER NOT NULL,
        data DATE NOT NULL,
        kandydat TEXT,
        checklist_json TEXT,
        notatki TEXT,
        FOREIGN KEY (osoba_id) REFERENCES osoby(id)
    );

    CREATE TABLE IF NOT EXISTS targety (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        stanowisko TEXT NOT NULL,
        kpi TEXT NOT NULL,
        wartosc REAL NOT NULL,
        okres TEXT DEFAULT 'tydzien',
        aktywny_od DATE DEFAULT CURRENT_DATE
    );

    CREATE TABLE IF NOT EXISTS analizy (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tydzien_id INTEGER,
        tresc TEXT,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (tydzien_id) REFERENCES tygodnie(id)
    );

    CREATE TABLE IF NOT EXISTS import_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        filename TEXT,
        detected_type TEXT,
        records_imported INTEGER,
        ai_summary TEXT,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP
    );
`);

// Insert default targets
const targetCount = db.prepare('SELECT COUNT(*) as cnt FROM targety').get();
if (targetCount.cnt === 0) {
    const insertTarget = db.prepare('INSERT INTO targety (stanowisko, kpi, wartosc, okres) VALUES (?, ?, ?, ?)');
    insertTarget.run('Sourcer', 'weryfikacje', 20, 'tydzien');
    insertTarget.run('Sourcer', 'rekomendacje', 15, 'tydzien');
    insertTarget.run('Rekruter', 'cv_dodane', 25, 'tydzien');
    insertTarget.run('all', 'placements', 1, 'miesiac');
    insertTarget.run('Delivery Lead', 'hit_ratio', 30, 'miesiac');
    insertTarget.run('SDR', 'leady', 10, 'tydzien');
    insertTarget.run('BDM', 'oferty', 1, 'tydzien');
    insertTarget.run('Head of Technology', 'mrr', 4000, 'tydzien');
}

// Middleware
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));
app.use(express.static('public'));

const upload = multer({ dest: 'uploads/', limits: { fileSize: 50 * 1024 * 1024 } });

// ============ HELPER FUNCTIONS ============

function getOrCreateTydzien(rok, tydzien) {
    let row = db.prepare('SELECT id FROM tygodnie WHERE rok = ? AND tydzien = ?').get(rok, tydzien);
    if (!row) {
        const info = db.prepare('INSERT INTO tygodnie (rok, tydzien) VALUES (?, ?)').run(rok, tydzien);
        return info.lastInsertRowid;
    }
    return row.id;
}

function getOrCreateOsoba(imie, stanowisko, dzial) {
    // Normalize stanowisko
    const stanLower = stanowisko.toLowerCase();
    let normalizedStan = stanowisko;
    let normalizedDzial = dzial;

    if (stanLower.includes('sourc')) normalizedStan = 'Sourcer';
    else if (stanLower.includes('rekrut') || stanLower.includes('recruit')) normalizedStan = 'Rekruter';
    else if (stanLower.includes('tac')) normalizedStan = 'TAC';
    else if (stanLower.includes('delivery') || stanLower.includes('dl')) normalizedStan = 'Delivery Lead';
    else if (stanLower.includes('sdr')) { normalizedStan = 'SDR'; normalizedDzial = 'sprzedaz'; }
    else if (stanLower.includes('bdm')) { normalizedStan = 'BDM'; normalizedDzial = 'sprzedaz'; }
    else if (stanLower.includes('head') || stanLower.includes('hot')) { normalizedStan = 'Head of Technology'; normalizedDzial = 'sprzedaz'; }

    let row = db.prepare('SELECT id FROM osoby WHERE imie = ? AND stanowisko = ?').get(imie, normalizedStan);
    if (!row) {
        const info = db.prepare('INSERT OR IGNORE INTO osoby (imie, stanowisko, dzial) VALUES (?, ?, ?)').run(imie, normalizedStan, normalizedDzial);
        if (info.lastInsertRowid) return info.lastInsertRowid;
        row = db.prepare('SELECT id FROM osoby WHERE imie = ? AND stanowisko = ?').get(imie, normalizedStan);
    }
    return row.id;
}

function readExcelFile(filePath) {
    const workbook = XLSX.readFile(filePath);
    const result = {};

    workbook.SheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
        result[sheetName] = data;
    });

    return result;
}

function excelDataToText(excelData, filename) {
    let text = `PLIK: ${filename}\n\n`;

    Object.keys(excelData).forEach(sheetName => {
        text += `=== ARKUSZ: ${sheetName} ===\n`;
        const rows = excelData[sheetName];

        // First 30 rows max
        rows.slice(0, 30).forEach((row, idx) => {
            const rowText = row.map(cell => String(cell || '')).join(' | ');
            text += `Wiersz ${idx + 1}: ${rowText}\n`;
        });

        if (rows.length > 30) {
            text += `... (${rows.length - 30} więcej wierszy)\n`;
        }
        text += '\n';
    });

    return text;
}

// ============ AI ANALYSIS ============

async function analyzeExcelWithAI(filesData, rok, tydzien) {
    if (!ANTHROPIC_API_KEY) {
        return { error: 'Brak klucza API Anthropic. Dodaj ANTHROPIC_API_KEY w ustawieniach.' };
    }

    // Build context from all files
    let filesContext = '';
    filesData.forEach(({ filename, data }) => {
        filesContext += excelDataToText(data, filename) + '\n---\n\n';
    });

    const prompt = `Jesteś ekspertem od analizy danych HR i sprzedaży. Przeanalizuj poniższe pliki Excel i wyciągnij z nich dane KPI.

KONTEKST:
- Firma rekrutacyjna InfraMinds
- Tydzień ${tydzien}, rok ${rok}
- Dział rekrutacji: Sourcer, Rekruter, TAC, Delivery Lead
- Dział sprzedaży: SDR, BDM, Head of Technology

STANOWISKA I ICH KPI:
- Sourcer: weryfikacje CV, rekomendacje, placements
- Rekruter: CV dodane do bazy, placements  
- TAC: placements
- Delivery Lead: placements, hit ratio (placements/zamknięte requesty), prep calls
- SDR: leady
- BDM: wysłane oferty
- Head of Technology: MRR (przychód)

PLIKI DO ANALIZY:
${filesContext}

ZADANIE:
Przeanalizuj dane i zwróć JSON w dokładnie takim formacie:

{
    "detected_data": [
        {
            "type": "rekrutacja|sprzedaz|hit_ratio|prep_calls",
            "source_file": "nazwa_pliku.xlsx",
            "source_sheet": "nazwa_arkusza",
            "records": [
                {
                    "imie": "Imię osoby",
                    "stanowisko": "Stanowisko (znormalizowane do: Sourcer/Rekruter/TAC/Delivery Lead/SDR/BDM/Head of Technology)",
                    "dni_pracy": 5,
                    "weryfikacje": 0,
                    "rekomendacje": 0,
                    "cv_dodane": 0,
                    "placements": 0,
                    "leady": 0,
                    "oferty": 0,
                    "mrr": 0,
                    "zamkniete_requesty": 0,
                    "hit_ratio": 0
                }
            ]
        },
        {
            "type": "prep_calls",
            "source_file": "nazwa.xlsx",
            "records": [
                {
                    "dl_imie": "Imię Delivery Leada",
                    "data": "2024-12-10",
                    "kandydat": "Imię kandydata",
                    "checklist": {"pytanie1": true, "pytanie2": false}
                }
            ]
        }
    ],
    "summary": "Krótkie podsumowanie co znaleziono w plikach",
    "warnings": ["Lista ostrzeżeń jeśli coś nie pasuje"]
}

WAŻNE:
- Rozpoznaj inteligentnie kolumny nawet jeśli nazywają się inaczej (np. "CV sprawdzone" = weryfikacje, "Zatrudnienia" = placements)
- Jeśli nie ma wartości, wstaw 0
- Dla hit_ratio: oblicz sam jeśli są dane (placements / zamkniete_requesty * 100)
- Dla prep_calls: wszystkie kolumny poza DL/Data/Kandydat traktuj jako checklist
- Zwróć TYLKO JSON, bez dodatkowego tekstu`;

    try {
        const response = await fetch('https://api.anthropic.com/v1/messages', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'x-api-key': ANTHROPIC_API_KEY,
                'anthropic-version': '2023-06-01'
            },
            body: JSON.stringify({
                model: 'claude-sonnet-4-20250514',
                max_tokens: 8192,
                messages: [{ role: 'user', content: prompt }]
            })
        });

        const result = await response.json();

        if (result.content && result.content[0]) {
            let text = result.content[0].text;

            // Extract JSON from response
            const jsonMatch = text.match(/\{[\s\S]*\}/);
            if (jsonMatch) {
                return JSON.parse(jsonMatch[0]);
            }
        }

        return { error: 'Nie udało się przeanalizować odpowiedzi AI' };

    } catch (err) {
        console.error('AI Analysis error:', err);
        return { error: err.message };
    }
}

function importAnalyzedData(analysisResult, rok, tydzien) {
    const tydzienId = getOrCreateTydzien(rok, tydzien);
    const miesiac = Math.ceil(tydzien / 4.33);
    let totalImported = 0;
    const importDetails = [];

    if (!analysisResult.detected_data) return { imported: 0, details: [] };

    analysisResult.detected_data.forEach(dataSet => {
        const type = dataSet.type;
        let imported = 0;

        if (type === 'rekrutacja' && dataSet.records) {
            dataSet.records.forEach(record => {
                if (!record.imie) return;

                const osobaId = getOrCreateOsoba(record.imie, record.stanowisko || 'Sourcer', 'rekrutacja');

                db.prepare(`
                    INSERT INTO kpi_rekrutacja (osoba_id, tydzien_id, dni_pracy, weryfikacje, rekomendacje, cv_dodane, placements)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(osoba_id, tydzien_id) DO UPDATE SET
                        dni_pracy = excluded.dni_pracy,
                        weryfikacje = excluded.weryfikacje,
                        rekomendacje = excluded.rekomendacje,
                        cv_dodane = excluded.cv_dodane,
                        placements = excluded.placements
                `).run(
                    osobaId, tydzienId,
                    record.dni_pracy || 5,
                    record.weryfikacje || 0,
                    record.rekomendacje || 0,
                    record.cv_dodane || 0,
                    record.placements || 0
                );
                imported++;
            });
        }

        if (type === 'sprzedaz' && dataSet.records) {
            dataSet.records.forEach(record => {
                if (!record.imie) return;

                const osobaId = getOrCreateOsoba(record.imie, record.stanowisko || 'SDR', 'sprzedaz');

                db.prepare(`
                    INSERT INTO kpi_sprzedaz (osoba_id, tydzien_id, dni_pracy, leady, oferty, mrr)
                    VALUES (?, ?, ?, ?, ?, ?)
                    ON CONFLICT(osoba_id, tydzien_id) DO UPDATE SET
                        dni_pracy = excluded.dni_pracy,
                        leady = excluded.leady,
                        oferty = excluded.oferty,
                        mrr = excluded.mrr
                `).run(
                    osobaId, tydzienId,
                    record.dni_pracy || 5,
                    record.leady || 0,
                    record.oferty || 0,
                    record.mrr || 0
                );
                imported++;
            });
        }

        if (type === 'hit_ratio' && dataSet.records) {
            dataSet.records.forEach(record => {
                if (!record.imie) return;

                const osobaId = getOrCreateOsoba(record.imie, 'Delivery Lead', 'rekrutacja');
                const hitRatio = record.zamkniete_requesty > 0
                    ? Math.round((record.placements || 0) / record.zamkniete_requesty * 100)
                    : record.hit_ratio || 0;

                db.prepare(`
                    INSERT INTO hit_ratio (osoba_id, rok, miesiac, zamkniete_requesty, placements, hit_ratio)
                    VALUES (?, ?, ?, ?, ?, ?)
                    ON CONFLICT(osoba_id, rok, miesiac) DO UPDATE SET
                        zamkniete_requesty = excluded.zamkniete_requesty,
                        placements = excluded.placements,
                        hit_ratio = excluded.hit_ratio
                `).run(
                    osobaId, rok, miesiac,
                    record.zamkniete_requesty || 0,
                    record.placements || 0,
                    hitRatio
                );
                imported++;
            });
        }

        if (type === 'prep_calls' && dataSet.records) {
            dataSet.records.forEach(record => {
                if (!record.dl_imie) return;

                const osobaId = getOrCreateOsoba(record.dl_imie, 'Delivery Lead', 'rekrutacja');

                db.prepare(`
                    INSERT INTO prep_calls (osoba_id, data, kandydat, checklist_json)
                    VALUES (?, ?, ?, ?)
                `).run(
                    osobaId,
                    record.data || new Date().toISOString().split('T')[0],
                    record.kandydat || '',
                    JSON.stringify(record.checklist || {})
                );
                imported++;
            });
        }

        if (imported > 0) {
            importDetails.push({
                type,
                file: dataSet.source_file,
                count: imported
            });
        }
        totalImported += imported;
    });

    return { imported: totalImported, details: importDetails };
}

async function generateDashboardAnalysis(rok, tydzien) {
    const weekData = getCurrentWeekData(rok, tydzien);
    const targets = getTargets();

    if (!ANTHROPIC_API_KEY) {
        return generateSimpleAnalysis(weekData, targets);
    }

    let context = `DANE TYGODNIA ${tydzien}/${rok}:\n\n`;

    context += `REKRUTACJA:\n`;
    weekData.rekrutacja.forEach(r => {
        context += `- ${r.imie} (${r.stanowisko}): ${r.dni_pracy}dni, Weryf:${r.weryfikacje}, Reco:${r.rekomendacje}, CV:${r.cv_dodane}, Place:${r.placements}\n`;
    });

    context += `\nSPRZEDAŻ:\n`;
    weekData.sprzedaz.forEach(s => {
        context += `- ${s.imie} (${s.stanowisko}): ${s.dni_pracy}dni, Leady:${s.leady}, Oferty:${s.oferty}, MRR:${s.mrr}zł\n`;
    });

    context += `\nHIT RATIO:\n`;
    weekData.hitRatio.forEach(h => {
        context += `- ${h.imie}: ${h.placements}/${h.zamkniete_requesty} = ${h.hit_ratio}%\n`;
    });

    context += `\nTARGETY:\n`;
    targets.forEach(t => {
        context += `- ${t.stanowisko}/${t.kpi}: ${t.wartosc}/${t.okres}\n`;
    });

    try {
        const response = await fetch('https://api.anthropic.com/v1/messages', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'x-api-key': ANTHROPIC_API_KEY,
                'anthropic-version': '2023-06-01'
            },
            body: JSON.stringify({
                model: 'claude-sonnet-4-20250514',
                max_tokens: 800,
                messages: [{
                    role: 'user',
                    content: `Jesteś analitykiem KPI dla InfraMinds. Przeanalizuj dane i napisz krótką analizę po polsku.

${context}

Format odpowiedzi:
## 📊 PODSUMOWANIE
(2-3 zdania)

## ✅ SUKCESY  
(max 3 punkty z imionami)

## ⚠️ DO POPRAWY
(max 3 punkty z imionami i liczbami)

## 🎯 HIT RATIO
(analiza Delivery Leadów)

## 💡 REKOMENDACJE
(2-3 konkretne działania)

Max 250 słów, używaj emoji i imion.`
                }]
            })
        });

        const result = await response.json();
        if (result.content && result.content[0]) {
            return result.content[0].text;
        }
    } catch (err) {
        console.error('Analysis error:', err);
    }

    return generateSimpleAnalysis(weekData, targets);
}

function generateSimpleAnalysis(weekData, targets) {
    let analysis = '## 📊 PODSUMOWANIE\n';
    analysis += `Rekrutacja: ${weekData.rekrutacja.length} osób | Sprzedaż: ${weekData.sprzedaz.length} osób\n\n`;
    analysis += '## ✅ SUKCESY\n- Dane załadowane\n\n';
    analysis += '## ⚠️ DO POPRAWY\n- Dodaj klucz API dla pełnej analizy\n\n';
    analysis += '## 💡 REKOMENDACJE\n- Sprawdź targety w panelu admin\n';
    return analysis;
}

function getCurrentWeekData(rok, tydzien) {
    const tydzienRow = db.prepare('SELECT id FROM tygodnie WHERE rok = ? AND tydzien = ?').get(rok, tydzien);
    if (!tydzienRow) return { rekrutacja: [], sprzedaz: [], hitRatio: [] };

    const tydzienId = tydzienRow.id;

    const rekrutacja = db.prepare(`
        SELECT o.imie, o.stanowisko, k.dni_pracy, k.weryfikacje, k.rekomendacje, k.cv_dodane, k.placements
        FROM kpi_rekrutacja k
        JOIN osoby o ON k.osoba_id = o.id
        WHERE k.tydzien_id = ?
    `).all(tydzienId);

    const sprzedaz = db.prepare(`
        SELECT o.imie, o.stanowisko, k.dni_pracy, k.leady, k.oferty, k.mrr
        FROM kpi_sprzedaz k
        JOIN osoby o ON k.osoba_id = o.id
        WHERE k.tydzien_id = ?
    `).all(tydzienId);

    const miesiac = Math.ceil(tydzien / 4.33);
    const hitRatio = db.prepare(`
        SELECT o.imie, h.zamkniete_requesty, h.placements, h.hit_ratio
        FROM hit_ratio h
        JOIN osoby o ON h.osoba_id = o.id
        WHERE h.rok = ? AND h.miesiac = ?
    `).all(rok, miesiac);

    return { rekrutacja, sprzedaz, hitRatio };
}

function getTargets() {
    return db.prepare('SELECT stanowisko, kpi, wartosc, okres FROM targety ORDER BY id DESC').all();
}

function getAverageData() {
    const rekrutacja = db.prepare(`
        SELECT o.imie, o.stanowisko,
            AVG(k.dni_pracy) as avg_dni,
            AVG(k.weryfikacje) as avg_weryfikacje,
            AVG(k.rekomendacje) as avg_rekomendacje,
            AVG(k.cv_dodane) as avg_cv,
            SUM(k.placements) as total_placements,
            COUNT(DISTINCT k.tydzien_id) as tygodni
        FROM kpi_rekrutacja k
        JOIN osoby o ON k.osoba_id = o.id
        GROUP BY o.id
    `).all();

    const sprzedaz = db.prepare(`
        SELECT o.imie, o.stanowisko,
            AVG(k.dni_pracy) as avg_dni,
            AVG(k.leady) as avg_leady,
            AVG(k.oferty) as avg_oferty,
            AVG(k.mrr) as avg_mrr,
            COUNT(DISTINCT k.tydzien_id) as tygodni
        FROM kpi_sprzedaz k
        JOIN osoby o ON k.osoba_id = o.id
        GROUP BY o.id
    `).all();

    return { rekrutacja, sprzedaz };
}

function getAvailableWeeks() {
    return db.prepare('SELECT rok, tydzien FROM tygodnie ORDER BY rok DESC, tydzien DESC LIMIT 100').all();
}

function getWeekNumber(d) {
    d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

// ============ ROUTES ============

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.get('/api/dashboard', (req, res) => {
    const rok = parseInt(req.query.rok) || new Date().getFullYear();
    const tydzien = parseInt(req.query.tydzien) || getWeekNumber(new Date());

    const weekData = getCurrentWeekData(rok, tydzien);

    // SORTING LOGIC: Placements > Rekomendacje > (CV + Weryfikacje)
    weekData.rekrutacja.sort((a, b) => {
        if (b.placements !== a.placements) return b.placements - a.placements;
        if (b.rekomendacje !== a.rekomendacje) return b.rekomendacje - a.rekomendacje;
        const scoreA = (a.cv_dodane || 0) + (a.weryfikacje || 0);
        const scoreB = (b.cv_dodane || 0) + (b.weryfikacje || 0);
        return scoreB - scoreA;
    });

    const avgData = getAverageData();
    const targets = getTargets();
    const weeks = getAvailableWeeks();

    // Calculate Team Counts for dynamic targets
    const teamCounts = {
        sourcer: weekData.rekrutacja.filter(p => p.stanowisko.toLowerCase().includes('sourc')).length || 1,
        rekruter: weekData.rekrutacja.filter(p => p.stanowisko.toLowerCase().includes('rekrut')).length || 1,
        sdr: weekData.sprzedaz.filter(p => p.stanowisko.toLowerCase().includes('sdr')).length || 1,
        bdm: weekData.sprzedaz.filter(p => p.stanowisko.toLowerCase().includes('bdm')).length || 1
    };
    // Safe minimum 1 to avoid division by zero or weird logic if needed, 
    // but for multiplication 0 is fine if no one exists. 
    // Actually for targets, if we have 0 people, target should be 0.
    // Let's use actual counts.
    teamCounts.sourcer = weekData.rekrutacja.filter(p => p.stanowisko.toLowerCase().includes('sourc')).length;
    teamCounts.rekruter = weekData.rekrutacja.filter(p => p.stanowisko.toLowerCase().includes('rekrut')).length;
    teamCounts.sdr = weekData.sprzedaz.filter(p => p.stanowisko.toLowerCase().includes('sdr')).length;
    teamCounts.bdm = weekData.sprzedaz.filter(p => p.stanowisko.toLowerCase().includes('bdm')).length;

    const latestAnalysis = db.prepare(`
        SELECT a.tresc, a.created_at 
        FROM analizy a 
        JOIN tygodnie t ON a.tydzien_id = t.id 
        WHERE t.rok = ? AND t.tydzien = ?
        ORDER BY a.created_at DESC LIMIT 1
    `).get(rok, tydzien);

    res.json({
        rok, tydzien,
        current: weekData,
        average: avgData,
        targets,
        weeks,
        teamCounts,
        analysis: latestAnalysis?.tresc || null
    });
});

app.post('/api/analyze', async (req, res) => {
    const rok = parseInt(req.body.rok) || new Date().getFullYear();
    const tydzien = parseInt(req.body.tydzien) || getWeekNumber(new Date());

    const analysis = await generateDashboardAnalysis(rok, tydzien);

    const tydzienId = getOrCreateTydzien(rok, tydzien);
    db.prepare('INSERT INTO analizy (tydzien_id, tresc) VALUES (?, ?)').run(tydzienId, analysis);

    res.json({ analysis });
});

app.post('/api/analyze-department', async (req, res) => {
    if (!ANTHROPIC_API_KEY) {
        return res.json({ analysis: "To use this feature, add ANTHROPIC_API_KEY to .env" });
    }

    const { department, score, kpis, rok, tydzien } = req.body;

    // Construct prompt
    const context = `
    KONTEKST:
    - Firma: InfraMinds
    - Dział: ${department === 'mindy' ? 'Rekrutacja (MINDY)' : 'Sprzedaż (INFRON)'}
    - Wynik ogólny: ${score}%
    - Okres: Tydzień ${tydzien}/${rok}
    
    SZCZEGÓŁY KPI:
    ${kpis.map(k => `- ${k.name}: ${k.value}/${k.target} (${k.percent}%)`).join('\n')}
    `;

    const prompt = `Jesteś doświadczonym managerem ${department === 'mindy' ? 'HR' : 'Sprzedaży'}.
    Twój zespół osiągnął wynik ${score}%. 
    
    ${context}
    
    Zanalizuj te wyniki i podaj 3 konkretne, motywujące porady co zrobić w tym tygodniu, żeby poprawić wynik.
    Jeśli wynik jest niski (<70%), bądź wspierający ale stanowczy.
    Jeśli wysoki (>100%), pogratuluj i zasugeruj jak to utrzymać.
    
    Odpowiedź sformatuj w HTML (użyj <ul>, <li>, <strong>). Nie dodawaj nagłówków h1-h2, tylko samą treść porad.
    Maksymalnie 150 słów.`;

    try {
        const response = await fetch('https://api.anthropic.com/v1/messages', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'x-api-key': ANTHROPIC_API_KEY,
                'anthropic-version': '2023-06-01'
            },
            body: JSON.stringify({
                model: 'claude-sonnet-4-20250514',
                max_tokens: 500,
                messages: [{ role: 'user', content: prompt }]
            })
        });

        const result = await response.json();
        if (result.content && result.content[0]) {
            res.json({ analysis: result.content[0].text });
        } else {
            res.status(500).json({ error: 'Błąd AI' });
        }
    } catch (err) {
        console.error('Dept Analysis Error:', err);
        res.status(500).json({ error: err.message });
    }
});

app.get('/admin', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'admin.html'));
});

// Multi-file upload with AI analysis
app.post('/admin/upload', upload.array('files', 10), async (req, res) => {
    if (req.body.password !== ADMIN_PASSWORD) {
        return res.status(401).json({ error: 'Nieprawidłowe hasło' });
    }

    if (!req.files || req.files.length === 0) {
        return res.status(400).json({ error: 'Brak plików' });
    }

    const rok = parseInt(req.body.rok) || new Date().getFullYear();
    const tydzien = parseInt(req.body.tydzien) || getWeekNumber(new Date());

    try {
        // Read all Excel files
        const filesData = [];
        for (const file of req.files) {
            const data = readExcelFile(file.path);
            filesData.push({
                filename: file.originalname,
                data
            });
        }

        // Analyze with AI
        const analysisResult = await analyzeExcelWithAI(filesData, rok, tydzien);

        if (analysisResult.error) {
            // Cleanup files
            req.files.forEach(f => fs.unlinkSync(f.path));
            return res.status(400).json({ error: analysisResult.error });
        }

        // Import data
        const importResult = importAnalyzedData(analysisResult, rok, tydzien);

        // Log import
        db.prepare(`
            INSERT INTO import_log (filename, detected_type, records_imported, ai_summary)
            VALUES (?, ?, ?, ?)
        `).run(
            req.files.map(f => f.originalname).join(', '),
            analysisResult.detected_data?.map(d => d.type).join(', ') || 'unknown',
            importResult.imported,
            analysisResult.summary || ''
        );

        // Cleanup files
        req.files.forEach(f => fs.unlinkSync(f.path));

        // Generate new analysis
        const dashboardAnalysis = await generateDashboardAnalysis(rok, tydzien);
        const tydzienId = getOrCreateTydzien(rok, tydzien);
        db.prepare('INSERT INTO analizy (tydzien_id, tresc) VALUES (?, ?)').run(tydzienId, dashboardAnalysis);

        res.json({
            success: true,
            message: `Zaimportowano ${importResult.imported} rekordów z ${req.files.length} plików`,
            summary: analysisResult.summary,
            details: importResult.details,
            warnings: analysisResult.warnings || [],
            analysis: dashboardAnalysis
        });

    } catch (err) {
        console.error('Upload error:', err);
        req.files.forEach(f => {
            try { fs.unlinkSync(f.path); } catch (e) { }
        });
        res.status(500).json({ error: err.message });
    }
});

// Get import history
app.get('/admin/history', (req, res) => {
    const history = db.prepare(`
        SELECT * FROM import_log 
        ORDER BY created_at DESC 
        LIMIT 50
    `).all();
    res.json(history);
});

// Update targets
app.post('/admin/targets', (req, res) => {
    if (req.body.password !== ADMIN_PASSWORD) {
        return res.status(401).json({ error: 'Nieprawidłowe hasło' });
    }

    const { stanowisko, kpi, wartosc, okres } = req.body;

    db.prepare(`
        INSERT INTO targety (stanowisko, kpi, wartosc, okres)
        VALUES (?, ?, ?, ?)
    `).run(stanowisko, kpi, wartosc, okres || 'tydzien');

    res.json({ success: true });
});

// Get targets
app.get('/api/targets', (req, res) => {
    const targets = getTargets();
    res.json(targets);
});

app.listen(PORT, () => {
    console.log(`
🤖 MINDY & INFRON Dashboard v3
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
🌐 Dashboard: http://localhost:${PORT}
🔐 Admin:     http://localhost:${PORT}/admin
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    `);
});
