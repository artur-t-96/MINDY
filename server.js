const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const Database = require('better-sqlite3');
const path = require('path');
const fs = require('fs');
const { jsonrepair } = require('jsonrepair');

const app = express();
const PORT = process.env.PORT || 3000;

const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || 'inframinds2025';
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;

if (!fs.existsSync('./data')) fs.mkdirSync('./data');
if (!fs.existsSync('./uploads')) fs.mkdirSync('./uploads');

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
        placements INTEGER DEFAULT 0,
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
        dzial TEXT,
        tresc TEXT,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (tydzien_id) REFERENCES tygodnie(id)
    );

    CREATE TABLE IF NOT EXISTS import_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        rok INTEGER,
        tydzien INTEGER,
        filename TEXT,
        detected_type TEXT,
        records_imported INTEGER,
        ai_summary TEXT,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP
    );
`);

// ============ MIGRATIONS ============
try { db.prepare("ALTER TABLE import_log ADD COLUMN rok INTEGER").run(); } catch (e) { }
try { db.prepare("ALTER TABLE import_log ADD COLUMN tydzien INTEGER").run(); } catch (e) { }

// Default targets (per person!)
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

app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));
app.use(express.static('public'));

const upload = multer({ dest: 'uploads/', limits: { fileSize: 50 * 1024 * 1024 } });

// ============ HELPERS ============

function getOrCreateTydzien(rok, tydzien) {
    let row = db.prepare('SELECT id FROM tygodnie WHERE rok = ? AND tydzien = ?').get(rok, tydzien);
    if (!row) {
        const info = db.prepare('INSERT INTO tygodnie (rok, tydzien) VALUES (?, ?)').run(rok, tydzien);
        return info.lastInsertRowid;
    }
    return row.id;
}

function getOrCreateOsoba(imie, stanowisko, dzial) {
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
        result[sheetName] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    });
    return result;
}

function excelDataToText(excelData, filename) {
    let text = `PLIK: ${filename}\n\n`;
    const MAX_ROWS_PER_SHEET = 50;  // Increased to capture more data
    const MAX_COLS = 15;  // Limit columns to prevent overly wide rows

    Object.keys(excelData).forEach(sheetName => {
        text += `=== ARKUSZ: ${sheetName} ===\n`;
        const rows = excelData[sheetName];

        // Take first rows up to limit
        rows.slice(0, MAX_ROWS_PER_SHEET).forEach((row, idx) => {
            // Limit columns and trim long cell values
            const limitedRow = row.slice(0, MAX_COLS).map(c => {
                const str = String(c || '');
                return str.length > 50 ? str.substring(0, 50) + '...' : str;
            });
            text += `R${idx + 1}: ${limitedRow.join(' | ')}\n`;
        });

        if (rows.length > MAX_ROWS_PER_SHEET) {
            text += `... (${rows.length - MAX_ROWS_PER_SHEET} więcej wierszy)\n`;
        }
        text += '\n';
    });
    return text;
}

function getWeekNumber(d) {
    d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

function getTargets() {
    return db.prepare('SELECT stanowisko, kpi, wartosc, okres FROM targety ORDER BY id DESC').all();
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
        SELECT o.imie, o.stanowisko, k.dni_pracy, k.leady, k.oferty, k.mrr, k.placements
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

function getAverageData() {
    // Per person average per working day
    const rekrutacja = db.prepare(`
        SELECT o.imie, o.stanowisko,
            SUM(k.dni_pracy) as total_dni,
            SUM(k.weryfikacje) as total_weryfikacje,
            SUM(k.rekomendacje) as total_rekomendacje,
            SUM(k.cv_dodane) as total_cv,
            SUM(k.placements) as total_placements,
            COUNT(DISTINCT k.tydzien_id) as tygodni,
            CASE WHEN SUM(k.dni_pracy) > 0 THEN ROUND(CAST(SUM(k.weryfikacje) AS FLOAT) / SUM(k.dni_pracy), 2) ELSE 0 END as weryf_per_day,
            CASE WHEN SUM(k.dni_pracy) > 0 THEN ROUND(CAST(SUM(k.rekomendacje) AS FLOAT) / SUM(k.dni_pracy), 2) ELSE 0 END as reco_per_day,
            CASE WHEN SUM(k.dni_pracy) > 0 THEN ROUND(CAST(SUM(k.cv_dodane) AS FLOAT) / SUM(k.dni_pracy), 2) ELSE 0 END as cv_per_day,
            CASE WHEN COUNT(DISTINCT k.tydzien_id) > 0 THEN ROUND(CAST(SUM(k.placements) AS FLOAT) / COUNT(DISTINCT k.tydzien_id), 2) ELSE 0 END as placements_per_week
        FROM kpi_rekrutacja k
        JOIN osoby o ON k.osoba_id = o.id
        GROUP BY o.id
        ORDER BY total_placements DESC, reco_per_day DESC, cv_per_day + weryf_per_day DESC
    `).all();

    const sprzedaz = db.prepare(`
        SELECT o.imie, o.stanowisko,
            SUM(k.dni_pracy) as total_dni,
            SUM(k.leady) as total_leady,
            SUM(k.oferty) as total_oferty,
            SUM(k.mrr) as total_mrr,
            SUM(k.placements) as total_placements,
            COUNT(DISTINCT k.tydzien_id) as tygodni,
            CASE WHEN SUM(k.dni_pracy) > 0 THEN ROUND(CAST(SUM(k.leady) AS FLOAT) / SUM(k.dni_pracy), 2) ELSE 0 END as leady_per_day,
            CASE WHEN SUM(k.dni_pracy) > 0 THEN ROUND(CAST(SUM(k.oferty) AS FLOAT) / SUM(k.dni_pracy), 2) ELSE 0 END as oferty_per_day,
            CASE WHEN COUNT(DISTINCT k.tydzien_id) > 0 THEN ROUND(CAST(SUM(k.mrr) AS FLOAT) / COUNT(DISTINCT k.tydzien_id), 0) ELSE 0 END as mrr_per_week
        FROM kpi_sprzedaz k
        JOIN osoby o ON k.osoba_id = o.id
        GROUP BY o.id
        ORDER BY mrr_per_week DESC, oferty_per_day DESC, leady_per_day DESC
    `).all();

    return { rekrutacja, sprzedaz };
}

function getAvailableWeeks() {
    return db.prepare('SELECT rok, tydzien FROM tygodnie ORDER BY rok DESC, tydzien DESC LIMIT 100').all();
}

// Calculate team targets (per person * count)
function calculateTeamTargets(data, targets) {
    const rekrutacjaCount = {
        Sourcer: data.rekrutacja.filter(r => r.stanowisko === 'Sourcer').length,
        Rekruter: data.rekrutacja.filter(r => r.stanowisko === 'Rekruter').length,
        'Delivery Lead': data.rekrutacja.filter(r => r.stanowisko === 'Delivery Lead').length,
        TAC: data.rekrutacja.filter(r => r.stanowisko === 'TAC').length,
        total: data.rekrutacja.length
    };

    const sprzedazCount = {
        SDR: data.sprzedaz.filter(r => r.stanowisko === 'SDR').length,
        BDM: data.sprzedaz.filter(r => r.stanowisko === 'BDM').length,
        'Head of Technology': data.sprzedaz.filter(r => r.stanowisko === 'Head of Technology').length,
        total: data.sprzedaz.length
    };

    const getTarget = (kpi) => {
        const t = targets.find(t => t.kpi === kpi);
        return t ? t.wartosc : 0;
    };

    return {
        rekrutacja: {
            weryfikacje: rekrutacjaCount.Sourcer * getTarget('weryfikacje'),
            rekomendacje: rekrutacjaCount.Sourcer * getTarget('rekomendacje'),
            cv_dodane: rekrutacjaCount.Rekruter * getTarget('cv_dodane'),
            placements: rekrutacjaCount.total * (getTarget('placements') / 4), // monthly / 4 = weekly
            peopleCount: rekrutacjaCount
        },
        sprzedaz: {
            leady: sprzedazCount.SDR * getTarget('leady'),
            oferty: sprzedazCount.BDM * getTarget('oferty'),
            mrr: sprzedazCount['Head of Technology'] * getTarget('mrr'),
            placements: sprzedazCount.total * (getTarget('placements') / 4),
            peopleCount: sprzedazCount
        }
    };
}

// ============ AI FUNCTIONS ============

async function analyzeExcelWithAI(filesData) {
    if (!ANTHROPIC_API_KEY) {
        return { error: 'Brak klucza API. Dodaj ANTHROPIC_API_KEY.' };
    }

    let filesContext = '';
    filesData.forEach(({ filename, data }) => {
        filesContext += excelDataToText(data, filename) + '\n---\n\n';
    });

    const prompt = `Przeanalizuj pliki Excel i wyciągnij dane KPI dla firmy rekrutacyjnej InfraMinds.

WAŻNE: Pliki mogą zawierać dane za WIELE tygodni/miesięcy. Musisz rozpoznać:
- Kolumnę z datą, tygodniem lub okresem (np. "Tydzień", "Data", "Okres", "Week", "W", "T")
- Lub arkusze nazwane po tygodniach/miesiącach
- Każdy wiersz przypisz do właściwego tygodnia i roku

STANOWISKA:
- Rekrutacja: Sourcer, Rekruter, TAC, Delivery Lead
- Sprzedaż: SDR, BDM, Head of Technology

KPI:
- Sourcer: weryfikacje, rekomendacje, placements
- Rekruter: cv_dodane, placements  
- TAC: placements
- Delivery Lead: placements, hit_ratio
- SDR: leady
- BDM: oferty
- Head of Technology: mrr

PLIKI:
${filesContext}

Zwróć TYLKO JSON:
{
    "detected_data": [
        {
            "type": "rekrutacja|sprzedaz|hit_ratio|prep_calls",
            "source_file": "nazwa.xlsx",
            "records": [
                {
                    "rok": 2024,
                    "tydzien": 50,
                    "miesiac": 12,
                    "imie": "Imię",
                    "stanowisko": "Sourcer|Rekruter|TAC|Delivery Lead|SDR|BDM|Head of Technology",
                    "dni_pracy": 5,
                    "weryfikacje": 0, "rekomendacje": 0, "cv_dodane": 0, "placements": 0,
                    "leady": 0, "oferty": 0, "mrr": 0,
                    "zamkniete_requesty": 0, "hit_ratio": 0
                }
            ]
        }
    ],
    "detected_periods": [
        {"rok": 2024, "tydzien": 49},
        {"rok": 2024, "tydzien": 50}
    ],
    "summary": "Znaleziono dane za X tygodni dla Y osób",
    "warnings": []
}

WAŻNE:
- Jeśli nie ma kolumny z datą/tygodniem, użyj bieżącego tygodnia: ${getWeekNumber(new Date())}/${new Date().getFullYear()}
- Jeśli jest data (np. 2024-12-10), oblicz tydzień z tej daty
- Jeśli jest "Tydzień 50" lub "W50" - użyj tego
- Każdy rekord MUSI mieć rok i tydzien
- W rekordach uwzględniaj TYLKO pola które mają wartości > 0 (pomiń zerowe wartości)
- Odpowiedź musi być kompletnym, poprawnym JSON - upewnij się że wszystkie nawiasy są zamknięte
- Jeśli danych jest dużo, skup się na najważniejszych (pierwsze 50 rekordów)`;

    try {
        const response = await fetch('https://api.anthropic.com/v1/messages', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'x-api-key': ANTHROPIC_API_KEY,
                'anthropic-version': '2023-06-01'
            },
            body: JSON.stringify({
                model: 'claude-3-haiku-20240307',
                max_tokens: 4096,  // Haiku max limit; jsonrepair handles truncation
                messages: [{ role: 'user', content: prompt }]
            })
        });

        const result = await response.json();

        if (!response.ok) {
            console.error('Anthropic API Error:', result);
            return { error: `Błąd API AI (${response.status}): ${result.error?.message || result.error?.type || 'Nieznany błąd'}` };
        }

        if (result.content && result.content[0]) {
            let jsonText = result.content[0].text;
            console.log("AI Response length:", jsonText.length);
            const jsonMatch = jsonText.match(/\{[\s\S]*\}/);

            if (jsonMatch) {
                jsonText = jsonMatch[0];
                try {
                    const parsed = JSON.parse(jsonText);
                    console.log("JSON parsed successfully. Records:", parsed.detected_data?.length || 0);
                    return parsed;
                } catch (parseErr) {
                    console.log("JSON Parse Error. Attempting repair with jsonrepair...");
                    console.log("First 500 chars:", jsonText.substring(0, 500));
                    console.log("Last 500 chars:", jsonText.substring(jsonText.length - 500));

                    // Use jsonrepair library for robust JSON repair
                    try {
                        const repairedJson = jsonrepair(jsonText);
                        console.log("JSON successfully repaired");
                        return JSON.parse(repairedJson);
                    } catch (repairErr) {
                        console.log("jsonrepair failed, trying manual repair...");

                        // Fallback: Try to cut off at the last valid closing sequence
                        const lastBrace = jsonText.lastIndexOf('}');
                        const lastBracket = jsonText.lastIndexOf(']');
                        const lastValidChar = Math.max(lastBrace, lastBracket);

                        if (lastValidChar !== -1) {
                            try {
                                const truncated = jsonText.substring(0, lastValidChar + 1);
                                const repairedTruncated = jsonrepair(truncated);
                                return JSON.parse(repairedTruncated);
                            } catch (e) { /* ignore */ }
                        }

                        console.error("All JSON repair attempts failed. Original text length:", jsonText.length);
                        console.error("Parse error:", parseErr.message);

                        return {
                            error: 'Otrzymano niepełne dane od AI (JSON Error). Spróbuj wgrać mniej plików naraz lub plik z mniejszą ilością danych.',
                            details: parseErr.message
                        };
                    }
                }
            }
        }
        return { error: 'Nie udało się przeanalizować odpowiedzi (brak contentu)' };
    } catch (err) {
        return { error: `Błąd połączenia: ${err.message}` };
    }
}

async function generateDepartmentHelp(dzial, rok, tydzien) {
    if (!ANTHROPIC_API_KEY) {
        return { error: 'Brak klucza API' };
    }

    const weekData = getCurrentWeekData(rok, tydzien);
    const targets = getTargets();
    const teamTargets = calculateTeamTargets(weekData, targets);

    let context = '';

    if (dzial === 'rekrutacja') {
        const data = weekData.rekrutacja;
        const tt = teamTargets.rekrutacja;

        const totals = {
            weryfikacje: data.reduce((s, r) => s + (r.weryfikacje || 0), 0),
            rekomendacje: data.reduce((s, r) => s + (r.rekomendacje || 0), 0),
            cv_dodane: data.reduce((s, r) => s + (r.cv_dodane || 0), 0),
            placements: data.reduce((s, r) => s + (r.placements || 0), 0)
        };

        context = `DZIAŁ REKRUTACJI (MINDY) - Tydzień ${tydzien}/${rok}

ZESPÓŁ (${data.length} osób):
${data.map(r => `- ${r.imie} (${r.stanowisko}): ${r.dni_pracy}dni, Weryf:${r.weryfikacje}, Reco:${r.rekomendacje}, CV:${r.cv_dodane}, Placements:${r.placements}`).join('\n')}

SUMY vs TARGETY ZESPOŁOWE:
- Weryfikacje: ${totals.weryfikacje} / ${tt.weryfikacje} (${tt.weryfikacje > 0 ? Math.round(totals.weryfikacje / tt.weryfikacje * 100) : 0}%)
- Rekomendacje: ${totals.rekomendacje} / ${tt.rekomendacje} (${tt.rekomendacje > 0 ? Math.round(totals.rekomendacje / tt.rekomendacje * 100) : 0}%)
- CV do bazy: ${totals.cv_dodane} / ${tt.cv_dodane} (${tt.cv_dodane > 0 ? Math.round(totals.cv_dodane / tt.cv_dodane * 100) : 0}%)
- Placements: ${totals.placements} / ${tt.placements.toFixed(1)} (${tt.placements > 0 ? Math.round(totals.placements / tt.placements * 100) : 0}%)

HIT RATIO (Delivery Leads):
${weekData.hitRatio.map(h => `- ${h.imie}: ${h.hit_ratio}% (${h.placements}/${h.zamkniete_requesty})`).join('\n') || 'Brak danych'}`;

    } else {
        const data = weekData.sprzedaz;
        const tt = teamTargets.sprzedaz;

        const totals = {
            leady: data.reduce((s, r) => s + (r.leady || 0), 0),
            oferty: data.reduce((s, r) => s + (r.oferty || 0), 0),
            mrr: data.reduce((s, r) => s + (r.mrr || 0), 0),
            placements: data.reduce((s, r) => s + (r.placements || 0), 0)
        };

        context = `DZIAŁ SPRZEDAŻY (INFRON) - Tydzień ${tydzien}/${rok}

ZESPÓŁ (${data.length} osób):
${data.map(r => `- ${r.imie} (${r.stanowisko}): ${r.dni_pracy}dni, Leady:${r.leady}, Oferty:${r.oferty}, MRR:${r.mrr}zł, Placements:${r.placements}`).join('\n')}

SUMY vs TARGETY ZESPOŁOWE:
- Leady: ${totals.leady} / ${tt.leady} (${tt.leady > 0 ? Math.round(totals.leady / tt.leady * 100) : 0}%)
- Oferty: ${totals.oferty} / ${tt.oferty} (${tt.oferty > 0 ? Math.round(totals.oferty / tt.oferty * 100) : 0}%)
- MRR: ${totals.mrr}zł / ${tt.mrr}zł (${tt.mrr > 0 ? Math.round(totals.mrr / tt.mrr * 100) : 0}%)
- Placements: ${totals.placements} / ${tt.placements.toFixed(1)} (${tt.placements > 0 ? Math.round(totals.placements / tt.placements * 100) : 0}%)`;
    }

    try {
        const response = await fetch('https://api.anthropic.com/v1/messages', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'x-api-key': ANTHROPIC_API_KEY,
                'anthropic-version': '2023-06-01'
            },
            body: JSON.stringify({
                model: 'claude-3-haiku-20240307',
                max_tokens: 600,
                messages: [{
                    role: 'user',
                    content: `${context}

Jesteś ${dzial === 'rekrutacja' ? 'MINDY - asystentką rekrutacji' : 'INFRON - asystentem sprzedaży'}.
Mów w pierwszej osobie, bądź przyjazny i konkretny.

Napisz krótką analizę (max 150 słów):
1. 💚 Co idzie dobrze (z imionami, max 2 punkty)
2. 🔴 Co wymaga uwagi (z imionami, max 2 punkty)  
3. 💡 Jedna konkretna rekomendacja na ten tydzień

Używaj emoji. Bądź bezpośredni, nie lej wody.`
                }]
            })
        });

        const result = await response.json();
        if (result.content && result.content[0]) {
            return { help: result.content[0].text };
        }
        return { error: 'Brak odpowiedzi' };
    } catch (err) {
        return { error: err.message };
    }
}

async function generateFullAnalysis(rok, tydzien) {
    if (!ANTHROPIC_API_KEY) {
        return 'Brak klucza API. Dodaj ANTHROPIC_API_KEY aby włączyć analizę AI.';
    }

    const weekData = getCurrentWeekData(rok, tydzien);
    const targets = getTargets();
    const teamTargets = calculateTeamTargets(weekData, targets);

    let context = `DANE TYGODNIA ${tydzien}/${rok}:\n\n`;

    // Rekrutacja
    const rekTotals = {
        weryfikacje: weekData.rekrutacja.reduce((s, r) => s + (r.weryfikacje || 0), 0),
        rekomendacje: weekData.rekrutacja.reduce((s, r) => s + (r.rekomendacje || 0), 0),
        cv_dodane: weekData.rekrutacja.reduce((s, r) => s + (r.cv_dodane || 0), 0),
        placements: weekData.rekrutacja.reduce((s, r) => s + (r.placements || 0), 0)
    };

    context += `REKRUTACJA (${weekData.rekrutacja.length} osób):\n`;
    weekData.rekrutacja.forEach(r => {
        context += `- ${r.imie} (${r.stanowisko}): Weryf:${r.weryfikacje}, Reco:${r.rekomendacje}, CV:${r.cv_dodane}, Place:${r.placements}\n`;
    });
    context += `SUMY: Weryf:${rekTotals.weryfikacje}/${teamTargets.rekrutacja.weryfikacje}, Reco:${rekTotals.rekomendacje}/${teamTargets.rekrutacja.rekomendacje}, CV:${rekTotals.cv_dodane}/${teamTargets.rekrutacja.cv_dodane}, Place:${rekTotals.placements}/${teamTargets.rekrutacja.placements.toFixed(1)}\n\n`;

    // Sprzedaż
    const salesTotals = {
        leady: weekData.sprzedaz.reduce((s, r) => s + (r.leady || 0), 0),
        oferty: weekData.sprzedaz.reduce((s, r) => s + (r.oferty || 0), 0),
        mrr: weekData.sprzedaz.reduce((s, r) => s + (r.mrr || 0), 0),
        placements: weekData.sprzedaz.reduce((s, r) => s + (r.placements || 0), 0)
    };

    context += `SPRZEDAŻ (${weekData.sprzedaz.length} osób):\n`;
    weekData.sprzedaz.forEach(s => {
        context += `- ${s.imie} (${s.stanowisko}): Leady:${s.leady}, Oferty:${s.oferty}, MRR:${s.mrr}zł, Place:${s.placements}\n`;
    });
    context += `SUMY: Leady:${salesTotals.leady}/${teamTargets.sprzedaz.leady}, Oferty:${salesTotals.oferty}/${teamTargets.sprzedaz.oferty}, MRR:${salesTotals.mrr}/${teamTargets.sprzedaz.mrr}zł, Place:${salesTotals.placements}/${teamTargets.sprzedaz.placements.toFixed(1)}\n\n`;

    // Hit Ratio
    context += `HIT RATIO:\n`;
    weekData.hitRatio.forEach(h => {
        context += `- ${h.imie}: ${h.hit_ratio}% (${h.placements}/${h.zamkniete_requesty})\n`;
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
                model: 'claude-3-haiku-20240307',
                max_tokens: 1000,
                messages: [{
                    role: 'user',
                    content: `${context}

WAŻNE: Targety są per osoba - więcej osób w zespole = wyższy target zespołowy.
Sortowanie ważności KPI: 1) Placements 2) Rekomendacje 3) CV do bazy/Weryfikacje

Napisz analizę (max 250 słów):

## 📊 PODSUMOWANIE
(2 zdania o obu działach)

## ✅ SUKCESY
(max 3, z imionami)

## ⚠️ DO POPRAWY
(max 3, z imionami i liczbami)

## 💡 REKOMENDACJE
(2 konkretne działania)

Używaj emoji, bądź konkretny.`
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

    return 'Błąd generowania analizy.';
}

function importAnalyzedData(analysisResult) {
    let totalImported = 0;
    const importDetails = [];
    const importedPeriods = new Set();

    console.log("=== Import Analysis ===");
    console.log("detected_data:", JSON.stringify(analysisResult.detected_data, null, 2).substring(0, 1000));

    if (!analysisResult.detected_data) {
        console.log("No detected_data found!");
        return { imported: 0, details: [], periods: [] };
    }

    analysisResult.detected_data.forEach(dataSet => {
        const type = dataSet.type;
        let imported = 0;

        if (type === 'rekrutacja' && dataSet.records) {
            dataSet.records.forEach(record => {
                if (!record.imie) return;

                const rok = record.rok || new Date().getFullYear();
                const tydzien = record.tydzien || getWeekNumber(new Date());
                const tydzienId = getOrCreateTydzien(rok, tydzien);

                const osobaId = getOrCreateOsoba(record.imie, record.stanowisko || 'Sourcer', 'rekrutacja');

                db.prepare(`
                    INSERT INTO kpi_rekrutacja (osoba_id, tydzien_id, dni_pracy, weryfikacje, rekomendacje, cv_dodane, placements)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(osoba_id, tydzien_id) DO UPDATE SET
                        dni_pracy = excluded.dni_pracy, weryfikacje = excluded.weryfikacje,
                        rekomendacje = excluded.rekomendacje, cv_dodane = excluded.cv_dodane, placements = excluded.placements
                `).run(osobaId, tydzienId, record.dni_pracy || 5, record.weryfikacje || 0, record.rekomendacje || 0, record.cv_dodane || 0, record.placements || 0);

                importedPeriods.add(`${rok}-${tydzien}`);
                imported++;
            });
        }

        if (type === 'sprzedaz' && dataSet.records) {
            dataSet.records.forEach(record => {
                if (!record.imie) return;

                const rok = record.rok || new Date().getFullYear();
                const tydzien = record.tydzien || getWeekNumber(new Date());
                const tydzienId = getOrCreateTydzien(rok, tydzien);

                const osobaId = getOrCreateOsoba(record.imie, record.stanowisko || 'SDR', 'sprzedaz');

                db.prepare(`
                    INSERT INTO kpi_sprzedaz (osoba_id, tydzien_id, dni_pracy, leady, oferty, mrr, placements)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(osoba_id, tydzien_id) DO UPDATE SET
                        dni_pracy = excluded.dni_pracy, leady = excluded.leady,
                        oferty = excluded.oferty, mrr = excluded.mrr, placements = excluded.placements
                `).run(osobaId, tydzienId, record.dni_pracy || 5, record.leady || 0, record.oferty || 0, record.mrr || 0, record.placements || 0);

                importedPeriods.add(`${rok}-${tydzien}`);
                imported++;
            });
        }

        if (type === 'hit_ratio' && dataSet.records) {
            dataSet.records.forEach(record => {
                if (!record.imie) return;

                const rok = record.rok || new Date().getFullYear();
                const miesiac = record.miesiac || Math.ceil((record.tydzien || getWeekNumber(new Date())) / 4.33);

                const osobaId = getOrCreateOsoba(record.imie, 'Delivery Lead', 'rekrutacja');
                const hitRatio = record.zamkniete_requesty > 0
                    ? Math.round((record.placements || 0) / record.zamkniete_requesty * 100)
                    : record.hit_ratio || 0;

                db.prepare(`
                    INSERT INTO hit_ratio (osoba_id, rok, miesiac, zamkniete_requesty, placements, hit_ratio)
                    VALUES (?, ?, ?, ?, ?, ?)
                    ON CONFLICT(osoba_id, rok, miesiac) DO UPDATE SET
                        zamkniete_requesty = excluded.zamkniete_requesty, placements = excluded.placements, hit_ratio = excluded.hit_ratio
                `).run(osobaId, rok, miesiac, record.zamkniete_requesty || 0, record.placements || 0, hitRatio);

                imported++;
            });
        }

        if (type === 'prep_calls' && dataSet.records) {
            dataSet.records.forEach(record => {
                if (!record.dl_imie) return;
                const osobaId = getOrCreateOsoba(record.dl_imie, 'Delivery Lead', 'rekrutacja');

                db.prepare(`INSERT INTO prep_calls (osoba_id, data, kandydat, checklist_json) VALUES (?, ?, ?, ?)`
                ).run(osobaId, record.data || new Date().toISOString().split('T')[0], record.kandydat || '', JSON.stringify(record.checklist || {}));
                imported++;
            });
        }

        if (imported > 0) {
            importDetails.push({ type, file: dataSet.source_file, count: imported });
        }
        totalImported += imported;
    });

    const periods = Array.from(importedPeriods).map(p => {
        const [rok, tydzien] = p.split('-');
        return { rok: parseInt(rok), tydzien: parseInt(tydzien) };
    });

    return { imported: totalImported, details: importDetails, periods };
}

// ============ ROUTES ============

app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));
app.get('/admin', (req, res) => res.sendFile(path.join(__dirname, 'public', 'admin.html')));

// Dashboard data
app.get('/api/dashboard', (req, res) => {
    const rok = parseInt(req.query.rok) || new Date().getFullYear();
    const tydzien = parseInt(req.query.tydzien) || getWeekNumber(new Date());

    const weekData = getCurrentWeekData(rok, tydzien);
    const avgData = getAverageData();
    const targets = getTargets();
    const weeks = getAvailableWeeks();
    const teamTargets = calculateTeamTargets(weekData, targets);

    const latestAnalysis = db.prepare(`
        SELECT a.tresc, a.created_at FROM analizy a 
        JOIN tygodnie t ON a.tydzien_id = t.id 
        WHERE t.rok = ? AND t.tydzien = ? AND a.dzial IS NULL
        ORDER BY a.created_at DESC LIMIT 1
    `).get(rok, tydzien);

    res.json({
        rok, tydzien,
        current: weekData,
        average: avgData,
        targets,
        teamTargets,
        weeks,
        analysis: latestAnalysis?.tresc || null
    });
});

// Upload files (AI rozpoznaje okresy automatycznie)
app.post('/admin/upload', upload.array('files', 10), async (req, res) => {
    if (req.body.password !== ADMIN_PASSWORD) {
        return res.status(401).json({ error: 'Nieprawidłowe hasło' });
    }

    if (!req.files || req.files.length === 0) {
        return res.status(400).json({ error: 'Brak plików' });
    }

    try {
        const filesData = [];
        for (const file of req.files) {
            const excelData = readExcelFile(file.path);
            console.log(`File: ${file.originalname}, Sheets: ${Object.keys(excelData).join(', ')}`);
            Object.keys(excelData).forEach(sheet => {
                console.log(`  Sheet "${sheet}": ${excelData[sheet].length} rows`);
            });
            filesData.push({ filename: file.originalname, data: excelData });
        }

        // Use AI to analyze structure (AI rozpozna okresy automatycznie)
        console.log("Sending to AI for analysis...");
        const analysisResult = await analyzeExcelWithAI(filesData);
        console.log("AI analysis complete. Summary:", analysisResult.summary || analysisResult.error || 'No summary');

        if (analysisResult.error) {
            req.files.forEach(f => fs.unlinkSync(f.path));
            return res.status(400).json({ error: analysisResult.error });
        }

        // Import data (każdy rekord ma swój rok/tydzien)
        const importResult = importAnalyzedData(analysisResult);

        console.log("Import result:", JSON.stringify(importResult));

        // Check if any data was imported
        if (importResult.imported === 0) {
            req.files.forEach(f => fs.unlinkSync(f.path));
            return res.status(400).json({
                error: 'AI nie znalazło danych do importu. Sprawdź czy plik zawiera dane KPI w rozpoznawalnym formacie (imiona, stanowiska, weryfikacje, rekomendacje, CV, placements, leady, oferty, MRR).',
                aiSummary: analysisResult.summary || 'Brak podsumowania'
            });
        }

        // Log
        const periodsStr = importResult.periods.map(p => `T${p.tydzien}/${p.rok}`).join(', ');
        db.prepare(`INSERT INTO import_log (rok, tydzien, filename, detected_type, records_imported, ai_summary) VALUES (?, ?, ?, ?, ?, ?)`
        ).run(
            importResult.periods[0]?.rok || new Date().getFullYear(),
            importResult.periods[0]?.tydzien || getWeekNumber(new Date()),
            req.files.map(f => f.originalname).join(', '),
            analysisResult.detected_data?.map(d => d.type).join(', ') || 'unknown',
            importResult.imported,
            `${analysisResult.summary || ''} Okresy: ${periodsStr}`
        );

        req.files.forEach(f => fs.unlinkSync(f.path));

        res.json({
            success: true,
            message: `Zaimportowano ${importResult.imported} rekordów`,
            summary: analysisResult.summary,
            details: importResult.details,
            periods: importResult.periods,
            warnings: analysisResult.warnings || []
        });

    } catch (err) {
        req.files.forEach(f => { try { fs.unlinkSync(f.path); } catch (e) { } });
        res.status(500).json({ error: err.message });
    }
});

// Generate AI analysis (separate endpoint - costs API credits)
app.post('/api/analyze', async (req, res) => {
    const rok = parseInt(req.body.rok) || new Date().getFullYear();
    const tydzien = parseInt(req.body.tydzien) || getWeekNumber(new Date());

    const analysis = await generateFullAnalysis(rok, tydzien);

    const tydzienId = getOrCreateTydzien(rok, tydzien);
    db.prepare('INSERT INTO analizy (tydzien_id, tresc) VALUES (?, ?)').run(tydzienId, analysis);

    res.json({ analysis });
});

// Department help (AI)
app.post('/api/help/:dzial', async (req, res) => {
    const dzial = req.params.dzial;
    const rok = parseInt(req.body.rok) || new Date().getFullYear();
    const tydzien = parseInt(req.body.tydzien) || getWeekNumber(new Date());

    const result = await generateDepartmentHelp(dzial, rok, tydzien);
    res.json(result);
});

// Delete week data
app.delete('/admin/data/:rok/:tydzien', (req, res) => {
    if (req.query.password !== ADMIN_PASSWORD) {
        return res.status(401).json({ error: 'Nieprawidłowe hasło' });
    }

    const rok = parseInt(req.params.rok);
    const tydzien = parseInt(req.params.tydzien);

    const tydzienRow = db.prepare('SELECT id FROM tygodnie WHERE rok = ? AND tydzien = ?').get(rok, tydzien);
    if (!tydzienRow) {
        return res.status(404).json({ error: 'Nie znaleziono danych' });
    }

    const tydzienId = tydzienRow.id;

    const rekDeleted = db.prepare('DELETE FROM kpi_rekrutacja WHERE tydzien_id = ?').run(tydzienId).changes;
    const salesDeleted = db.prepare('DELETE FROM kpi_sprzedaz WHERE tydzien_id = ?').run(tydzienId).changes;
    db.prepare('DELETE FROM analizy WHERE tydzien_id = ?').run(tydzienId);

    res.json({
        success: true,
        message: `Usunięto ${rekDeleted + salesDeleted} rekordów z tygodnia ${tydzien}/${rok}`
    });
});

// Delete specific import log
app.delete('/admin/import/:id', (req, res) => {
    if (req.query.password !== ADMIN_PASSWORD) {
        return res.status(401).json({ error: 'Nieprawidłowe hasło' });
    }

    db.prepare('DELETE FROM import_log WHERE id = ?').run(parseInt(req.params.id));
    res.json({ success: true });
});

// Get import history
app.get('/admin/history', (req, res) => {
    const history = db.prepare('SELECT * FROM import_log ORDER BY created_at DESC LIMIT 50').all();
    res.json(history);
});

// Targets
app.get('/api/targets', (req, res) => res.json(getTargets()));

app.post('/admin/targets', (req, res) => {
    if (req.body.password !== ADMIN_PASSWORD) {
        return res.status(401).json({ error: 'Nieprawidłowe hasło' });
    }
    const { stanowisko, kpi, wartosc, okres } = req.body;
    db.prepare('INSERT INTO targety (stanowisko, kpi, wartosc, okres) VALUES (?, ?, ?, ?)').run(stanowisko, kpi, wartosc, okres || 'tydzien');
    res.json({ success: true });
});

app.listen(PORT, () => {
    console.log(`
🤖 MINDY & INFRON Dashboard v4
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
🌐 Dashboard: http://localhost:${PORT}
🔐 Admin:     http://localhost:${PORT}/admin
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    `);
});
