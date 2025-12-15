const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const Database = require('better-sqlite3');
const path = require('path');
const fs = require('fs');

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

    CREATE TABLE IF NOT EXISTS targety (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        stanowisko TEXT NOT NULL,
        kpi TEXT NOT NULL,
        wartosc REAL NOT NULL,
        okres TEXT DEFAULT 'tydzien'
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
        records_imported INTEGER,
        periods TEXT,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP
    );
`);

// Default targets
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

function getWeekNumber(d) {
    d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

function getOrCreateTydzien(rok, tydzien) {
    let row = db.prepare('SELECT id FROM tygodnie WHERE rok = ? AND tydzien = ?').get(rok, tydzien);
    if (!row) {
        const info = db.prepare('INSERT INTO tygodnie (rok, tydzien) VALUES (?, ?)').run(rok, tydzien);
        return info.lastInsertRowid;
    }
    return row.id;
}

function normalizeStanowisko(stanowisko) {
    if (!stanowisko) return 'Sourcer';
    const s = String(stanowisko).toLowerCase();
    
    if (s.includes('sourc')) return 'Sourcer';
    if (s.includes('rekrut') || s.includes('recruit')) return 'Rekruter';
    if (s.includes('tac')) return 'TAC';
    if (s.includes('delivery') || s === 'dl') return 'Delivery Lead';
    if (s.includes('sdr')) return 'SDR';
    if (s.includes('bdm')) return 'BDM';
    if (s.includes('head') || s.includes('hot')) return 'Head of Technology';
    
    return stanowisko;
}

function getDzialForStanowisko(stanowisko) {
    const salesRoles = ['SDR', 'BDM', 'Head of Technology'];
    return salesRoles.includes(stanowisko) ? 'sprzedaz' : 'rekrutacja';
}

function getOrCreateOsoba(imie, stanowisko) {
    const normalizedStan = normalizeStanowisko(stanowisko);
    const dzial = getDzialForStanowisko(normalizedStan);
    
    let row = db.prepare('SELECT id FROM osoby WHERE imie = ? AND stanowisko = ?').get(imie, normalizedStan);
    if (!row) {
        const info = db.prepare('INSERT OR IGNORE INTO osoby (imie, stanowisko, dzial) VALUES (?, ?, ?)').run(imie, normalizedStan, dzial);
        if (info.lastInsertRowid) return info.lastInsertRowid;
        row = db.prepare('SELECT id FROM osoby WHERE imie = ? AND stanowisko = ?').get(imie, normalizedStan);
    }
    return row.id;
}

function getTargets() {
    return db.prepare('SELECT stanowisko, kpi, wartosc, okres FROM targety ORDER BY id DESC').all();
}

// ============ EXCEL PARSER (stały szablon - BEZ AI) ============

function parseExcelFiles(filePaths) {
    const result = {
        rekrutacja: [],
        sprzedaz: [],
        hitRatio: [],
        periods: new Set()
    };
    
    filePaths.forEach(filePath => {
        const workbook = XLSX.readFile(filePath);
        
        // Arkusz Rekrutacja
        const rekSheet = workbook.SheetNames.find(n => 
            n.toLowerCase().includes('rekrutacja') || n.toLowerCase() === 'recruitment'
        );
        if (rekSheet) {
            const data = XLSX.utils.sheet_to_json(workbook.Sheets[rekSheet]);
            data.forEach(row => {
                const tydzien = row['Tydzień'] || row['Tydzien'] || row['Week'] || row['T'];
                const rok = row['Rok'] || row['Year'] || new Date().getFullYear();
                const imie = row['Imię'] || row['Imie'] || row['Name'] || row['Pracownik'];
                const stanowisko = row['Stanowisko'] || row['Rola'] || row['Role'] || row['Position'];
                
                if (!imie || !tydzien) return;
                
                result.rekrutacja.push({
                    rok: parseInt(rok),
                    tydzien: parseInt(tydzien),
                    imie: String(imie).trim(),
                    stanowisko: normalizeStanowisko(stanowisko),
                    dni_pracy: parseFloat(row['Dni pracy'] || row['Dni'] || row['Days'] || 5),
                    weryfikacje: parseInt(row['Weryfikacje'] || row['Weryf'] || 0),
                    rekomendacje: parseInt(row['Rekomendacje'] || row['Reco'] || 0),
                    cv_dodane: parseInt(row['CV do bazy'] || row['CV'] || 0),
                    placements: parseInt(row['Placements'] || row['Placement'] || 0)
                });
                result.periods.add(`${rok}-${tydzien}`);
            });
        }
        
        // Arkusz Sprzedaż
        const salesSheet = workbook.SheetNames.find(n => 
            n.toLowerCase().includes('sprzeda') || n.toLowerCase() === 'sales'
        );
        if (salesSheet) {
            const data = XLSX.utils.sheet_to_json(workbook.Sheets[salesSheet]);
            data.forEach(row => {
                const tydzien = row['Tydzień'] || row['Tydzien'] || row['Week'] || row['T'];
                const rok = row['Rok'] || row['Year'] || new Date().getFullYear();
                const imie = row['Imię'] || row['Imie'] || row['Name'] || row['Pracownik'];
                const stanowisko = row['Stanowisko'] || row['Rola'] || row['Role'] || row['Position'];
                
                if (!imie || !tydzien) return;
                
                result.sprzedaz.push({
                    rok: parseInt(rok),
                    tydzien: parseInt(tydzien),
                    imie: String(imie).trim(),
                    stanowisko: normalizeStanowisko(stanowisko),
                    dni_pracy: parseFloat(row['Dni pracy'] || row['Dni'] || row['Days'] || 5),
                    leady: parseInt(row['Leady'] || row['Leads'] || 0),
                    oferty: parseInt(row['Wysłane oferty'] || row['Oferty'] || row['Offers'] || 0),
                    mrr: parseFloat(row['MRR'] || row['Revenue'] || 0)
                });
                result.periods.add(`${rok}-${tydzien}`);
            });
        }
        
        // Arkusz HitRatio
        const hitSheet = workbook.SheetNames.find(n => 
            n.toLowerCase().includes('hit') || n.toLowerCase().includes('ratio')
        );
        if (hitSheet) {
            const data = XLSX.utils.sheet_to_json(workbook.Sheets[hitSheet]);
            data.forEach(row => {
                const miesiac = row['Miesiąc'] || row['Miesiac'] || row['Month'] || row['M'];
                const rok = row['Rok'] || row['Year'] || new Date().getFullYear();
                const imie = row['Delivery Lead'] || row['DL'] || row['Imię'] || row['Imie'];
                
                if (!imie || !miesiac) return;
                
                const zamkniete = parseInt(row['Zamknięte Requesty'] || row['Zamkniete'] || row['Closed'] || 0);
                const placements = parseInt(row['Placements'] || row['Placement'] || 0);
                
                result.hitRatio.push({
                    rok: parseInt(rok),
                    miesiac: parseInt(miesiac),
                    imie: String(imie).trim(),
                    zamkniete_requesty: zamkniete,
                    placements: placements,
                    hit_ratio: zamkniete > 0 ? Math.round((placements / zamkniete) * 100) : 0
                });
            });
        }
    });
    
    return result;
}

function importParsedData(parsedData) {
    let totalImported = 0;
    const importDetails = [];
    
    // Import rekrutacja
    if (parsedData.rekrutacja.length > 0) {
        parsedData.rekrutacja.forEach(record => {
            const tydzienId = getOrCreateTydzien(record.rok, record.tydzien);
            const osobaId = getOrCreateOsoba(record.imie, record.stanowisko);
            
            db.prepare(`
                INSERT INTO kpi_rekrutacja (osoba_id, tydzien_id, dni_pracy, weryfikacje, rekomendacje, cv_dodane, placements)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(osoba_id, tydzien_id) DO UPDATE SET
                    dni_pracy = excluded.dni_pracy, weryfikacje = excluded.weryfikacje,
                    rekomendacje = excluded.rekomendacje, cv_dodane = excluded.cv_dodane, placements = excluded.placements
            `).run(osobaId, tydzienId, record.dni_pracy, record.weryfikacje, record.rekomendacje, record.cv_dodane, record.placements);
        });
        importDetails.push({ type: 'rekrutacja', count: parsedData.rekrutacja.length });
        totalImported += parsedData.rekrutacja.length;
    }
    
    // Import sprzedaż
    if (parsedData.sprzedaz.length > 0) {
        parsedData.sprzedaz.forEach(record => {
            const tydzienId = getOrCreateTydzien(record.rok, record.tydzien);
            const osobaId = getOrCreateOsoba(record.imie, record.stanowisko);
            
            db.prepare(`
                INSERT INTO kpi_sprzedaz (osoba_id, tydzien_id, dni_pracy, leady, oferty, mrr)
                VALUES (?, ?, ?, ?, ?, ?)
                ON CONFLICT(osoba_id, tydzien_id) DO UPDATE SET
                    dni_pracy = excluded.dni_pracy, leady = excluded.leady,
                    oferty = excluded.oferty, mrr = excluded.mrr
            `).run(osobaId, tydzienId, record.dni_pracy, record.leady, record.oferty, record.mrr);
        });
        importDetails.push({ type: 'sprzedaż', count: parsedData.sprzedaz.length });
        totalImported += parsedData.sprzedaz.length;
    }
    
    // Import hit ratio
    if (parsedData.hitRatio.length > 0) {
        parsedData.hitRatio.forEach(record => {
            const osobaId = getOrCreateOsoba(record.imie, 'Delivery Lead');
            
            db.prepare(`
                INSERT INTO hit_ratio (osoba_id, rok, miesiac, zamkniete_requesty, placements, hit_ratio)
                VALUES (?, ?, ?, ?, ?, ?)
                ON CONFLICT(osoba_id, rok, miesiac) DO UPDATE SET
                    zamkniete_requesty = excluded.zamkniete_requesty, placements = excluded.placements, hit_ratio = excluded.hit_ratio
            `).run(osobaId, record.rok, record.miesiac, record.zamkniete_requesty, record.placements, record.hit_ratio);
        });
        importDetails.push({ type: 'hit_ratio', count: parsedData.hitRatio.length });
        totalImported += parsedData.hitRatio.length;
    }
    
    const periods = Array.from(parsedData.periods).map(p => {
        const [rok, tydzien] = p.split('-');
        return { rok: parseInt(rok), tydzien: parseInt(tydzien) };
    }).sort((a, b) => a.rok === b.rok ? a.tydzien - b.tydzien : a.rok - b.rok);
    
    return { imported: totalImported, details: importDetails, periods };
}

// ============ DATA GETTERS ============

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

function getAverageData() {
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
            CASE WHEN SUM(k.dni_pracy) > 0 THEN ROUND(CAST(SUM(k.cv_dodane) AS FLOAT) / SUM(k.dni_pracy), 2) ELSE 0 END as cv_per_day
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
            placements: rekrutacjaCount.total * (getTarget('placements') / 4),
            peopleCount: rekrutacjaCount
        },
        sprzedaz: {
            leady: sprzedazCount.SDR * getTarget('leady'),
            oferty: sprzedazCount.BDM * getTarget('oferty'),
            mrr: sprzedazCount['Head of Technology'] * getTarget('mrr'),
            peopleCount: sprzedazCount
        }
    };
}

// ============ AI FUNCTIONS (tylko analiza - NIE upload) ============

async function generateDepartmentHelp(dzial, rok, tydzien) {
    if (!ANTHROPIC_API_KEY) return { error: 'Brak klucza API' };
    
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
        
        context = `DZIAŁ REKRUTACJI - Tydzień ${tydzien}/${rok}

ZESPÓŁ (${data.length} osób):
${data.map(r => `- ${r.imie} (${r.stanowisko}): ${r.dni_pracy}dni, Weryf:${r.weryfikacje}, Reco:${r.rekomendacje}, CV:${r.cv_dodane}, Placements:${r.placements}`).join('\n')}

SUMY vs TARGETY:
- Weryfikacje: ${totals.weryfikacje} / ${tt.weryfikacje} (${tt.weryfikacje > 0 ? Math.round(totals.weryfikacje/tt.weryfikacje*100) : 0}%)
- Rekomendacje: ${totals.rekomendacje} / ${tt.rekomendacje} (${tt.rekomendacje > 0 ? Math.round(totals.rekomendacje/tt.rekomendacje*100) : 0}%)
- CV: ${totals.cv_dodane} / ${tt.cv_dodane} (${tt.cv_dodane > 0 ? Math.round(totals.cv_dodane/tt.cv_dodane*100) : 0}%)
- Placements: ${totals.placements}`;
    } else {
        const data = weekData.sprzedaz;
        const tt = teamTargets.sprzedaz;
        const totals = {
            leady: data.reduce((s, r) => s + (r.leady || 0), 0),
            oferty: data.reduce((s, r) => s + (r.oferty || 0), 0),
            mrr: data.reduce((s, r) => s + (r.mrr || 0), 0)
        };
        
        context = `DZIAŁ SPRZEDAŻY - Tydzień ${tydzien}/${rok}

ZESPÓŁ (${data.length} osób):
${data.map(r => `- ${r.imie} (${r.stanowisko}): ${r.dni_pracy}dni, MRR:${r.mrr}zł, Oferty:${r.oferty}, Leady:${r.leady}`).join('\n')}

SUMY vs TARGETY:
- MRR: ${totals.mrr}zł / ${tt.mrr}zł (${tt.mrr > 0 ? Math.round(totals.mrr/tt.mrr*100) : 0}%)
- Oferty: ${totals.oferty} / ${tt.oferty} (${tt.oferty > 0 ? Math.round(totals.oferty/tt.oferty*100) : 0}%)
- Leady: ${totals.leady} / ${tt.leady} (${tt.leady > 0 ? Math.round(totals.leady/tt.leady*100) : 0}%)`;
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
                model: 'claude-sonnet-4-20250514',
                max_tokens: 500,
                messages: [{
                    role: 'user',
                    content: `${context}

Jesteś ${dzial === 'rekrutacja' ? 'MINDY - asystentką rekrutacji' : 'INFRON - asystentem sprzedaży'}.
Napisz krótką analizę (max 120 słów):
1. 💚 Co idzie dobrze (max 2 punkty z imionami)
2. 🔴 Co wymaga uwagi (max 2 punkty z imionami)
3. 💡 Jedna rekomendacja

Bądź konkretny, używaj emoji.`
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
    if (!ANTHROPIC_API_KEY) return 'Brak klucza API. Dodaj ANTHROPIC_API_KEY.';
    
    const weekData = getCurrentWeekData(rok, tydzien);
    const targets = getTargets();
    const teamTargets = calculateTeamTargets(weekData, targets);
    
    const rekTotals = {
        weryfikacje: weekData.rekrutacja.reduce((s, r) => s + (r.weryfikacje || 0), 0),
        rekomendacje: weekData.rekrutacja.reduce((s, r) => s + (r.rekomendacje || 0), 0),
        cv_dodane: weekData.rekrutacja.reduce((s, r) => s + (r.cv_dodane || 0), 0),
        placements: weekData.rekrutacja.reduce((s, r) => s + (r.placements || 0), 0)
    };
    
    const salesTotals = {
        leady: weekData.sprzedaz.reduce((s, r) => s + (r.leady || 0), 0),
        oferty: weekData.sprzedaz.reduce((s, r) => s + (r.oferty || 0), 0),
        mrr: weekData.sprzedaz.reduce((s, r) => s + (r.mrr || 0), 0)
    };
    
    let context = `TYDZIEŃ ${tydzien}/${rok}

REKRUTACJA (${weekData.rekrutacja.length} osób):
${weekData.rekrutacja.map(r => `- ${r.imie} (${r.stanowisko}): Weryf:${r.weryfikacje}, Reco:${r.rekomendacje}, CV:${r.cv_dodane}, Place:${r.placements}`).join('\n')}
SUMY: Weryf:${rekTotals.weryfikacje}/${teamTargets.rekrutacja.weryfikacje}, Reco:${rekTotals.rekomendacje}/${teamTargets.rekrutacja.rekomendacje}, CV:${rekTotals.cv_dodane}/${teamTargets.rekrutacja.cv_dodane}

SPRZEDAŻ (${weekData.sprzedaz.length} osób):
${weekData.sprzedaz.map(s => `- ${s.imie} (${s.stanowisko}): MRR:${s.mrr}zł, Oferty:${s.oferty}, Leady:${s.leady}`).join('\n')}
SUMY: MRR:${salesTotals.mrr}/${teamTargets.sprzedaz.mrr}zł, Oferty:${salesTotals.oferty}/${teamTargets.sprzedaz.oferty}, Leady:${salesTotals.leady}/${teamTargets.sprzedaz.leady}`;

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
                max_tokens: 700,
                messages: [{
                    role: 'user',
                    content: `${context}

Napisz analizę (max 200 słów):
## 📊 PODSUMOWANIE
## ✅ SUKCESY (max 3)
## ⚠️ DO POPRAWY (max 3)
## 💡 REKOMENDACJE (2)

Używaj imion i konkretnych liczb.`
                }]
            })
        });

        const result = await response.json();
        if (result.content && result.content[0]) return result.content[0].text;
    } catch (err) {
        console.error('Analysis error:', err);
    }
    return 'Błąd generowania analizy.';
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
        SELECT a.tresc FROM analizy a 
        JOIN tygodnie t ON a.tydzien_id = t.id 
        WHERE t.rok = ? AND t.tydzien = ?
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

// Upload files (BEZ AI - prosty parser stałego szablonu)
app.post('/admin/upload', upload.array('files', 10), (req, res) => {
    if (req.body.password !== ADMIN_PASSWORD) {
        return res.status(401).json({ error: 'Nieprawidłowe hasło' });
    }
    
    if (!req.files || req.files.length === 0) {
        return res.status(400).json({ error: 'Brak plików' });
    }
    
    try {
        const filePaths = req.files.map(f => f.path);
        const parsedData = parseExcelFiles(filePaths);
        const importResult = importParsedData(parsedData);
        
        // Log
        const periodsStr = importResult.periods.map(p => `T${p.tydzien}/${p.rok}`).join(', ');
        db.prepare('INSERT INTO import_log (filename, records_imported, periods) VALUES (?, ?, ?)')
            .run(req.files.map(f => f.originalname).join(', '), importResult.imported, periodsStr);
        
        // Cleanup
        req.files.forEach(f => fs.unlinkSync(f.path));
        
        res.json({
            success: true,
            message: `Zaimportowano ${importResult.imported} rekordów`,
            details: importResult.details,
            periods: importResult.periods
        });
        
    } catch (err) {
        req.files.forEach(f => { try { fs.unlinkSync(f.path); } catch(e) {} });
        res.status(500).json({ error: err.message });
    }
});

// AI Analysis (osobny przycisk - kosztuje kredyty)
app.post('/api/analyze', async (req, res) => {
    const rok = parseInt(req.body.rok) || new Date().getFullYear();
    const tydzien = parseInt(req.body.tydzien) || getWeekNumber(new Date());
    
    const analysis = await generateFullAnalysis(rok, tydzien);
    
    const tydzienId = getOrCreateTydzien(rok, tydzien);
    db.prepare('INSERT INTO analizy (tydzien_id, tresc) VALUES (?, ?)').run(tydzienId, analysis);
    
    res.json({ analysis });
});

// AI Help (przycisk "Jak mogę pomóc?" - kosztuje kredyty)
app.post('/api/help/:dzial', async (req, res) => {
    const dzial = req.params.dzial;
    const rok = parseInt(req.body.rok) || new Date().getFullYear();
    const tydzien = parseInt(req.body.tydzien) || getWeekNumber(new Date());
    
    const result = await generateDepartmentHelp(dzial, rok, tydzien);
    res.json(result);
});

// Delete week
app.delete('/admin/data/:rok/:tydzien', (req, res) => {
    if (req.query.password !== ADMIN_PASSWORD) {
        return res.status(401).json({ error: 'Nieprawidłowe hasło' });
    }
    
    const rok = parseInt(req.params.rok);
    const tydzien = parseInt(req.params.tydzien);
    
    const tydzienRow = db.prepare('SELECT id FROM tygodnie WHERE rok = ? AND tydzien = ?').get(rok, tydzien);
    if (!tydzienRow) return res.status(404).json({ error: 'Brak danych' });
    
    const deleted = db.prepare('DELETE FROM kpi_rekrutacja WHERE tydzien_id = ?').run(tydzienRow.id).changes +
                    db.prepare('DELETE FROM kpi_sprzedaz WHERE tydzien_id = ?').run(tydzienRow.id).changes;
    db.prepare('DELETE FROM analizy WHERE tydzien_id = ?').run(tydzienRow.id);
    
    res.json({ success: true, message: `Usunięto ${deleted} rekordów` });
});

// History
app.get('/admin/history', (req, res) => {
    res.json(db.prepare('SELECT * FROM import_log ORDER BY created_at DESC LIMIT 50').all());
});

// Targets
app.get('/api/targets', (req, res) => res.json(getTargets()));

app.listen(PORT, () => {
    console.log(`
🤖 MINDY & INFRON Dashboard v4
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
🌐 Dashboard: http://localhost:${PORT}
🔐 Admin:     http://localhost:${PORT}/admin
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📤 Upload: BEZ AI (stały szablon)
🤖 AI: tylko analiza i pomoc
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    `);
});
