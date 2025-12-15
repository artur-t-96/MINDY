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

    -- Rada Nadzorcza KPI (matematyczne)
    CREATE TABLE IF NOT EXISTS kpi_rada_matematyczne (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        rok INTEGER NOT NULL,
        miesiac INTEGER NOT NULL,
        tydzien INTEGER,
        kpi_code TEXT NOT NULL,
        wartosc REAL DEFAULT 0,
        target REAL DEFAULT 0,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(rok, miesiac, tydzien, kpi_code)
    );

    -- Rada Nadzorcza KPI (opisowe)
    CREATE TABLE IF NOT EXISTS kpi_rada_opisowe (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        rok INTEGER NOT NULL,
        miesiac INTEGER NOT NULL,
        kpi_code TEXT NOT NULL,
        opis TEXT,
        status TEXT CHECK(status IN ('green', 'yellow', 'red')) DEFAULT 'yellow',
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(rok, miesiac, kpi_code)
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
        dzial TEXT,
        tresc TEXT,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (tydzien_id) REFERENCES tygodnie(id)
    );

    CREATE TABLE IF NOT EXISTS import_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        filename TEXT,
        records_imported INTEGER,
        periods TEXT,
        panel TEXT DEFAULT 'all',
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP
    );
`);

// Migrations
try { db.exec(`ALTER TABLE import_log ADD COLUMN periods TEXT`); } catch (e) {}
try { db.exec(`ALTER TABLE import_log ADD COLUMN panel TEXT DEFAULT 'all'`); } catch (e) {}

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

// KPI definitions for Rada Nadzorcza
const RADA_KPI = {
    matematyczne: {
        tygodniowe: [
            { code: 'TD-01', name: 'Sourcer Daily Verification Amount', target: 20 },
            { code: 'TD-02', name: 'Job Post Coverage', target: 80 },
            { code: 'TD-03', name: 'Recruiter Weekly New CV Upload', target: 25 }
        ],
        miesieczne: [
            { code: 'CS-02', name: 'Technical Verification Rate', target: 90 },
            { code: 'CS-03', name: 'Champion Advertisement Rate', target: 50 },
            { code: 'CS-04', name: 'Candidate Follow-up Frequency', target: 100 },
            { code: 'IM-02', name: 'Prep Call Completion Rate', target: 95 },
            { code: 'IM-05', name: 'Feedback After Interview', target: 80 },
            { code: 'DL-HR', name: 'DL Hit Ratio', target: 30 }
        ]
    },
    opisowe: [
        { code: 'DD-01', name: 'Profile Completion Rate' },
        { code: 'CS-01', name: 'Rejection Justification Timeliness' }
    ]
};

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

// ============ EXCEL PARSERS ============

function parseBodyLeasingExcel(filePaths) {
    const result = { rekrutacja: [], periods: new Set() };
    
    filePaths.forEach(filePath => {
        const workbook = XLSX.readFile(filePath);
        const rekSheet = workbook.SheetNames.find(n => 
            n.toLowerCase().includes('rekrutacja') || n.toLowerCase() === 'recruitment' || n.toLowerCase() === 'body'
        ) || workbook.SheetNames[0];
        
        if (rekSheet) {
            const data = XLSX.utils.sheet_to_json(workbook.Sheets[rekSheet]);
            data.forEach(row => {
                const tydzien = row['Tydzien'] || row['Tydzień'] || row['Week'] || row['T'];
                const rok = row['Rok'] || row['Year'] || new Date().getFullYear();
                const imie = row['Imie'] || row['Imię'] || row['Name'] || row['Pracownik'];
                const stanowisko = row['Stanowisko'] || row['Rola'] || row['Role'];
                
                if (!imie || !tydzien) return;
                
                result.rekrutacja.push({
                    rok: parseInt(rok), tydzien: parseInt(tydzien),
                    imie: String(imie).trim(), stanowisko: normalizeStanowisko(stanowisko),
                    dni_pracy: parseFloat(row['Dni pracy'] || row['Dni'] || 5),
                    weryfikacje: parseInt(row['Weryfikacje'] || row['Weryf'] || 0),
                    rekomendacje: parseInt(row['Rekomendacje'] || row['Reco'] || 0),
                    cv_dodane: parseInt(row['CV do bazy'] || row['CV'] || 0),
                    placements: parseInt(row['Placements'] || 0)
                });
                result.periods.add(`${rok}-${tydzien}`);
            });
        }
    });
    return result;
}

function parseSprzedazExcel(filePaths) {
    const result = { sprzedaz: [], periods: new Set() };
    
    filePaths.forEach(filePath => {
        const workbook = XLSX.readFile(filePath);
        const salesSheet = workbook.SheetNames.find(n => 
            n.toLowerCase().includes('sprzeda') || n.toLowerCase() === 'sales'
        ) || workbook.SheetNames[0];
        
        if (salesSheet) {
            const data = XLSX.utils.sheet_to_json(workbook.Sheets[salesSheet]);
            data.forEach(row => {
                const tydzien = row['Tydzien'] || row['Tydzień'] || row['Week'] || row['T'];
                const rok = row['Rok'] || row['Year'] || new Date().getFullYear();
                const imie = row['Imie'] || row['Imię'] || row['Name'] || row['Pracownik'];
                const stanowisko = row['Stanowisko'] || row['Rola'] || row['Role'];
                
                if (!imie || !tydzien) return;
                
                result.sprzedaz.push({
                    rok: parseInt(rok), tydzien: parseInt(tydzien),
                    imie: String(imie).trim(), stanowisko: normalizeStanowisko(stanowisko),
                    dni_pracy: parseFloat(row['Dni pracy'] || row['Dni'] || 5),
                    leady: parseInt(row['Leady'] || row['Leads'] || 0),
                    oferty: parseInt(row['Oferty'] || row['Wysłane oferty'] || 0),
                    mrr: parseFloat(row['MRR'] || row['Revenue'] || 0)
                });
                result.periods.add(`${rok}-${tydzien}`);
            });
        }
    });
    return result;
}

function parseRadaExcel(filePaths) {
    const result = { hitRatio: [], kpiMat: [], kpiOpisowe: [], periods: new Set() };
    
    filePaths.forEach(filePath => {
        const workbook = XLSX.readFile(filePath);
        
        // Hit Ratio sheet
        const hitSheet = workbook.SheetNames.find(n => n.toLowerCase().includes('hit') || n.toLowerCase().includes('ratio'));
        if (hitSheet) {
            const data = XLSX.utils.sheet_to_json(workbook.Sheets[hitSheet]);
            data.forEach(row => {
                const miesiac = row['Miesiac'] || row['Miesiąc'] || row['Month'];
                const rok = row['Rok'] || row['Year'] || new Date().getFullYear();
                const imie = row['Delivery Lead'] || row['DL'] || row['Imie'] || row['Imię'];
                if (!imie || !miesiac) return;
                
                const zamkniete = parseInt(row['Zamkniete Requesty'] || row['Zamknięte Requesty'] || row['Closed'] || 0);
                const placements = parseInt(row['Placements'] || 0);
                
                result.hitRatio.push({
                    rok: parseInt(rok), miesiac: parseInt(miesiac),
                    imie: String(imie).trim(), zamkniete_requesty: zamkniete, placements,
                    hit_ratio: zamkniete > 0 ? Math.round((placements / zamkniete) * 100) : 0
                });
            });
        }
        
        // KPI Matematyczne sheet
        const kpiMatSheet = workbook.SheetNames.find(n => n.toLowerCase().includes('kpi') && n.toLowerCase().includes('mat'));
        if (kpiMatSheet) {
            const data = XLSX.utils.sheet_to_json(workbook.Sheets[kpiMatSheet]);
            data.forEach(row => {
                const code = row['KPI'] || row['Code'] || row['Kod'];
                if (!code) return;
                
                result.kpiMat.push({
                    rok: parseInt(row['Rok'] || new Date().getFullYear()),
                    miesiac: parseInt(row['Miesiac'] || row['Miesiąc'] || 1),
                    tydzien: row['Tydzien'] || row['Tydzień'] ? parseInt(row['Tydzien'] || row['Tydzień']) : null,
                    kpi_code: code,
                    wartosc: parseFloat(row['Wartosc'] || row['Wartość'] || row['Value'] || 0),
                    target: parseFloat(row['Target'] || row['Cel'] || 0)
                });
            });
        }
        
        // KPI Opisowe sheet
        const kpiOpisSheet = workbook.SheetNames.find(n => n.toLowerCase().includes('kpi') && n.toLowerCase().includes('opis'));
        if (kpiOpisSheet) {
            const data = XLSX.utils.sheet_to_json(workbook.Sheets[kpiOpisSheet]);
            data.forEach(row => {
                const code = row['KPI'] || row['Code'] || row['Kod'];
                if (!code) return;
                
                result.kpiOpisowe.push({
                    rok: parseInt(row['Rok'] || new Date().getFullYear()),
                    miesiac: parseInt(row['Miesiac'] || row['Miesiąc'] || 1),
                    kpi_code: code,
                    opis: row['Opis'] || row['Description'] || '',
                    status: row['Status'] || 'yellow'
                });
            });
        }
    });
    return result;
}

// ============ IMPORT FUNCTIONS ============

function importBodyLeasingData(parsedData) {
    let imported = 0;
    parsedData.rekrutacja.forEach(r => {
        const tydzienId = getOrCreateTydzien(r.rok, r.tydzien);
        const osobaId = getOrCreateOsoba(r.imie, r.stanowisko);
        db.prepare(`
            INSERT INTO kpi_rekrutacja (osoba_id, tydzien_id, dni_pracy, weryfikacje, rekomendacje, cv_dodane, placements)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(osoba_id, tydzien_id) DO UPDATE SET
                dni_pracy=excluded.dni_pracy, weryfikacje=excluded.weryfikacje,
                rekomendacje=excluded.rekomendacje, cv_dodane=excluded.cv_dodane, placements=excluded.placements
        `).run(osobaId, tydzienId, r.dni_pracy, r.weryfikacje, r.rekomendacje, r.cv_dodane, r.placements);
        imported++;
    });
    const periods = Array.from(parsedData.periods).map(p => { const [rok, tydzien] = p.split('-'); return { rok: parseInt(rok), tydzien: parseInt(tydzien) }; });
    return { imported, periods };
}

function importSprzedazData(parsedData) {
    let imported = 0;
    parsedData.sprzedaz.forEach(r => {
        const tydzienId = getOrCreateTydzien(r.rok, r.tydzien);
        const osobaId = getOrCreateOsoba(r.imie, r.stanowisko);
        db.prepare(`
            INSERT INTO kpi_sprzedaz (osoba_id, tydzien_id, dni_pracy, leady, oferty, mrr)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(osoba_id, tydzien_id) DO UPDATE SET
                dni_pracy=excluded.dni_pracy, leady=excluded.leady, oferty=excluded.oferty, mrr=excluded.mrr
        `).run(osobaId, tydzienId, r.dni_pracy, r.leady, r.oferty, r.mrr);
        imported++;
    });
    const periods = Array.from(parsedData.periods).map(p => { const [rok, tydzien] = p.split('-'); return { rok: parseInt(rok), tydzien: parseInt(tydzien) }; });
    return { imported, periods };
}

function importRadaData(parsedData) {
    let imported = 0;
    
    // Hit Ratio
    parsedData.hitRatio.forEach(r => {
        const osobaId = getOrCreateOsoba(r.imie, 'Delivery Lead');
        db.prepare(`
            INSERT INTO hit_ratio (osoba_id, rok, miesiac, zamkniete_requesty, placements, hit_ratio)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(osoba_id, rok, miesiac) DO UPDATE SET
                zamkniete_requesty=excluded.zamkniete_requesty, placements=excluded.placements, hit_ratio=excluded.hit_ratio
        `).run(osobaId, r.rok, r.miesiac, r.zamkniete_requesty, r.placements, r.hit_ratio);
        imported++;
    });
    
    // KPI Matematyczne
    parsedData.kpiMat.forEach(r => {
        db.prepare(`
            INSERT INTO kpi_rada_matematyczne (rok, miesiac, tydzien, kpi_code, wartosc, target)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(rok, miesiac, tydzien, kpi_code) DO UPDATE SET wartosc=excluded.wartosc, target=excluded.target
        `).run(r.rok, r.miesiac, r.tydzien, r.kpi_code, r.wartosc, r.target);
        imported++;
    });
    
    // KPI Opisowe
    parsedData.kpiOpisowe.forEach(r => {
        db.prepare(`
            INSERT INTO kpi_rada_opisowe (rok, miesiac, kpi_code, opis, status)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(rok, miesiac, kpi_code) DO UPDATE SET opis=excluded.opis, status=excluded.status
        `).run(r.rok, r.miesiac, r.kpi_code, r.opis, r.status);
        imported++;
    });
    
    return { imported };
}

// ============ DATA GETTERS ============

function getBodyLeasingData(rok, tydzien) {
    const tydzienRow = db.prepare('SELECT id FROM tygodnie WHERE rok = ? AND tydzien = ?').get(rok, tydzien);
    if (!tydzienRow) return { rekrutacja: [], hitRatio: [] };
    
    const rekrutacja = db.prepare(`
        SELECT o.imie, o.stanowisko, k.dni_pracy, k.weryfikacje, k.rekomendacje, k.cv_dodane, k.placements
        FROM kpi_rekrutacja k JOIN osoby o ON k.osoba_id = o.id
        WHERE k.tydzien_id = ?
    `).all(tydzienRow.id);
    
    const miesiac = Math.ceil(tydzien / 4.33);
    const hitRatio = db.prepare(`
        SELECT o.imie, h.zamkniete_requesty, h.placements, h.hit_ratio
        FROM hit_ratio h JOIN osoby o ON h.osoba_id = o.id
        WHERE h.rok = ? AND h.miesiac = ?
    `).all(rok, miesiac);
    
    return { rekrutacja, hitRatio };
}

function getSprzedazData(rok, tydzien) {
    const tydzienRow = db.prepare('SELECT id FROM tygodnie WHERE rok = ? AND tydzien = ?').get(rok, tydzien);
    if (!tydzienRow) return { sprzedaz: [] };
    
    const sprzedaz = db.prepare(`
        SELECT o.imie, o.stanowisko, k.dni_pracy, k.leady, k.oferty, k.mrr
        FROM kpi_sprzedaz k JOIN osoby o ON k.osoba_id = o.id
        WHERE k.tydzien_id = ?
    `).all(tydzienRow.id);
    
    return { sprzedaz };
}

function getRadaData(rok, miesiac) {
    const hitRatio = db.prepare(`
        SELECT o.imie, h.zamkniete_requesty, h.placements, h.hit_ratio
        FROM hit_ratio h JOIN osoby o ON h.osoba_id = o.id
        WHERE h.rok = ? AND h.miesiac = ?
    `).all(rok, miesiac);
    
    const kpiMat = db.prepare(`SELECT * FROM kpi_rada_matematyczne WHERE rok = ? AND miesiac = ?`).all(rok, miesiac);
    const kpiOpisowe = db.prepare(`SELECT * FROM kpi_rada_opisowe WHERE rok = ? AND miesiac = ?`).all(rok, miesiac);
    
    return { hitRatio, kpiMat, kpiOpisowe };
}

function getAverageData(dzial) {
    if (dzial === 'rekrutacja') {
        return db.prepare(`
            SELECT o.imie, o.stanowisko,
                SUM(k.placements) as total_placements, COUNT(DISTINCT k.tydzien_id) as tygodni,
                CASE WHEN SUM(k.dni_pracy) > 0 THEN ROUND(CAST(SUM(k.weryfikacje) AS FLOAT) / SUM(k.dni_pracy), 2) ELSE 0 END as weryf_per_day,
                CASE WHEN SUM(k.dni_pracy) > 0 THEN ROUND(CAST(SUM(k.rekomendacje) AS FLOAT) / SUM(k.dni_pracy), 2) ELSE 0 END as reco_per_day,
                CASE WHEN SUM(k.dni_pracy) > 0 THEN ROUND(CAST(SUM(k.cv_dodane) AS FLOAT) / SUM(k.dni_pracy), 2) ELSE 0 END as cv_per_day
            FROM kpi_rekrutacja k JOIN osoby o ON k.osoba_id = o.id
            GROUP BY o.id ORDER BY total_placements DESC, reco_per_day DESC
        `).all();
    } else {
        return db.prepare(`
            SELECT o.imie, o.stanowisko,
                COUNT(DISTINCT k.tydzien_id) as tygodni,
                CASE WHEN SUM(k.dni_pracy) > 0 THEN ROUND(CAST(SUM(k.leady) AS FLOAT) / SUM(k.dni_pracy), 2) ELSE 0 END as leady_per_day,
                CASE WHEN SUM(k.dni_pracy) > 0 THEN ROUND(CAST(SUM(k.oferty) AS FLOAT) / SUM(k.dni_pracy), 2) ELSE 0 END as oferty_per_day,
                CASE WHEN COUNT(DISTINCT k.tydzien_id) > 0 THEN ROUND(CAST(SUM(k.mrr) AS FLOAT) / COUNT(DISTINCT k.tydzien_id), 0) ELSE 0 END as mrr_per_week
            FROM kpi_sprzedaz k JOIN osoby o ON k.osoba_id = o.id
            GROUP BY o.id ORDER BY mrr_per_week DESC
        `).all();
    }
}

function getAvailableWeeks() {
    return db.prepare('SELECT rok, tydzien FROM tygodnie ORDER BY rok DESC, tydzien DESC LIMIT 100').all();
}

function calculateTeamTargets(data, targets, type) {
    if (type === 'rekrutacja') {
        const count = {
            Sourcer: data.filter(r => r.stanowisko === 'Sourcer').length,
            Rekruter: data.filter(r => r.stanowisko === 'Rekruter').length,
            total: data.length
        };
        const getTarget = (kpi) => (targets.find(t => t.kpi === kpi)?.wartosc || 0);
        return {
            weryfikacje: count.Sourcer * getTarget('weryfikacje'),
            rekomendacje: count.Sourcer * getTarget('rekomendacje'),
            cv_dodane: count.Rekruter * getTarget('cv_dodane'),
            placements: count.total * (getTarget('placements') / 4),
            peopleCount: count
        };
    } else {
        const count = {
            SDR: data.filter(r => r.stanowisko === 'SDR').length,
            BDM: data.filter(r => r.stanowisko === 'BDM').length,
            'Head of Technology': data.filter(r => r.stanowisko === 'Head of Technology').length,
            total: data.length
        };
        const getTarget = (kpi) => (targets.find(t => t.kpi === kpi)?.wartosc || 0);
        return {
            leady: count.SDR * getTarget('leady'),
            oferty: count.BDM * getTarget('oferty'),
            mrr: count['Head of Technology'] * getTarget('mrr'),
            peopleCount: count
        };
    }
}

// ============ ROUTES ============

// Landing page
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

// Panels
app.get('/body-leasing', (req, res) => res.sendFile(path.join(__dirname, 'public', 'body-leasing.html')));
app.get('/sprzedaz', (req, res) => res.sendFile(path.join(__dirname, 'public', 'sprzedaz.html')));
app.get('/rada-nadzorcza', (req, res) => res.sendFile(path.join(__dirname, 'public', 'rada-nadzorcza.html')));
app.get('/admin', (req, res) => res.sendFile(path.join(__dirname, 'public', 'admin.html')));

// API - Body Leasing
app.get('/api/body-leasing', (req, res) => {
    const rok = parseInt(req.query.rok) || new Date().getFullYear();
    const tydzien = parseInt(req.query.tydzien) || getWeekNumber(new Date());
    const data = getBodyLeasingData(rok, tydzien);
    const avgData = getAverageData('rekrutacja');
    const targets = getTargets();
    const teamTargets = calculateTeamTargets(data.rekrutacja, targets, 'rekrutacja');
    const weeks = getAvailableWeeks();
    res.json({ rok, tydzien, current: data, average: avgData, targets, teamTargets, weeks });
});

// API - Sprzedaz
app.get('/api/sprzedaz', (req, res) => {
    const rok = parseInt(req.query.rok) || new Date().getFullYear();
    const tydzien = parseInt(req.query.tydzien) || getWeekNumber(new Date());
    const data = getSprzedazData(rok, tydzien);
    const avgData = getAverageData('sprzedaz');
    const targets = getTargets();
    const teamTargets = calculateTeamTargets(data.sprzedaz, targets, 'sprzedaz');
    const weeks = getAvailableWeeks();
    res.json({ rok, tydzien, current: data, average: avgData, targets, teamTargets, weeks });
});

// API - Rada Nadzorcza
app.get('/api/rada-nadzorcza', (req, res) => {
    const rok = parseInt(req.query.rok) || new Date().getFullYear();
    const miesiac = parseInt(req.query.miesiac) || new Date().getMonth() + 1;
    const data = getRadaData(rok, miesiac);
    res.json({ rok, miesiac, data, kpiDefinitions: RADA_KPI });
});

// Upload endpoints
app.post('/admin/upload/body-leasing', upload.array('files', 10), (req, res) => {
    if (req.body.password !== ADMIN_PASSWORD) return res.status(401).json({ error: 'Nieprawidłowe hasło' });
    if (!req.files?.length) return res.status(400).json({ error: 'Brak plików' });
    
    try {
        const parsedData = parseBodyLeasingExcel(req.files.map(f => f.path));
        const result = importBodyLeasingData(parsedData);
        const periodsStr = result.periods.map(p => `T${p.tydzien}/${p.rok}`).join(', ');
        db.prepare('INSERT INTO import_log (filename, records_imported, periods, panel) VALUES (?, ?, ?, ?)')
            .run(req.files.map(f => f.originalname).join(', '), result.imported, periodsStr, 'body-leasing');
        req.files.forEach(f => fs.unlinkSync(f.path));
        res.json({ success: true, message: `Zaimportowano ${result.imported} rekordów`, periods: result.periods });
    } catch (err) {
        req.files.forEach(f => { try { fs.unlinkSync(f.path); } catch(e) {} });
        res.status(500).json({ error: err.message });
    }
});

app.post('/admin/upload/sprzedaz', upload.array('files', 10), (req, res) => {
    if (req.body.password !== ADMIN_PASSWORD) return res.status(401).json({ error: 'Nieprawidłowe hasło' });
    if (!req.files?.length) return res.status(400).json({ error: 'Brak plików' });
    
    try {
        const parsedData = parseSprzedazExcel(req.files.map(f => f.path));
        const result = importSprzedazData(parsedData);
        const periodsStr = result.periods.map(p => `T${p.tydzien}/${p.rok}`).join(', ');
        db.prepare('INSERT INTO import_log (filename, records_imported, periods, panel) VALUES (?, ?, ?, ?)')
            .run(req.files.map(f => f.originalname).join(', '), result.imported, periodsStr, 'sprzedaz');
        req.files.forEach(f => fs.unlinkSync(f.path));
        res.json({ success: true, message: `Zaimportowano ${result.imported} rekordów`, periods: result.periods });
    } catch (err) {
        req.files.forEach(f => { try { fs.unlinkSync(f.path); } catch(e) {} });
        res.status(500).json({ error: err.message });
    }
});

app.post('/admin/upload/rada-nadzorcza', upload.array('files', 10), (req, res) => {
    if (req.body.password !== ADMIN_PASSWORD) return res.status(401).json({ error: 'Nieprawidłowe hasło' });
    if (!req.files?.length) return res.status(400).json({ error: 'Brak plików' });
    
    try {
        const parsedData = parseRadaExcel(req.files.map(f => f.path));
        const result = importRadaData(parsedData);
        db.prepare('INSERT INTO import_log (filename, records_imported, periods, panel) VALUES (?, ?, ?, ?)')
            .run(req.files.map(f => f.originalname).join(', '), result.imported, '', 'rada-nadzorcza');
        req.files.forEach(f => fs.unlinkSync(f.path));
        res.json({ success: true, message: `Zaimportowano ${result.imported} rekordów` });
    } catch (err) {
        req.files.forEach(f => { try { fs.unlinkSync(f.path); } catch(e) {} });
        res.status(500).json({ error: err.message });
    }
});

// AI Help endpoints
app.post('/api/help/:panel', async (req, res) => {
    if (!ANTHROPIC_API_KEY) return res.json({ error: 'Brak klucza API' });
    
    const panel = req.params.panel;
    const rok = parseInt(req.body.rok) || new Date().getFullYear();
    const tydzien = parseInt(req.body.tydzien) || getWeekNumber(new Date());
    
    let context = '';
    let robotName = '';
    
    if (panel === 'body-leasing') {
        const data = getBodyLeasingData(rok, tydzien);
        robotName = 'MINDY - asystentka Body Leasing';
        context = `BODY LEASING - Tydzień ${tydzien}/${rok}\n${data.rekrutacja.map(r => `- ${r.imie} (${r.stanowisko}): Weryf:${r.weryfikacje}, Reco:${r.rekomendacje}, CV:${r.cv_dodane}, Place:${r.placements}`).join('\n')}`;
    } else if (panel === 'sprzedaz') {
        const data = getSprzedazData(rok, tydzien);
        robotName = 'INFRON - asystent Sprzedaży';
        context = `SPRZEDAŻ - Tydzień ${tydzien}/${rok}\n${data.sprzedaz.map(r => `- ${r.imie} (${r.stanowisko}): MRR:${r.mrr}zł, Oferty:${r.oferty}, Leady:${r.leady}`).join('\n')}`;
    }
    
    try {
        const response = await fetch('https://api.anthropic.com/v1/messages', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json', 'x-api-key': ANTHROPIC_API_KEY, 'anthropic-version': '2023-06-01' },
            body: JSON.stringify({
                model: 'claude-sonnet-4-20250514', max_tokens: 500,
                messages: [{ role: 'user', content: `${context}\n\nJesteś ${robotName}.\nNapisz krótką analizę (max 120 słów):\n1. 💚 Co idzie dobrze (max 2)\n2. 🔴 Co wymaga uwagi (max 2)\n3. 💡 Jedna rekomendacja\n\nUżywaj emoji.` }]
            })
        });
        const result = await response.json();
        res.json({ help: result.content?.[0]?.text || 'Brak odpowiedzi' });
    } catch (err) {
        res.json({ error: err.message });
    }
});

// History & misc
app.get('/admin/history', (req, res) => {
    res.json(db.prepare('SELECT * FROM import_log ORDER BY created_at DESC LIMIT 50').all());
});

app.get('/api/targets', (req, res) => res.json(getTargets()));

app.listen(PORT, () => {
    console.log(`
🤖 MINDY & INFRON Dashboard v5
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
🌐 Home:           http://localhost:${PORT}
💼 Body Leasing:   http://localhost:${PORT}/body-leasing
💰 Sprzedaż:       http://localhost:${PORT}/sprzedaz
👔 Rada Nadzorcza: http://localhost:${PORT}/rada-nadzorcza
🔐 Admin:          http://localhost:${PORT}/admin
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    `);
});
