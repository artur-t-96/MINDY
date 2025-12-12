# 🤖 MINDY & INFRON Dashboard v3

**Inteligentny dashboard KPI z AI-powered rozpoznawaniem danych Excel**

## ✨ Co nowego w v3?

- 🧠 **AI rozpoznaje strukturę** — wgrywasz dowolny Excel, Claude sam rozpozna co to za dane
- 📤 **Multi-file upload** — wgraj kilka plików naraz za jeden okres
- 🎯 **Automatyczna kategoryzacja** — AI przypisuje dane do właściwych tabel
- 💾 **SQLite na lata** — wszystkie dane historyczne zapisane permanentnie

---

## 🚀 Deploy na Render.com

### 1. Wgraj do GitHub

```
mindy-v3/
├── server.js
├── package.json
├── public/
│   ├── index.html
│   └── admin.html
```

### 2. Stwórz Web Service na Render

- **Build Command**: `npm install`
- **Start Command**: `npm start`

### 3. Dodaj Environment Variables

```
ADMIN_PASSWORD = TwojeHasloAdmina
ANTHROPIC_API_KEY = sk-ant-api03-...
```

⚠️ **ANTHROPIC_API_KEY jest wymagany** — bez niego AI nie będzie działać!

### 4. Dodaj Disk (ważne!)

- **Mount Path**: `/opt/render/project/src/data`
- **Size**: 1 GB

---

## 📤 Jak używać?

### 1. Wejdź na /admin

### 2. Wybierz tydzień i rok

### 3. Przeciągnij pliki Excel

Możesz wgrać **dowolne pliki** — AI sam rozpozna:
- Dane rekrutacji (Sourcer, Rekruter, TAC, Delivery Lead)
- Dane sprzedaży (SDR, BDM, Head of Technology)
- Hit Ratio Delivery Leadów
- Prep Calls z dynamiczną checklistą

### 4. Kliknij "Analizuj i importuj"

AI:
1. Przeczyta wszystkie pliki i arkusze
2. Rozpozna strukturę kolumn
3. Zaimportuje dane do właściwych tabel
4. Wygeneruje analizę dashboardu

---

## 🧠 Przykłady rozpoznawania

AI rozpozna kolumny nawet jeśli nazywają się inaczej:

| W Excelu | AI rozpozna jako |
|----------|------------------|
| CV sprawdzone | weryfikacje |
| Rekomendacje wysłane | rekomendacje |
| CV do bazy | cv_dodane |
| Zatrudnienia | placements |
| Nowe leady | leady |
| Wysłane propozycje | oferty |
| Przychód | mrr |
| src | Sourcer |
| DL | Delivery Lead |

---

## 🤖 Maskotki

| | MINDY | INFRON |
|---|---|---|
| **Dział** | Rekrutacja | Sprzedaż |
| **Kolor** | 💙 Niebieski | 🧡 Pomarańczowy |
| **Styl** | Żeński | Męski |

### KPI MINDY (Rekrutacja)

| Stanowisko | KPI | Target |
|------------|-----|--------|
| Sourcer | Weryfikacje | 20/tydzień |
| Sourcer | Rekomendacje | 15/tydzień |
| Rekruter | CV do bazy | 25/tydzień |
| Wszyscy | Placements | 1/miesiąc |
| Delivery Lead | Hit Ratio | min 30% |

### KPI INFRON (Sprzedaż)

| Stanowisko | KPI | Target |
|------------|-----|--------|
| SDR | Leady | 10/tydzień |
| BDM | Oferty | 1/tydzień |
| Head of Technology | MRR | 4000 zł/tydzień |

---

## 📊 Baza danych

SQLite przechowuje wszystko na lata:

- **osoby** — pracownicy
- **tygodnie** — kalendarz
- **kpi_rekrutacja** — dane rekrutacji per osoba/tydzień
- **kpi_sprzedaz** — dane sprzedaży per osoba/tydzień
- **hit_ratio** — miesięczne dane DL
- **prep_calls** — wszystkie prep calls
- **targety** — historia targetów
- **analizy** — historia AI analiz
- **import_log** — historia importów

---

## 🔐 Bezpieczeństwo

- Dashboard `/` — publiczny (cały zespół)
- Admin `/admin` — chroniony hasłem

---

Made with 💙🧡 for InfraMinds
