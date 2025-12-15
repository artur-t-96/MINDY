# 🤖 MINDY & INFRON Dashboard v4

Dashboard KPI z emocjonalnymi robotami i AI analizą.

## ✨ Funkcje

- 🎭 **Emocjonalne roboty** — twarz zmienia się zależnie od % targetu
- 💬 **Przycisk "Jak mogę pomóc?"** — AI analiza per dział
- 📤 **Multi-file upload** — wiele plików na raz
- 📅 **Auto-rozpoznawanie okresów** — AI sam rozpozna tygodnie/daty z Excel
- 🗑️ **Usuwanie danych** — można kasować stare uploady
- 🎯 **Targety zespołowe** — per osoba × liczba osób
- 📊 **Dwa rankingi**: Tydzień + Average (per working day)

## 📊 Rankingi

Dashboard pokazuje **dwa rankingi**:

1. **Ranking za wybrany tydzień** — aktualne wyniki
2. **Ranking Average** — historyczne średnie per working day

### Sortowanie Rekrutacja:
1. 🏆 Placements (najważniejsze)
2. Rekomendacje
3. CV + Weryfikacje

### Sortowanie Sprzedaż:
1. 💎 MRR (najważniejsze)
2. Wysłane oferty
3. Leady

## 📅 Automatyczne rozpoznawanie okresów

Excel może zawierać dane za **wiele tygodni/miesięcy**. AI automatycznie:
- Rozpoznaje kolumny z datami (np. "Tydzień", "Data", "Okres", "Week")
- Lub arkusze nazwane po tygodniach
- Każdy wiersz przypisuje do właściwego tygodnia

## 🚀 Deploy

### Render.com

**Environment Variables:**
```
ADMIN_PASSWORD = TwojeHasło
ANTHROPIC_API_KEY = sk-ant-...
```

**Disk (WAŻNE!):**
- Mount Path: `/opt/render/project/src/data`
- Size: 1 GB

## 🤖 Emocje robotów

| Wynik | Emocja |
|-------|--------|
| ≥120% | 🤩 Jestem niesamowita! |
| ≥100% | 😊 Świetnie się spisujemy! |
| ≥85% | 🙂 Idzie dobrze! |
| ≥70% | 😐 Może być lepiej... |
| ≥50% | 😟 Potrzebuję wsparcia |
| <50% | 😢 To trudny tydzień... |

## 🔐 Bezpieczeństwo

- Dashboard `/` — publiczny
- Admin `/admin` — wymaga hasła

---

Made with 💙🧡 for InfraMinds
