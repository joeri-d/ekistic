# Lightfoot — CSV Mirror Structuur

**Project Existenzminimum 2.0** · Vlaamse context · 2026

---

## Doel van de CSV-mirrors

De `.xlsx`-bestanden zijn de primaire bron van het model.  
CSV-mirrors dienen als:
- auditeerbare tekstversie (leesbaar zonder Excel),
- versiehistorie in Git (diff-vergelijking mogelijk),
- export voor verdere verwerking of import in andere tools.

---

## Mapstructuur

```
lightfoot/spreadsheets/
│
├── leven_inventaris.xlsx            ← masterfile (dashboard + totalen)
├── kleding_systeem.xlsx
├── slaapsysteem.xlsx
├── eetsysteem.xlsx
├── hygienesysteem.xlsx
├── werksysteem.xlsx
├── beweging_sport.xlsx
├── mobiliteit.xlsx
├── wonen_shelter.xlsx
│
├── <functie>/
│   └── csv_mirror_snapshots/
│       └── YYYY-MM-DD_HHhMM/      ← snapshot met datum + uur
│           ├── filosofie.csv
│           ├── inventaris.csv
│           ├── verbruiksgoederen.csv
│           ├── onderhoud.csv
│           ├── budget.csv
│           └── afhankelijkheden.csv
│
├── csv_mirror/
│   └── README.md                  ← dit bestand
│
└── tools/
    ├── helpers.py                 ← gedeelde stijlfuncties
    ├── gen_kleding.py
    ├── gen_slaap.py
    ├── gen_eten.py
    ├── gen_hygiene.py
    ├── gen_werk.py
    ├── gen_beweging.py
    ├── gen_mobiliteit.py
    ├── gen_wonen.py
    ├── gen_leven_inventaris.py
    └── export_csv_mirrors.py      ← exportscript voor alle snapshots
```

---

## Snapshots aanmaken

```bash
# Met huidige datum en tijd:
python3 lightfoot/spreadsheets/tools/export_csv_mirrors.py

# Met specifieke datum en tijdlabel:
python3 lightfoot/spreadsheets/tools/export_csv_mirrors.py 2026-04-17 15h30
```

## Excel-bestanden regenereren

```bash
python3 lightfoot/spreadsheets/tools/gen_kleding.py
python3 lightfoot/spreadsheets/tools/gen_slaap.py
python3 lightfoot/spreadsheets/tools/gen_eten.py
python3 lightfoot/spreadsheets/tools/gen_hygiene.py
python3 lightfoot/spreadsheets/tools/gen_werk.py
python3 lightfoot/spreadsheets/tools/gen_beweging.py
python3 lightfoot/spreadsheets/tools/gen_mobiliteit.py
python3 lightfoot/spreadsheets/tools/gen_wonen.py
python3 lightfoot/spreadsheets/tools/gen_leven_inventaris.py
```

---

## Versiebeleid

- Nieuwe iteraties altijd als **nieuwe snapshot** toevoegen — nooit bestaande overschrijven.
- Tijdlabel in mapnaam: `YYYY-MM-DD_HHhMM` (bv. `2026-04-17_04h45`).
- De `.xlsx`-bestanden in de root van `spreadsheets/` zijn altijd de **meest recente versie**.

---

## Inhoud per functiebestand

Elk systeembestand bevat minimaal de tabbladen:

| Tabblad | Inhoud |
|---|---|
| Filosofie | Ontwerpprincipes, normen, uitzonderingen |
| Inventaris | Minimale itemlijst met TCO-berekening |
| Verbruiksgoederen | Maandelijkse verbruikskosten |
| Onderhoud | Onderhoudsschema met frequentie en duur |
| Budget | Gecombineerd TCO-overzicht + tijdsbelasting |
| Afhankelijkheden | Koppelingen met andere leeffuncties |

De masterfile `leven_inventaris.xlsx` bevat:

| Tabblad | Inhoud |
|---|---|
| Dashboard | Overzicht alle functies met investering + maandkost + tijd |
| Budget_master | Geconsolideerd budgetoverzicht + benchmark |
| Tijdsbelasting | Tijdsbesteding per functie |
| Afhankelijkheden | Kritieke koppelingen cross-functioneel |
