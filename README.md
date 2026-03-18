# 🤖 Quip Spreadsheet Automation — Amazon Connect Defects

**Automated Python pipeline that replaces a 50-minute daily manual process with a 3-minute script execution — saving ~205 hours per year.**

---

## 📋 Overview

The GO-AI team at Amazon VAR managed Connect defects through a shared Quip Spreadsheet updated manually once per day, 5 days a week. A team member would download a CSV from QuickSight, clean it, compare it against the Quip, and paste rows manually — a process prone to human error, data lag, and single-person dependency.

This script fully automates that pipeline: it reads the raw QuickSight export, intelligently compares it against the live Quip sheet, and produces a clean Excel file ready to import — in under 3 minutes.

---

## 📊 Business Impact

| Metric | Before | After |
|---|---|---|
| Time per execution | ~50 minutes | ~3 minutes |
| Frequency | 5× per week (daily) | On-demand |
| Weekly time cost | ~250 minutes (~4.2 hours) | ~15 minutes |
| Annual time cost | ~218 hours | ~13 hours |
| **Time saved per year** | **~205 hours** | — |
| Error rate | High (manual copy-paste) | Near zero |
| Team dependency | 1 person | Anyone on the team |
| Data freshness | Updated once per day | On-demand |

> **94% reduction in preparation time per execution. ~205 hours saved per year. 100% elimination of copy-paste errors.**

---

## ⚙️ How It Works

The script runs in 3 phases:

### Phase 1 — CSV Cleaning
- Reads the latest CSV exported from Amazon QuickSight
- Filters only real defects (`Real Defects_1 = 1`)
- Drops unnecessary columns and renames them to match Quip format
- Auto-detects the week number from the `Week` column

### Phase 2 — Intelligent Quip Comparison
- Connects to the Quip API and fetches the spreadsheet HTML
- Auto-identifies the correct weekly sheet (e.g. `Week 11 (2026)`) by matching the week from the CSV
- Compares data using 3 logical rules:
  1. **New rows** — Job IDs in CSV but not in Quip → added to Excel for import
  2. **Missing SharePoint links** — Job IDs already in Quip, CSV has a SharePoint link but Quip doesn't → listed separately for manual update
  3. **Removed defects** — Job IDs no longer in CSV → safely ignored

### Phase 3 — Excel Output
Generates a clean `.xlsx` file with 2 sheets:
- **Sheet 1 (`Week X 2026`)** — New rows to add to Quip, sorted A→Z by OSE login
- **Sheet 2 (`SharePoint Faltante`)** — Existing rows missing a SharePoint link, sorted A→Z by OSE login

---

## 🛠️ Tech Stack

| Component | Detail |
|---|---|
| Language | Python 3.x |
| Libraries | `pandas`, `requests`, `beautifulsoup4`, `lxml`, `openpyxl` |
| Data source | CSV export from Amazon QuickSight |
| Destination | Quip Spreadsheet via Quip API |
| Authentication | Quip Personal Token |
| Merge key | Job ID (unique UUID per defect) |
| Sheet detection | `title` attribute of Quip HTML table |

---

## 🚀 Setup

### Prerequisites

```bash
pip install pandas requests beautifulsoup4 lxml openpyxl
```

### Configuration

Open `connect_script.py` and update the following constants:

```python
QUIP_TOKEN    = "YOUR_TOKEN_HERE"      # From https://quip-amazon.com/api/personal-token
QUIP_THREAD_ID = "YOUR_THREAD_ID"     # From your Quip spreadsheet URL
CSV_FOLDER    = r"C:\Users\YOUR_USER\Desktop\CSV Connect"
OUTPUT_FOLDER = r"C:\Users\YOUR_USER\Desktop\CSV Connect"
```

### Folder Structure

```
Desktop/
└── CSV Connect/
    ├── connect_script.py
    ├── <latest_quicksight_export>.csv    ← place here before running
    └── Connect_Week_X_2026.xlsx          ← generated output
```

---

## ▶️ Usage

1. Download the latest CSV from QuickSight and place it in the `CSV Connect` folder
2. Open `cmd` (Windows + R → `cmd`)
3. Run:

```bash
python "C:\Users\YOUR_USER\Desktop\CSV Connect\connect_script.py"
```

4. Expected output:

```
🚀 Iniciando proceso...

─── FASE 1: Limpiando CSV ───
📥 Leyendo archivo: connect_export.csv
   Filas originales: 342
   Semana detectada: 11
   Filas después de filtrar Real Defects_1=1: 198
✅ CSV limpio con 198 filas y 6 columnas

─── FASE 2: Cargando Quip ───
   Buscando hoja: Week 11 (2026)
✅ Hoja 'Week 11 (2026)' cargada: 174 filas

─── FASE 3: Generando Excel ───
➕ Filas nuevas para agregar al Quip: 24
🔴 Filas con SharePoint faltante: 11
✅ Excel guardado en: ...\Connect_Week_11_2026.xlsx
   Hoja 1 - Week 11 (2026): 24 filas nuevas
   Hoja 2 - SharePoint Faltante: 11 filas
```

---

## 🐛 Troubleshooting

| Error | Cause | Fix |
|---|---|---|
| `❌ No se encontró ningún archivo CSV` | No CSV in folder | Download the QuickSight export and place it in `CSV Connect` |
| `❌ Error conectando al Quip: 401` | Token expired or invalid | Generate a new token at `quip-amazon.com/api/personal-token` |
| `❌ No se encontró la hoja 'Week X'` | Sheet doesn't exist in Quip | Create the sheet manually with the exact name `Week X (2026)` |
| `KeyError: 'column_name'` | CSV column names changed | Verify the QuickSight export has the expected column names |

---

## 📁 Project Structure

```
quip-automation/
├── connect_script.py    # Main automation script
├── README.md            # This file
└── .env.example         # Token config template (never commit your actual token)
```

---

## 🔐 Security Note

Never commit your `QUIP_TOKEN` to a public repository. Use a `.env` file or environment variables and add it to `.gitignore`:

```
# .gitignore
.env
```

---

## 💡 Key Learnings

- **Quip API** returns spreadsheet data as raw HTML — parsing it with `BeautifulSoup` and targeting the `title` attribute of each `<table>` was the most reliable approach to handle multi-sheet documents
- **Smart merge logic** based on Job ID (UUID) ensures zero duplicates and clean diffs between data sources
- **Pandas** handles all CSV-to-DataFrame transformations with minimal code and maximum clarity
- Designing for **team independence** (any member can run it) was a core requirement that shaped the UX of the script output

---

## 👤 Author

**Koichi Rodriguez** — Data Analyst @ Amazon VAR | GO-AI Team  
Built with Python + Atlas (GO-AI Team AI Assistant) | March 2026
