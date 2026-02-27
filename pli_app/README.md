# PLI / RPLI vs APT Comparison Dashboard

A local web application that compares McCamish (PLI/RPLI) reports against APT
data, stores everything in a SQLite database, and lets you re-run comparisons
without re-uploading files.

---

## Quick Start

### 1. Install Python dependency

```bash
pip install flask
```

### 2. Run the server

```bash
python app.py
```

You will see:
```
==================================================
 PLI/RPLI Comparison Server
 http://localhost:5000
==================================================
```

### 3. Open the dashboard

Open your browser and go to:  **http://localhost:5000**

The database file `pli_comparison.db` is created automatically in the same
folder as `app.py` on first run.

---

## File Structure

```
pli_app/
├── app.py                  ← Flask backend (run this)
├── requirements.txt
├── pli_comparison.db       ← SQLite DB (auto-created)
├── README.md
└── templates/
    └── index.html          ← Frontend dashboard
```

---

## How to Use

### Step 1 — Load & Save to DB  (one-time per data refresh)

Upload these three files and click **⬆ Load All & Save to DB**:

| # | File | Key columns needed |
|---|------|--------------------|
| 1 | **Master Data** | MCC Name, APT Name, APT Code, Office Type |
| 2 | **Account Code Classification** | Category (PLI/RPLI), Account_Code |
| 3 | **APT Report** | Office ID, Source Transaction Date, GL Code, Total Amount, Credit/Debit |

All data is saved into the SQLite database and persists across sessions.
Re-uploading replaces the existing data for that table entirely.

### Step 2 — Upload Report & Compare

1. Select **PLI** or **RPLI** using the toggle at the top
2. Upload the PLI or RPLI report (+ optional BODI Duplicate Office file)
3. Click **⚡ Save & Run Comparison**

The report rows are saved to the DB (primary key = receipt number), then
comparison runs via SQL against the APT data filtered by GL Code classification.

### Re-run Without Re-uploading

Once data is in the DB, click **↺ Re-run** on the results panel any time —
no file upload needed. Change date/month filters, office type, or
"Show Differences Only" and re-run instantly.

---

## Primary Keys

| Table | Primary Key |
|-------|-------------|
| `pli` | `receipt_number` |
| `rpli` | `receipt_number` |
| `apt` | `(posting_date, office_id, gl_code, total_amount)` |
| `master` | `(mcc_name, apt_code)` |
| `acc_classification` | `account_code` |

---

## Database Tab

Switch to the **🗄 Database** tab to:

- See row counts and last upload time for every table
- **View** — browse table contents (paginated, 100 rows/page)
- **Clear** — delete all rows from a specific table
- **↓ Export .db File** — download the full SQLite database for backup

---

## Browse Data Tab

Switch to **🔍 Browse Data**, pick any table from the dropdown, and click
**Load** to page through the stored rows.

---

## Export Options

| Button | Output |
|--------|--------|
| ↓ Full Report (XLSX) | All rows for current report type |
| ↓ Differences Only (XLSX) | Only rows where Difference ≠ 0 |
| ↓ PLI + RPLI Differences (CSV) | Both types, differences only, two sections |

---

## Column Mapping

### PLI / RPLI Report
| Column header in file | Stored as |
|-----------------------|-----------|
| OFFICE NAME | office_name |
| RECEIPT (number) | receipt_number ← **primary key** |
| EFFECTIVE DATE | effective_date |
| GROSS | gross_amount |

### APT Report
| Column header in file | Stored as |
|-----------------------|-----------|
| OFFICE ID | office_id |
| SOURCE TRANSACTION DATE | posting_date |
| GL Code (col H) | gl_code |
| TOTAL AMOUNT | total_amount |
| CREDIT/DEBIT | credit_debit |

### Account Code Classification
| Column header | Purpose |
|---------------|---------|
| Category | "PLI" or "RPLI" — determines which GL codes belong to which type |
| Account_Code | Matched against GL Code in APT rows |

---

## BODI Duplicate Office Logic

When the PLI/RPLI report contains office **BODINAICKENPATTI** and the master
file maps it to two APT codes, receipts are split:
- Receipts found in the BODI Duplicate Office file → APT code `29103265`
- Remaining receipts → APT code `29103266`
