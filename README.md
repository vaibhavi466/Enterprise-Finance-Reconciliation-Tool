# 🏦 Enterprise Finance Reconciliation Tool

A production-style Python automation tool that reconciles financial records between an **ERP / System of Record** and an **External / Bank file** — classifying every record with a clear status and generating a business-friendly multi-sheet Excel report.

---

## 📌 Overview

Manual reconciliation between ERP systems and bank statements is one of the most time-consuming and error-prone tasks in finance operations. This tool automates that process end-to-end: ingesting two data sources, sanitizing input formatting, isolating data quality issues, intelligently matching records, and producing a structured Excel output ready for executive audit.

---

## ✅ Classification Categories

| Category | Description |
|---|---|
| ✅ **Matches** | Record is fully aligned across all shared fields in both systems |
| ⚠️ **Mismatches** | Record exists in both systems but differs in Amount, Currency, Date, etc. |
| ❌ **Missing** | Record is present in one system but absent in the other |
| 🔁 **Data Issues** | Duplicate primary keys detected and isolated prior to matching |

---

## 🧠 Key Architecture & Design Decisions

- **Strict Duplicate Isolation** — Primary key duplicates are flagged and isolated *prior* to merging. This prevents false positive matches and duplicate fan-out in outer joins.
- **Smart Amount Tolerance Parsing** — Currency values (e.g., `"$1,500.00"` vs `1500.0`) are automatically parsed and compared using a configurable floating-point tolerance threshold (default: `±0.01`).
- **Targeted Field Comparison** — Numeric tolerance is scoped exclusively to amount/financial fields (`amount`, `price`, `fee`, `total`), preventing false matches on numerical IDs or codes (`customer_id`, `zip_code`).
- **Human-Readable Mismatch Annotations** — Each mismatch row generates an explicit reason string detailing exactly *which* fields diverged (e.g., `amount diff (ERP: 3200.0, Bank: $3,250.00)`).
- **Clean Audit Reports** — Low-level pandas implementation details (such as `_merge` columns) are mapped to business terms (`Present in ERP only (Missing in Bank)`).

---

## 📊 Output Excel Structure

The tool outputs a single timestamped Excel file (`reconciliation_output_YYYYMMDD_HHMMSS.xlsx`) with the following worksheets:

| Sheet Name | Description |
|---|---|
| `summary` | Executive metrics dashboard (Total records ingested, counts per category) |
| `matches` | All fully reconciled records |
| `mismatches` | Records with field-level differences + human-readable mismatch reasons |
| `missing` | Unmatched records annotated with source of absence |
| `data_issues` | Isolated duplicate records from ERP or Bank |
| `detailed` | Complete merged dataset with reconciliation status annotations |

---

## 🛠️ Stack & Dependencies

- **Python 3.x**
- **Pandas** — Data ingestion, cleaning, deduplication, and vector merging
- **OpenPyXL** — Excel workbook generation

---

## 🚀 How to Run

### 1. Install dependencies

```bash
pip install pandas openpyxl
```

### 2. Generate sample data (Optional)

```bash
python generate_sample_data.py
```

### 3. Run reconciliation

```bash
python reconciliation.py
```

The CLI tool will prompt you for:
- Path to the ERP / System of Record file (e.g., `sample_data/erp_data.csv`)
- Path to the External / Bank file (e.g., `sample_data/bank_data.xlsx`)
- Output directory path (e.g., `output`)
- Primary key column name (default: `invoice_id`)

---

## 🧪 Automated Testing

Run the included `unittest` test suite to verify ingestion, numeric parsing, duplicate isolation, and Excel generation:

```bash
python test_reconciliation.py
```

---

## 📁 File Structure

```
├── reconciliation.py          # Main enterprise reconciliation engine & CLI
├── test_reconciliation.py     # Automated unit test suite
├── generate_sample_data.py    # Helper script to generate mock financial files
├── sample_data/               # Sample input datasets (CSV & XLSX)
│   ├── erp_data.csv
│   └── bank_data.xlsx
└── README.md                  # System documentation
```

---

## 👤 Author & Architecture Note

Engineered as a production-style finance automation tool with defensive data handling, clear auditability, and automated verification.
