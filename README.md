# 🏦 Enterprise Finance Reconciliation Tool

A production-style Python automation tool that reconciles financial records between an **ERP / System of Record** and an **External / Bank file** — classifying every record with a clear status and generating a business-friendly Excel report.

---

## 📌 Overview

Manual reconciliation between ERP systems and bank statements is one of the most time-consuming tasks in finance operations. This tool automates that process end-to-end: ingesting two data sources, comparing them intelligently, and producing a structured output that a finance team can act on immediately.

---

## ✅ Classification Categories

| Status | Description |
|---|---|
| ✅ **Match** | Record is fully aligned across both systems |
| ⚠️ **Mismatch** | Record exists in both but differs in Amount, Currency, or Date |
| ❌ **Missing** | Record is present in one system but absent in the other |
| 🔁 **Data Issues** | Duplicate records detected before reconciliation |

---

## 🧠 Key Design Decisions

- **Configurable amount tolerance** — handles real-world rounding differences (default: ±0.01)
- **Dynamic field-level comparison** — not hardcoded; works across any shared columns between the two files
- **Duplicate isolation** — duplicates are detected and removed *before* the merge step, preventing false matches
- **Mismatch reasons** — each mismatch row includes a human-readable explanation of *which* fields diverged
- **Scalable input handling** — accepts `.csv`, `.xlsx`, or `.xls` for both input files

---

## 📊 Output

A single timestamped Excel file with multiple sheets:

| Sheet | Contents |
|---|---|
| `summary` | Count of Matches, Mismatches, Missing, Data Issues |
| `matches` | All fully reconciled records |
| `mismatches` | Records with field-level differences + reason |
| `missing` | Records not found in one of the systems |
| `data_issues` | Duplicate records from either source |
| `detailed` | Full merged dataset with all statuses |

---

## 🛠️ Stack

- **Python 3.x**
- **Pandas** — data ingestion, merging, comparison
- **OpenPyXL** — Excel output

---

## 🚀 How to Run

### 1. Install dependencies

```bash
pip install pandas openpyxl
```

### 2. Run the script

```bash
python reconciliation.py
```

The tool will prompt you for:
- Path to the ERP / System of Record file
- Path to the External / Bank file
- Output folder path

### 3. Collect your report

The output file is saved as:
```
reconciliation_output_YYYYMMDD_HHMMSS.xlsx
```

---

## 📁 File Structure

```
├── reconciliation.py       # Main script
├── README.md               # This file
└── sample_data/            # (Optional) Sample input files for testing
```

---

## 💡 Use Cases

This pattern applies directly to:
- **Financial close reconciliation** (ERP vs bank)
- **Data quality governance** across systems
- **Fraud detection** — surfacing discrepancies automatically
- **Audit trail generation** for finance teams

---

## 👤 Author

Built as part of a hands-on data governance and finance automation exercise, with a focus on production-level thinking over just making it work.
