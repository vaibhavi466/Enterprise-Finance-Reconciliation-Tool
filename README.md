# 🏦 Enterprise Finance Reconciliation Tool

A production-grade, rule-driven Python reconciliation engine that reconciles financial records between an **Enterprise System of Record (ERP)** and an **External Source (Bank / Payment Gateway / Partner)** — producing a deterministic, auditable multi-sheet Excel workbook with executive KPI dashboards and granular exception details.

---

## 📌 Business Problem & Importance

Manual financial reconciliation between ERP ledgers and external bank statements is one of the most risk-prone, labor-intensive operations in finance and accounting. Misaligned currency formats, floating-point rounding errors, ambiguous duplicate keys, and silent merge fan-outs frequently corrupt audit trails. 

This engine automates end-to-end financial reconciliation with **strict data-quality quarantine rules**, **exact `Decimal` monetary precision**, **typed field comparisons**, **source row traceability**, and **protection against Excel formula injection**.

---

## ✨ Key Features

- **Exact `Decimal` Monetary Precision**: Replaces binary floating-point representation (`float`) with `decimal.Decimal` to prevent rounding inaccuracies during tolerance evaluation.
- **Robust Financial Value Parsing**: Safely handles formats such as `$1,500.00`, `₹1,500.00`, `€1,500.00`, `USD 1500.00`, `-1500.00`, and parenthesized negative amounts `(500.00)`. Corrupted text formats (e.g. `ERROR123`, `12.3.4`) are explicitly tagged as invalid data exceptions rather than false matches.
- **Strict Primary Key Quarantine**: Rows with missing, blank, `NaN`, or whitespace-only primary keys are isolated as `MISSING_PRIMARY_KEY` prior to matching.
- **Ambiguous Duplicate Isolation**: Duplicate keys on either source (e.g. ERP duplicate key with 1 External row) are quarantined into `data_issues`, preventing false `MISSING_IN_ERP` classifications.
- **Guaranteed One-to-One Merges**: Clean datasets are merged with pandas `validate="one_to_one"` enforcement, failing loudly if cardinality assumptions are broken.
- **Explicit Schema & Typed Field Comparison**: Typed field rules (`money`, `text`, `date`) enforce explicit header validation and header collision detection (`Invoice ID` vs `invoice_id`).
- **Currency-Aware Exposure KPIs**: Executive KPIs group monetary exposures by currency (e.g. USD exposure vs EUR exposure), avoiding invalid multi-currency aggregation.
- **Full Source Traceability & Auditability**: Tracks original source row numbers (`_source_row_erp`, `_source_row_external`), file paths, SHA-256 file hashes, and unique run IDs.
- **Excel Formula Injection Protection**: Pre-sanitizes untrusted text starting with `=`, `+`, `-`, `@` to protect downstream spreadsheet viewers.

---

## 🏷️ Machine-Readable Status Taxonomy

| Status Category | Description |
|---|---|
| `MATCH` | Record present in both sources and fully aligned across all comparison fields within tolerance |
| `MISMATCH` | Record present in both sources but diverges on monetary, text, or date comparison fields |
| `MISSING_IN_EXTERNAL` | Valid unique key present in ERP, but absent in External source |
| `MISSING_IN_ERP` | Valid unique key present in External source, but absent in ERP |
| `MISSING_PRIMARY_KEY` | Record with missing, blank, `NaN`, or whitespace-only primary key |
| `DUPLICATE_KEY_IN_ERP` | Primary key appears multiple times in ERP (quarantined prior to merge) |
| `DUPLICATE_KEY_IN_EXTERNAL` | Primary key appears multiple times in External source (quarantined prior to merge) |
| `DUPLICATE_KEY_IN_BOTH` | Primary key appears multiple times in both ERP and External sources |
| `INVALID_FIELD_VALUE` | Record where field value failed parsing (e.g. corrupted monetary format `ERROR123`) |

---

## ⚙️ Reconciliation Rules & Matching Flow

```
   ┌──────────────────────┐         ┌──────────────────────────┐
   │    ERP Input File    │         │  External Source File    │
   └──────────┬───────────┘         └────────────┬─────────────┘
              │                                  │
              ▼                                  ▼
   ┌───────────────────────────────────────────────────────────┐
   │ 1. Ingestion, Source Row Tagging & SHA-256 Hash Calculation│
   └──────────────────────────┬────────────────────────────────┘
                              │
                              ▼
   ┌───────────────────────────────────────────────────────────┐
   │ 2. Header Normalization & Collision Validation (snake_case)│
   └──────────────────────────┬────────────────────────────────┘
                              │
                              ▼
   ┌───────────────────────────────────────────────────────────┐
   │ 3. Missing Primary Key Quarantine (MISSING_PRIMARY_KEY)   │
   └──────────────────────────┬────────────────────────────────┘
                              │
                              ▼
   ┌───────────────────────────────────────────────────────────┐
   │ 4. Ambiguous Duplicate Key Quarantine (DUPLICATE_KEY_*)   │
   └──────────────────────────┬────────────────────────────────┘
                              │
                              ▼
   ┌───────────────────────────────────────────────────────────┐
   │ 5. One-to-One Vector Outer Merge on Clean Keys            │
   └──────────────────────────┬────────────────────────────────┘
                              │
                              ▼
   ┌───────────────────────────────────────────────────────────┐
   │ 6. Typed Field Comparison (Money/Decimal, Text, Date)     │
   └──────────────────────────┬────────────────────────────────┘
                              │
                              ▼
   ┌───────────────────────────────────────────────────────────┐
   │ 7. Executive Summary KPI & Auditable Excel Export         │
   └───────────────────────────────────────────────────────────┘
```

---

## 📊 Output Excel Structure

The tool generates a timestamped Excel workbook (`reconciliation_output_YYYYMMDD_HHMMSS.xlsx`) with 7 explicitly ordered worksheets:

| Sheet Name | Contents & Business Purpose |
|---|---|
| `summary` | Executive KPI dashboard (counts, match rate %, exception rate %, currency-grouped exposures) |
| `run_metadata` | Audit trail parameters (Run ID, timestamp, file paths, SHA-256 hashes, tolerance, compared fields) |
| `matches` | All fully reconciled records (`MATCH`) |
| `mismatches` | Records with field-level discrepancies + structured diff details (`amount_difference`, `mismatch_fields`) |
| `missing` | Unmatched records annotated with origin (`MISSING_IN_EXTERNAL`, `MISSING_IN_ERP`) |
| `data_issues` | Quarantined duplicate keys, blank primary keys, and unparseable data exceptions |
| `all_records` | Master consolidated view containing every ingested row and final reconciliation status |

---

## 🛠️ Tech Stack & Requirements

- **Python 3.10+**
- **Pandas** — Vectorized data ingestion, normalization, and merge operations
- **OpenPyXL** — Formatted Excel workbook generation with frozen panes and auto-fit columns

---

## 🚀 Installation & Usage

### 1. Installation

Clone the repository and install dependencies:

```bash
git clone https://github.com/vaibhavi466/Enterprise-Finance-Reconciliation-Tool.git
cd Enterprise-Finance-Reconciliation-Tool
pip install .
```

### 2. Generate Deterministic Sample Data

Generate reproducible test files in `sample_data/`:

```bash
python generate_sample_data.py
```

### 3. Run via Command Line Interface (CLI)

```bash
python reconciliation.py \
  --erp sample_data/erp_data.csv \
  --external sample_data/bank_data.xlsx \
  --output output \
  --key invoice_id \
  --amount-tolerance 0.01
```

### 4. Interactive Mode Fallback

If run without file parameters, the CLI safely prompts for inputs interactively:

```bash
python reconciliation.py
```

---

## 🧪 Automated Testing & CI

Run the complete unit and integration test suite:

```bash
python -m unittest discover tests
```

### GitHub Actions CI
Automated testing is configured via [.github/workflows/tests.yml](file:///.github/workflows/tests.yml), running matrix tests across Python 3.10, 3.11, 3.12, and 3.13 on every `push` and `pull_request`.

---

## 📁 Repository Structure

```
Enterprise-Finance-Reconciliation-Tool/
├── .github/
│   └── workflows/
│       └── tests.yml            # GitHub Actions CI workflow
├── sample_data/
│   ├── erp_data.csv             # Sample ERP dataset
│   └── bank_data.xlsx           # Sample External/Bank dataset
├── tests/
│   └── test_reconciliation.py   # Comprehensive automated test suite
├── reconciliation.py            # Main reconciliation engine & CLI
├── generate_sample_data.py      # Deterministic sample dataset generator
├── pyproject.toml               # Python package & project configuration
├── .gitignore                   # Workspace gitignore rules
└── README.md                    # Project documentation
```

---

## 📐 Assumptions & Limitations

1. **Key-Based Reconciliation**: Matching is driven by an explicit primary key column (e.g. `invoice_id`). Fuzzy matching across unstructured entity names is not performed.
2. **Defined Date Formats**: Dates are parsed using standard ISO/configured formats (`YYYY-MM-DD`).
3. **No Dynamic FX Conversion**: Amounts in differing currencies (e.g. USD vs EUR) are flagged as mismatches rather than dynamically converted using real-time exchange rates.
4. **Duplicate Quarantine**: Duplicate primary keys are quarantined for operational human review rather than arbitrarily resolved.

---

## 👤 Author & Engineering Standard

Engineered with defensive data quality controls, explicit financial precision (`Decimal`), full source traceability, and automated CI verification suitable for finance operations and enterprise portfolio review.
