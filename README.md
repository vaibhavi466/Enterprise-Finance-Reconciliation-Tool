# 🏦 Enterprise Finance Reconciliation Tool

A rule-driven Python financial reconciliation engine designed with production-oriented validation, auditability, and automated testing. It reconciles financial records between an **Enterprise System of Record (ERP)** and an **External Source (Bank / Payment Gateway / Partner)** — producing a deterministic, auditable multi-sheet Excel workbook with executive KPI dashboards and granular exception details.

---

## 📌 Business Problem & Importance

Manual financial reconciliation between ERP ledgers and external bank statements is a risk-prone, labor-intensive operation. Misaligned currency formats, floating-point rounding errors, ambiguous duplicate keys, and silent merge fan-outs frequently corrupt audit trails. 

This engine automates end-to-end financial reconciliation with **strict data-quality quarantine rules**, **exact `Decimal` monetary precision**, **explicit typed field rules**, **source row traceability**, and **protection against Excel formula injection**.

---

## ✨ Key Features

- **Zero Automatic Substring Inference**: Replaces unsafe column-name substring searching with explicit, validated configuration (`money`, `text`, `date`).
- **Exact `Decimal` Monetary Precision**: Uses `decimal.Decimal` and strict monetary syntax rules supporting `$1,500.00`, `₹1,500.00`, `€1,500.00`, `USD 1500.00`, `-1500.00`, parenthesized negatives `(500.00)`, and Indian grouping (`1,23,456.78`).
- **Malformed Data Defense**: Corrupted strings (e.g. `ERROR123`, `ABC500XYZ`, `12.3.4`, `1,2,3`) fail numeric syntax validation and trigger `INVALID_FIELD_VALUE` data exceptions, preventing false matches.
- **Typed Date Parsing**: Supports ISO dates, Excel datetime objects, and custom date formats with explicit invalid date quarantine.
- **Raw Identifier Preservation**: Preserves raw textual representations (e.g. leading zero keys `"00123"`) during ingestion.
- **Strict Primary Key Quarantine**: Rows with missing, blank, `NaN`, or whitespace-only primary keys are isolated as `MISSING_PRIMARY_KEY`.
- **Ambiguous Duplicate Isolation**: Duplicate keys on either source (e.g. ERP duplicate key with 1 External row) quarantine all associated records as `DUPLICATE_KEY_IN_ERP`, `DUPLICATE_KEY_IN_EXTERNAL`, or `DUPLICATE_KEY_IN_BOTH`.
- **Guaranteed One-to-One Merges**: Clean datasets are merged under pandas `validate="one_to_one"` enforcement.
- **Multi-Money Field & Currency Exposure**: Computes per-field monetary variances (`amount_variance`, `tax_amount_variance`) and isolates cross-currency mismatches without invalid cross-currency variance aggregation.
- **Full Source Traceability & Audit Trail**: Tracks original source row numbers (`_recon_source_row_erp`, `_recon_source_row_external`), file paths, SHA-256 file hashes, config SHA-256 hash, UTC ISO timestamps, and unique Run IDs.
- **Excel Formula Injection Protection**: Pre-sanitizes untrusted text starting with `=`, `+`, `-`, `@` to protect downstream spreadsheet viewers.

---

## 🏷️ Machine-Readable Status Taxonomy

| Status Category | Description |
|---|---|
| `MATCH` | Record present in both sources and fully aligned across all configured comparison fields |
| `MISMATCH` | Record present in both sources but diverges on monetary, text, or date comparison fields |
| `MISSING_IN_EXTERNAL` | Valid unique key present in ERP, but absent in External source |
| `MISSING_IN_ERP` | Valid unique key present in External source, but absent in ERP |
| `MISSING_PRIMARY_KEY` | Record with missing, blank, `NaN`, or whitespace-only primary key |
| `DUPLICATE_KEY_IN_ERP` | Primary key appears multiple times in ERP (quarantined prior to merge) |
| `DUPLICATE_KEY_IN_EXTERNAL` | Primary key appears multiple times in External source (quarantined prior to merge) |
| `DUPLICATE_KEY_IN_BOTH` | Primary key appears multiple times in both ERP and External sources |
| `INVALID_FIELD_VALUE` | Record where field value failed parsing (e.g. corrupted monetary format `ERROR123` or unparseable date) |

---

## ⚙️ Reconciliation Matching Pipeline

```
   ┌──────────────────────┐         ┌──────────────────────────┐
   │    ERP Input File    │         │  External Source File    │
   └──────────┬───────────┘         └────────────┬─────────────┘
              │                                  │
              ▼                                  ▼
   ┌───────────────────────────────────────────────────────────┐
   │ 1. Ingestion & Raw String Identifier Preservation          │
   └──────────────────────────┬────────────────────────────────┘
                              │
                              ▼
   ┌───────────────────────────────────────────────────────────┐
   │ 2. Header Normalization & Collision Validation (snake_case)│
   └──────────────────────────┬────────────────────────────────┘
                              │
                              ▼
   ┌───────────────────────────────────────────────────────────┐
   │ 3. Explicit Field Rule Normalization & Mandatory Schema Check│
   └──────────────────────────┬────────────────────────────────┘
                              │
                              ▼
   ┌───────────────────────────────────────────────────────────┐
   │ 4. Missing Primary Key Quarantine (MISSING_PRIMARY_KEY)   │
   └──────────────────────────┬────────────────────────────────┘
                              │
                              ▼
   ┌───────────────────────────────────────────────────────────┐
   │ 5. Deterministic Duplicate Key Quarantine (DUPLICATE_KEY_*)│
   └──────────────────────────┬────────────────────────────────┘
                              │
                              ▼
   ┌───────────────────────────────────────────────────────────┐
   │ 6. One-to-One Outer Merge on Clean Records                │
   └──────────────────────────┬────────────────────────────────┘
                              │
                              ▼
   ┌───────────────────────────────────────────────────────────┐
   │ 7. Typed Field Comparison & Per-Field Variance Calculation│
   └──────────────────────────┬────────────────────────────────┘
                              │
                              ▼
   ┌───────────────────────────────────────────────────────────┐
   │ 8. Executive KPI Summary & Auditable Excel Workbook Export│
   └──────────────────────────┘────────────────────────────────┘
```

---

## 📄 Explicit Field Rule Configuration

Comparison rules are explicitly declared via JSON configuration (or dictionary), eliminating keyword substring inference:

```json
{
  "key_column": "invoice_id",
  "field_rules": {
    "amount": {
      "type": "money",
      "tolerance": "0.01"
    },
    "currency": {
      "type": "text",
      "case_sensitive": false
    },
    "customer_name": {
      "type": "text",
      "case_sensitive": false
    },
    "date": {
      "type": "date",
      "format": "%Y-%m-%d"
    }
  }
}
```

---

## 📊 Output Excel Structure

The tool generates a timestamped Excel workbook (`reconciliation_YYYYMMDDTHHMMSS_<short_id>.xlsx`) with 7 explicitly ordered worksheets:

| Sheet Name | Contents & Business Purpose |
|---|---|
| `summary` | Executive KPI dashboard (counts, match rate %, exception rate %, currency-grouped exposures) |
| `run_metadata` | Audit trail parameters (Run ID, UTC timestamp, file paths, SHA-256 hashes, tolerance, serialized rules) |
| `matches` | All fully reconciled records (`MATCH`) |
| `mismatches` | Records with field discrepancies + per-field variances (`amount_variance`, `mismatch_fields`) |
| `missing` | Unmatched records annotated with origin (`MISSING_IN_EXTERNAL`, `MISSING_IN_ERP`) |
| `data_issues` | Quarantined duplicate keys, blank primary keys, and unparseable data exceptions |
| `all_records` | Master consolidated view containing every reconciliation unit and final status |

---

## 🛠️ Tech Stack & Requirements

- **Python 3.10+**
- **Pandas** — Vectorized data ingestion, normalization, and merge operations
- **OpenPyXL** — Formatted Excel workbook generation with frozen panes and auto-fit columns
- **Ruff** — Code linting and formatting enforcement

---

## 🚀 Installation & Usage

### 1. Installation

Clone the repository and install in editable mode with development dependencies:

```bash
git clone https://github.com/vaibhavi466/Enterprise-Finance-Reconciliation-Tool.git
cd Enterprise-Finance-Reconciliation-Tool
pip install -e ".[dev]"
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
  --config reconciliation_config.json \
  --output output \
  --key invoice_id \
  --amount-tolerance 0.01 \
  --overwrite
```

Or run via installed project entry point:

```bash
finance-reconcile \
  --erp sample_data/erp_data.csv \
  --external sample_data/bank_data.xlsx \
  --config reconciliation_config.json \
  --output output
```

---

## 🧪 Automated Testing & CI

Run the complete unit and integration test suite:

```bash
python -m unittest discover tests
```

Run code formatting and lint checks:

```bash
ruff check .
ruff format --check .
```

### GitHub Actions CI
Automated testing is configured via [.github/workflows/tests.yml](.github/workflows/tests.yml), executing Ruff lint checks, Ruff format checks, unit tests, and CLI end-to-end runs across Python 3.10, 3.11, 3.12, and 3.13 on every `push` and `pull_request`.

---

## 📐 Assumptions & Limitations

1. **Key-Based Reconciliation**: Matching is driven by an explicit primary key column (e.g. `invoice_id`). Fuzzy matching across unstructured entity names is not performed.
2. **Defined Date Formats**: Dates are parsed using configured formats (`%Y-%m-%d`).
3. **No Dynamic FX Conversion**: Amounts in differing currencies (e.g. USD vs EUR) are flagged as currency mismatches rather than dynamically converted using live FX APIs.
4. **Duplicate Quarantine**: Duplicate primary keys are quarantined for operational human review rather than arbitrarily resolved.
5. **Excel Row Limits**: Output reports are generated as standard Excel files suitable for moderate enterprise datasets.

---

## 👤 Author & Portfolio Note

Engineered as a rule-driven Python financial reconciliation tool with defensive data validation, explicit field configuration, `Decimal` precision, full source traceability, and automated CI verification.
