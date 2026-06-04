# AuditIQ — AI-Powered GST Audit & Reconciliation Platform

> **Built for Chartered Accountants and finance teams who can't afford to guess.**
> AuditIQ automates the two most time-consuming parts of GST compliance: invoice extraction & ITC classification, and GSTR-2B reconciliation against Tally books.

**Live App → [auditiq-app.streamlit.app](https://auditiq-app.streamlit.app)**

---

## What AuditIQ Does

Manual GST reconciliation means a CA spending hours (sometimes days) cross-checking vendor invoices against GSTR-2B data, hunting for mismatches, and deciding which ITC claims are eligible. AuditIQ automates this pipeline end-to-end — the CA reviews and signs off, the machine does the grunt work.

The platform has two modules:

### 1. Invoice Processing
Upload invoice images (JPG/PNG) → AuditIQ extracts every GST field, verifies the math, classifies ITC eligibility, and outputs an audit-ready Excel report.

- **LLM extraction** via Gemini 2.5 Pro — pulls vendor name, GSTIN, line items, tax amounts, HSN codes from unstructured invoice images
- **Math verification** — confirms that line-item totals, tax calculations, and grand totals are internally consistent; flags mismatches explicitly
- **ITC classification** — rule-based engine classifies each invoice as `ELIGIBLE`, `BLOCKED`, `RCM`, or `REVIEW_NEEDED` under GST Section 17(5)
- **Excel audit report** — structured output per invoice with status, flags, and extracted fields
- **Tally XML export** *(in development)* — direct import into Tally ledger, no manual re-entry

### 2. GST Reconciliation
Upload your Tally export + GSTR-2B from the GST portal → get a complete reconciliation report with claimable ITC calculated.

- Parses GSTR-2B (downloaded from GST Portal) and Tally books (IGST Input + CGST Input + RCM sheets)
- Runs multi-strategy matching: exact match → substring match → fuzzy match
- Buckets every entry: `MATCHED` / `MISSING_IN_BOOKS` / `MISSING_IN_2B` / `INELIGIBLE` / `RCM` / `PREVIOUS_MONTH` / `REVIEW`
- Calculates claimable ITC for the period (IGST + CGST + SGST breakdown)
- **Direct Tally import for "Missing in Books"** *(in development)* — entries missing from your books can be pushed directly into Tally using stored ledger name mappings, no manual lookup

---

## Design Philosophy

**Fail loudly. Never guess silently.**

AuditIQ is built around one principle: in financial compliance, a wrong answer presented with confidence is worse than no answer at all. Every low-confidence extraction is labelled. Every mismatch is surfaced. The system produces a reviewed draft — the CA finalises.

This human-in-the-loop model is intentional and non-negotiable.

---

## Tech Stack

| Layer | Technology |
|---|---|
| Frontend / App | Python, Streamlit |
| LLM Extraction | Gemini 2.5 Pro (Google AI) |
| Data Processing | Pandas |
| Reconciliation | Custom fuzzy matching engine (exact / substring / fuzzy) |
| Output | Excel (openpyxl), Tally XML *(in dev)* |
| Deduplication | SHA-256 hash per invoice |
| Deployment | Streamlit Cloud |

---

## Project Structure

```
auditiq/
├── app.py                          # Main Streamlit app (two-page navigation)
├── requirements.txt
├── config/                         # Config files
└── core/
    ├── gstr2b_parser.py            # Parses GSTR-2B Excel from GST Portal
    ├── tally_parser.py             # Parses Tally export (IGST/CGST/RCM sheets)
    ├── reconciler.py               # Core reconciliation logic + bucket classification
    ├── excel_output.py             # Reconciliation Excel report generation
    └── invoice_processing/
        ├── vision.py               # Gemini API extraction from invoice images
        ├── cleaner.py              # Field normalisation and cleanup
        ├── validator.py            # Math verification + invoice ID generation
        ├── itc_classifier.py       # ITC eligibility classification (Section 17(5))
        └── invoice_excel.py        # Invoice audit Excel report generation
```

---

## Getting Started (Run Locally)

**Prerequisites:** Python 3.10+, a Google AI (Gemini) API key

```bash
# Clone the repo
git clone https://github.com/kanishk-debug/auditiq.git
cd auditiq

# Install dependencies
pip install -r requirements.txt

# Set your Gemini API key
export GOOGLE_API_KEY="your_key_here"

# Run the app
streamlit run app.py
```

App opens at `http://localhost:8501`

---

## How to Use

### Invoice Processing
1. Navigate to **Invoice Processing** in the sidebar
2. Upload one or more invoice images (JPG or PNG)
3. Enter client name (optional, used for the output filename)
4. Click **Extract & Classify**
5. Download the **Invoice Audit Excel** — one row per invoice with extracted fields, math check result, and ITC status

### Reconciliation
1. Navigate to **Reconciliation** in the sidebar
2. Upload your **Tally export** (Excel with IGST Input, CGST Input, and RCM sheets)
3. Upload your **GSTR-2B** (downloaded from the GST Portal as Excel)
4. Enter client name and period
5. Click **Run Reconciliation**
6. Review the summary (matched count, missing entries, claimable ITC)
7. Download the **Reconciliation Excel**

---

## ITC Classification Buckets

| Status | Meaning |
|---|---|
| `ELIGIBLE` | ITC can be claimed |
| `BLOCKED` | Blocked under Section 17(5) (e.g. food, motor vehicles, personal use) |
| `RCM` | Reverse Charge Mechanism applies |
| `REVIEW_NEEDED` | Ambiguous — flagged for CA review |

## Reconciliation Buckets

| Bucket | Meaning |
|---|---|
| `MATCHED` | Entry present in both Tally and GSTR-2B |
| `MISSING_IN_BOOKS` | Present in GSTR-2B but not in Tally *(can be imported to Tally)* |
| `MISSING_IN_2B` | Present in Tally but not in GSTR-2B |
| `INELIGIBLE` | ITC not claimable |
| `RCM` | Reverse charge entry |
| `PREVIOUS_MONTH` | Likely belongs to a different period |
| `REVIEW` / `PROBABLE_MATCH` | Fuzzy match found — needs human confirmation |

---

## Roadmap

- [x] Gemini-based invoice extraction (images)
- [x] ITC classification engine
- [x] Math verification
- [x] GSTR-2B ↔ Tally reconciliation
- [x] Excel reports for both modules
- [ ] Tally XML export for invoice processing
- [ ] Direct "Missing in Books" import to Tally with stored ledger name mappings
- [ ] PDF invoice support
- [ ] Multi-month reconciliation
- [ ] CA firm multi-client dashboard

---

## Status

**MVP — Live.** Validated with a practising Chartered Accountant on real reconciliation datasets.

Active development. Tally XML export and direct ledger import are in progress.

---

## Author

**Kanishk Arora** — B.Tech CSE, NIT Hamirpur (2024–2028)

[godparticle8105@gmail.com](mailto:godparticle8105@gmail.com) · [github.com/kanishk-debug](https://github.com/kanishk-debug)

---

> *AuditIQ is built as a real-world, production-grade tool — not a hackathon prototype. Every design decision prioritises trust and accuracy over speed.*
