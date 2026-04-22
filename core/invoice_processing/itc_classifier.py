# core/invoice_processing/itc_classifier.py
# Cells 13A + 13 + 14 + 15 from notebook
# Change from Colab: CONFIG_PATH now points to config/itc_rules.json
# relative to this file's location, not /content/

import json
import os
import re

# ── Load rules ─────────────────────────────────────────────────
# Goes up two levels: invoice_processing/ → core/ → auditiq/ → config/
_HERE       = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(_HERE, "..", "..", "config", "itc_rules.json")

if not os.path.exists(CONFIG_PATH):
    raise FileNotFoundError(
        f"itc_rules.json not found at {CONFIG_PATH}. "
        "Run Cell 13A in Colab once, download the file, "
        "and place it at auditiq/config/itc_rules.json"
    )

with open(CONFIG_PATH, "r") as f:
    _rules = json.load(f)

BLOCKED_KEYWORDS      = _rules["blocked_keywords"]
SOFT_BLOCK_KEYWORDS   = _rules["soft_block_keywords"]
ELIGIBLE_KEYWORDS     = _rules["eligible_keywords"]
BLOCKED_HSN_CODES     = _rules["blocked_hsn_codes"]
KNOWN_BLOCKED_VENDORS = _rules["known_blocked_vendors"]


# ── Vendor name cleaner ────────────────────────────────────────
def clean_vendor_name(vendor: str) -> str:
    noise = [
        "private limited", "pvt ltd", "pvt. ltd.", "pvt. ltd",
        "limited", "ltd", "ltd.", "m/s", "m/s.", "llp"
    ]
    vendor = vendor.lower()
    for term in noise:
        vendor = vendor.replace(term, " ")
    return " ".join(vendor.split())


# ── ITC classifier ─────────────────────────────────────────────
def classify_itc(invoice: dict) -> dict:
    vendor_raw   = str(invoice.get("vendor_name", "") or "")
    vendor_clean = clean_vendor_name(vendor_raw)
    line_items   = invoice.get("line_items") or []
    description  = " ".join([str(item.get("description") or "") for item in line_items])
    search_text  = re.sub(r"[^\w\s]", " ", description).lower()
    search_text  = " ".join(search_text.split())

    hsn           = re.sub(r'\D', '', (invoice.get("hsn_sac_code") or ""))
    line_item_hsns= [re.sub(r'\D', '', (item.get("hsn_sac") or "")) for item in line_items]
    all_hsns      = [h for h in [hsn] + line_item_hsns if h]

    gst_paid = round(
        (invoice.get("cgst") or 0) +
        (invoice.get("sgst") or 0) +
        (invoice.get("igst") or 0) +
        (invoice.get("cess") or 0),
        2
    )

    # Step 0: RCM
    if invoice.get("is_rcm") == True:
        return {
            "itc_status": "RCM",
            "reason":     ("Reverse Charge Mechanism — supplier has not charged GST. "
                           "Pay GST directly to govt, then claim ITC."),
            "category":   "rcm",
            "itc_amount": None,
            "gst_paid":   gst_paid,
            "confidence": "HIGH"
        }

    # Step 1: B2C
    if invoice.get("invoice_type") == "B2C":
        return {
            "itc_status": "REVIEW_NEEDED",
            "reason":     "B2C invoice — buyer GSTIN not detected. Verify before claiming ITC.",
            "category":   "b2c_invoice",
            "itc_amount": None,
            "gst_paid":   gst_paid,
            "confidence": "MEDIUM"
        }

    # Step 2: Known blocked vendors
    for known in KNOWN_BLOCKED_VENDORS:
        if known.lower() in vendor_clean:
            return {
                "itc_status": "REVIEW_NEEDED",
                "reason":     "Known consumer platform — verify business purpose.",
                "category":   "known_vendor_review",
                "itc_amount": None,
                "gst_paid":   gst_paid,
                "confidence": "MEDIUM"
            }

    # Step 3: Blocked HSN codes
    for hsn_code in all_hsns:
        for blocked_hsn in BLOCKED_HSN_CODES:
            if hsn_code.startswith(blocked_hsn):
                return {
                    "itc_status": "REVIEW_NEEDED",
                    "reason":     f"HSN {hsn_code} may fall under Sec 17(5) — verify.",
                    "category":   "hsn_review",
                    "itc_amount": None,
                    "gst_paid":   gst_paid,
                    "confidence": "MEDIUM"
                }

    # Step 4: Hard blocked keywords
    for category, keywords in BLOCKED_KEYWORDS.items():
        for keyword in keywords:
            if re.search(r'\b' + re.escape(keyword) + r'\b', search_text):
                return {
                    "itc_status": "BLOCKED",
                    "reason":     f"Blocked under Sec 17(5) — {category}",
                    "category":   category,
                    "itc_amount": 0,
                    "gst_paid":   gst_paid,
                    "confidence": "HIGH"
                }

    # Step 5: Soft block
    for keyword in SOFT_BLOCK_KEYWORDS:
        if re.search(r'\b' + re.escape(keyword) + r'\b', search_text):
            return {
                "itc_status": "REVIEW_NEEDED",
                "reason":     f"Possible blocked expense — '{keyword}' detected. CA review required.",
                "category":   "soft_blocked",
                "itc_amount": None,
                "gst_paid":   gst_paid,
                "confidence": "MEDIUM"
            }

    # Step 6: Eligible keywords
    for keyword in ELIGIBLE_KEYWORDS:
        if re.search(r'\b' + re.escape(keyword) + r'\b', search_text):
            return {
                "itc_status": "ELIGIBLE",
                "reason":     "Eligible ITC — matches office/business expense",
                "category":   "eligible_expense",
                "itc_amount": gst_paid,
                "gst_paid":   gst_paid,
                "confidence": "MEDIUM"
            }

    # Step 7: Fallback
    return {
        "itc_status": "REVIEW_NEEDED",
        "reason":     "Cannot auto-classify — needs manual CA review",
        "category":   "unknown",
        "itc_amount": None,
        "gst_paid":   gst_paid,
        "confidence": "LOW"
    }


# ── Audit report ───────────────────────────────────────────────
def generate_audit_report(invoice: dict, validation: dict) -> dict:
    itc = classify_itc(invoice)
    return {
        "invoice_number"   : invoice.get("invoice_number"),
        "invoice_date"     : invoice.get("invoice_date"),
        "vendor_name"      : invoice.get("vendor_name"),
        "vendor_gstin"     : invoice.get("vendor_gstin"),
        "taxable_value"    : invoice.get("taxable_value"),
        "cgst"             : invoice.get("cgst"),
        "sgst"             : invoice.get("sgst"),
        "igst"             : invoice.get("igst"),
        "total_amount"     : invoice.get("total_amount"),
        "invoice_type"     : invoice.get("invoice_type"),
        "is_rcm"           : invoice.get("is_rcm", False),
        "itc_status"       : itc["itc_status"],
        "itc_amount"       : itc["itc_amount"],
        "itc_reason"       : itc["reason"],
        "itc_category"     : itc["category"],
        "itc_confidence"   : itc["confidence"],
        "gst_paid"          : itc["gst_paid"],
        "extraction_confidence": invoice.get("confidence_score"),
        "validation_status" : validation["overall_status"],
        "flags"            : ", ".join(validation["flags"]) if validation["flags"] else "None",
        "disclaimer"       : (
            "ITC Recommendation only. "
            "Final decision subject to CA review. "
            "Not a substitute for professional tax advice."
        )
    }