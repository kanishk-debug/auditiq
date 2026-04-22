# core/invoice_processing/validator.py
# Cells 11 + 11B from notebook — no changes except removing print statements

import re
import hashlib

GSTIN_PATTERN = re.compile(
    r"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$"
)


def validate_gstin(gstin: str) -> dict:
    if not gstin:
        return {"valid": None, "reason": "GSTIN not present (possibly B2C invoice)"}
    if GSTIN_PATTERN.match(gstin):
        return {"valid": True, "reason": "GSTIN format is correct"}
    return {"valid": False, "reason": f"Invalid GSTIN format: {gstin}"}


def validate_math(data: dict) -> dict:
    taxable = data.get("taxable_value")
    cgst    = data.get("cgst") or 0
    sgst    = data.get("sgst") or 0
    igst    = data.get("igst") or 0
    total   = data.get("total_amount")

    if taxable is None or total is None:
        return {
            "valid": None,
            "reason": "Cannot validate — taxable_value or total_amount is missing",
            "expected": None, "actual": total
        }
    if taxable < 0 or total < 0:
        return {
            "valid": False,
            "reason": "Negative amount detected — flag for manual review",
            "expected": None, "actual": total
        }

    gst_total      = round(cgst + sgst + igst, 2)
    expected_total = round(taxable + gst_total, 2)
    actual_total   = round(total, 2)

    if abs(expected_total - actual_total) <= 0.5:
        return {"valid": True, "reason": "Math checks out",
                "expected": expected_total, "actual": actual_total}
    return {"valid": False,
            "reason": f"Math mismatch — expected {expected_total}, got {actual_total}",
            "expected": expected_total, "actual": actual_total}


def validate_line_item_total(data: dict) -> dict:
    items = data.get("line_items", [])
    total = data.get("total_amount")

    if (data.get("cgst") is not None or
            data.get("sgst") is not None or
            data.get("igst") is not None):
        return {"valid": None, "reason": "Skipping — GST fields present"}

    if not items or total is None:
        return {"valid": None, "reason": "Line items or total missing"}

    try:
        valid_items = [i for i in items if i.get("amount") is not None]
        if not valid_items:
            return {"valid": None, "reason": "No usable line item amounts"}
        items_total = round(sum(float(i["amount"]) for i in valid_items), 2)
        if abs(items_total - total) <= 0.5:
            return {"valid": True, "reason": "Line item total matches invoice total",
                    "items_total": items_total, "invoice_total": total}
        return {"valid": False, "reason": "Line items total does not match invoice total",
                "items_total": items_total, "invoice_total": total}
    except Exception:
        return {"valid": None, "reason": "Could not compute line item total"}


def validate_invoice(data: dict) -> dict:
    invoice_type = data.get("invoice_type", "B2C")

    report = {
        "vendor_gstin_check": validate_gstin(data.get("vendor_gstin")),
        "buyer_gstin_check":  validate_gstin(data.get("buyer_gstin")),
        "math_check":         validate_math(data),
        "line_item_check":    validate_line_item_total(data),
        "confidence_check": {
            "valid":  (data.get("confidence_score") or 0) >= 0.75,
            "score":  data.get("confidence_score"),
            "reason": "Above threshold" if (data.get("confidence_score") or 0) >= 0.75
                      else "Low confidence — flag for manual review"
        }
    }

    gstin_valid = report["vendor_gstin_check"]["valid"]
    gstin_ok    = (gstin_valid is True) if invoice_type == "B2B" else \
                  (gstin_valid is True or gstin_valid is None)
    conf_valid  = report["confidence_check"]["valid"]
    math_valid  = report["math_check"]["valid"]

    report["flags"] = []
    if math_valid is False:
        report["flags"].append("MATH_MISMATCH")
    if conf_valid is False:
        report["flags"].append("LOW_CONFIDENCE")
    if gstin_valid is False:
        report["flags"].append("INVALID_GSTIN")
    if report["line_item_check"]["valid"] is False:
        report["flags"].append("LINE_TOTAL_MISMATCH")
    if (data.get("total_amount") or 0) > 100000:
        report["flags"].append("HIGH_VALUE_INVOICE")
    if invoice_type == "B2B" and not data.get("buyer_gstin"):
        report["flags"].append("BUYER_GSTIN_MISSING")

    gst_missing = (data.get("cgst") is None and
                   data.get("sgst") is None and
                   data.get("igst") is None)
    taxable = data.get("taxable_value")
    total   = data.get("total_amount")
    if gst_missing and taxable and total and total > taxable:
        report["flags"].append("GST_BREAKUP_MISSING")

    serious = {"MATH_MISMATCH", "INVALID_GSTIN", "LINE_TOTAL_MISMATCH"}
    if any(f in report["flags"] for f in serious):
        report["confidence_check"]["valid"] = False
        if "LOW_CONFIDENCE" not in report["flags"]:
            report["flags"].append("LOW_CONFIDENCE")

    report["overall_status"] = (
        "PASS" if (gstin_ok and conf_valid and math_valid != False)
        else "NEEDS_REVIEW"
    )
    return report


def generate_invoice_id(data: dict) -> str:
    key = (
        str(data.get("vendor_gstin",   "")).strip().upper() +
        str(data.get("invoice_number", "")).strip().upper() +
        str(data.get("invoice_date",   "")).strip() +
        str(data.get("total_amount",   "")) +
        str(data.get("vendor_name",    "")).strip().upper()
    )
    return hashlib.sha256(key.encode()).hexdigest()