# core/invoice_processing/cleaner.py
# Cell 10 from notebook — no changes needed except removing print statements

import re
import copy
from datetime import datetime


def clean_amount(value) -> float:
    if value is None:
        return None
    try:
        value = str(value)
        value = value.replace("O", "0").replace("o", "0")
        cleaned = re.sub(r'[₹$,\s]', '', value)
        cleaned = re.sub(r'[^\d.\-]', '', cleaned)
        if not cleaned or cleaned in [".", "-", ""]:
            return None
        if cleaned.count("-") > 1:
            return None
        cleaned = cleaned.rstrip(".")
        if not cleaned:
            return None
        amount = float(cleaned)
        if amount > 10_000_000:
            return None
        return round(amount, 2)
    except Exception:
        return None


def clean_date(value) -> str:
    if value is None:
        return None
    value = str(value).strip()
    formats = [
        "%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d",
        "%d %b %Y", "%d %B %Y", "%d-%b-%Y",
        "%d.%m.%Y", "%d-%m-%y", "%d/%m/%y",
    ]
    for fmt in formats:
        try:
            parsed = datetime.strptime(value, fmt)
            return parsed.strftime("%d-%m-%Y")
        except Exception:
            continue
    return value


def clean_string(value) -> str:
    if value is None:
        return None
    value = re.sub(r'[^\w\s&.\-/]', '', str(value))
    value = value[:200]
    return " ".join(value.strip().split())


def clean_gstin(value) -> str:
    if value is None:
        return None
    return str(value).strip().upper().replace(" ", "")


def clean_invoice(data: dict) -> dict:
    cleaned = copy.deepcopy(data)

    for field in ["taxable_value", "cgst", "sgst", "igst", "total_amount"]:
        cleaned[field] = clean_amount(cleaned.get(field))

    if cleaned.get("hsn_sac_code"):
        cleaned["hsn_sac_code"] = str(cleaned["hsn_sac_code"]).strip()

    for item in cleaned.get("line_items", []):
        item["quantity"] = clean_amount(item.get("quantity"))
        item["rate"]     = clean_amount(item.get("rate"))
        item["amount"]   = clean_amount(item.get("amount"))
        item["hsn_sac"]  = str(item.get("hsn_sac", "")).strip() or None

    cleaned["invoice_date"]    = clean_date(cleaned.get("invoice_date"))
    cleaned["vendor_name"]     = clean_string(cleaned.get("vendor_name"))
    cleaned["place_of_supply"] = clean_string(cleaned.get("place_of_supply"))
    cleaned["vendor_gstin"]    = clean_gstin(cleaned.get("vendor_gstin"))
    cleaned["buyer_gstin"]     = clean_gstin(cleaned.get("buyer_gstin"))

    if cleaned.get("buyer_gstin"):
        cleaned["invoice_type"] = "B2B"
    else:
        cleaned["invoice_type"] = "B2C"

    return cleaned