# core/invoice_processing/vision.py
# Cells 6 + 7 + 8 from notebook
# Changes from Colab version:
#   - Removed: from google.colab import userdata
#   - Added: _get_client() reads GEMINI_API_KEY from environment variable
#   - Added: extract_invoice_from_bytes() for Streamlit file uploads

import copy
import json
import os
import tempfile

import PIL.Image
from google import genai


# ── Schema ─────────────────────────────────────────────────────
INVOICE_SCHEMA = {
    "invoice_number"  : None,
    "invoice_date"    : None,
    "vendor_name"     : None,
    "vendor_gstin"    : None,
    "buyer_gstin"     : None,
    "hsn_sac_code"    : None,
    "taxable_value"   : None,
    "cgst"            : None,
    "sgst"            : None,
    "igst"            : None,
    "total_amount"    : None,
    "place_of_supply" : None,
    "invoice_type"    : None,
    "line_items"      : [],
    "confidence_score": None,
    "is_rcm"          : False,
}


def enforce_schema(data: dict) -> dict:
    validated = copy.deepcopy(INVOICE_SCHEMA)
    for key in validated:
        validated[key] = data.get(key, validated[key])
    validated["is_rcm"] = str(data.get("is_rcm", "false")).strip().lower() in ("true", "yes", "1", "y")
    return validated


# ── Prompt ─────────────────────────────────────────────────────
EXTRACTION_PROMPT = """
You are a GST Invoice Data Extraction Engine for Indian tax compliance.

Analyze the invoice image and extract ALL fields below.

STRICT RULES:
1. Return ONLY a valid JSON object. No explanation, no markdown, no code blocks.
2. Return valid JSON that strictly follows the schema below. Do not include any extra keys.
3. If a field is not found, return null for that field.
4. All monetary values must be returned as floats only. Example: 1180.50
   Never use ₹ symbol, commas, or string formatting.
   taxable_value is the amount BEFORE GST is added.
   If labeled as "Taxable Amount", "Taxable Value", "Net Amount",
   "Basic Value", or "Taxable" on the invoice, extract that as taxable_value.
5. Dates must be in DD-MM-YYYY format.
6. GSTIN must be returned exactly as printed, uppercase.
7. invoice_type must be either "B2B" (has buyer GSTIN) or "B2C" (no buyer GSTIN).
8. confidence_score must be a decimal between 0.0 and 1.0 reflecting your
   certainty about the overall extraction accuracy.
9. line_items must be a list of objects, each with:
   {"description": "", "hsn_sac": "", "quantity": null, "rate": null, "amount": null}
10. If multiple values appear for a field, choose the one labeled as
    "Invoice No", "Invoice Date", or "GSTIN".
11. If the image contains multiple invoices, extract ONLY the first invoice
    located in the TOP-LEFT area of the image. Ignore all other invoices.
12. is_rcm must be true if the invoice explicitly contains any of these phrases:
    "Reverse Charge: Yes", "RCM Applicable", "Tax payable under Reverse Charge",
    "GST payable under RCM", "Reverse Charge Mechanism applies".
    Otherwise set is_rcm to false.

Return exactly this JSON structure and nothing else:
{
    "invoice_number"  : null,
    "invoice_date"    : null,
    "vendor_name"     : null,
    "vendor_gstin"    : null,
    "buyer_gstin"     : null,
    "hsn_sac_code"    : null,
    "taxable_value"   : null,
    "cgst"            : null,
    "sgst"            : null,
    "igst"            : null,
    "total_amount"    : null,
    "place_of_supply" : null,
    "invoice_type"    : null,
    "line_items"      : [],
    "confidence_score": null,
    "is_rcm"          : false
}
"""


# ── Client ─────────────────────────────────────────────────────
def _get_client():
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise ValueError(
            "GEMINI_API_KEY environment variable is not set. "
            "Set it in your terminal before running: set GEMINI_API_KEY=your_key"
        )
    return genai.Client(api_key=api_key)


# ── Core extraction ────────────────────────────────────────────
def extract_invoice(image_path: str) -> dict:
    # Step 1: Load image
    try:
        image = PIL.Image.open(image_path).convert("RGB")
    except Exception as e:
        raise ValueError(f"Could not load image: {e}")

    # Step 2: Send to Gemini with retry
    client   = _get_client()
    raw_text = None
    for attempt in range(2):
        try:
            response = client.models.generate_content(
                model="models/gemini-2.5-flash",
                contents=[EXTRACTION_PROMPT, image],
                config={
                    "temperature": 0.1,
                    "max_output_tokens": 8192
                }
            )
            raw_text = response.text
            break
        except Exception as e:
            if attempt == 1:
                raise ValueError(f"Gemini API failed after 2 attempts: {e}")

    if not raw_text:
        raise ValueError("Empty response from Gemini.")

    # Step 3: Parse JSON
    try:
        cleaned = raw_text.strip()
        if cleaned.startswith("```"):
            cleaned = cleaned.replace("```json", "").replace("```", "").strip()
        if cleaned.count("{") > cleaned.count("}"):
            cleaned = cleaned[:cleaned.rfind("}") + 1]
        extracted_data = json.loads(cleaned)
    except json.JSONDecodeError as e:
        raise ValueError(f"JSON parsing failed: {e}. Raw: {raw_text[:200]}")

    return enforce_schema(extracted_data)


# ── Streamlit helper — takes bytes, not a file path ────────────
def extract_invoice_from_bytes(image_bytes: bytes) -> dict:
    """
    Used by the Streamlit app.
    Streamlit gives us bytes from file_uploader, not a path.
    We write to a temp file, extract, then clean up.
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp:
        tmp.write(image_bytes)
        tmp_path = tmp.name
    try:
        result = extract_invoice(tmp_path)
    finally:
        os.unlink(tmp_path)
    return result