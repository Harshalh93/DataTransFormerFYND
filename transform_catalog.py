import re
import json
import shutil
import hashlib
from pathlib import Path
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill

# ==============================================================================
# 1. EXACT 104 OUTPUT HEADERS & MANDATORY COLUMNS (FROM OutPut_Format.xlsx)
# ==============================================================================
OUTPUT_HEADERS = [
    "Name", "Slug", "Item Code", "Brand", "Category", "Description",
    "Short Description", "Tax Rule Name", "HS Code", "Country of Origin",
    "Media", "Multi Size", "Gtin Type", "Gtin Value", "Seller Identifier",
    "Meta", "Size Meta", "Size", "Actual Price", "Selling Price", "Currency",
    "Length (cm)", "Width (cm)", "Height (cm)", "Product Dead Weight (gram)",
    "Size Guide", "Available", "Highlights", "Unlisted Product", "Variant Type",
    "Variant Group ID", "Variant Media", "Trader Type", "Trader Name",
    "Trader Address", "Track Inventory", "Teaser Tag Name", "No of Boxes",
    "Manufacturing Time", "Manufacturing Time Unit", "Return Time Limit",
    "Return Time Unit", "Product Publishing Date", "Tags", "Net Quantity Value",
    "Net Quantity Unit", "Product Bundle", "Colour", "Material",
    "Quantity Factor", "Priority",
    *[f"Custom Attribute {i}" for i in range(1, 51)],
    "Threshold", "Gender", "Package Contents"
]

# 21 Green-Highlighted (FFC5FFB0) Mandatory Columns in OutPut_Format.xlsx
MANDATORY_HEADERS = {
    "Name", "Item Code", "Brand", "Category", "Tax Rule Name", "HS Code",
    "Country of Origin", "Gtin Type", "Gtin Value", "Seller Identifier",
    "Size", "Actual Price", "Selling Price", "Currency", "Length (cm)",
    "Width (cm)", "Height (cm)", "Product Dead Weight (gram)",
    "Trader Type", "Trader Name", "Trader Address"
}

# ==============================================================================
# 2. DATA DECLARATIONS & NORMALIZATION MAPS (FROM DataDeclarations.xlsx)
# ==============================================================================
ALLOWED_BRANDS = {
    "Generic",
    "Teafields Reserve",
    "Himalayan Brew",
    "Artisan Leaf",
    "Gourmet Crumb",
}

BRAND_ALIAS_MAP = {
    "generic": "Generic",
    "teafields reserve": "Teafields Reserve",
    "teafields": "Teafields Reserve",
    "teafields house brand": "Teafields Reserve",
    "himalayan brew": "Himalayan Brew",
    "artisan leaf": "Artisan Leaf",
    "gourmet crumb": "Gourmet Crumb",
}

DEFAULT_BRAND_FALLBACK = "Generic"

# Exact case-sensitive Net Quantity Units from DataDeclarations.xlsx:
# ['kg', 'g', 'L', 'ml', 'm', 'cm', 'nos', 'pack', 'ea', 'set']
UOM_MAP = {
    "PCS": "nos",
    "PC": "nos",
    "PIECE": "nos",
    "PIECES": "nos",
    "NOS": "nos",
    "NO": "nos",
    "EA": "ea",
    "EACH": "ea",
    "PACK": "pack",
    "PKT": "pack",
    "SET": "set",
    "KG": "kg",
    "KGS": "kg",
    "KILOGRAM": "kg",
    "KILOGRAMS": "kg",
    "GM": "g",
    "GMS": "g",
    "G": "g",
    "GRAM": "g",
    "GRAMS": "g",
    "L": "L",
    "LTR": "L",
    "LITRE": "L",
    "LITER": "L",
    "ML": "ml",
    "MILLILITRE": "ml",
    "M": "m",
    "MTR": "m",
    "CM": "cm",
}

# Exact Tax Rule Names from DataDeclarations.xlsx (preserving en-dash '\u2013')
TAX_RULE_MAP = {
    "0": "Zero-Rated Tax Rule \u2013 0% (Eff. 22 Sep 2025)",
    "0.0": "Zero-Rated Tax Rule \u2013 0% (Eff. 22 Sep 2025)",
    "0.25": "Flat tax rule - 0.25%",
    "3": "Flat Tax Rule \u2013 3%",
    "3.0": "Flat Tax Rule \u2013 3%",
    "5": "Merit Tax Rule \u2013 5% (Eff. 22 Sep 2025)",
    "5.0": "Merit Tax Rule \u2013 5% (Eff. 22 Sep 2025)",
    "12": "Flat Tax Rule 3 \u2013 12%",
    "12.0": "Flat Tax Rule 3 \u2013 12%",
    "18": "Standard Tax Rule \u2013 18% (Eff. 22 Sep 2025)",
    "18.0": "Standard Tax Rule \u2013 18% (Eff. 22 Sep 2025)",
    "28": "Flat Tax Rule 4 \u2013 28.0%",
    "28.0": "Flat Tax Rule 4 \u2013 28.0%",
    "40": "De-merit Tax Rule \u2013 40% (Eff. 22 Sep 2025)",
    "40.0": "De-merit Tax Rule \u2013 40% (Eff. 22 Sep 2025)",
}
DEFAULT_TAX_RULE = "Merit Tax Rule \u2013 5% (Eff. 22 Sep 2025)"

# Category-level fallback HSN for Specialty Tea / Food when missing (~8% rows)
CATEGORY_HSN_FALLBACK = {
    "tea": "09024020",
    "chai": "09109100",
    "tisane": "21069099",
    "teapot": "73239190",
    "mug": "70134200",
    "biscotti": "19053100",
    "honey": "04090000",
}

# Default packaging dimensions (Definitions.xlsx requires Length > Width)
DEFAULT_LENGTH_CM = 15
DEFAULT_WIDTH_CM = 10
DEFAULT_HEIGHT_CM = 10
DEFAULT_DEAD_WEIGHT_G = 100

MANIFEST_PATH = Path("idempotency_manifest.json")
TEMPLATE_PATH = Path("OutPut_Format.xlsx")


def clean_string(val):
    if pd.isna(val):
        return ""
    return re.sub(r"\s+", " ", str(val).strip())


def normalize_title(val):
    s = clean_string(val)
    return s.title() if s else ""


def resolve_brand(raw_val):
    """
    Normalizes brand string and enforces strict membership in DataDeclarations.xlsx Brand list.
    """
    cleaned = clean_string(raw_val)
    if not cleaned or cleaned.upper() in {"N/A", "NULL", "NONE", "NA", "-"}:
        return DEFAULT_BRAND_FALLBACK
    alias_hit = BRAND_ALIAS_MAP.get(cleaned.lower())
    if alias_hit:
        return alias_hit
    titled = cleaned.title()
    if titled in ALLOWED_BRANDS:
        return titled
    return DEFAULT_BRAND_FALLBACK


def resolve_tax_rule(raw_gst_val):
    """
    Maps a numeric GST percentage or tax string to an exact Tax Rule Name from DataDeclarations.xlsx.
    """
    s = clean_string(raw_gst_val)
    if not s:
        return DEFAULT_TAX_RULE
    if s in TAX_RULE_MAP.values():
        return s
    cleaned_num = re.sub(r"(?i)(gst|vat|%|@|\s)", "", s)
    return TAX_RULE_MAP.get(cleaned_num, DEFAULT_TAX_RULE)


def resolve_category(item_name, raw_cat=""):
    """
    Maps product to a valid Category from DataDeclarations.xlsx.
    Defaults to 'Others level 3' for Supplementary items.
    """
    c = clean_string(raw_cat)
    if c in {"Others level 3", "Bottle", "Carry Bag"}:
        return c
    lower_name = item_name.lower()
    if "bottle" in lower_name or "flask" in lower_name:
        return "Bottle"
    if "carry bag" in lower_name or "tote" in lower_name:
        return "Carry Bag"
    return "Others level 3"


def parse_mrp(val):
    """
    Extracts numeric MRP from noisy text strings like 'Rs. 1,250.00', '₹899/-'.
    Returns (numeric_val, error_reason). Casts whole floats to int to satisfy
    OutPut_Format.xlsx whole-number Data Validation on Actual Price and Selling Price.
    """
    s = clean_string(val)
    if not s:
        return None, "Missing or blank MRP"
    if re.search(r"-\s*\d+", s):
        return None, f"Negative MRP not allowed: '{s}'"
    cleaned = re.sub(r"(?i)(rs\.?|inr|₹|/-|,|\s)", "", s)
    try:
        price = float(cleaned)
        if price <= 0 or price > 100000:
            return None, f"MRP out of valid commercial bounds ({price}): '{s}'"
        rounded = round(price, 2)
        return (int(rounded) if rounded.is_integer() else rounded), None
    except ValueError:
        return None, f"Unparseable non-numeric MRP string: '{s}'"


def extract_packed_attributes(spec_text):
    """
    Splits free-text description string into structured attributes:
    net_weight_volume, dead_weight_grams, size, color, pack_quantity, packaging_type
    """
    s = clean_string(spec_text)
    attrs = {
        "net_weight_volume": "Standard",
        "dead_weight_grams": DEFAULT_DEAD_WEIGHT_G,
        "size": "OS",  # One Size default for supplementary items
        "color": "Multi",
        "pack_quantity": 1,
        "packaging_type": "Standard Pack",
    }
    if not s:
        return attrs

    # Weight / Volume (e.g., 250g, 500 GM, 1kg, 800ml)
    wt_match = re.search(r"\b(\d+(?:\.\d+)?)\s*(g|gm|gms|kg|kgs|ml|l)\b", s, re.I)
    if wt_match:
        num_val = float(wt_match.group(1))
        unit_val = wt_match.group(2).upper()
        attrs["net_weight_volume"] = f"{wt_match.group(1)}{unit_val}"
        if unit_val in {"KG", "KGS", "L"}:
            attrs["dead_weight_grams"] = max(1, int(round(num_val * 1000)))
        elif unit_val in {"G", "GM", "GMS", "ML"}:
            attrs["dead_weight_grams"] = max(1, int(round(num_val)))

    # Size (Small, Medium, Large, OS)
    sz_match = re.search(r"Size:\s*([A-Za-z0-9]+)", s, re.I)
    if sz_match:
        attrs["size"] = sz_match.group(1).title()
    elif attrs["net_weight_volume"] != "Standard":
        attrs["size"] = attrs["net_weight_volume"]

    # Color
    col_match = re.search(r"Color:\s*([A-Za-z\s]+?)(?:\||,|-|/|Pack|$)", s, re.I)
    if col_match:
        attrs["color"] = col_match.group(1).strip().title()

    # Pack count
    pk_match = re.search(r"Pack\s*of\s*(\d+)", s, re.I)
    if pk_match:
        attrs["pack_quantity"] = max(1, int(pk_match.group(1)))

    # Packaging type
    for pkg in ["Tin Caddy", "Kraft Pouch", "Zipper Bag", "Glass Jar", "Pyramid Bags", "Box", "Glass Bottle"]:
        if pkg.lower() in s.lower():
            attrs["packaging_type"] = pkg
            break

    return attrs


def build_short_description(item_name, attrs):
    """
    Enforces the <= 50 character limit defined in Definitions.xlsx for Short Description.
    """
    candidate = f"{item_name} ({attrs['net_weight_volume']}, {attrs['packaging_type']})"
    if len(candidate) <= 50:
        return candidate
    compact = f"{item_name} ({attrs['net_weight_volume']})"
    if len(compact) <= 50:
        return compact
    return item_name[:47].rstrip() + "..." if len(item_name) > 50 else item_name


def compute_row_hash(payload_dict):
    """Deterministic SHA-256 hash of transformed product payload for re-run safety."""
    canonical = json.dumps(payload_dict, sort_keys=True, default=str)
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()


def write_formatted_output_excel(out_df, output_excel, template_path=TEMPLATE_PATH):
    """
    Writes the 104-column DataFrame to an .xlsx workbook matching OutPut_Format.xlsx:
    - Sheet name: 'Sheet1'
    - Clones OutPut_Format.xlsx when present to preserve header fills (FFC5FFB0),
      bold fonts, and Excel Data Validation rules (dynamically scaled to row count).
    """
    out_df = out_df.reindex(columns=OUTPUT_HEADERS)

    if template_path and Path(template_path).exists():
        shutil.copyfile(template_path, output_excel)
        wb = openpyxl.load_workbook(output_excel)
        ws = wb["Sheet1"] if "Sheet1" in wb.sheetnames else wb.active
        ws.title = "Sheet1"

        max_target_row = max(101, len(out_df) + 1)
        if ws.data_validations and ws.data_validations.dataValidation:
            for dv in ws.data_validations.dataValidation:
                dv.sqref = re.sub(r":([A-Z]+)101\b", rf":\g<1>{max_target_row}", str(dv.sqref))

        for r_idx, row_values in enumerate(out_df.itertuples(index=False, name=None), start=2):
            for c_idx, val in enumerate(row_values, start=1):
                cell_val = None if (pd.isna(val) or val == "") else val
                ws.cell(row=r_idx, column=c_idx, value=cell_val)

        wb.save(output_excel)
    else:
        with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
            out_df.to_excel(writer, sheet_name="Sheet1", index=False)
            ws = writer.sheets["Sheet1"]
            mandatory_fill = PatternFill(start_color="FFC5FFB0", end_color="FFC5FFB0", fill_type="solid")
            bold_font = Font(bold=True)
            for c_idx, col_name in enumerate(OUTPUT_HEADERS, start=1):
                cell = ws.cell(row=1, column=c_idx)
                cell.font = bold_font
                if col_name in MANDATORY_HEADERS:
                    cell.fill = mandatory_fill


def transform_catalog(
    input_csv="legacy_item_master_420_rows.csv",
    output_excel="fynd_bulk_import_supplementary.xlsx",
    exceptions_excel="catalog_exceptions_log.xlsx",
    template_path=TEMPLATE_PATH,
):
    df = pd.read_csv(input_csv, dtype=str)
    exceptions = []
    sync_audit_rows = []

    if MANIFEST_PATH.exists():
        manifest = json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))
    else:
        manifest = {}

    # Step 1: Pre-scan for duplicate item codes to detect Exact vs Conflicting duplicates
    seen_raw_rows = {}
    conflicting_codes = set()

    for idx, row in df.iterrows():
        excel_row_num = idx + 2  # 1-based header + 1
        code = clean_string(row.get("LEGACY_ITM_CD", "")).upper()
        if not code:
            continue

        row_sig = (
            normalize_title(row.get("ITM_DESC_MAIN", "")),
            normalize_title(row.get("BRAND_TXT", "")),
            clean_string(row.get("FREE_TXT_SPEC", "")),
            parse_mrp(row.get("MRP_STR", ""))[0],
            clean_string(row.get("UOM_CD", "")).upper(),
        )

        if code in seen_raw_rows:
            prev_row_num, prev_sig = seen_raw_rows[code]
            if row_sig != prev_sig:
                conflicting_codes.add(code)
                exceptions.append({
                    "source_row": excel_row_num,
                    "item_code": code,
                    "field": "LEGACY_ITM_CD",
                    "rejection_reason": (
                        f"Conflicting duplicate item_code (conflicts with row {prev_row_num}). "
                        f"Both records quarantined to prevent corrupt price/attribute overwrite."
                    ),
                })
            else:
                exceptions.append({
                    "source_row": excel_row_num,
                    "item_code": code,
                    "field": "LEGACY_ITM_CD",
                    "rejection_reason": f"Exact duplicate of row {prev_row_num}; deduplicated (first occurrence retained).",
                })
        else:
            seen_raw_rows[code] = (excel_row_num, row_sig)

    for code in conflicting_codes:
        first_row_num, _ = seen_raw_rows[code]
        exceptions.append({
            "source_row": first_row_num,
            "item_code": code,
            "field": "LEGACY_ITM_CD",
            "rejection_reason": "Primary record of a conflicting duplicate item_code; quarantined pending merchant sign-off.",
        })

    # Step 2: Validate and transform remaining records into the 104-column schema
    processed_codes = set()
    transformed_rows = []

    for idx, row in df.iterrows():
        excel_row_num = idx + 2
        code = clean_string(row.get("LEGACY_ITM_CD", "")).upper()

        if not code:
            exceptions.append({
                "source_row": excel_row_num,
                "item_code": "MISSING",
                "field": "LEGACY_ITM_CD",
                "rejection_reason": "Missing primary item_code (LEGACY_ITM_CD is blank).",
            })
            continue

        if code in conflicting_codes or code in processed_codes:
            continue
        processed_codes.add(code)

        # Item Name Validation
        item_name = normalize_title(row.get("ITM_DESC_MAIN", ""))
        if not item_name or "drop table" in item_name.lower() or len(item_name) < 3:
            exceptions.append({
                "source_row": excel_row_num,
                "item_code": code,
                "field": "ITM_DESC_MAIN",
                "rejection_reason": f"Invalid or missing item name: '{row.get('ITM_DESC_MAIN')}'",
            })
            continue

        # MRP Validation
        mrp_val, mrp_err = parse_mrp(row.get("MRP_STR", ""))
        if mrp_err:
            exceptions.append({
                "source_row": excel_row_num,
                "item_code": code,
                "field": "MRP_STR",
                "rejection_reason": mrp_err,
            })
            continue

        # UOM Normalization (mapped to DataDeclarations.xlsx 'Net Quantity Unit')
        raw_uom = clean_string(row.get("UOM_CD", "")).upper()
        norm_uom = UOM_MAP.get(raw_uom)
        if not norm_uom:
            exceptions.append({
                "source_row": excel_row_num,
                "item_code": code,
                "field": "UOM_CD",
                "rejection_reason": f"Unrecognized Unit of Measure '{raw_uom}'.",
            })
            continue

        # Brand Normalization (validated against DataDeclarations.xlsx 'Brand')
        brand_name = resolve_brand(row.get("BRAND_TXT", ""))

        # HSN Validation & Controlled Fallback (~8% missing)
        raw_hsn = re.sub(r"\D", "", clean_string(row.get("HSN_SAC_CD", "")))
        if len(raw_hsn) not in {4, 6, 8}:
            inferred_hsn = None
            for kw, hsn_code in CATEGORY_HSN_FALLBACK.items():
                if kw in item_name.lower():
                    inferred_hsn = hsn_code
                    break
            if inferred_hsn:
                raw_hsn = inferred_hsn
            else:
                exceptions.append({
                    "source_row": excel_row_num,
                    "item_code": code,
                    "field": "HSN_SAC_CD",
                    "rejection_reason": "Missing/invalid HSN code and product name did not match deterministic tax fallback rules.",
                })
                continue

        # Extract structured attributes from free-text spec
        attrs = extract_packed_attributes(row.get("FREE_TXT_SPEC", ""))
        tax_rule_name = resolve_tax_rule(row.get("GST_PCT", "5"))
        category_name = resolve_category(item_name, row.get("CATEGORY_TXT", ""))
        seller_sku = f"{code}-{attrs['size']}"
        primary_sku = f"{code}-{attrs['size']}"

        # Initialize all 104 columns in exact OutPut_Format.xlsx sequence
        fynd_record = {col: None for col in OUTPUT_HEADERS}

        fynd_record.update({
            "Name": item_name,
            "Slug": re.sub(r"[^a-z0-9]+", "-", f"{brand_name}-{item_name}-{code}".lower()).strip("-"),
            "Item Code": code,
            "Brand": brand_name,
            "Category": category_name,
            "Description": (
                f"{item_name} by {brand_name}. Net Weight/Volume: {attrs['net_weight_volume']}, "
                f"Packaging: {attrs['packaging_type']}, Pack of {attrs['pack_quantity']}."
            ),
            "Short Description": build_short_description(item_name, attrs),
            "Tax Rule Name": tax_rule_name,
            "HS Code": raw_hsn,
            "Country of Origin": "India",
            "Media": clean_string(row.get("MEDIA_URL", "")) or None,
            "Multi Size": "No",
            "Gtin Type": "SKU",
            "Gtin Value": primary_sku,
            "Seller Identifier": primary_sku,
            "Size": attrs["size"],
            "Actual Price": mrp_val,
            "Selling Price": mrp_val,
            "Currency": "INR",
            "Length (cm)": DEFAULT_LENGTH_CM,
            "Width (cm)": DEFAULT_WIDTH_CM,
            "Height (cm)": DEFAULT_HEIGHT_CM,
            "Product Dead Weight (gram)": attrs["dead_weight_grams"],
            "Available": "Yes",
            "Highlights": (
                f"1. Net Weight/Volume : {attrs['net_weight_volume']}\n"
                f"2. Packaging Type : {attrs['packaging_type']}\n"
                f"3. Pack Quantity : {attrs['pack_quantity']}"
            ),
            "Unlisted Product": "No",
            "Trader Type": "Manufacturer",
            "Trader Name": "Teafields Retail Pvt Ltd",
            "Trader Address": "Plot 14, MIDC Industrial Area, Navi Mumbai, MH 400705",
            "Track Inventory": "Yes",
            "No of Boxes": 1,
            "Manufacturing Time": 2,
            "Manufacturing Time Unit": "Days",
            "Return Time Limit": 7,
            "Return Time Unit": "Days",
            "Net Quantity Value": attrs["pack_quantity"],
            "Net Quantity Unit": norm_uom,
            "Colour": attrs["color"],
            "Custom Attribute 1": attrs["net_weight_volume"],
            "Custom Attribute 2": str(attrs["pack_quantity"]),
            "Custom Attribute 3": attrs["packaging_type"],
            "Custom Attribute 6": attrs["pack_quantity"],
            "Package Contents": f"Pack of {attrs['pack_quantity']} - {item_name} ({attrs['packaging_type']})",
        })

        row_hash = compute_row_hash(fynd_record)
        sync_action = (
            "SKIP_UNCHANGED" if manifest.get(code) == row_hash else ("UPDATE" if code in manifest else "CREATE")
        )
        manifest[code] = row_hash
        transformed_rows.append(fynd_record)
        sync_audit_rows.append({
            "Item Code": code,
            "Seller Identifier": primary_sku,
            "Sync Action": sync_action,
            "Idempotency Hash": row_hash,
        })

    # Write Fynd Bulk Import Excel matching OutPut_Format.xlsx (104 columns in 'Sheet1')
    out_df = pd.DataFrame(transformed_rows, columns=OUTPUT_HEADERS)
    write_formatted_output_excel(out_df, output_excel=output_excel, template_path=template_path)

    # Write Exceptions & Idempotency Audit Report
    exc_df = pd.DataFrame(exceptions, columns=["source_row", "item_code", "field", "rejection_reason"])
    if not exc_df.empty:
        exc_df = exc_df.sort_values(by="source_row")
    audit_df = pd.DataFrame(
        sync_audit_rows,
        columns=["Item Code", "Seller Identifier", "Sync Action", "Idempotency Hash"],
    )

    with pd.ExcelWriter(exceptions_excel, engine="openpyxl") as writer:
        exc_df.to_excel(writer, sheet_name="Rejected_Rows", index=False)
        audit_df.to_excel(writer, sheet_name="Sync_Audit_Log", index=False)

    # Save updated idempotency manifest
    MANIFEST_PATH.write_text(json.dumps(manifest, indent=2), encoding="utf-8")

    print(f"Transformation Complete:")
    print(f"  Total Input Rows      : {len(df)}")
    print(f"  Clean Transformed SKUs: {len(out_df)}")
    print(f"  Logged Exceptions     : {len(exc_df)}")


if __name__ == "__main__":
    transform_catalog()
