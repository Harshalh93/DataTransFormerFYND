import random
import pandas as pd
import numpy as np

random.seed(42)
np.random.seed(42)

def generate_defective_legacy_export(filename="legacy_item_master_420_rows.csv"):
    """
    Generates a 420-row defective legacy ERP export simulating a 10-year patched database.
    Includes:
    - Casing/whitespace defects in item_name and brand
    - Exact duplicate rows & conflicting attribute duplicate item_codes
    - Messy text MRP (currency symbols, commas, spaces, blanks, zero/negative)
    - Packed attributes (size, weight, colour, pack count) inside free-text descriptions
    - ~8% missing HSN codes & ~4% missing brands
    - Mixed UOMs (PCS, Pcs, NOS, EA, KG, Kgs, GM, Grams)
    - Genuinely unusable rows that must be rejected (corrupt prices, missing names, negative MRP)
    """
    brands = [
        "Teafields Reserve", "  teafields reserve ", "TEAFIELDS RESERVE",
        "Himalayan Brew", "himalayan brew  ", "HIMALAYAN BREW",
        "Artisan Leaf", " artisan leaf", "Gourmet Crumb", "gourmet crumb "
    ]
    
    base_items = [
        ("Darjeeling First Flush Loose Tea", "250g | Tin Caddy | Color: Gold | Pack of 1", "09022020"),
        ("Assam Orthodox Black Tea", "500 GM - Kraft Pouch - Color: Brown - Pack of 2", "09024020"),
        ("Nilgiri Frost Green Tea", "100g, Zipper Bag, Color: Green, Pack of 1", "09021090"),
        ("Masala Chai Spice Blend", "200g / Glass Jar / Color: Red / Pack of 1", "09109100"),
        ("Earl Grey Bergamot Infusion", "50 Pyramid Bags (125g) Color: Blue Pack of 1", "09023020"),
        ("Chamomile Honey Herbal Tisane", "150g Box Color: Yellow Pack of 3", "21069099"),
        ("Cast Iron Japanese Teapot", "Size: Medium | 800ml | Color: Matte Black | Pack of 1", "73239190"),
        ("Double Wall Borosilicate Infuser Mug", "Size: Large 350ml Color: Clear Pack of 2", "70134200"),
        ("Artisanal Almond Biscotti", "300g Pack of 1 Color: Beige", "19053100"),
        ("Organic Wild Forest Honey", "500g Glass Bottle Color: Amber Pack of 1", "04090000"),
    ]
    
    uoms = ["PCS", "Pcs", "NOS", "EA", "KG", "Kgs", "pcs ", " Nos", "GM", "Gms"]
    mrp_formats = [
        lambda p: f"Rs. {p:,.2f}",
        lambda p: f"INR {p:,}",
        lambda p: f"₹{p:,.0f}/-",
        lambda p: f"  {p} ",
        lambda p: f"Rs {p}"
    ]

    rows = []
    # 1. Generate 385 normal + mildly defective rows
    for i in range(1, 386):
        sku_code = f"TF-SKU-{1000 + i}"
        base_name, desc_pattern, default_hsn = random.choice(base_items)
        
        # Introduce casing & whitespace defects in item name
        name_variant = random.choice([
            f"  {base_name.lower()}  ",
            f"{base_name.upper()} ",
            f" {base_name.title()}   ",
            base_name
        ])
        
        # ~4% missing brand
        if random.random() < 0.042:
            brand_val = random.choice(["", "   ", None, "N/A"])
        else:
            brand_val = random.choice(brands)
            
        # ~8% missing HSN
        if random.random() < 0.082:
            hsn_val = random.choice(["", None, "   "])
        else:
            hsn_val = default_hsn
            
        base_price = random.choice([250, 399, 450, 650, 899, 1250, 1499, 2100])
        # Occasional blank MRP (~1.5% here + explicit corrupt rows below)
        if random.random() < 0.015:
            mrp_val = ""
        else:
            mrp_val = random.choice(mrp_formats)(base_price)
            
        uom_val = random.choice(uoms)
        
        # Legacy mystery columns from a 10-year-old patched DB
        rows.append({
            "LEGACY_ITM_CD": sku_code,
            "ITM_DESC_MAIN": name_variant,
            "BRAND_TXT": brand_val,
            "FREE_TXT_SPEC": desc_pattern,
            "MRP_STR": mrp_val,
            "UOM_CD": uom_val,
            "HSN_SAC_CD": hsn_val,
            "GST_PCT": random.choice([5, 12, 18]),
            "OLD_FOXPRO_FLAG": random.choice(["Y", "N", "1", None, "LEGACY_99"]),
            "TEMP_PATCH_2018": random.choice([" NULL ", "0", "MIG_OK", ""])
        })

    # 2. Inject 10 Exact Duplicate Rows
    for i in range(10):
        rows.append(rows[i].copy())

    # 3. Inject 10 Conflicting Duplicate Item Codes (same LEGACY_ITM_CD, different MRP or attributes)
    for i in range(10, 20):
        conflict_row = rows[i].copy()
        conflict_row["MRP_STR"] = "Rs. 9,999.00"  # Conflicting price on same item code
        conflict_row["FREE_TXT_SPEC"] = "1kg | Bulk Sack | Color: White | Pack of 5"
        rows.append(conflict_row)

    # 4. Inject 15 Genuinely Unusable Rows (Must be rejected, not guessed)
    unusable_cases = [
        {"LEGACY_ITM_CD": "", "ITM_DESC_MAIN": "Orphan Tea Item Without Code", "BRAND_TXT": "Teafields Reserve", "FREE_TXT_SPEC": "250g", "MRP_STR": "Rs. 450", "UOM_CD": "PCS", "HSN_SAC_CD": "09022020", "GST_PCT": 5, "OLD_FOXPRO_FLAG": "N", "TEMP_PATCH_2018": ""},
        {"LEGACY_ITM_CD": "TF-BAD-001", "ITM_DESC_MAIN": "   ", "BRAND_TXT": "Teafields Reserve", "FREE_TXT_SPEC": "250g", "MRP_STR": "Rs. 450", "UOM_CD": "PCS", "HSN_SAC_CD": "09022020", "GST_PCT": 5, "OLD_FOXPRO_FLAG": "N", "TEMP_PATCH_2018": ""},
        {"LEGACY_ITM_CD": "TF-BAD-002", "ITM_DESC_MAIN": "Corrupt Negative Price Tea", "BRAND_TXT": "Himalayan Brew", "FREE_TXT_SPEC": "100g", "MRP_STR": "Rs. -450.00", "UOM_CD": "PCS", "HSN_SAC_CD": "09021090", "GST_PCT": 5, "OLD_FOXPRO_FLAG": "N", "TEMP_PATCH_2018": ""},
        {"LEGACY_ITM_CD": "TF-BAD-003", "ITM_DESC_MAIN": "Zero Price Promo Glitch", "BRAND_TXT": "Himalayan Brew", "FREE_TXT_SPEC": "100g", "MRP_STR": "INR 0.00", "UOM_CD": "PCS", "HSN_SAC_CD": "09021090", "GST_PCT": 5, "OLD_FOXPRO_FLAG": "N", "TEMP_PATCH_2018": ""},
        {"LEGACY_ITM_CD": "TF-BAD-004", "ITM_DESC_MAIN": "Unparseable Text MRP Item", "BRAND_TXT": "Artisan Leaf", "FREE_TXT_SPEC": "500g", "MRP_STR": "ASK MANAGER", "UOM_CD": "KG", "HSN_SAC_CD": "09024020", "GST_PCT": 5, "OLD_FOXPRO_FLAG": "N", "TEMP_PATCH_2018": ""},
        {"LEGACY_ITM_CD": "TF-BAD-005", "ITM_DESC_MAIN": "SQL Injection / Corrupt Row ); DROP TABLE--", "BRAND_TXT": None, "FREE_TXT_SPEC": "???", "MRP_STR": "NaN", "UOM_CD": "UNKNOWN_UOM", "HSN_SAC_CD": "INVALID", "GST_PCT": 99, "OLD_FOXPRO_FLAG": "ERR", "TEMP_PATCH_2018": "ERR"},
    ]
    # Pad out to 15 unusable rows
    for idx in range(6, 16):
        unusable_cases.append({
            "LEGACY_ITM_CD": f"TF-BAD-{idx:03d}",
            "ITM_DESC_MAIN": f"Legacy Corrupt Record #{idx}" if idx % 2 == 0 else None,
            "BRAND_TXT": "Teafields Reserve",
            "FREE_TXT_SPEC": "Corrupt spec string",
            "MRP_STR": "" if idx % 3 == 0 else "TBD",
            "UOM_CD": "PCS",
            "HSN_SAC_CD": "09022020",
            "GST_PCT": 5,
            "OLD_FOXPRO_FLAG": "ERR",
            "TEMP_PATCH_2018": ""
        })
    rows.extend(unusable_cases)

    df = pd.DataFrame(rows)
    df.to_csv(filename, index=False)
    print(f"Generated {len(df)} rows into {filename}")

if __name__ == "__main__":
    generate_defective_legacy_export()