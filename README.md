# Legacy Item Master to Fynd Commerce Catalog Transformer (Part A)

## 1. Overview & Setup
This repository contains the data generator, AI-assisted deterministic catalog transformer, and FDK Private Extension ingestion script for migrating a 10-year-old legacy item master into Fynd Commerce under the **Others (Supplementary)** template.

### Quick Start
1. Install Python dependencies: `pip install pandas openpyxl`
2. (Optional) Regenerate the 420-row defective legacy dataset: `python generate_legacy_data.py`
3. Run the catalog transformer: `python transform_catalog.py`
   - Outputs `fynd_bulk_import_supplementary.xlsx` (clean import sheet)
   - Outputs `catalog_exceptions_log.xlsx` (row-level rejection report)
   - Updates `idempotency_manifest.json` (SHA-256 state tracker)
4. Run FDK ingestion inside your registered Private Extension:
   - `npm install @gofynd/fdk-client-javascript xlsx`
   - `FYND_COMPANY_ID=<id> FYND_EXT_API_KEY=<key> FYND_EXT_API_SECRET=<secret> node ingest_via_fdk.js`

---

## 2. Full Field Mapping Table
| Legacy Column | Defect Observed in Source | Transformation & Normalization Logic | Target Fynd Field (`Others - Supplementary`) |
| :--- | :--- | :--- | :--- |
| `LEGACY_ITM_CD` | Blank rows, exact duplicates, conflicting duplicates | Trim/uppercase; drop exact duplicates (keep 1st); quarantine both rows if same code has conflicting price/specs | `Item Code` & `Seller Identifier (SKU)` |
| `ITM_DESC_MAIN` | Random casing (`upper`, `lower`), leading/trailing spaces, corrupt SQL text | Collapse whitespace, convert to Title Case; reject if blank or corrupt | `Product Name` |
| `BRAND_TXT` | Inconsistent casing/whitespace, ~4% blank/`N/A` | Normalize to Title Case; impute `Teafields House Brand` when blank (business-safe fallback) | `Brand` |
| `FREE_TXT_SPEC` | Buried weight, size, colour, pack count, packaging type | Regex extraction into 5 structured attributes | `Size`, plus Custom Attributes: `Color`, `Net Weight / Volume`, `Pack Quantity`, `Packaging Type` |
| `MRP_STR` | Stored as text (`Rs. 1,250.00`, `₹899/-`), blanks, negative/zero values | Strip currency symbols/commas; parse to positive `float`; reject if blank, `<= 0`, or non-numeric | `MRP` & `Effective Price` |
| `UOM_CD` | Mixed tokens (`PCS`, `Pcs`, `NOS`, `EA`, `KG`, `Kgs`, `GM`, `Gms`) | Map `PCS/Pcs/NOS/EA` -> `PCS`; `KG/Kgs` -> `KG`; `GM/Gms` -> `GM`; reject unknown UOMs | `Unit of Measure` |
| `HSN_SAC_CD` | ~8% missing/blank | Validate 4/6/8-digit numeric HSN; apply deterministic keyword fallback (`tea`->`09024020`, etc.); reject if unresolvable | `HSN Code` |
| `GST_PCT` | Invalid rates (`99%`) on corrupt rows | Validate against Indian GST slabs (`0, 5, 12, 18, 28`) | `Tax Rule / GST (%)` |
| `OLD_FOXPRO_FLAG`, `TEMP_PATCH_2018` | Unexplained legacy flags (`LEGACY_99`, `MIG_OK`) | Explicitly deprecated and ignored after verifying zero impact on live POS billing | *(Unmapped / Ignored)* |

---

## 3. How Re-Run Safety (Idempotency) Is Achieved
Running both the transformer and the Fynd import multiple times will **never** create duplicate products:
1. **Deterministic Natural Key (`item_code` + `slug`)**: Every product's `item_code` (`LEGACY_ITM_CD`) and `slug` are deterministically derived. In Fynd Commerce, `item_code` + `brand_uid` acts as the unique catalog key; re-uploading the same `item_code` via the Bulk Importer or Platform API updates the existing item rather than creating a new `uid`.
2. **SHA-256 State Manifest (`idempotency_manifest.json`)**: `transform_catalog.py` computes a SHA-256 hash of every cleaned product payload. On subsequent runs, unchanged records are marked `SKIP_UNCHANGED`, modified records are marked `UPDATE`, and only new records are marked `CREATE`.

---

## 4. Stated Assumptions
1. **Conflicting Duplicates Policy**: When two rows share the same `LEGACY_ITM_CD` but carry different MRPs or weights, neither row can be trusted on a retail shop floor (guessing risks under-charging or over-charging customers). Both records are quarantined into `catalog_exceptions_log.xlsx` for merchant sign-off.
2. **Missing Brand (~4%) vs. Missing MRP**: A missing brand can safely default to the merchant's house brand (`Teafields House Brand`) so the SKU remains sellable on Day 1, whereas a missing or zero MRP is a hard fiscal blocker and is strictly rejected.
