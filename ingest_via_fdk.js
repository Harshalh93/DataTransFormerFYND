/**
 * File: ingest_via_fdk.js
 * Vehicle: Private Fynd Extension built via FDK CLI (@gofynd/fdk-client-javascript)
 * Purpose: Reads the 104-column OutPut_Format.xlsx compliant workbook and performs
 *          idempotent upsert (createProduct / editProduct) into Fynd Commerce.
 */
const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");
const { PlatformConfig, PlatformClient } = require("@gofynd/fdk-client-javascript");

// ==============================================================================
// 1. CONFIGURATION & DATA DECLARATIONS LOOKUP MAPS
// ==============================================================================
const config = new PlatformConfig({
  companyId: parseInt(process.env.FYND_COMPANY_ID, 10),
  apiKey: process.env.FYND_EXT_API_KEY,
  apiSecret: process.env.FYND_EXT_API_SECRET,
  domain: process.env.FYND_DOMAIN || "https://api.fynd.com",
});

// Map DataDeclarations.xlsx Brands to Fynd brand_uid values (configurable via env)
const BRAND_UID_MAP = {
  "Generic": parseInt(process.env.FYND_BRAND_UID_GENERIC || process.env.FYND_DEFAULT_BRAND_UID || "1", 10),
  "Teafields Reserve": parseInt(process.env.FYND_BRAND_UID_TEAFIELDS || process.env.FYND_DEFAULT_BRAND_UID || "1", 10),
  "Himalayan Brew": parseInt(process.env.FYND_BRAND_UID_HIMALAYAN || process.env.FYND_DEFAULT_BRAND_UID || "1", 10),
  "Artisan Leaf": parseInt(process.env.FYND_BRAND_UID_ARTISAN || process.env.FYND_DEFAULT_BRAND_UID || "1", 10),
  "Gourmet Crumb": parseInt(process.env.FYND_BRAND_UID_GOURMET || process.env.FYND_DEFAULT_BRAND_UID || "1", 10),
};

// Map DataDeclarations.xlsx Categories to Fynd category slugs
const CATEGORY_SLUG_MAP = {
  "Others level 3": "others-level-3",
  "Bottle": "bottle",
  "Carry Bag": "carry-bag",
};

// Allowed Net Quantity Units from DataDeclarations.xlsx:
// ['kg', 'g', 'L', 'ml', 'm', 'cm', 'nos', 'pack', 'ea', 'set']
const ALLOWED_UNITS = new Set(["kg", "g", "L", "ml", "m", "cm", "nos", "pack", "ea", "set"]);

// ==============================================================================
// 2. NORMALIZATION & HELPER FUNCTIONS
// ==============================================================================
function cleanStr(val, fallback = "") {
  if (val === undefined || val === null) return fallback;
  return String(val).replace(/\s+/g, " ").trim() || fallback;
}

/**
 * Restores leading zero if Excel numeric coercion turned an 8-digit HSN like
 * '09024020' into 7-digit '9024020'.
 */
function normalizeHsnCode(rawHsn) {
  const digits = cleanStr(rawHsn).replace(/\D/g, "");
  if (digits.length === 3 || digits.length === 5 || digits.length === 7) {
    return "0" + digits;
  }
  return digits;
}

/**
 * Preserves uppercase 'L' from DataDeclarations.xlsx while keeping other units lowercase.
 */
function normalizeNetQuantityUnit(rawUnit) {
  const u = cleanStr(rawUnit, "nos");
  if (u.toUpperCase() === "L") return "L";
  const lower = u.toLowerCase();
  return ALLOWED_UNITS.has(lower) ? lower : "nos";
}

/**
 * Loads Sync Action and Idempotency Hash from catalog_exceptions_log.xlsx ('Sync_Audit_Log')
 * and idempotency_manifest.json so Gate 1 works cleanly with the 104-column Sheet1 layout.
 */
function loadSyncAuditMap(exceptionsExcelPath, manifestJsonPath) {
  const auditMap = new Map();

  if (exceptionsExcelPath && fs.existsSync(exceptionsExcelPath)) {
    try {
      const excWb = xlsx.readFile(exceptionsExcelPath);
      const auditSheet = excWb.Sheets["Sync_Audit_Log"];
      if (auditSheet) {
        const auditRows = xlsx.utils.sheet_to_json(auditSheet, { defval: "" });
        for (const r of auditRows) {
          const code = cleanStr(r["Item Code"]).toUpperCase();
          if (code) {
            auditMap.set(code, {
              syncAction: cleanStr(r["Sync Action"], "CREATE").toUpperCase(),
              idempotencyHash: cleanStr(r["Idempotency Hash"]),
            });
          }
        }
      }
    } catch (err) {
      console.warn(`Warning: Could not read Sync_Audit_Log from ${exceptionsExcelPath}: ${err.message}`);
    }
  }

  if (manifestJsonPath && fs.existsSync(manifestJsonPath)) {
    try {
      const manifest = JSON.parse(fs.readFileSync(manifestJsonPath, "utf-8"));
      for (const [code, hash] of Object.entries(manifest)) {
        const upperCode = cleanStr(code).toUpperCase();
        if (!auditMap.has(upperCode)) {
          auditMap.set(upperCode, { syncAction: "CREATE", idempotencyHash: String(hash) });
        } else if (!auditMap.get(upperCode).idempotencyHash) {
          auditMap.get(upperCode).idempotencyHash = String(hash);
        }
      }
    } catch (err) {
      console.warn(`Warning: Could not read manifest ${manifestJsonPath}: ${err.message}`);
    }
  }

  return auditMap;
}

/**
 * Queries Fynd Catalog API by item_code to check if the product already exists.
 * Returns existing product uid (number/string) if found, otherwise null.
 */
async function findExistingProductUid(platformClient, itemCode) {
  if (typeof platformClient.catalog?.getProducts !== "function") {
    return null;
  }
  try {
    const response = await platformClient.catalog.getProducts({ itemCode: [itemCode] });
    const items = response?.items || response?.data || [];
    const match = items.find(
      (item) => cleanStr(item.item_code).toUpperCase() === itemCode.toUpperCase()
    );
    return match ? match.uid || match.id || null : null;
  } catch (_) {
    return null;
  }
}

// ==============================================================================
// 3. MAIN CATALOG INGESTION WORKFLOW
// ==============================================================================
async function ingestCatalog(
  excelPath = "./fynd_bulk_import_supplementary.xlsx",
  exceptionsExcelPath = "./catalog_exceptions_log.xlsx",
  manifestJsonPath = "./idempotency_manifest.json"
) {
  // 1. Obtain OAuth token for the development company
  const tokenData = await config.oauthClient.getAccesstokenObj({
    grant_type: "client_credentials",
  });
  config.oauthClient.setToken(tokenData);
  const platformClient = new PlatformClient(config);

  // 2. Read transformed 104-column Excel workbook ('Sheet1' from OutPut_Format.xlsx)
  const resolvedExcelPath = path.resolve(excelPath);
  const workbook = xlsx.readFile(resolvedExcelPath, { raw: false });
  const targetSheetName =
    workbook.Sheets["Sheet1"]
      ? "Sheet1"
      : workbook.Sheets["Supplementary"]
      ? "Supplementary"
      : workbook.SheetNames[0];

  const rows = xlsx.utils.sheet_to_json(workbook.Sheets[targetSheetName], { defval: "" });
  const syncAuditMap = loadSyncAuditMap(exceptionsExcelPath, manifestJsonPath);

  console.log(
    `Loaded ${rows.length} validated rows from sheet '${targetSheetName}' for Fynd ingestion.`
  );

  const results = {
    created: 0,
    updated: 0,
    skipped_unchanged: 0,
    api_errors: [],
  };

  for (const [index, row] of rows.entries()) {
    const excelRowNumber = index + 2;
    const itemCode = cleanStr(row["Item Code"]).toUpperCase();

    if (!itemCode) {
      results.api_errors.push({
        row_number: excelRowNumber,
        item_code: "MISSING",
        field: "Item Code",
        reason: "Blank Item Code in worksheet row.",
      });
      continue;
    }

    // Resolve Sync Action & Idempotency Hash from Sync_Audit_Log or inline columns
    const auditEntry = syncAuditMap.get(itemCode) || {};
    const syncAction = cleanStr(row["Sync Action"] || auditEntry.syncAction, "CREATE").toUpperCase();
    const idempotencyHash = cleanStr(row["Idempotency Hash"] || auditEntry.idempotencyHash);

    // Re-run safety Gate 1: Local hash manifest check
    if (syncAction === "SKIP_UNCHANGED") {
      results.skipped_unchanged++;
      continue;
    }

    // Extract & normalize fields from the 104-column OutPut_Format.xlsx schema
    const productName = cleanStr(row["Name"] || row["Product Name"]);
    const brandName = cleanStr(row["Brand"], "Generic");
    const categoryName = cleanStr(row["Category"] || row["Category L3"], "Others level 3");
    const sizeVal = cleanStr(row["Size"], "OS");

    // CRITICAL PRIMARY IDENTIFIER RESOLUTION:
    // Ensures Seller Identifier == Gtin Value and marks primary: true in identifiers[]
    // to prevent: "Product with item code '...' and size '...' should have one primary identifer"
    const primaryIdentifier = cleanStr(
      row["Seller Identifier"] || row["Gtin Value"] || row["Seller Identifier (SKU)"] || `${itemCode}-${sizeVal}`
    );
    const gtinType = cleanStr(row["Gtin Type"], "SKU");

    const actualPrice = Number(row["Actual Price"] !== "" ? row["Actual Price"] : row["MRP"]);
    const sellingPrice = Number(
      row["Selling Price"] !== "" ? row["Selling Price"] : row["Effective Price"] || actualPrice
    );

    // Enforce Length > Width per Definitions.xlsx (Cols V, W, X, Y)
    const lengthCm = Math.max(1, Number(row["Length (cm)"] || 15));
    const rawWidthCm = Math.max(1, Number(row["Width (cm)"] || 10));
    const widthCm = lengthCm > rawWidthCm ? rawWidthCm : Math.max(0.5, lengthCm - 1);
    const heightCm = Math.max(1, Number(row["Height (cm)"] || 10));
    const deadWeightGram = Math.max(1, Number(row["Product Dead Weight (gram)"] || 100));

    const netQtyValue = Math.max(1, Number(row["Net Quantity Value"] || row["Net Quantity"] || 1));
    const netQtyUnit = normalizeNetQuantityUnit(row["Net Quantity Unit"] || row["Unit of Measure"]);
    const hsnCode = normalizeHsnCode(row["HS Code"] || row["HSN Code"]);
    const shortDesc = cleanStr(row["Short Description"]).slice(0, 50);
    const fullDesc = cleanStr(row["Description"] || shortDesc);
    const mediaUrl = cleanStr(row["Media"]);

    const highlightsList = cleanStr(row["Highlights"])
      .split(/\r?\n/)
      .map((line) => line.replace(/^\d+\.\s*/, "").trim())
      .filter(Boolean);

    const payload = {
      item_code: itemCode,
      slug: cleanStr(row["Slug"]),
      name: productName,
      brand: brandName,
      brand_uid: BRAND_UID_MAP[brandName] || parseInt(process.env.FYND_DEFAULT_BRAND_UID || "1", 10),
      department_uid: parseInt(process.env.FYND_OTHERS_DEPT_UID || "1", 10),
      category_slug: CATEGORY_SLUG_MAP[categoryName] || "others-level-3",
      category: categoryName,
      template_tag: "supplementary",
      item_type: "standard",
      multi_size: cleanStr(row["Multi Size"], "No").toLowerCase() === "yes",
      is_active: cleanStr(row["Available"], "Yes").toLowerCase() === "yes",
      is_dependent: cleanStr(row["Unlisted Product"], "No").toLowerCase() === "yes",
      hsn_code: hsnCode,
      hs_code: hsnCode,
      tax_rule_name: cleanStr(row["Tax Rule Name"], "Merit Tax Rule \u2013 5% (Eff. 22 Sep 2025)"),
      country_of_origin: cleanStr(row["Country of Origin"], "India"),
      currency: cleanStr(row["Currency"], "INR"),
      description: fullDesc,
      short_description: shortDesc,
      highlights: highlightsList,
      media: mediaUrl ? [{ type: "image", url: mediaUrl }] : [],
      trader: [
        {
          type: cleanStr(row["Trader Type"], "Manufacturer"),
          name: cleanStr(row["Trader Name"], "Teafields Retail Pvt Ltd"),
          address: [
            cleanStr(
              row["Trader Address"],
              "Plot 14, MIDC Industrial Area, Navi Mumbai, MH 400705"
            ),
          ],
        },
      ],
      custom_json: {
        color: cleanStr(row["Colour"] || row["Color (Custom Attribute)"]),
        net_weight_volume: cleanStr(
          row["Custom Attribute 1"] || row["Net Weight / Volume (Custom Attribute)"]
        ),
        pack_quantity: cleanStr(
          row["Custom Attribute 2"] || row["Pack Quantity (Custom Attribute)"],
          String(netQtyValue)
        ),
        packaging_type: cleanStr(
          row["Custom Attribute 3"] || row["Packaging Type (Custom Attribute)"]
        ),
        pack_quantity_int: Number(row["Custom Attribute 6"] || netQtyValue),
        package_contents: cleanStr(row["Package Contents"]),
        unit_of_measure: netQtyUnit,
        idempotency_hash: idempotencyHash,
      },
      sizes: [
        {
          size: sizeVal,
          price: actualPrice,
          price_effective: sellingPrice,
          seller_identifier: primaryIdentifier,
          primary_identifier: primaryIdentifier,
          identifiers: [
            {
              gtin_type: gtinType,
              gtin_value: primaryIdentifier,
              primary: true,
            },
          ],
          track_inventory: cleanStr(row["Track Inventory"], "Yes").toLowerCase() === "yes",
          quantity: Math.max(1, parseInt(row["No of Boxes"] || "1", 10)),
          item_length: lengthCm,
          item_width: widthCm,
          item_height: heightCm,
          item_weight: deadWeightGram,
          length: lengthCm,
          width: widthCm,
          height: heightCm,
          weight: deadWeightGram,
          net_quantity: {
            value: netQtyValue,
            unit: netQtyUnit,
          },
          manufacturing_time: Math.max(0, parseInt(row["Manufacturing Time"] || "2", 10)),
          manufacturing_time_unit: cleanStr(row["Manufacturing Time Unit"], "Days").toLowerCase(),
          return_config: {
            returnable: Number(row["Return Time Limit"] || 7) > 0,
            time: Math.max(0, parseInt(row["Return Time Limit"] || "7", 10)),
            unit: cleanStr(row["Return Time Unit"], "Days").toLowerCase(),
          },
        },
      ],
    };

    try {
      // Re-run safety Gate 2: Query Fynd Catalog by item_code before write
      const existingUid = await findExistingProductUid(platformClient, itemCode);

      if (existingUid && typeof platformClient.catalog?.editProduct === "function") {
        await platformClient.catalog.editProduct({
          itemId: existingUid,
          body: payload,
        });
        results.updated++;
      } else {
        await platformClient.catalog.createProduct({ body: payload });
        results.created++;
      }
    } catch (err) {
      results.api_errors.push({
        row_number: excelRowNumber,
        item_code: itemCode,
        seller_identifier: primaryIdentifier,
        field: "Catalog API Payload",
        reason: err?.response?.data
          ? JSON.stringify(err.response.data)
          : err.message || JSON.stringify(err),
      });
    }
  }

  console.log("Ingestion Summary:", JSON.stringify(results, null, 2));
}

ingestCatalog().catch(console.error);