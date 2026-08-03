/* ==========================================================
 * outbound-availability.js — Outbound: Combined Real-Time Inventory
 * (Bundle-Expanded Availability Report)
 *
 * Purpose: produce ONE combined report from the Real-Time Detailed
 * Inventory file (the source of truth for what currently exists) plus
 * the Bundle Child Product file (a mapping used only to expand bundle
 * cartons that are CURRENTLY present and available in the real-time
 * inventory). This is not a bundle audit/selection tool — processing
 * always starts from the real-time inventory rows, never from the
 * bundle mapping file.
 *
 * Per available (Available Qty > 0) inventory row:
 *  - if it matches a Bundle SKU in the mapping -> expand into one row
 *    per Child SKU (EAN Qty = SL from the mapping; Available Qty is
 *    NEVER shown/derived on these rows, since 1 Bundle SKU already =
 *    1 physical available carton and the child qty already represents
 *    what's inside it — no multiplication, no "Available PCS");
 *  - else if it looks like a bundle/MIX row but has no mapping -> the
 *    row is dropped and logged as a BUNDLE_MAPPING_NOT_FOUND diagnostic
 *    (its Child SKU composition can't be determined, so it must not be
 *    kept as an ordinary item);
 *  - else -> kept as-is, a normal inventory row.
 * A Bundle SKU that exists only in the mapping file (not currently in
 * inventory) never appears in the output at all.
 *
 * 100% client-side (SheetJS, already loaded globally as `XLSX` in
 * index.html) — files never leave the browser.
 * ========================================================== */
window.WOPOutboundAvailability = (function () {
  "use strict";

  // ----------------------------------------------------------------
  // 1. Normalization helpers
  // ----------------------------------------------------------------
  function normalizeHeaderText(h) {
    if (h === null || h === undefined) return "";
    let s = String(h).toLowerCase();
    s = s.replace(/đ/g, "d");
    s = s.normalize("NFD").replace(/[̀-ͯ]/g, "");
    s = s.replace(/[^a-z0-9]+/g, " ").trim().replace(/\s+/g, " ");
    return s;
  }

  function normalizeEan(v) {
    if (v === null || v === undefined) return "";
    if (typeof v === "number") {
      if (!isFinite(v)) return "";
      return Number.isInteger(v) ? String(v) : String(Math.round(v));
    }
    let s = String(v).trim();
    if (s === "") return "";
    s = s.replace(/\s+/g, "");
    s = s.replace(/\.0+$/, "");
    return s;
  }

  // Bundle SKU (alphanumeric identifier, e.g. "TP-251223-09"): string-safe,
  // trims spaces + NBSP, strips a stray Excel-generated trailing ".0",
  // preserves original casing for display; matching key is case-insensitive.
  function normalizeBundleSkuDisplay(v) {
    if (v === null || v === undefined) return "";
    let s = String(v).replace(/ /g, " ").trim();
    if (s === "") return "";
    s = s.replace(/\.0+$/, "");
    return s;
  }
  function normalizeBundleSkuKey(v) { return normalizeBundleSkuDisplay(v).toLowerCase(); }

  function toNumberOrNull(v) {
    if (v === null || v === undefined) return null;
    if (typeof v === "number") return isFinite(v) ? v : null;
    let s = String(v).trim();
    if (s === "") return null;
    s = s.replace(/,/g, "");
    if (!/^-?\d+(\.\d+)?$/.test(s)) return null;
    const n = Number(s);
    return isFinite(n) ? n : null;
  }

  function cleanStr(v) {
    if (v === null || v === undefined) return "";
    return String(v).trim();
  }

  function isMixUnit(unit) {
    return normalizeHeaderText(unit) === "mix";
  }

  // Numeric-aware comparator for product/bundle codes (Warehouse -> Barcode
  // sort key): compares as numbers when both sides are pure digit strings
  // (typical EANs), otherwise falls back to plain string order (typical
  // alphanumeric Bundle SKUs / Partner SKUs).
  function compareCodes(a, b) {
    const as = a === null || a === undefined ? "" : String(a);
    const bs = b === null || b === undefined ? "" : String(b);
    if (as === bs) return 0;
    const an = /^\d+$/.test(as), bn = /^\d+$/.test(bs);
    if (an && bn) { const na = Number(as), nb = Number(bs); return na < nb ? -1 : (na > nb ? 1 : 0); }
    return as < bs ? -1 : 1;
  }

  // ----------------------------------------------------------------
  // Final EAN Qty rule (single source of truth for BOTH the web preview
  // and the Excel export — never computed separately in either place):
  //   - expanded MIX bundle child row -> EAN Qty = Child SL from the
  //     bundle mapping (Stock Qty is left untouched: the original MIX
  //     source row's Stock Qty, normally 1 carton).
  //   - Unit/UOM === "PCS"            -> EAN Qty = Stock Qty
  //   - Unit/UOM === CARTON_<N>PCS    -> EAN Qty = Stock Qty x N
  //     (case/space-insensitive: "carton_40pcs", "CARTON_40 PCS",
  //     "CARTON_40PCS " all resolve to N = 40; a bare "CARTON_PCS" with
  //     no parseable N yields CARTON_PCS_SIZE_NOT_DETECTED, never a
  //     guessed size)
  //   - any other/unknown UOM         -> keep the source file's own EAN
  //     Qty value if it looks numeric/trustworthy, else leave blank and
  //     flag EAN_QTY_RULE_NOT_DETECTED
  //   - Stock Qty blank/invalid on a PCS or CARTON row -> EAN Qty blank,
  //     flagged EAN_QTY_STOCK_INVALID (never silently coerced to 0)
  // Returns { eanQty: number|null, diagnostic: string|null }.
  // ----------------------------------------------------------------
  function deriveEanQty(input) {
    input = input || {};
    const isExpandedBundle = !!input.isExpandedBundle;
    const stockQtyNum = toNumberOrNull(input.stockQty);
    const childQtyNum = toNumberOrNull(input.childQty);
    const sourceEanQtyNum = toNumberOrNull(input.sourceEanQty);

    if (isExpandedBundle) {
      // Rule 3 — MIX bundle child: EAN Qty = Child SL only. Never
      // multiplied by Stock Qty/Available Qty; Stock Qty itself is not
      // touched here at all (caller keeps the original source value).
      return { eanQty: childQtyNum, diagnostic: null };
    }

    const normUnit = String(input.unit === null || input.unit === undefined ? "" : input.unit).trim().toUpperCase();

    if (normUnit === "PCS") {
      // Rule 1
      if (stockQtyNum === null) return { eanQty: null, diagnostic: "EAN_QTY_STOCK_INVALID" };
      return { eanQty: stockQtyNum, diagnostic: null };
    }

    if (normUnit.indexOf("CARTON") === 0) {
      // Rule 2 — normalize away all whitespace so "CARTON_40 PCS",
      // "CARTON_40PCS ", "carton_40pcs" all match the same pattern.
      const compact = normUnit.replace(/\s+/g, "");
      const m = compact.match(/^CARTON_(\d+)PCS$/);
      if (!m) return { eanQty: null, diagnostic: "CARTON_PCS_SIZE_NOT_DETECTED" };
      if (stockQtyNum === null) return { eanQty: null, diagnostic: "EAN_QTY_STOCK_INVALID" };
      return { eanQty: stockQtyNum * Number(m[1]), diagnostic: null };
    }

    // Unknown/unsupported UOM: fall back to the source file's own EAN Qty
    // value only if it is present and numeric; otherwise leave blank.
    if (sourceEanQtyNum !== null) return { eanQty: sourceEanQtyNum, diagnostic: null };
    return { eanQty: null, diagnostic: "EAN_QTY_RULE_NOT_DETECTED" };
  }

  function escapeHtml(s) {
    return String(s === null || s === undefined ? "" : s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  // ----------------------------------------------------------------
  // 2. Column definitions
  // ----------------------------------------------------------------
  const BUNDLE_COLUMN_DEFS = [
    { key: "bundleSku", required: true, aliases: ["Bundle SKU (Barcode)", "Bundle SKU", "Parent SKU", "Mã Bundle", "SKU Bundle", "Bundle Barcode"] },
    { key: "bundleName", required: false, aliases: ["Tên Bundle", "Bundle Name"] },
    { key: "childSku", required: true, aliases: ["Child SKU (Barcode)", "Child EAN", "Child Barcode", "EAN con", "Mã EAN con", "Child SKU"] },
    { key: "childProduct", required: false, aliases: ["Tên sản phẩm con", "Child Product", "Child Product Name", "Product Name"] },
    { key: "unitCode", required: false, aliases: ["Mã ĐVT", "Unit Code"] },
    { key: "unitName", required: false, aliases: ["Tên ĐVT", "Unit Name", "Unit", "ĐVT", "Đơn vị tính"] },
    { key: "qty", required: false, aliases: ["SL", "EAN Qty", "Số lượng", "Qty/Carton", "PCS per carton", "Quantity"] },
    { key: "availabilityRatio", required: false, aliases: ["Tỷ Lệ Khả Dụng (%)", "Availability Ratio", "Ty Le Kha Dung"] },
  ];

  // "invProductCode" / "partnerSku" are the two candidate columns a Bundle
  // SKU can be found under in the inventory export — auto-detected per file.
  const INVENTORY_COLUMN_DEFS = [
    { key: "warehouse", required: false, aliases: ["Warehouse", "Kho"] },
    { key: "invProductCode", required: true, aliases: ["Mã sản phẩm", "Child EAN", "Product Code", "Barcode", "Mã vạch"] },
    { key: "partnerSku", required: false, aliases: ["Partner SKU", "Mã SKU đối tác"] },
    { key: "product", required: false, aliases: ["Product", "Tên sản phẩm"] },
    { key: "category", required: false, aliases: ["Category", "Danh mục"] },
    { key: "unit", required: false, aliases: ["Unit", "ĐVT", "Đơn vị tính"] },
    { key: "conditionType", required: false, aliases: ["Condition type", "Condition Type", "Tình trạng hàng hoá", "Tình trạng"] },
    { key: "stockQty", required: false, aliases: ["Stock Qty", "Tồn kho"] },
    { key: "freezeQty", required: false, aliases: ["Freeze Qty", "Phong tỏa", "Đóng băng"] },
    { key: "availableQty", required: true, aliases: ["Available Qty", "Khả dụng"] },
    { key: "pendingIn", required: false, aliases: ["Pending In", "Chờ nhập"] },
    { key: "pendingOut", required: false, aliases: ["Pending Out", "Chờ xuất"] },
    { key: "weightKg", required: false, aliases: ["Weight (Kg)", "Weight", "Khối lượng (Kg)", "Khối lượng"] },
    { key: "cbm", required: false, aliases: ["CBM"] },
    { key: "lastStockoutDate", required: false, aliases: ["Last Stockout Date", "Ngày hết tồn gần nhất", "Ngày hết hàng gần nhất"] },
    { key: "lastOutboundDate", required: false, aliases: ["Last Outbound Date", "Ngày xuất kho gần nhất", "Ngày xuất gần nhất"] },
    { key: "lastInboundDate", required: false, aliases: ["Last Inbound Date", "Ngày nhập kho gần nhất", "Ngày nhập gần nhất"] },
    { key: "invEanQty", required: false, aliases: ["EAN Qty", "Số lượng EAN"] },
  ];

  function detectColumns(headerRow, defs) {
    const normalizedHeaders = (headerRow || []).map(normalizeHeaderText);
    const used = new Set();
    const mapping = {};
    const detectedHeader = {};
    const unresolvedRequired = [];

    defs.forEach(function (def) {
      const normAliases = def.aliases.map(normalizeHeaderText);
      let bestIdx = -1;
      for (let i = 0; i < normalizedHeaders.length; i++) {
        if (used.has(i) || !normalizedHeaders[i]) continue;
        if (normAliases.indexOf(normalizedHeaders[i]) !== -1) { bestIdx = i; break; }
      }
      if (bestIdx === -1) {
        for (let i = 0; i < normalizedHeaders.length; i++) {
          if (used.has(i) || !normalizedHeaders[i]) continue;
          const hit = normAliases.some(function (a) {
            return a && (normalizedHeaders[i].indexOf(a) !== -1 || a.indexOf(normalizedHeaders[i]) !== -1);
          });
          if (hit) { bestIdx = i; break; }
        }
      }
      if (bestIdx !== -1) {
        used.add(bestIdx);
        mapping[def.key] = bestIdx;
        detectedHeader[def.key] = headerRow[bestIdx];
      } else if (def.required) {
        unresolvedRequired.push(def.key);
      }
    });

    const seen = new Map();
    const duplicateHeaders = [];
    (headerRow || []).forEach(function (h) {
      const norm = normalizeHeaderText(h);
      if (!norm) return;
      seen.set(norm, (seen.get(norm) || 0) + 1);
    });
    seen.forEach(function (count, norm) { if (count > 1) duplicateHeaders.push(norm); });

    return { mapping: mapping, detectedHeader: detectedHeader, unresolvedRequired: unresolvedRequired, duplicateHeaders: duplicateHeaders };
  }

  // ----------------------------------------------------------------
  // 3. Workbook / sheet reading
  // ----------------------------------------------------------------
  function fixSheetRange(ws) {
    let maxR = -1, maxC = -1, minR = Infinity, minC = Infinity;
    Object.keys(ws).forEach(function (key) {
      if (key.charAt(0) === "!") return;
      const cell = XLSX.utils.decode_cell(key);
      if (cell.r > maxR) maxR = cell.r;
      if (cell.c > maxC) maxC = cell.c;
      if (cell.r < minR) minR = cell.r;
      if (cell.c < minC) minC = cell.c;
    });
    if (maxR < 0) return;
    const declared = ws["!ref"] ? XLSX.utils.decode_range(ws["!ref"]) : null;
    if (!declared || maxR > declared.e.r || maxC > declared.e.c) {
      ws["!ref"] = XLSX.utils.encode_range({
        s: { r: Math.min(minR, declared ? declared.s.r : minR), c: Math.min(minC, declared ? declared.s.c : minC) },
        e: { r: maxR, c: maxC },
      });
    }
  }

  function pickFirstSheetRows(workbook, preferredName) {
    if (!workbook.SheetNames || !workbook.SheetNames.length) throw new Error("File Excel rỗng — không có sheet nào.");
    const name = workbook.SheetNames.find(function (n) { return n.toLowerCase() === (preferredName || "").toLowerCase(); }) || workbook.SheetNames[0];
    const ws = workbook.Sheets[name];
    if (!ws) throw new Error('Sheet "' + name + '" not found');
    fixSheetRange(ws);
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    return { sheetName: name, rows: rows };
  }

  // ----------------------------------------------------------------
  // 4. Parse Bundle Child Product File -> grouped bundles
  //    (1 Bundle SKU = 1 physical carton; its rows are the components
  //    packed inside it — used only as a mapping/expansion source).
  // ----------------------------------------------------------------
  function parseBundleFile(rows, colMapOverride) {
    if (!rows || !rows.length) throw new Error("File Bundle Child Product rỗng — không đọc được dòng nào.");
    const headerRow = rows[0] || [];
    const det = detectColumns(headerRow, BUNDLE_COLUMN_DEFS);
    const mapping = Object.assign({}, det.mapping, colMapOverride || {});
    const missing = BUNDLE_COLUMN_DEFS.filter(function (d) { return d.required && (mapping[d.key] === undefined || mapping[d.key] === null || mapping[d.key] === "" || mapping[d.key] === -1); }).map(function (d) { return d.key; });
    if (missing.length) {
      throw new Error("Bundle Child Product File thiếu cột bắt buộc: " + missing.join(", ") + ". Vui lòng kiểm tra lại file hoặc chọn cột thủ công ở mục 'File Mapping'.");
    }

    const childRows = [];
    let missingBundleSkuRows = 0;
    let invalidChildSkuRows = 0;
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row) continue;
      const bundleSkuDisplay = normalizeBundleSkuDisplay(row[mapping.bundleSku]);
      const childSku = mapping.childSku !== undefined ? normalizeEan(row[mapping.childSku]) : "";
      if (!bundleSkuDisplay && !childSku) continue; // fully blank row
      if (!bundleSkuDisplay) { missingBundleSkuRows++; continue; } // can't attribute to any carton
      if (!childSku) { invalidChildSkuRows++; continue; } // no usable Child SKU on this mapping row
      childRows.push({
        rowIndex: i + 1,
        bundleSku: bundleSkuDisplay,
        bundleSkuKey: normalizeBundleSkuKey(bundleSkuDisplay),
        bundleName: mapping.bundleName !== undefined ? cleanStr(row[mapping.bundleName]) : "",
        childSku: childSku,
        childProduct: mapping.childProduct !== undefined ? cleanStr(row[mapping.childProduct]) : "",
        unitCode: mapping.unitCode !== undefined ? cleanStr(row[mapping.unitCode]) : "",
        unitName: mapping.unitName !== undefined ? cleanStr(row[mapping.unitName]) : "",
        qty: mapping.qty !== undefined ? toNumberOrNull(row[mapping.qty]) : null,
        availabilityRatio: mapping.availabilityRatio !== undefined ? row[mapping.availabilityRatio] : null,
      });
    }

    const bundleMap = new Map(); // bundleSkuKey -> { bundleSku, bundleName, children:[], totalChildPcs }
    childRows.forEach(function (r) {
      if (!bundleMap.has(r.bundleSkuKey)) {
        bundleMap.set(r.bundleSkuKey, { bundleSku: r.bundleSku, bundleName: r.bundleName, children: [], totalChildPcs: 0 });
      }
      const b = bundleMap.get(r.bundleSkuKey);
      if (!b.bundleName && r.bundleName) b.bundleName = r.bundleName;
      b.children.push(r);
      if (typeof r.qty === "number") b.totalChildPcs += r.qty;
    });

    const bundles = Array.from(bundleMap.values()).sort(function (a, b) { return a.bundleSku < b.bundleSku ? -1 : (a.bundleSku > b.bundleSku ? 1 : 0); });

    return {
      headerRow: headerRow, mapping: mapping, detected: det, childRows: childRows, bundles: bundles, bundleMap: bundleMap,
      missingBundleSkuRows: missingBundleSkuRows, invalidChildSkuRows: invalidChildSkuRows,
    };
  }

  // ----------------------------------------------------------------
  // 5. Parse Real-Time Detailed Inventory File
  //    (the source of truth — every output row starts here).
  // ----------------------------------------------------------------
  function parseInventoryFile(rows, colMapOverride) {
    if (!rows || !rows.length) throw new Error("File Real-Time Detailed Inventory rỗng — không đọc được dòng nào.");
    const headerRow = rows[0] || [];
    const det = detectColumns(headerRow, INVENTORY_COLUMN_DEFS);
    const mapping = Object.assign({}, det.mapping, colMapOverride || {});
    const missing = INVENTORY_COLUMN_DEFS.filter(function (d) { return d.required && (mapping[d.key] === undefined || mapping[d.key] === null || mapping[d.key] === "" || mapping[d.key] === -1); }).map(function (d) { return d.key; });
    if (missing.length) {
      throw new Error("Real-Time Detailed Inventory File thiếu cột bắt buộc: " + missing.join(", ") + ". Vui lòng kiểm tra lại file hoặc chọn cột thủ công ở mục 'File Mapping'.");
    }

    const get = function (row, key) { return mapping[key] !== undefined ? row[mapping[key]] : null; };
    const records = [];
    const warehouses = new Set();
    let invalidSourceIdentifierRows = 0;
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row) continue;
      const invProductCode = cleanStr(get(row, "invProductCode"));
      const partnerSku = cleanStr(get(row, "partnerSku"));
      if (!invProductCode && !partnerSku) { invalidSourceIdentifierRows++; continue; } // no usable identifier on this row
      const warehouse = cleanStr(get(row, "warehouse")) || "(N/A)";
      warehouses.add(warehouse);
      records.push({
        rowIndex: i + 1,
        warehouse: warehouse,
        invProductCode: invProductCode,
        partnerSku: partnerSku,
        product: cleanStr(get(row, "product")),
        category: cleanStr(get(row, "category")),
        unit: cleanStr(get(row, "unit")),
        conditionType: cleanStr(get(row, "conditionType")),
        stockQty: toNumberOrNull(get(row, "stockQty")),
        freezeQty: toNumberOrNull(get(row, "freezeQty")),
        availableQtyRaw: get(row, "availableQty"),
        availableQty: toNumberOrNull(get(row, "availableQty")),
        pendingIn: toNumberOrNull(get(row, "pendingIn")),
        pendingOut: toNumberOrNull(get(row, "pendingOut")),
        weightKg: toNumberOrNull(get(row, "weightKg")),
        cbm: toNumberOrNull(get(row, "cbm")),
        lastStockoutDate: get(row, "lastStockoutDate"),
        lastOutboundDate: get(row, "lastOutboundDate"),
        lastInboundDate: get(row, "lastInboundDate"),
        invEanQty: toNumberOrNull(get(row, "invEanQty")),
      });
    }
    return {
      headerRow: headerRow, mapping: mapping, detected: det, records: records, warehouses: Array.from(warehouses).sort(),
      invalidSourceIdentifierRows: invalidSourceIdentifierRows,
    };
  }

  // ----------------------------------------------------------------
  // 6. Bundle SKU inventory-key detection — never hardcoded. Compares
  //    normalized overlap between the bundle map and each candidate
  //    inventory column ("Mã sản phẩm" / "Mã SKU đối tác"); manual
  //    override is supported from the UI.
  // ----------------------------------------------------------------
  function detectBundleSkuKeySource(bundleMap, invRecords, forcedField) {
    const candidates = ["invProductCode", "partnerSku"];
    function overlapFor(field) {
      let n = 0;
      const seen = new Set();
      invRecords.forEach(function (r) {
        const k = normalizeBundleSkuKey(r[field]);
        if (!k || seen.has(k)) return;
        if (bundleMap.has(k)) { n++; seen.add(k); }
      });
      return n;
    }
    const tried = candidates.map(function (f) { return { field: f, overlap: overlapFor(f) }; });
    if (forcedField) {
      const forced = tried.find(function (t) { return t.field === forcedField; }) || { field: forcedField, overlap: overlapFor(forcedField) };
      return { field: forced.field, overlap: forced.overlap, tried: tried };
    }
    let best = tried[0];
    tried.forEach(function (t) { if (t.overlap > best.overlap) best = t; });
    return { field: best.field, overlap: best.overlap, tried: tried };
  }

  // ----------------------------------------------------------------
  // 7. Core transform — single pass over the real-time inventory rows
  //    (the source of truth). Produces the ONE combined output array.
  // ----------------------------------------------------------------
  const DIAGNOSTIC_LABELS = {
    AVAILABLE_QTY_ZERO: "Available Qty = 0 — loại khỏi báo cáo",
    AVAILABLE_QTY_NEGATIVE: "Available Qty < 0 — loại khỏi báo cáo",
    AVAILABLE_QTY_INVALID: "Available Qty blank/invalid — loại khỏi báo cáo",
    WAREHOUSE_FILTERED_OUT: "Bị loại bởi bộ lọc Warehouse",
    BUNDLE_MAPPING_NOT_FOUND: "Dòng MIX/bundle đang tồn nhưng không có trong file mapping — không xác định được Child SKU",
    INVALID_SOURCE_IDENTIFIER: "Dòng tồn kho không có mã định danh hợp lệ (Mã sản phẩm / Mã SKU đối tác đều trống)",
    INVALID_CHILD_SKU_IN_MAPPING: "Dòng trong file bundle mapping thiếu Child SKU hợp lệ",
    EAN_QTY_RULE_NOT_DETECTED: "Không xác định được quy tắc tính EAN Qty cho UOM này — giữ nguyên dòng, để trống EAN Qty",
    CARTON_PCS_SIZE_NOT_DETECTED: "UOM dạng CARTON nhưng không đọc được số PCS/carton — giữ nguyên dòng, để trống EAN Qty",
    EAN_QTY_STOCK_INVALID: "Stock Qty trống/không hợp lệ nên không tính được EAN Qty — giữ nguyên dòng, để trống EAN Qty",
  };

  function generateCombinedReport(opts) {
    const invRecords = opts.invRecords || [];
    const bundleMap = opts.bundleMap; // Map bundleSkuKey -> { bundleSku, bundleName, children, totalChildPcs }
    const bundleKeyField = opts.bundleKeyField; // "invProductCode" | "partnerSku"
    const warehouseFilter = opts.warehouseFilter; // Set|null

    const combined = [];
    const diagnostics = [];

    let removedUnavailableRows = 0;
    let warehouseFilteredRows = 0;
    let availableNormalRows = 0;
    let availableBundlesExpanded = 0;
    let bundleChildRowsGenerated = 0;
    let unmappedBundleRowsExcluded = 0;

    invRecords.forEach(function (ir) {
      // Available Qty is a SOURCE-ROW FILTER ONLY — it is never displayed
      // or used in any calculation on the final output rows.
      const avail = ir.availableQty;
      let qtyReason = null;
      if (avail === null) qtyReason = "AVAILABLE_QTY_INVALID";
      else if (avail < 0) qtyReason = "AVAILABLE_QTY_NEGATIVE";
      else if (avail === 0) qtyReason = "AVAILABLE_QTY_ZERO";
      if (qtyReason) {
        removedUnavailableRows++;
        diagnostics.push({ warehouse: ir.warehouse, code: ir.invProductCode || ir.partnerSku, product: ir.product, reason: qtyReason });
        return;
      }

      if (warehouseFilter && warehouseFilter.size && !warehouseFilter.has(ir.warehouse)) {
        warehouseFilteredRows++;
        diagnostics.push({ warehouse: ir.warehouse, code: ir.invProductCode || ir.partnerSku, product: ir.product, reason: "WAREHOUSE_FILTERED_OUT" });
        return;
      }

      const keyRaw = bundleKeyField === "partnerSku" ? ir.partnerSku : ir.invProductCode;
      const key = normalizeBundleSkuKey(keyRaw);
      const bundle = key ? bundleMap.get(key) : null;

      if (bundle) {
        // CASE 2 — available bundle carton with a valid mapping: drop the
        // 1-line summary row, expand into 1 row per Child SKU. EAN Qty
        // comes straight from the mapping's SL — never Available Qty,
        // never a multiplication (1 Bundle SKU already = 1 available carton).
        availableBundlesExpanded++;
        bundle.children.forEach(function (child, childIdx) {
          bundleChildRowsGenerated++;
          const derivedChild = deriveEanQty({ unit: ir.unit, stockQty: ir.stockQty, isExpandedBundle: true, childQty: child.qty, sourceEanQty: null });
          if (derivedChild.diagnostic) {
            diagnostics.push({ warehouse: ir.warehouse, code: child.childSku, product: child.childProduct, reason: derivedChild.diagnostic });
          }
          combined.push({
            warehouse: ir.warehouse, barcode: bundle.bundleSku, partnerSku: ir.partnerSku, childEan: child.childSku,
            product: child.childProduct, category: ir.category, unit: ir.unit, conditionType: ir.conditionType,
            stockQty: ir.stockQty, freezeQty: ir.freezeQty, pendingIn: ir.pendingIn, pendingOut: ir.pendingOut,
            weightKg: ir.weightKg, cbm: ir.cbm,
            lastStockoutDate: ir.lastStockoutDate, lastOutboundDate: ir.lastOutboundDate, lastInboundDate: ir.lastInboundDate,
            eanQty: derivedChild.eanQty, _rowType: "bundleChild", _bundleSku: bundle.bundleSku, _childOrder: childIdx,
          });
        });
        return;
      }

      if (isMixUnit(ir.unit)) {
        // CASE 4 — currently available bundle/MIX row with NO mapping:
        // composition can't be determined, so it is excluded entirely
        // (not kept as an ordinary item) and logged as a diagnostic.
        unmappedBundleRowsExcluded++;
        diagnostics.push({ warehouse: ir.warehouse, code: ir.invProductCode || ir.partnerSku, product: ir.product, reason: "BUNDLE_MAPPING_NOT_FOUND" });
        return;
      }

      // CASE 1 — ordinary available inventory row: keep as-is. EAN Qty
      // is derived from Stock Qty per the Unit/UOM rule (PCS / CARTON_NPCS
      // / unknown), never left hardcoded and never computed a second time
      // anywhere else (UI preview and Excel export both read this exact
      // field).
      availableNormalRows++;
      const derivedNormal = deriveEanQty({ unit: ir.unit, stockQty: ir.stockQty, isExpandedBundle: false, childQty: null, sourceEanQty: ir.invEanQty });
      if (derivedNormal.diagnostic) {
        diagnostics.push({ warehouse: ir.warehouse, code: ir.invProductCode || ir.partnerSku, product: ir.product, reason: derivedNormal.diagnostic });
      }
      combined.push({
        warehouse: ir.warehouse, barcode: ir.invProductCode, partnerSku: ir.partnerSku, childEan: ir.invProductCode,
        product: ir.product, category: ir.category, unit: ir.unit, conditionType: ir.conditionType,
        stockQty: ir.stockQty, freezeQty: ir.freezeQty, pendingIn: ir.pendingIn, pendingOut: ir.pendingOut,
        weightKg: ir.weightKg, cbm: ir.cbm,
        lastStockoutDate: ir.lastStockoutDate, lastOutboundDate: ir.lastOutboundDate, lastInboundDate: ir.lastInboundDate,
        eanQty: derivedNormal.eanQty, _rowType: "normal", _childOrder: 0,
      });
    });

    // Default output order (applied automatically, before any interactive
    // on-screen sort click): Warehouse -> Product Code. "Product Code" is
    // the Barcode field, which for a normal row IS its own code, and for a
    // bundle-expanded row is the shared Bundle SKU — so grouping by
    // (warehouse, barcode) automatically clusters every bundle's child
    // rows together, and _childOrder (their original position in the
    // bundle mapping file) keeps them in that original order within the
    // group. Array.prototype.sort is stable (ES2019+), so this is
    // deterministic.
    combined.sort(function (a, b) {
      const w = compareCodes(a.warehouse, b.warehouse);
      if (w) return w;
      const c = compareCodes(a.barcode, b.barcode);
      if (c) return c;
      return (a._childOrder || 0) - (b._childOrder || 0);
    });

    return {
      combined: combined,
      diagnostics: diagnostics,
      summary: {
        sourceRealTimeInventoryRows: invRecords.length,
        availableNormalRows: availableNormalRows,
        availableBundlesExpanded: availableBundlesExpanded,
        bundleChildRowsGenerated: bundleChildRowsGenerated,
        finalCombinedRows: combined.length,
        removedUnavailableRows: removedUnavailableRows,
        unmappedBundleRowsExcluded: unmappedBundleRowsExcluded,
        warehouseFilteredRows: warehouseFilteredRows,
      },
    };
  }

  // ----------------------------------------------------------------
  // 8. Excel export — exactly ONE worksheet, identifiers forced to text.
  // ----------------------------------------------------------------
  function asText(v) { return v === null || v === undefined ? "" : String(v); }

  const EXPORT_HEADER = [
    "#", "Warehouse", "Barcode", "Partner SKU", "Child EAN", "Product", "Category", "Unit", "Condition Type",
    "Stock Qty", "Freeze Qty", "Pending In", "Pending Out", "Weight (Kg)", "CBM",
    "Last Stockout Date", "Last Outbound Date", "Last Inbound Date", "EAN Qty",
  ];
  const EXPORT_COL_WIDTHS = [5, 20, 18, 16, 16, 34, 14, 10, 12, 10, 10, 10, 11, 12, 10, 16, 16, 16, 10];

  function buildExportWorkbook(combinedRows) {
    const aoa = [EXPORT_HEADER].concat(combinedRows.map(function (r, i) {
      return [
        i + 1, asText(r.warehouse), asText(r.barcode), asText(r.partnerSku), asText(r.childEan), asText(r.product),
        asText(r.category), asText(r.unit), asText(r.conditionType),
        r.stockQty, r.freezeQty, r.pendingIn, r.pendingOut, r.weightKg, r.cbm,
        asText(r.lastStockoutDate), asText(r.lastOutboundDate), asText(r.lastInboundDate), r.eanQty,
      ];
    }));
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = EXPORT_COL_WIDTHS.map(function (w) { return { wch: w }; });
    ws["!autofilter"] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(aoa.length - 1, 0), c: EXPORT_HEADER.length - 1 } }) };
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Combined Real-Time Inventory");
    const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    return new Blob([out], { type: "application/octet-stream" });
  }

  function exportFilename() {
    const d = new Date();
    const pad = function (n) { return String(n).padStart(2, "0"); };
    return "Combined_Real_Time_Inventory_" + d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) +
      "_" + pad(d.getHours()) + pad(d.getMinutes()) + ".xlsx";
  }

  // ----------------------------------------------------------------
  // 9. UI wiring
  // ----------------------------------------------------------------
  const PREVIEW_COLS = [
    ["warehouse", "Warehouse"], ["barcode", "Barcode"], ["partnerSku", "Partner SKU"], ["childEan", "Child EAN"],
    ["product", "Product"], ["category", "Category"], ["unit", "Unit"], ["conditionType", "Condition Type"],
    ["stockQty", "Stock Qty"], ["freezeQty", "Freeze Qty"], ["pendingIn", "Pending In"], ["pendingOut", "Pending Out"],
    ["weightKg", "Weight (Kg)"], ["cbm", "CBM"],
    ["lastStockoutDate", "Last Stockout Date"], ["lastOutboundDate", "Last Outbound Date"], ["lastInboundDate", "Last Inbound Date"],
    ["eanQty", "EAN Qty"],
  ];
  const SEARCH_FIELDS = ["barcode", "partnerSku", "childEan", "product"]; // per spec: search these 4 fields simultaneously (Warehouse has its own dedicated filter)
  const PAGE_SIZE = 100;

  function init() {
    const bundleInput = document.getElementById("rtoaBundleInput");
    if (!bundleInput) return; // markup not present on this page

    const invInput = document.getElementById("rtoaInventoryInput");
    const bundleStatus = document.getElementById("rtoaBundleStatus");
    const invStatus = document.getElementById("rtoaInventoryStatus");
    const mappingBox = document.getElementById("rtoaMappingBody");
    const warehouseFilterEl = document.getElementById("rtoaWarehouseFilter");
    const runBtn = document.getElementById("rtoaRunBtn");
    const resetBtn = document.getElementById("rtoaResetBtn");
    const logEl = document.getElementById("rtoaLog");
    const resultsWrap = document.getElementById("rtoaResultsWrap");
    const summaryCardsEl = document.getElementById("rtoaSummaryCards");
    const tableWrap = document.getElementById("rtoaAvailableTableWrap");
    const searchInput = document.getElementById("rtoaTableSearch");
    const pagerInfo = document.getElementById("rtoaPagerInfo");
    const pagerPrev = document.getElementById("rtoaPagerPrev");
    const pagerNext = document.getElementById("rtoaPagerNext");
    const diagnosticsWrap = document.getElementById("rtoaExcludedTableWrap");
    const diagnosticsCountEl = document.getElementById("rtoaExcludedCount");
    const exportBtn = document.getElementById("rtoaExportBtn");

    let bundleParsed = null;
    let invParsed = null;
    let bundleColOverride = {};
    let invColOverride = {};
    let bundleSkuKeyFieldOverride = null;
    let lastCombined = []; // the ONE final array — export always uses this, unfiltered
    let sortState = { col: null, dir: 1 };
    let searchQuery = "";
    let currentPage = 0;

    function log(msg) { logEl.textContent += msg + "\n"; logEl.scrollTop = logEl.scrollHeight; }
    function tick() { return new Promise(function (r) { setTimeout(r, 0); }); }
    function refreshRunEnabled() { runBtn.disabled = !(bundleParsed && invParsed); }

    // ---- File Mapping panel ----
    function renderMappingRow(fieldLabel, key, detected, headerRow, currentOverride, onChange) {
      const row = document.createElement("div");
      row.className = "rtoa-map-row";
      const label = document.createElement("div");
      label.className = "rtoa-map-label";
      label.textContent = fieldLabel;
      const sel = document.createElement("select");
      const noneOpt = document.createElement("option");
      noneOpt.value = "";
      noneOpt.textContent = "-- không dùng --";
      sel.appendChild(noneOpt);
      headerRow.forEach(function (h, i) {
        const opt = document.createElement("option");
        opt.value = String(i);
        opt.textContent = "[" + i + "] " + (h === null || h === undefined || h === "" ? "(trống)" : String(h));
        sel.appendChild(opt);
      });
      const chosenIdx = currentOverride[key] !== undefined ? currentOverride[key] : detected.mapping[key];
      sel.value = chosenIdx !== undefined && chosenIdx !== null ? String(chosenIdx) : "";
      const badge = document.createElement("span");
      badge.className = "rtoa-map-badge " + (detected.mapping[key] !== undefined ? "ok" : "warn");
      badge.textContent = detected.mapping[key] !== undefined ? "tự nhận diện" : "chưa nhận diện";
      sel.addEventListener("change", function () {
        onChange(key, sel.value === "" ? undefined : parseInt(sel.value, 10));
      });
      row.appendChild(label); row.appendChild(sel); row.appendChild(badge);
      return row;
    }

    function renderBundleKeySourceRow() {
      if (!bundleParsed || !invParsed) return null;
      const detection = detectBundleSkuKeySource(bundleParsed.bundleMap, invParsed.records, bundleSkuKeyFieldOverride);

      const wrap = document.createElement("div");
      wrap.className = "rtoa-map-row";
      const label = document.createElement("div");
      label.className = "rtoa-map-label";
      label.textContent = "Bundle SKU khớp với cột (Inventory)";
      const sel = document.createElement("select");
      ["", "invProductCode", "partnerSku"].forEach(function (f) {
        const opt = document.createElement("option");
        opt.value = f;
        if (f === "") { opt.textContent = "-- tự động --"; }
        else {
          const headerName = invParsed.mapping[f] !== undefined ? invParsed.headerRow[invParsed.mapping[f]] : "(chưa nhận diện)";
          const tried = detection.tried.find(function (t) { return t.field === f; });
          opt.textContent = (f === "invProductCode" ? "Mã sản phẩm" : "Mã SKU đối tác") + " = '" + headerName + "' (khớp " + (tried ? tried.overlap : 0) + " Bundle SKU)";
        }
        sel.appendChild(opt);
      });
      sel.value = bundleSkuKeyFieldOverride || "";
      sel.addEventListener("change", function () {
        bundleSkuKeyFieldOverride = sel.value || null;
        renderMapping();
      });
      const badge = document.createElement("span");
      if (detection.overlap === 0) { badge.className = "rtoa-map-badge warn"; badge.textContent = "0 khớp!"; }
      else { badge.className = "rtoa-map-badge ok"; badge.textContent = detection.overlap + " Bundle SKU khớp"; }
      wrap.appendChild(label); wrap.appendChild(sel); wrap.appendChild(badge);

      if (detection.overlap === 0) {
        const warn = document.createElement("div");
        warn.className = "rtoa-warning-banner";
        warn.textContent = "No Bundle SKU from the bundle file was found in the selected real-time inventory key column. Please verify the inventory file or manually select the correct Bundle SKU column. (Không tìm thấy Bundle SKU nào trong cột đã chọn — vui lòng kiểm tra lại file tồn kho hoặc chọn cột khác ở trên.)";
        wrap.appendChild(warn);
      }
      return wrap;
    }

    function renderMapping() {
      mappingBox.innerHTML = "";
      if (bundleParsed) {
        const h = document.createElement("h5");
        h.textContent = "Bundle Child Product File";
        mappingBox.appendChild(h);
        BUNDLE_COLUMN_DEFS.forEach(function (def) {
          mappingBox.appendChild(renderMappingRow(def.key + (def.required ? " *" : ""), def.key, bundleParsed.detected, bundleParsed.headerRow, bundleColOverride, function (key, idx) {
            if (idx === undefined) delete bundleColOverride[key]; else bundleColOverride[key] = idx;
            reparseBundle();
          }));
        });
      }
      if (invParsed) {
        const h2 = document.createElement("h5");
        h2.textContent = "Real-Time Detailed Inventory File";
        mappingBox.appendChild(h2);
        INVENTORY_COLUMN_DEFS.forEach(function (def) {
          mappingBox.appendChild(renderMappingRow(def.key + (def.required ? " *" : ""), def.key, invParsed.detected, invParsed.headerRow, invColOverride, function (key, idx) {
            if (idx === undefined) delete invColOverride[key]; else invColOverride[key] = idx;
            reparseInventory();
          }));
        });
      }
      if (bundleParsed && invParsed) {
        const h3 = document.createElement("h5");
        h3.textContent = "Bundle SKU ⇄ Inventory matching key";
        mappingBox.appendChild(h3);
        const row = renderBundleKeySourceRow();
        if (row) mappingBox.appendChild(row);
      }
    }

    function renderWarehouseFilter() {
      warehouseFilterEl.innerHTML = "";
      if (!invParsed || invParsed.warehouses.length <= 1) { warehouseFilterEl.style.display = "none"; return; }
      warehouseFilterEl.style.display = "";
      const title = document.createElement("div");
      title.className = "small";
      title.textContent = "Lọc theo Warehouse (bỏ chọn = loại khỏi kết quả, áp dụng cho cả dòng thường và dòng bundle bung ra):";
      warehouseFilterEl.appendChild(title);
      invParsed.warehouses.forEach(function (wh) {
        const row = document.createElement("label");
        row.className = "rtoa-wh-chip";
        const chk = document.createElement("input");
        chk.type = "checkbox";
        chk.checked = true;
        chk.setAttribute("data-wh", wh);
        const txt = document.createElement("span");
        txt.textContent = wh;
        row.appendChild(chk); row.appendChild(txt);
        warehouseFilterEl.appendChild(row);
      });
    }

    function getWarehouseFilter() {
      if (!invParsed || invParsed.warehouses.length <= 1) return null;
      const boxes = Array.from(warehouseFilterEl.querySelectorAll("input[type=checkbox]"));
      if (!boxes.length) return null;
      const checked = boxes.filter(function (b) { return b.checked; }).map(function (b) { return b.getAttribute("data-wh"); });
      if (checked.length === boxes.length) return null;
      return new Set(checked);
    }

    // ---- File loading ----
    function reparseBundle() {
      if (!bundleInput.files.length) return;
      try {
        const raw = bundleInput._rows;
        bundleParsed = parseBundleFile(raw, bundleColOverride);
        let msg = "Đã đọc " + bundleParsed.childRows.length + " dòng child SKU — " + bundleParsed.bundles.length + " Bundle SKU (carton) distinct.";
        if (bundleParsed.missingBundleSkuRows) msg += " (" + bundleParsed.missingBundleSkuRows + " dòng thiếu Bundle SKU đã bị bỏ qua.)";
        if (bundleParsed.invalidChildSkuRows) msg += " (" + bundleParsed.invalidChildSkuRows + " dòng thiếu Child SKU đã bị bỏ qua.)";
        bundleStatus.textContent = msg;
        renderMapping();
        refreshRunEnabled();
      } catch (e) {
        bundleParsed = null;
        bundleStatus.textContent = "❌ " + e.message;
        refreshRunEnabled();
      }
    }

    function reparseInventory() {
      if (!invInput.files.length) return;
      try {
        const raw = invInput._rows;
        invParsed = parseInventoryFile(raw, invColOverride);
        invStatus.textContent = "Đã đọc " + invParsed.records.length + " dòng tồn kho — " + invParsed.warehouses.length + " warehouse.";
        renderWarehouseFilter();
        renderMapping();
        refreshRunEnabled();
      } catch (e) {
        invParsed = null;
        invStatus.textContent = "❌ " + e.message;
        refreshRunEnabled();
      }
    }

    bundleInput.addEventListener("change", async function () {
      const file = bundleInput.files[0];
      if (!file) return;
      bundleStatus.textContent = "Đang đọc file...";
      try {
        const wb = await WOPUtils.readWorkbookFromFile(file);
        const picked = pickFirstSheetRows(wb, "Data");
        bundleInput._rows = picked.rows;
        bundleColOverride = {};
        reparseBundle();
      } catch (e) {
        bundleStatus.textContent = "❌ Lỗi đọc file: " + e.message;
        bundleParsed = null; refreshRunEnabled();
      }
    });

    invInput.addEventListener("change", async function () {
      const file = invInput.files[0];
      if (!file) return;
      invStatus.textContent = "Đang đọc file...";
      try {
        const wb = await WOPUtils.readWorkbookFromFile(file);
        const picked = pickFirstSheetRows(wb, "Sheet1");
        invInput._rows = picked.rows;
        invColOverride = {};
        bundleSkuKeyFieldOverride = null;
        reparseInventory();
      } catch (e) {
        invStatus.textContent = "❌ Lỗi đọc file: " + e.message;
        invParsed = null; refreshRunEnabled();
      }
    });

    // ---- Results rendering ----
    function renderSummaryCards(s) {
      summaryCardsEl.innerHTML = "";
      const cards = [
        ["Source Real-Time Inventory Rows", s.sourceRealTimeInventoryRows],
        ["Available Normal Rows", s.availableNormalRows],
        ["Available Bundles Expanded", s.availableBundlesExpanded],
        ["Bundle Child Rows Generated", s.bundleChildRowsGenerated],
        ["Final Combined Rows", s.finalCombinedRows],
        ["Removed Unavailable Rows", s.removedUnavailableRows],
        ["Unmapped Bundle Rows Excluded", s.unmappedBundleRowsExcluded],
      ];
      cards.forEach(function (c) {
        const div = document.createElement("div");
        div.className = "summary-card";
        div.innerHTML = '<div class="summary-card-value">' + escapeHtml(c[1]) + '</div><div class="summary-card-label">' + escapeHtml(c[0]) + '</div>';
        summaryCardsEl.appendChild(div);
      });
    }

    function getFilteredSortedRows() {
      let rows = lastCombined;
      const q = searchQuery.trim().toLowerCase();
      if (q) {
        rows = rows.filter(function (r) {
          return SEARCH_FIELDS.some(function (f) { return String(r[f] === null || r[f] === undefined ? "" : r[f]).toLowerCase().indexOf(q) !== -1; });
        });
      }
      if (sortState.col) {
        const col = sortState.col, dir = sortState.dir;
        rows = rows.slice().sort(function (a, b) {
          const av = a[col], bv = b[col];
          const an = typeof av === "number", bn = typeof bv === "number";
          if (an && bn) return (av - bv) * dir;
          const as = av === null || av === undefined ? "" : String(av);
          const bs = bv === null || bv === undefined ? "" : String(bv);
          return as.localeCompare(bs) * dir;
        });
      }
      return rows;
    }

    function renderTable() {
      tableWrap.innerHTML = "";
      const rows = getFilteredSortedRows();
      const totalPages = Math.max(1, Math.ceil(rows.length / PAGE_SIZE));
      if (currentPage >= totalPages) currentPage = totalPages - 1;
      if (currentPage < 0) currentPage = 0;
      const startIdx = currentPage * PAGE_SIZE;
      const pageRows = rows.slice(startIdx, startIdx + PAGE_SIZE);

      // Always shows the exact count of rows currently matching the filter
      // (searchQuery) — this is the number that also gets exported to Excel.
      const filterNote = searchQuery.trim() ? (" — tổng " + lastCombined.length + " dòng trước khi lọc") : "";
      pagerInfo.textContent = rows.length === 0
        ? ("0 rows" + filterNote)
        : ("Hiển thị " + (startIdx + 1) + "–" + Math.min(startIdx + PAGE_SIZE, rows.length) + " / " + rows.length + " rows sau lọc" + filterNote);
      pagerPrev.disabled = currentPage <= 0;
      pagerNext.disabled = currentPage >= totalPages - 1;

      if (!rows.length) {
        const p = document.createElement("p");
        p.className = "status";
        p.textContent = lastCombined.length ? "Không có dòng nào khớp với tìm kiếm." : "Không có dòng nào trong báo cáo kết hợp.";
        tableWrap.appendChild(p);
        return;
      }

      const table = document.createElement("table");
      table.className = "data-table";
      const thead = document.createElement("thead");
      const headTr = document.createElement("tr");
      PREVIEW_COLS.forEach(function (c) {
        const th = document.createElement("th");
        th.className = "rtoa-sortable";
        th.textContent = c[1];
        if (sortState.col === c[0]) {
          const arrow = document.createElement("span");
          arrow.className = "rtoa-sort-arrow";
          arrow.textContent = sortState.dir === 1 ? "▲" : "▼";
          th.appendChild(arrow);
        }
        th.addEventListener("click", function () {
          if (sortState.col === c[0]) sortState.dir = -sortState.dir;
          else { sortState.col = c[0]; sortState.dir = 1; }
          currentPage = 0;
          renderTable();
        });
        headTr.appendChild(th);
      });
      thead.appendChild(headTr);
      table.appendChild(thead);

      const tbody = document.createElement("tbody");
      pageRows.forEach(function (r) {
        const tr = document.createElement("tr");
        PREVIEW_COLS.forEach(function (c) {
          const td = document.createElement("td");
          const v = r[c[0]];
          td.textContent = v === null || v === undefined ? "" : String(v);
          tr.appendChild(td);
        });
        tbody.appendChild(tr);
      });
      table.appendChild(tbody);
      tableWrap.appendChild(table);
    }

    searchInput.addEventListener("input", function () { searchQuery = searchInput.value; currentPage = 0; renderTable(); });
    pagerPrev.addEventListener("click", function () { currentPage--; renderTable(); });
    pagerNext.addEventListener("click", function () { currentPage++; renderTable(); });

    function renderDiagnosticsTable(diagnostics) {
      diagnosticsCountEl.textContent = String(diagnostics.length);
      if (!diagnostics.length) { diagnosticsWrap.innerHTML = '<p class="status">Không có cảnh báo xử lý nào.</p>'; return; }
      let html = '<table class="data-table"><thead><tr><th>Warehouse</th><th>Code</th><th>Product</th><th>Reason</th></tr></thead><tbody>';
      diagnostics.forEach(function (r) {
        html += "<tr><td>" + escapeHtml(r.warehouse) + "</td><td>" + escapeHtml(r.code) + "</td><td>" + escapeHtml(r.product) +
          '</td><td><span class="rtoa-reason-badge">' + escapeHtml(DIAGNOSTIC_LABELS[r.reason] || r.reason) + "</span></td></tr>";
      });
      html += "</tbody></table>";
      diagnosticsWrap.innerHTML = html;
    }

    runBtn.addEventListener("click", async function () {
      if (!bundleParsed || !invParsed) return;
      runBtn.disabled = true;
      logEl.textContent = "";
      log("[INFO] Đang tạo Combined Real-Time Inventory Report...");
      await tick();
      try {
        const detection = detectBundleSkuKeySource(bundleParsed.bundleMap, invParsed.records, bundleSkuKeyFieldOverride);
        log("[INFO] Bundle SKU key column: " + detection.field + " (khớp " + detection.overlap + " Bundle SKU trên toàn bộ file tồn kho).");
        if (detection.overlap === 0) {
          log("⚠️ No Bundle SKU from the bundle file was found in the selected real-time inventory key column. Please verify the inventory file or manually select the correct Bundle SKU column.");
        }
        const result = generateCombinedReport({
          invRecords: invParsed.records,
          bundleMap: bundleParsed.bundleMap,
          bundleKeyField: detection.field,
          warehouseFilter: getWarehouseFilter(),
        });
        // surface file-level parse diagnostics alongside row-level ones
        const extraDiagnostics = [];
        for (let i = 0; i < (invParsed.invalidSourceIdentifierRows || 0); i++) {
          extraDiagnostics.push({ warehouse: "", code: "", product: "", reason: "INVALID_SOURCE_IDENTIFIER" });
        }
        for (let i = 0; i < (bundleParsed.invalidChildSkuRows || 0); i++) {
          extraDiagnostics.push({ warehouse: "", code: "", product: "", reason: "INVALID_CHILD_SKU_IN_MAPPING" });
        }
        const allDiagnostics = result.diagnostics.concat(extraDiagnostics);

        lastCombined = result.combined;
        searchQuery = ""; searchInput.value = ""; sortState = { col: null, dir: 1 }; currentPage = 0;

        log("[INFO] Final combined rows: " + result.summary.finalCombinedRows +
          " (normal: " + result.summary.availableNormalRows + ", bundle child rows: " + result.summary.bundleChildRowsGenerated + " from " + result.summary.availableBundlesExpanded + " bundles)");
        renderSummaryCards(result.summary);
        renderTable();
        renderDiagnosticsTable(allDiagnostics);
        resultsWrap.style.display = "";
        exportBtn.disabled = lastCombined.length === 0;
        log("✅ Hoàn tất.");
      } catch (e) {
        log("❌ Lỗi: " + e.message);
      } finally {
        runBtn.disabled = false;
      }
    });

    exportBtn.addEventListener("click", function () {
      if (!lastCombined.length) return;
      const blob = buildExportWorkbook(lastCombined); // always the FULL final array, never the filtered/paginated view
      WOPUtils.downloadBlob(blob, exportFilename());
    });

    resetBtn.addEventListener("click", function () {
      bundleInput.value = ""; invInput.value = "";
      bundleParsed = null; invParsed = null;
      bundleColOverride = {}; invColOverride = {}; bundleSkuKeyFieldOverride = null;
      lastCombined = []; sortState = { col: null, dir: 1 }; searchQuery = ""; currentPage = 0;
      bundleStatus.textContent = "Chưa chọn file."; invStatus.textContent = "Chưa chọn file.";
      mappingBox.innerHTML = ""; warehouseFilterEl.innerHTML = ""; warehouseFilterEl.style.display = "none";
      logEl.textContent = ""; resultsWrap.style.display = "none"; exportBtn.disabled = true;
      searchInput.value = "";
      refreshRunEnabled();
    });
  }

  const api = {
    init: init,
    parseBundleFile: parseBundleFile,
    parseInventoryFile: parseInventoryFile,
    detectBundleSkuKeySource: detectBundleSkuKeySource,
    generateCombinedReport: generateCombinedReport,
    buildExportWorkbook: buildExportWorkbook,
    exportFilename: exportFilename,
    EXPORT_HEADER: EXPORT_HEADER,
    DIAGNOSTIC_LABELS: DIAGNOSTIC_LABELS,
    BUNDLE_COLUMN_DEFS: BUNDLE_COLUMN_DEFS,
    INVENTORY_COLUMN_DEFS: INVENTORY_COLUMN_DEFS,
    _internal: {
      normalizeHeaderText: normalizeHeaderText, normalizeEan: normalizeEan,
      normalizeBundleSkuDisplay: normalizeBundleSkuDisplay, normalizeBundleSkuKey: normalizeBundleSkuKey,
      toNumberOrNull: toNumberOrNull, detectColumns: detectColumns, fixSheetRange: fixSheetRange, isMixUnit: isMixUnit,
      compareCodes: compareCodes,
      deriveEanQty: deriveEanQty,
    },
  };

  if (typeof module !== "undefined" && module.exports) module.exports = api;
  return api;
})();
