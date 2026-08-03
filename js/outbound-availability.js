/* ==========================================================
 * outbound-availability.js — Outbound: Real-Time Outbound Availability
 * (Bundle-Aware Inventory Report)
 *
 * Business model (confirmed against real files):
 *   A Bundle SKU is the unique inventory identity of ONE physical bundle
 *   carton. The "bundle child product" file lists many rows per Bundle SKU,
 *   but those rows are the components packed INSIDE that one carton, not
 *   separate cartons. The bundle carton itself shows up as its OWN line
 *   item in the Real-Time Detailed Inventory export (its code appears in
 *   the "Mã sản phẩm" / "Mã SKU đối tác" column, Unit often "MIX"), with
 *   its own Available Qty ("Khả dụng"). THAT row is what determines whether
 *   the physical carton currently exists in stock — not the availability
 *   of its individual child EANs.
 *
 * Matching key: Bundle SKU -> looked up directly against the Real-Time
 * Inventory file (never against Child SKU/EAN). Only after a Bundle SKU
 * passes the Available Qty > 0 check are its child lines expanded/shown.
 *
 * 100% client-side (SheetJS, already loaded globally as `XLSX` in
 * index.html) — files never leave the browser.
 * ========================================================== */
window.WOPOutboundAvailability = (function () {
  "use strict";

  // ----------------------------------------------------------------
  // 1. Normalization helpers
  // ----------------------------------------------------------------

  // Normalize a header cell (or alias string) for column detection.
  function normalizeHeaderText(h) {
    if (h === null || h === undefined) return "";
    let s = String(h).toLowerCase();
    s = s.replace(/đ/g, "d");
    s = s.normalize("NFD").replace(/[̀-ͯ]/g, "");
    s = s.replace(/[^a-z0-9]+/g, " ").trim().replace(/\s+/g, " ");
    return s;
  }

  // Normalize a numeric-style identifier (Child SKU / EAN / Barcode) into a
  // stable string: strips Excel-generated trailing ".0", never round-trips
  // through Number() for display, preserves leading zeroes from string
  // sources.
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

  // Normalize a Bundle SKU (alphanumeric identifier, e.g. "TP-251223-09"):
  // string-safe, trims spaces + non-breaking spaces, strips a stray
  // Excel-generated trailing ".0", preserves original casing for display.
  function normalizeBundleSkuDisplay(v) {
    if (v === null || v === undefined) return "";
    let s = String(v).replace(/ /g, " ").trim();
    if (s === "") return "";
    s = s.replace(/\.0+$/, "");
    return s;
  }
  // Case-insensitive matching key derived from the display form above.
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
  // SKU can be found under in the inventory export — which one actually
  // holds it is auto-detected per file (see detectBundleSkuKeySource).
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
  ];

  // Detect which column index best matches each definition (exact match
  // first, "contains" fallback second). Each column index can be claimed
  // by at most one field.
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
  // 3. Workbook / sheet reading (real exports sometimes declare a stale
  //    "!ref" dimension smaller than the actual data).
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
  // 4. Parse Bundle Child Product File -> child rows + grouped bundles
  //    (1 Bundle SKU = 1 physical carton; its rows are the components
  //    packed inside it, NOT separate cartons).
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
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row) continue;
      const bundleSkuDisplay = normalizeBundleSkuDisplay(row[mapping.bundleSku]);
      const childSku = mapping.childSku !== undefined ? normalizeEan(row[mapping.childSku]) : "";
      if (!bundleSkuDisplay && !childSku) continue; // fully blank row
      if (!bundleSkuDisplay) { missingBundleSkuRows++; continue; } // can't attribute to any carton
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

    // Group child rows into one record per unique Bundle SKU (case-insensitive
    // key) — this is the "1 Bundle SKU = 1 physical carton" grouping.
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

    return { headerRow: headerRow, mapping: mapping, detected: det, childRows: childRows, bundles: bundles, missingBundleSkuRows: missingBundleSkuRows };
  }

  // ----------------------------------------------------------------
  // 5. Parse Real-Time Detailed Inventory File
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
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row) continue;
      const invProductCode = cleanStr(get(row, "invProductCode"));
      const partnerSku = cleanStr(get(row, "partnerSku"));
      if (!invProductCode && !partnerSku) continue; // no usable key on this row
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
      });
    }
    return { headerRow: headerRow, mapping: mapping, detected: det, records: records, warehouses: Array.from(warehouses).sort() };
  }

  // ----------------------------------------------------------------
  // 6. Bundle SKU inventory-key detection
  //    A Bundle SKU can show up under either "Mã sản phẩm" (invProductCode)
  //    or "Mã SKU đối tác" (partnerSku) depending on the export — pick
  //    whichever column actually contains the most of the selected Bundle
  //    SKUs (never hardcode one column), with a manual override available.
  // ----------------------------------------------------------------
  function buildBundleKeyIndex(invRecords, field) {
    const idx = new Map(); // normalizedKeyCI -> [invRecord...]
    invRecords.forEach(function (r) {
      const key = normalizeBundleSkuKey(r[field]);
      if (!key) return;
      if (!idx.has(key)) idx.set(key, []);
      idx.get(key).push(r);
    });
    return idx;
  }

  function countOverlap(bundleSkuKeys, index) {
    let n = 0;
    bundleSkuKeys.forEach(function (k) { if (index.has(k)) n++; });
    return n;
  }

  // Returns { field, label, index, overlap, tried:[{field,overlap}] }.
  // `field` is "invProductCode" or "partnerSku" (or the override field
  // name if the caller forces one).
  function detectBundleSkuKeySource(bundleSkuKeys, invRecords, forcedField) {
    const candidates = ["invProductCode", "partnerSku"];
    const tried = candidates.map(function (f) {
      const index = buildBundleKeyIndex(invRecords, f);
      return { field: f, index: index, overlap: countOverlap(bundleSkuKeys, index) };
    });
    if (forcedField) {
      const forced = tried.find(function (t) { return t.field === forcedField; }) ||
        { field: forcedField, index: buildBundleKeyIndex(invRecords, forcedField), overlap: 0 };
      forced.overlap = countOverlap(bundleSkuKeys, forced.index);
      return { field: forced.field, index: forced.index, overlap: forced.overlap, tried: tried.map(function (t) { return { field: t.field, overlap: t.overlap }; }) };
    }
    let best = tried[0];
    tried.forEach(function (t) { if (t.overlap > best.overlap) best = t; });
    return { field: best.field, index: best.index, overlap: best.overlap, tried: tried.map(function (t) { return { field: t.field, overlap: t.overlap }; }) };
  }

  // ----------------------------------------------------------------
  // 7. Core matching engine — Bundle SKU is the inventory matching key.
  //    A bundle carton is available only when its Bundle SKU exists in the
  //    inventory file AND Available Qty is numeric and > 0. Child rows are
  //    never used to decide availability — only shown once their parent
  //    bundle carton has already passed the check.
  // ----------------------------------------------------------------
  const EXCLUSION_LABELS = {
    AVAILABLE_QTY_ZERO: "Available Qty = 0",
    AVAILABLE_QTY_NEGATIVE: "Available Qty < 0",
    AVAILABLE_QTY_INVALID: "Available Qty blank/invalid",
    BUNDLE_SKU_NOT_FOUND_IN_INVENTORY: "Bundle SKU không có trong file tồn kho realtime",
    BUNDLE_SKU_MISSING: "Bundle SKU bị thiếu/không hợp lệ trong file bundle",
    WAREHOUSE_FILTERED_OUT: "Bị loại bởi bộ lọc Warehouse",
  };

  function runMatch(opts) {
    const bundles = opts.bundles || []; // grouped bundles from parseBundleFile
    const bundleKeyIndex = opts.bundleKeyIndex; // Map from detectBundleSkuKeySource
    const selectedBundleSkuKeys = opts.selectedBundleSkuKeys; // Set|null (of bundleSkuKey)
    const warehouseFilter = opts.warehouseFilter; // Set|null

    const available = []; // one entry per available bundle-carton (bundle x warehouse)
    const excluded = []; // one entry per excluded bundle-carton (or per not-found bundle)

    let uniqueBundleSkusFound = 0;
    let matchedInInventory = 0;
    let notFoundInInventory = 0;
    let availableQtyGt0Count = 0;
    let availableQtyLte0Count = 0;
    let excludedByWarehouseFilter = 0;

    const availableBundleSkuKeys = new Set();

    bundles.forEach(function (b) {
      const key = normalizeBundleSkuKey(b.bundleSku);
      if (selectedBundleSkuKeys && !selectedBundleSkuKeys.has(key)) return;
      uniqueBundleSkusFound++;

      const invRows = bundleKeyIndex ? bundleKeyIndex.get(key) : null;
      if (!invRows || !invRows.length) {
        notFoundInInventory++;
        excluded.push({
          bundleSku: b.bundleSku, bundleName: b.bundleName, warehouse: "",
          childCount: b.children.length, totalChildPcs: b.totalChildPcs, reason: "BUNDLE_SKU_NOT_FOUND_IN_INVENTORY",
        });
        return;
      }

      matchedInInventory++;
      invRows.forEach(function (ir) {
        if (warehouseFilter && warehouseFilter.size && !warehouseFilter.has(ir.warehouse)) {
          excludedByWarehouseFilter++;
          excluded.push({
            bundleSku: b.bundleSku, bundleName: b.bundleName, warehouse: ir.warehouse,
            childCount: b.children.length, totalChildPcs: b.totalChildPcs, reason: "WAREHOUSE_FILTERED_OUT",
          });
          return;
        }

        const avail = ir.availableQty;
        let reason = null;
        if (avail === null) reason = "AVAILABLE_QTY_INVALID";
        else if (avail < 0) reason = "AVAILABLE_QTY_NEGATIVE";
        else if (avail === 0) reason = "AVAILABLE_QTY_ZERO";

        if (reason) {
          availableQtyLte0Count++;
          excluded.push({
            bundleSku: b.bundleSku, bundleName: b.bundleName, warehouse: ir.warehouse,
            childCount: b.children.length, totalChildPcs: b.totalChildPcs, reason: reason,
          });
          return;
        }

        availableQtyGt0Count++;
        availableBundleSkuKeys.add(key);
        available.push({
          bundleSku: b.bundleSku, bundleName: b.bundleName, warehouse: ir.warehouse,
          invProductCode: ir.invProductCode, partnerSku: ir.partnerSku, invProductName: ir.product,
          category: ir.category, unit: ir.unit, conditionType: ir.conditionType,
          stockQty: ir.stockQty, freezeQty: ir.freezeQty, availableQty: avail,
          pendingIn: ir.pendingIn, pendingOut: ir.pendingOut, weightKg: ir.weightKg, cbm: ir.cbm,
          lastStockoutDate: ir.lastStockoutDate, lastInboundDate: ir.lastInboundDate, lastOutboundDate: ir.lastOutboundDate,
          childCount: b.children.length, totalChildPcs: b.totalChildPcs, children: b.children,
        });
      });
    });

    // Total child SKU lines / PCS are bundle-definition properties, counted
    // once per distinct AVAILABLE Bundle SKU (not per warehouse instance,
    // so a bundle stocked in 2 warehouses doesn't double its child totals).
    let totalChildSkuLinesInAvailableBundles = 0;
    let totalChildPcsInAvailableBundles = 0;
    bundles.forEach(function (b) {
      const key = normalizeBundleSkuKey(b.bundleSku);
      if (!availableBundleSkuKeys.has(key)) return;
      totalChildSkuLinesInAvailableBundles += b.children.length;
      totalChildPcsInAvailableBundles += b.totalChildPcs;
    });

    return {
      available: available,
      excluded: excluded,
      reconciliation: {
        uniqueBundleSkusFound: uniqueBundleSkusFound,
        matchedInInventory: matchedInInventory,
        availableQtyGt0: availableQtyGt0Count,
        availableQtyLte0: availableQtyLte0Count,
        notFoundInInventory: notFoundInInventory,
        invalidOrMissingIdentifiers: opts.missingBundleSkuRows || 0,
        excludedByWarehouseFilter: excludedByWarehouseFilter,
      },
      summary: {
        totalUniqueBundleCartonsInFile: uniqueBundleSkusFound,
        availableBundleCartons: availableBundleSkuKeys.size,
        unavailableBundleCartons: uniqueBundleSkusFound - availableBundleSkuKeys.size,
        totalChildSkuLinesInAvailableBundles: totalChildSkuLinesInAvailableBundles,
        totalChildPcsInAvailableBundles: totalChildPcsInAvailableBundles,
      },
    };
  }

  // ----------------------------------------------------------------
  // 8. Excel export (4-sheet workbook, identifiers forced to text)
  // ----------------------------------------------------------------
  function asText(v) { return v === null || v === undefined ? "" : String(v); }

  function buildExportWorkbook(result) {
    const cartonHeader = [
      "Bundle SKU", "Bundle Name", "Warehouse", "Inventory Product Code", "Partner SKU", "Inventory Product Name",
      "Category", "Unit", "Condition", "Stock Qty", "Freeze Qty", "Available Qty", "Pending In", "Pending Out",
      "Weight (Kg)", "CBM", "Last Stockout Date", "Last Inbound Date", "Last Outbound Date", "Child SKU Count", "Total Child PCS",
    ];
    const cartonAoa = [cartonHeader].concat(result.available.map(function (r) {
      return [
        asText(r.bundleSku), asText(r.bundleName), asText(r.warehouse), asText(r.invProductCode), asText(r.partnerSku),
        asText(r.invProductName), asText(r.category), asText(r.unit), asText(r.conditionType),
        r.stockQty, r.freezeQty, r.availableQty, r.pendingIn, r.pendingOut, r.weightKg, r.cbm,
        asText(r.lastStockoutDate), asText(r.lastInboundDate), asText(r.lastOutboundDate), r.childCount, r.totalChildPcs,
      ];
    }));

    const childHeader = ["Bundle SKU", "Bundle Name", "Child SKU", "Child Product Name", "Unit Code", "Unit Name", "Qty", "Availability Ratio (%)"];
    const childRows = [];
    const seenBundleForChildSheet = new Set();
    result.available.forEach(function (r) {
      if (seenBundleForChildSheet.has(r.bundleSku)) return; // list children once per available bundle, not per warehouse
      seenBundleForChildSheet.add(r.bundleSku);
      (r.children || []).forEach(function (c) {
        childRows.push([asText(r.bundleSku), asText(r.bundleName), asText(c.childSku), asText(c.childProduct), asText(c.unitCode), asText(c.unitName), c.qty, c.availabilityRatio]);
      });
    });
    const childAoa = [childHeader].concat(childRows);

    const exclHeader = ["Bundle SKU", "Bundle Name", "Warehouse", "Child SKU Count", "Total Child PCS", "Exclusion Reason"];
    const exclAoa = [exclHeader].concat(result.excluded.map(function (r) {
      return [asText(r.bundleSku), asText(r.bundleName), asText(r.warehouse), r.childCount, r.totalChildPcs, r.reason];
    }));

    const s = result.summary, rc = result.reconciliation;
    const reconAoa = [
      ["Reconciliation Summary", ""],
      ["Total Unique Bundle Cartons in Bundle File", s.totalUniqueBundleCartonsInFile],
      ["Available Bundle Cartons", s.availableBundleCartons],
      ["Unavailable / Missing Bundle Cartons", s.unavailableBundleCartons],
      ["Total Child SKU Lines in Available Bundles", s.totalChildSkuLinesInAvailableBundles],
      ["Total Child PCS in Available Bundles", s.totalChildPcsInAvailableBundles],
      ["", ""],
      ["Unique Bundle SKUs found in bundle file", rc.uniqueBundleSkusFound],
      ["Bundle SKUs matched in real-time inventory", rc.matchedInInventory],
      ["Bundle SKUs with Available Qty > 0", rc.availableQtyGt0],
      ["Bundle SKUs with Available Qty ≤ 0 / invalid", rc.availableQtyLte0],
      ["Bundle SKUs not found in real-time inventory", rc.notFoundInInventory],
      ["Bundle SKUs with invalid or missing identifiers", rc.invalidOrMissingIdentifiers],
      ["Excluded by warehouse filter", rc.excludedByWarehouseFilter],
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(cartonAoa), "Available Bundle Cartons");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(childAoa), "Available Bundle Child Details");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(exclAoa), "Unavailable Bundle Cartons");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(reconAoa), "Reconciliation Summary");
    const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    return new Blob([out], { type: "application/octet-stream" });
  }

  function exportFilename() {
    const d = new Date();
    const pad = function (n) { return String(n).padStart(2, "0"); };
    return "Real_Time_Outbound_Availability_" + d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) +
      "_" + pad(d.getHours()) + pad(d.getMinutes()) + ".xlsx";
  }

  // ----------------------------------------------------------------
  // 9. UI wiring
  // ----------------------------------------------------------------
  function init() {
    const bundleInput = document.getElementById("rtoaBundleInput");
    if (!bundleInput) return; // markup not present on this page

    const invInput = document.getElementById("rtoaInventoryInput");
    const bundleStatus = document.getElementById("rtoaBundleStatus");
    const invStatus = document.getElementById("rtoaInventoryStatus");
    const mappingBox = document.getElementById("rtoaMappingBody");
    const bundleSearch = document.getElementById("rtoaBundleSearch");
    const selectAllChk = document.getElementById("rtoaSelectAll");
    const bundleListEl = document.getElementById("rtoaBundleList");
    const warehouseFilterEl = document.getElementById("rtoaWarehouseFilter");
    const runBtn = document.getElementById("rtoaRunBtn");
    const resetBtn = document.getElementById("rtoaResetBtn");
    const logEl = document.getElementById("rtoaLog");
    const resultsWrap = document.getElementById("rtoaResultsWrap");
    const summaryCardsEl = document.getElementById("rtoaSummaryCards");
    const reconTableEl = document.getElementById("rtoaReconTable");
    const availableTableWrap = document.getElementById("rtoaAvailableTableWrap");
    const excludedTableWrap = document.getElementById("rtoaExcludedTableWrap");
    const excludedCountEl = document.getElementById("rtoaExcludedCount");
    const exportBtn = document.getElementById("rtoaExportBtn");

    let bundleParsed = null;
    let invParsed = null;
    let bundleColOverride = {};
    let invColOverride = {};
    let bundleSkuKeyFieldOverride = null; // "invProductCode" | "partnerSku" | null(auto)
    let lastResult = null;
    let checkedBundleKeys = new Set();

    function log(msg) { logEl.textContent += msg + "\n"; logEl.scrollTop = logEl.scrollHeight; }
    function tick() { return new Promise(function (r) { setTimeout(r, 0); }); }

    function refreshRunEnabled() {
      runBtn.disabled = !(bundleParsed && invParsed);
    }

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
      const bundleKeys = new Set(bundleParsed.bundles.map(function (b) { return normalizeBundleSkuKey(b.bundleSku); }));
      const detection = detectBundleSkuKeySource(bundleKeys, invParsed.records, bundleSkuKeyFieldOverride);

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
      if (detection.overlap === 0) {
        badge.className = "rtoa-map-badge warn";
        badge.textContent = "0 khớp!";
      } else {
        badge.className = "rtoa-map-badge ok";
        badge.textContent = detection.overlap + " Bundle SKU khớp";
      }
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

    // ---- Bundle SKU picker ----
    function renderBundleList() {
      bundleListEl.innerHTML = "";
      if (!bundleParsed) return;
      const q = (bundleSearch.value || "").trim().toLowerCase();
      bundleParsed.bundles.forEach(function (b) {
        if (q && b.bundleSku.toLowerCase().indexOf(q) === -1 && (b.bundleName || "").toLowerCase().indexOf(q) === -1) return;
        const key = normalizeBundleSkuKey(b.bundleSku);
        const row = document.createElement("label");
        row.className = "rtoa-bundle-row";
        const chk = document.createElement("input");
        chk.type = "checkbox";
        chk.checked = checkedBundleKeys.has(key);
        chk.addEventListener("change", function () {
          if (chk.checked) checkedBundleKeys.add(key); else checkedBundleKeys.delete(key);
        });
        const txt = document.createElement("span");
        txt.textContent = b.bundleSku + (b.bundleName ? " — " + b.bundleName : "") + " (" + b.children.length + " child SKU, " + b.totalChildPcs + " PCS)";
        row.appendChild(chk); row.appendChild(txt);
        bundleListEl.appendChild(row);
      });
    }

    function renderWarehouseFilter() {
      warehouseFilterEl.innerHTML = "";
      if (!invParsed || invParsed.warehouses.length <= 1) { warehouseFilterEl.style.display = "none"; return; }
      warehouseFilterEl.style.display = "";
      const title = document.createElement("div");
      title.className = "small";
      title.textContent = "Lọc theo Warehouse (bỏ chọn = loại khỏi kết quả):";
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

    bundleSearch.addEventListener("input", renderBundleList);
    selectAllChk.addEventListener("change", function () {
      if (!bundleParsed) return;
      if (selectAllChk.checked) bundleParsed.bundles.forEach(function (b) { checkedBundleKeys.add(normalizeBundleSkuKey(b.bundleSku)); });
      else checkedBundleKeys.clear();
      renderBundleList();
    });

    // ---- File loading ----
    function reparseBundle() {
      if (!bundleInput.files.length) return;
      try {
        const raw = bundleInput._rows;
        bundleParsed = parseBundleFile(raw, bundleColOverride);
        bundleStatus.textContent = "Đã đọc " + bundleParsed.childRows.length + " dòng child SKU — " + bundleParsed.bundles.length + " Bundle SKU (carton) distinct." +
          (bundleParsed.missingBundleSkuRows ? " (" + bundleParsed.missingBundleSkuRows + " dòng thiếu Bundle SKU đã bị bỏ qua.)" : "");
        checkedBundleKeys = new Set();
        renderBundleList();
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

    // ---- Rendering results ----
    function renderSummaryCards(s) {
      summaryCardsEl.innerHTML = "";
      const cards = [
        ["Total Unique Bundle Cartons in Bundle File", s.totalUniqueBundleCartonsInFile],
        ["Available Bundle Cartons", s.availableBundleCartons],
        ["Unavailable / Missing Bundle Cartons", s.unavailableBundleCartons],
        ["Total Child SKU Lines in Available Bundles", s.totalChildSkuLinesInAvailableBundles],
        ["Total Child PCS in Available Bundles", s.totalChildPcsInAvailableBundles],
      ];
      cards.forEach(function (c) {
        const div = document.createElement("div");
        div.className = "summary-card";
        div.innerHTML = '<div class="summary-card-value">' + escapeHtml(c[1]) + '</div><div class="summary-card-label">' + escapeHtml(c[0]) + '</div>';
        summaryCardsEl.appendChild(div);
      });
    }

    function renderRecon(rc) {
      reconTableEl.innerHTML =
        "<tr><th>Mục</th><th>Số lượng</th></tr>" +
        "<tr><td>Unique Bundle SKUs found in bundle file</td><td>" + rc.uniqueBundleSkusFound + "</td></tr>" +
        "<tr><td>Bundle SKUs matched in real-time inventory</td><td>" + rc.matchedInInventory + "</td></tr>" +
        "<tr><td>Bundle SKUs with Available Qty &gt; 0</td><td>" + rc.availableQtyGt0 + "</td></tr>" +
        "<tr><td>Bundle SKUs with Available Qty ≤ 0 / invalid</td><td>" + rc.availableQtyLte0 + "</td></tr>" +
        "<tr><td>Bundle SKUs not found in real-time inventory</td><td>" + rc.notFoundInInventory + "</td></tr>" +
        "<tr><td>Bundle SKUs with invalid or missing identifiers</td><td>" + rc.invalidOrMissingIdentifiers + "</td></tr>" +
        "<tr><td>Excluded by warehouse filter</td><td>" + rc.excludedByWarehouseFilter + "</td></tr>";
    }

    // Bundle-carton-centric table with an expandable child-detail row
    // beneath each carton (Option B: one parent bundle row + expandable
    // child table), built with real DOM nodes so the toggle buttons work.
    function renderAvailableTable(rows) {
      availableTableWrap.innerHTML = "";
      if (!rows.length) {
        const p = document.createElement("p");
        p.className = "status";
        p.textContent = "Không có bundle carton nào khả dụng (Available Qty > 0) cho lựa chọn hiện tại.";
        availableTableWrap.appendChild(p);
        return;
      }
      const cartonCols = [
        ["bundleSku", "Bundle SKU"], ["bundleName", "Bundle Name"], ["warehouse", "Warehouse"],
        ["invProductCode", "Inventory Product Code"], ["partnerSku", "Partner SKU"], ["invProductName", "Inventory Product Name"],
        ["category", "Category"], ["unit", "Unit"], ["conditionType", "Condition"],
        ["stockQty", "Stock Qty"], ["freezeQty", "Freeze Qty"], ["availableQty", "Available Qty"],
        ["pendingIn", "Pending In"], ["pendingOut", "Pending Out"], ["weightKg", "Weight (Kg)"], ["cbm", "CBM"],
        ["lastStockoutDate", "Last Stockout Date"], ["lastInboundDate", "Last Inbound Date"], ["lastOutboundDate", "Last Outbound Date"],
        ["childCount", "Child SKU Count"], ["totalChildPcs", "Total Child PCS"],
      ];
      const childCols = [
        ["childSku", "Child SKU"], ["childProduct", "Child Product Name"], ["unitCode", "Unit Code"],
        ["unitName", "Unit Name"], ["qty", "Qty"], ["availabilityRatio", "Availability Ratio (%)"],
      ];

      const table = document.createElement("table");
      table.className = "data-table";
      const thead = document.createElement("thead");
      const headTr = document.createElement("tr");
      const expandTh = document.createElement("th"); expandTh.textContent = "";
      headTr.appendChild(expandTh);
      cartonCols.forEach(function (c) { const th = document.createElement("th"); th.textContent = c[1]; headTr.appendChild(th); });
      thead.appendChild(headTr);
      table.appendChild(thead);

      const tbody = document.createElement("tbody");
      rows.forEach(function (r, idx) {
        const tr = document.createElement("tr");
        const toggleTd = document.createElement("td");
        const toggleBtn = document.createElement("button");
        toggleBtn.type = "button";
        toggleBtn.className = "rtoa-toggle-btn";
        toggleBtn.textContent = "▸ " + r.childCount;
        toggleTd.appendChild(toggleBtn);
        tr.appendChild(toggleTd);
        cartonCols.forEach(function (c) {
          const td = document.createElement("td");
          const v = r[c[0]];
          td.textContent = v === null || v === undefined ? "" : String(v);
          tr.appendChild(td);
        });
        tbody.appendChild(tr);

        const childTr = document.createElement("tr");
        childTr.className = "rtoa-child-row";
        childTr.style.display = "none";
        const childTd = document.createElement("td");
        childTd.colSpan = cartonCols.length + 1;
        const childTable = document.createElement("table");
        childTable.className = "data-table rtoa-child-table";
        const cThead = document.createElement("thead");
        const cHeadTr = document.createElement("tr");
        childCols.forEach(function (c) { const th = document.createElement("th"); th.textContent = c[1]; cHeadTr.appendChild(th); });
        cThead.appendChild(cHeadTr);
        childTable.appendChild(cThead);
        const cTbody = document.createElement("tbody");
        (r.children || []).forEach(function (child) {
          const cTr = document.createElement("tr");
          childCols.forEach(function (c) {
            const cTd = document.createElement("td");
            const v = child[c[0]];
            cTd.textContent = v === null || v === undefined ? "" : String(v);
            cTr.appendChild(cTd);
          });
          cTbody.appendChild(cTr);
        });
        childTable.appendChild(cTbody);
        childTd.appendChild(childTable);
        childTr.appendChild(childTd);
        tbody.appendChild(childTr);

        toggleBtn.addEventListener("click", function () {
          const open = childTr.style.display !== "none";
          childTr.style.display = open ? "none" : "table-row";
          toggleBtn.textContent = (open ? "▸ " : "▾ ") + r.childCount;
        });
      });
      table.appendChild(tbody);

      availableTableWrap.appendChild(table);
    }

    function renderExcludedTable(rows) {
      excludedCountEl.textContent = String(rows.length);
      if (!rows.length) { excludedTableWrap.innerHTML = '<p class="status">Không có bundle carton nào bị loại.</p>'; return; }
      let html = '<table class="data-table"><thead><tr><th>Bundle SKU</th><th>Bundle Name</th><th>Warehouse</th><th>Child SKU Count</th><th>Total Child PCS</th><th>Exclusion Reason</th></tr></thead><tbody>';
      rows.forEach(function (r) {
        html += "<tr><td>" + escapeHtml(r.bundleSku) + "</td><td>" + escapeHtml(r.bundleName) + "</td><td>" + escapeHtml(r.warehouse) +
          "</td><td>" + escapeHtml(r.childCount) + "</td><td>" + escapeHtml(r.totalChildPcs) +
          '</td><td><span class="rtoa-reason-badge">' + escapeHtml(EXCLUSION_LABELS[r.reason] || r.reason) + "</span></td></tr>";
      });
      html += "</tbody></table>";
      excludedTableWrap.innerHTML = html;
    }

    runBtn.addEventListener("click", async function () {
      if (!bundleParsed || !invParsed) return;
      if (!checkedBundleKeys.size) { log("⚠️ Vui lòng chọn ít nhất 1 Bundle SKU."); return; }
      runBtn.disabled = true;
      logEl.textContent = "";
      log("[INFO] Đang chạy đối chiếu Bundle SKU x Real-Time Inventory...");
      await tick();
      try {
        const bundleKeys = new Set(bundleParsed.bundles.map(function (b) { return normalizeBundleSkuKey(b.bundleSku); }));
        const detection = detectBundleSkuKeySource(bundleKeys, invParsed.records, bundleSkuKeyFieldOverride);
        log("[INFO] Bundle SKU key column: " + detection.field + " (khớp " + detection.overlap + "/" + checkedBundleKeys.size + " bundle đã chọn trên toàn bộ file).");
        if (detection.overlap === 0) {
          log("⚠️ No Bundle SKU from the bundle file was found in the selected real-time inventory key column. Please verify the inventory file or manually select the correct Bundle SKU column.");
        }
        const result = runMatch({
          bundles: bundleParsed.bundles,
          bundleKeyIndex: detection.index,
          selectedBundleSkuKeys: checkedBundleKeys,
          warehouseFilter: getWarehouseFilter(),
          missingBundleSkuRows: bundleParsed.missingBundleSkuRows,
        });
        lastResult = result;
        log("[INFO] Available bundle cartons: " + result.available.length + " rows | Excluded: " + result.excluded.length + " rows");
        renderSummaryCards(result.summary);
        renderRecon(result.reconciliation);
        renderAvailableTable(result.available);
        renderExcludedTable(result.excluded);
        resultsWrap.style.display = "";
        exportBtn.disabled = result.available.length === 0;
        log("✅ Hoàn tất.");
      } catch (e) {
        log("❌ Lỗi: " + e.message);
      } finally {
        runBtn.disabled = false;
      }
    });

    exportBtn.addEventListener("click", function () {
      if (!lastResult) return;
      const blob = buildExportWorkbook(lastResult);
      WOPUtils.downloadBlob(blob, exportFilename());
    });

    resetBtn.addEventListener("click", function () {
      bundleInput.value = ""; invInput.value = "";
      bundleParsed = null; invParsed = null;
      bundleColOverride = {}; invColOverride = {}; bundleSkuKeyFieldOverride = null;
      checkedBundleKeys = new Set(); lastResult = null;
      bundleStatus.textContent = "Chưa chọn file."; invStatus.textContent = "Chưa chọn file.";
      mappingBox.innerHTML = ""; bundleListEl.innerHTML = ""; warehouseFilterEl.innerHTML = ""; warehouseFilterEl.style.display = "none";
      logEl.textContent = ""; resultsWrap.style.display = "none"; exportBtn.disabled = true;
      bundleSearch.value = ""; selectAllChk.checked = false;
      refreshRunEnabled();
    });
  }

  const api = {
    init: init,
    parseBundleFile: parseBundleFile,
    parseInventoryFile: parseInventoryFile,
    detectBundleSkuKeySource: detectBundleSkuKeySource,
    buildBundleKeyIndex: buildBundleKeyIndex,
    runMatch: runMatch,
    buildExportWorkbook: buildExportWorkbook,
    exportFilename: exportFilename,
    EXCLUSION_LABELS: EXCLUSION_LABELS,
    BUNDLE_COLUMN_DEFS: BUNDLE_COLUMN_DEFS,
    INVENTORY_COLUMN_DEFS: INVENTORY_COLUMN_DEFS,
    _internal: {
      normalizeHeaderText: normalizeHeaderText, normalizeEan: normalizeEan,
      normalizeBundleSkuDisplay: normalizeBundleSkuDisplay, normalizeBundleSkuKey: normalizeBundleSkuKey,
      toNumberOrNull: toNumberOrNull, detectColumns: detectColumns, fixSheetRange: fixSheetRange,
    },
  };

  if (typeof module !== "undefined" && module.exports) module.exports = api;
  return api;
})();
