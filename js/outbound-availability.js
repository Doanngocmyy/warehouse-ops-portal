/* ==========================================================
 * outbound-availability.js — Outbound: Real-Time Outbound Availability
 * (Bundle-Aware Inventory Report)
 *
 * Business problem this solves:
 *   The "bundle child product" export lists EVERY Child EAN that can ever
 *   belong to a bundle, regardless of whether that carton currently has
 *   stock. It must NOT be treated as the final list of sellable/outboundable
 *   cartons. This tool reverse-checks every Child EAN of the selected
 *   Bundle SKU(s) against the latest Real-Time Detailed Inventory export
 *   and keeps only cartons where Available Qty > 0 right now.
 *
 * 100% client-side (SheetJS, already loaded globally as `XLSX` in
 * index.html) — files never leave the browser, same philosophy as the
 * rest of this portal.
 * ========================================================== */
window.WOPOutboundAvailability = (function () {
  "use strict";

  // ----------------------------------------------------------------
  // 1. Normalization helpers
  // ----------------------------------------------------------------

  // Normalize a header cell (or alias string) for column detection:
  // lower-case, "đ"->"d", strip accents, collapse to single spaces.
  // Both real headers and the alias dictionary below go through this
  // same function, so "Khả dụng" and "kha dung" compare equal.
  function normalizeHeaderText(h) {
    if (h === null || h === undefined) return "";
    let s = String(h).toLowerCase();
    s = s.replace(/đ/g, "d");
    s = s.normalize("NFD").replace(/[̀-ͯ]/g, "");
    s = s.replace(/[^a-z0-9]+/g, " ").trim().replace(/\s+/g, " ");
    return s;
  }

  // Normalize a Child EAN / Barcode value into a stable string key.
  //  - numbers: printed without any decimal point (Excel/SheetJS keeps
  //    13-digit EANs as exact integers when read with raw:true, so no
  //    floating-point precision is lost here — this just guards against
  //    a stray ".0" if some upstream export coerced it to text first).
  //  - strings: trimmed, internal ".0"/".00" suffix stripped, whitespace
  //    removed. Leading zeroes are preserved whenever the source value
  //    was already a string (we never round-trip through Number()).
  function normalizeEan(v) {
    if (v === null || v === undefined) return "";
    if (typeof v === "number") {
      if (!isFinite(v)) return "";
      return Number.isInteger(v) ? String(v) : String(Math.round(v));
    }
    let s = String(v).trim();
    if (s === "") return "";
    s = s.replace(/\s+/g, "");
    s = s.replace(/\.0+$/, ""); // strip Excel-generated trailing ".0"/".00"
    return s;
  }

  // Parse a quantity cell into a finite number, or null if it is blank,
  // text, or otherwise not a valid number. Never uses floating "loose"
  // coercion (e.g. "" or "abc" must NOT silently become 0).
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
  // 2. Column definitions (key -> list of accepted header aliases)
  //    Order matters: earlier defs claim a column before later ones,
  //    which avoids e.g. "EAN Qty" being mis-claimed by the "Child EAN"
  //    detector.
  // ----------------------------------------------------------------
  const BUNDLE_COLUMN_DEFS = [
    { key: "bundleSku", required: true, aliases: ["Bundle SKU (Barcode)", "Bundle SKU", "Parent SKU", "Mã Bundle", "SKU Bundle", "Bundle Barcode"] },
    { key: "bundleName", required: false, aliases: ["Tên Bundle", "Bundle Name"] },
    { key: "childEan", required: true, aliases: ["Child SKU (Barcode)", "Child EAN", "Child Barcode", "EAN con", "Mã EAN con", "Child SKU"] },
    { key: "childProduct", required: false, aliases: ["Tên sản phẩm con", "Child Product", "Child Product Name", "Product Name"] },
    { key: "unit", required: false, aliases: ["Mã ĐVT", "Tên ĐVT", "Unit", "ĐVT", "Đơn vị tính"] },
    { key: "eanQty", required: false, aliases: ["SL", "EAN Qty", "Số lượng", "Qty/Carton", "PCS per carton", "Quantity"] },
  ];

  const INVENTORY_COLUMN_DEFS = [
    { key: "warehouse", required: false, aliases: ["Warehouse", "Kho"] },
    { key: "barcode", required: false, aliases: ["Barcode", "Mã vạch"] },
    { key: "partnerSku", required: false, aliases: ["Partner SKU", "Mã SKU đối tác"] },
    { key: "childEan", required: true, aliases: ["Child EAN", "Mã sản phẩm", "Child Barcode", "EAN con"] },
    { key: "product", required: false, aliases: ["Product", "Tên sản phẩm"] },
    { key: "category", required: false, aliases: ["Category", "Danh mục"] },
    { key: "unit", required: false, aliases: ["Unit", "ĐVT", "Đơn vị tính"] },
    { key: "conditionType", required: false, aliases: ["Condition type", "Condition Type", "Tình trạng"] },
    { key: "stockQty", required: false, aliases: ["Stock Qty", "Tồn kho"] },
    { key: "freezeQty", required: false, aliases: ["Freeze Qty", "Đóng băng"] },
    { key: "availableQty", required: true, aliases: ["Available Qty", "Khả dụng"] },
    { key: "pendingIn", required: false, aliases: ["Pending In", "Chờ nhập"] },
    { key: "pendingOut", required: false, aliases: ["Pending Out", "Chờ xuất"] },
    { key: "weightKg", required: false, aliases: ["Weight (Kg)", "Weight", "Khối lượng (Kg)", "Khối lượng"] },
    { key: "cbm", required: false, aliases: ["CBM"] },
    { key: "lastStockoutDate", required: false, aliases: ["Last Stockout Date", "Ngày xuất kho gần nhất"] },
    { key: "lastOutboundDate", required: false, aliases: ["Last Outbound Date", "Ngày xuất gần nhất"] },
    { key: "lastInboundDate", required: false, aliases: ["Last Inbound Date", "Ngày nhập gần nhất"] },
    { key: "eanQty", required: false, aliases: ["EAN Qty", "Số lượng EAN"] },
  ];

  // Detect which column index best matches each definition. Exact
  // (normalized) matches are tried first; a "contains" fallback is used
  // only if no exact match exists. Each column index can be claimed by
  // at most one field.
  function detectColumns(headerRow, defs) {
    const normalizedHeaders = (headerRow || []).map(normalizeHeaderText);
    const used = new Set();
    const mapping = {}; // key -> column index
    const detectedHeader = {}; // key -> raw header text
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

    // Duplicate-header warning: same raw header text used more than once.
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
  // 3. Workbook / sheet reading (mirrors outbound-uom.js: some real
  //    exports declare a stale "!ref" dimension smaller than the actual
  //    data — recompute the true range from real cell addresses first).
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
  // 4. Parse Bundle Child Product File
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

    const records = [];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row) continue;
      const bundleSku = cleanStr(row[mapping.bundleSku]);
      const childEanRaw = mapping.childEan !== undefined ? row[mapping.childEan] : null;
      if (!bundleSku && normalizeEan(childEanRaw) === "") continue; // fully blank row
      records.push({
        rowIndex: i + 1,
        bundleSku: bundleSku,
        bundleName: mapping.bundleName !== undefined ? cleanStr(row[mapping.bundleName]) : "",
        childEanRaw: childEanRaw,
        childEan: normalizeEan(childEanRaw),
        childProduct: mapping.childProduct !== undefined ? cleanStr(row[mapping.childProduct]) : "",
        unit: mapping.unit !== undefined ? cleanStr(row[mapping.unit]) : "",
        eanQty: mapping.eanQty !== undefined ? toNumberOrNull(row[mapping.eanQty]) : null,
      });
    }

    // distinct bundle list for the SKU picker
    const bundleMap = new Map();
    records.forEach(function (r) {
      if (!r.bundleSku) return;
      if (!bundleMap.has(r.bundleSku)) bundleMap.set(r.bundleSku, { sku: r.bundleSku, name: r.bundleName, count: 0 });
      bundleMap.get(r.bundleSku).count++;
    });
    const bundles = Array.from(bundleMap.values()).sort(function (a, b) { return a.sku < b.sku ? -1 : (a.sku > b.sku ? 1 : 0); });

    return { headerRow: headerRow, mapping: mapping, detected: det, records: records, bundles: bundles };
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
      const childEan = normalizeEan(get(row, "childEan"));
      if (!childEan) continue; // no usable key on this row
      const warehouse = cleanStr(get(row, "warehouse")) || "(N/A)";
      warehouses.add(warehouse);
      records.push({
        rowIndex: i + 1,
        warehouse: warehouse,
        childEan: childEan,
        barcode: cleanStr(get(row, "barcode")),
        partnerSku: cleanStr(get(row, "partnerSku")),
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
        invEanQty: toNumberOrNull(get(row, "eanQty")),
        lastStockoutDate: get(row, "lastStockoutDate"),
        lastOutboundDate: get(row, "lastOutboundDate"),
        lastInboundDate: get(row, "lastInboundDate"),
      });
    }
    return { headerRow: headerRow, mapping: mapping, detected: det, records: records, warehouses: Array.from(warehouses).sort() };
  }

  function buildInventoryIndex(invRecords) {
    const idx = new Map();
    invRecords.forEach(function (r) {
      if (!idx.has(r.childEan)) idx.set(r.childEan, []);
      idx.get(r.childEan).push(r);
    });
    return idx;
  }

  // ----------------------------------------------------------------
  // 6. Core matching engine
  //    Bundle SKU -> Child EANs (bundle file) -> reverse-check against
  //    inventory index -> keep only Available Qty > 0 (+ warehouse filter).
  // ----------------------------------------------------------------
  const EXCLUSION_LABELS = {
    AVAILABLE_QTY_ZERO: "Available Qty = 0",
    AVAILABLE_QTY_NEGATIVE: "Available Qty < 0",
    AVAILABLE_QTY_INVALID: "Available Qty blank/invalid",
    CHILD_EAN_NOT_FOUND_IN_INVENTORY: "Child EAN không có trong file tồn kho realtime",
    CHILD_EAN_MISSING: "Child EAN bị thiếu/không hợp lệ trong file bundle",
    WAREHOUSE_FILTERED_OUT: "Bị loại bởi bộ lọc Warehouse",
  };

  function runMatch(opts) {
    const bundleRecords = opts.bundleRecords || [];
    const inventoryIndex = opts.inventoryIndex || new Map();
    const selectedBundleSkus = opts.selectedBundleSkus; // Set|null (null/undefined = all)
    const warehouseFilter = opts.warehouseFilter; // Set|null (null/undefined = no filter)

    const available = [];
    const excluded = [];

    let totalBundleChildRows = 0;
    let matchedCount = 0;
    let notFoundCount = 0;
    let invalidMissingCount = 0;
    let qtyZeroOrInvalidCount = 0;
    let warehouseFilteredCount = 0;

    // tracks whether a given (bundleSku|childEan) pair ended up with at
    // least one included row, so summary cards don't double-count a pair
    // that is available from warehouse A but excluded from warehouse B.
    const pairHasAvailable = new Set();
    const pairSeen = new Set();

    bundleRecords.forEach(function (br) {
      if (selectedBundleSkus && !selectedBundleSkus.has(br.bundleSku)) return;
      totalBundleChildRows++;
      const pairKey = br.bundleSku + "" + br.childEan;
      pairSeen.add(pairKey);

      if (!br.childEan) {
        invalidMissingCount++;
        excluded.push({
          bundleSku: br.bundleSku, bundleName: br.bundleName, childEan: cleanStr(br.childEanRaw),
          product: br.childProduct, eanQty: br.eanQty, warehouse: "", reason: "CHILD_EAN_MISSING",
        });
        return;
      }

      const invRows = inventoryIndex.get(br.childEan);
      if (!invRows || !invRows.length) {
        notFoundCount++;
        excluded.push({
          bundleSku: br.bundleSku, bundleName: br.bundleName, childEan: br.childEan,
          product: br.childProduct, eanQty: br.eanQty, warehouse: "", reason: "CHILD_EAN_NOT_FOUND_IN_INVENTORY",
        });
        return;
      }

      matchedCount++;
      invRows.forEach(function (ir) {
        if (warehouseFilter && warehouseFilter.size && !warehouseFilter.has(ir.warehouse)) {
          warehouseFilteredCount++;
          excluded.push({
            bundleSku: br.bundleSku, bundleName: br.bundleName, childEan: br.childEan,
            product: br.childProduct || ir.product, eanQty: br.eanQty, warehouse: ir.warehouse, reason: "WAREHOUSE_FILTERED_OUT",
          });
          return;
        }

        const avail = ir.availableQty;
        let reason = null;
        if (avail === null) reason = "AVAILABLE_QTY_INVALID";
        else if (avail < 0) reason = "AVAILABLE_QTY_NEGATIVE";
        else if (avail === 0) reason = "AVAILABLE_QTY_ZERO";

        if (reason) {
          qtyZeroOrInvalidCount++;
          excluded.push({
            bundleSku: br.bundleSku, bundleName: br.bundleName, childEan: br.childEan,
            product: br.childProduct || ir.product, eanQty: br.eanQty, warehouse: ir.warehouse, reason: reason,
          });
          return;
        }

        // Included. EAN Qty (PCS/carton) is taken from the BUNDLE file —
        // it is the qty that defines how many PCS this Child EAN
        // contributes to ONE unit of the bundle — never from Stock/EAN
        // Qty in the inventory file (that number is not used to gate
        // availability, only Available Qty from inventory is).
        const eanQty = br.eanQty;
        const availablePcs = (avail !== null && eanQty !== null) ? avail * eanQty : null;
        pairHasAvailable.add(pairKey);
        available.push({
          bundleSku: br.bundleSku, bundleName: br.bundleName,
          warehouse: ir.warehouse, childEan: br.childEan, barcode: ir.barcode, partnerSku: ir.partnerSku,
          product: br.childProduct || ir.product, category: ir.category, unit: ir.unit || br.unit,
          conditionType: ir.conditionType, stockQty: ir.stockQty, freezeQty: ir.freezeQty,
          availableQty: avail, pendingIn: ir.pendingIn, pendingOut: ir.pendingOut, weightKg: ir.weightKg, cbm: ir.cbm,
          eanQty: eanQty, availablePcs: availablePcs,
          lastStockoutDate: ir.lastStockoutDate, lastOutboundDate: ir.lastOutboundDate, lastInboundDate: ir.lastInboundDate,
        });
      });
    });

    const totalAvailableQty = available.reduce(function (s, r) { return s + (r.availableQty || 0); }, 0);
    const totalAvailablePcs = available.reduce(function (s, r) { return s + (r.availablePcs || 0); }, 0);

    let availablePairCount = 0, excludedPairCount = 0;
    pairSeen.forEach(function (k) { if (pairHasAvailable.has(k)) availablePairCount++; else excludedPairCount++; });

    return {
      available: available,
      excluded: excluded,
      reconciliation: {
        foundInBundleFile: totalBundleChildRows,
        matchedInInventory: matchedCount,
        excludedQtyZeroOrInvalid: qtyZeroOrInvalidCount,
        notFoundInInventory: notFoundCount,
        invalidOrMissingChildEan: invalidMissingCount,
        excludedByWarehouseFilter: warehouseFilteredCount,
      },
      summary: {
        selectedBundles: selectedBundleSkus ? selectedBundleSkus.size : new Set(bundleRecords.map(function (r) { return r.bundleSku; })).size,
        totalChildEansInBundleFile: totalBundleChildRows,
        availableChildEans: availablePairCount,
        excludedChildEans: excludedPairCount,
        totalAvailableQty: totalAvailableQty,
        totalAvailablePcs: totalAvailablePcs,
      },
    };
  }

  // ----------------------------------------------------------------
  // 7. Excel export (3-sheet workbook, EAN/barcode forced to text)
  // ----------------------------------------------------------------
  function asText(v) { return v === null || v === undefined ? "" : String(v); }

  function buildExportWorkbook(result) {
    const availHeader = [
      "Bundle SKU", "Warehouse", "Child EAN", "Barcode", "Partner SKU", "Product", "Category", "Unit",
      "Condition Type", "Stock Qty", "Freeze Qty", "Available Qty", "Pending In", "Pending Out", "Weight (Kg)", "CBM",
      "EAN Qty / PCS per Carton", "Total Available PCS", "Last Stockout Date", "Last Outbound Date", "Last Inbound Date",
    ];
    const availAoa = [availHeader].concat(result.available.map(function (r) {
      return [
        asText(r.bundleSku), asText(r.warehouse), asText(r.childEan), asText(r.barcode), asText(r.partnerSku),
        asText(r.product), asText(r.category), asText(r.unit), asText(r.conditionType),
        r.stockQty, r.freezeQty, r.availableQty, r.pendingIn, r.pendingOut, r.weightKg, r.cbm,
        r.eanQty, r.availablePcs, asText(r.lastStockoutDate), asText(r.lastOutboundDate), asText(r.lastInboundDate),
      ];
    }));

    const exclHeader = ["Bundle SKU", "Child EAN", "Product", "EAN Qty", "Warehouse", "Exclusion Reason"];
    const exclAoa = [exclHeader].concat(result.excluded.map(function (r) {
      return [asText(r.bundleSku), asText(r.childEan), asText(r.product), r.eanQty, asText(r.warehouse), r.reason];
    }));

    const s = result.summary, rc = result.reconciliation;
    const reconAoa = [
      ["Reconciliation Summary", ""],
      ["Selected Bundles", s.selectedBundles],
      ["Total Child EANs in Bundle File", s.totalChildEansInBundleFile],
      ["Available Child EANs", s.availableChildEans],
      ["Excluded / Unavailable Child EANs", s.excludedChildEans],
      ["Total Available Qty", s.totalAvailableQty],
      ["Total Available PCS", s.totalAvailablePcs],
      ["", ""],
      ["Found in bundle file", rc.foundInBundleFile],
      ["Matched in inventory", rc.matchedInInventory],
      ["Excluded because Available Qty ≤ 0 / invalid", rc.excludedQtyZeroOrInvalid],
      ["Not found in inventory", rc.notFoundInInventory],
      ["Invalid or missing Child EAN", rc.invalidOrMissingChildEan],
      ["Excluded by warehouse filter", rc.excludedByWarehouseFilter],
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(availAoa), "Available Inventory");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(exclAoa), "Excluded Child EANs");
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
  // 8. UI wiring
  // ----------------------------------------------------------------
  function init() {
    const bundleInput = document.getElementById("rtoaBundleInput");
    if (!bundleInput) return; // markup not present on this page

    const invInput = document.getElementById("rtoaInventoryInput");
    const bundleStatus = document.getElementById("rtoaBundleStatus");
    const invStatus = document.getElementById("rtoaInventoryStatus");
    const mappingBox = document.getElementById("rtoaMappingBody");
    const filterPanel = document.getElementById("rtoaFilterPanel");
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

    let bundleParsed = null; // { headerRow, mapping, detected, records, bundles }
    let invParsed = null;    // { headerRow, mapping, detected, records, warehouses }
    let invIndex = null;
    let bundleColOverride = {};
    let invColOverride = {};
    let lastResult = null;
    let checkedBundles = new Set();

    function log(msg) { logEl.textContent += msg + "\n"; logEl.scrollTop = logEl.scrollHeight; }
    function tick() { return new Promise(function (r) { setTimeout(r, 0); }); }

    function refreshRunEnabled() {
      runBtn.disabled = !(bundleParsed && invParsed);
    }

    // ---- File Mapping panel (shared for both files) ----
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
    }

    // ---- Bundle SKU picker ----
    function renderBundleList() {
      bundleListEl.innerHTML = "";
      if (!bundleParsed) return;
      const q = (bundleSearch.value || "").trim().toLowerCase();
      bundleParsed.bundles.forEach(function (b) {
        if (q && b.sku.toLowerCase().indexOf(q) === -1 && (b.name || "").toLowerCase().indexOf(q) === -1) return;
        const row = document.createElement("label");
        row.className = "rtoa-bundle-row";
        const chk = document.createElement("input");
        chk.type = "checkbox";
        chk.checked = checkedBundles.has(b.sku);
        chk.addEventListener("change", function () {
          if (chk.checked) checkedBundles.add(b.sku); else checkedBundles.delete(b.sku);
        });
        const txt = document.createElement("span");
        txt.textContent = b.sku + (b.name ? " — " + b.name : "") + " (" + b.count + " child EAN)";
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
      if (checked.length === boxes.length) return null; // all checked = no filtering
      return new Set(checked);
    }

    bundleSearch.addEventListener("input", renderBundleList);
    selectAllChk.addEventListener("change", function () {
      if (!bundleParsed) return;
      if (selectAllChk.checked) bundleParsed.bundles.forEach(function (b) { checkedBundles.add(b.sku); });
      else checkedBundles.clear();
      renderBundleList();
    });

    // ---- File loading ----
    function reparseBundle() {
      if (!bundleInput.files.length) return;
      try {
        const raw = bundleInput._rows;
        bundleParsed = parseBundleFile(raw, bundleColOverride);
        bundleStatus.textContent = "Đã đọc " + bundleParsed.records.length + " dòng child EAN — " + bundleParsed.bundles.length + " Bundle SKU distinct.";
        checkedBundles = new Set();
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
        invIndex = buildInventoryIndex(invParsed.records);
        invStatus.textContent = "Đã đọc " + invParsed.records.length + " dòng tồn kho — " + invParsed.warehouses.length + " warehouse.";
        renderWarehouseFilter();
        renderMapping();
        refreshRunEnabled();
      } catch (e) {
        invParsed = null; invIndex = null;
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
        reparseInventory();
      } catch (e) {
        invStatus.textContent = "❌ Lỗi đọc file: " + e.message;
        invParsed = null; invIndex = null; refreshRunEnabled();
      }
    });

    // ---- Rendering results ----
    function renderSummaryCards(s) {
      summaryCardsEl.innerHTML = "";
      const cards = [
        ["Selected Bundles", s.selectedBundles],
        ["Total Child EANs in Bundle File", s.totalChildEansInBundleFile],
        ["Available Child EANs", s.availableChildEans],
        ["Excluded / Unavailable Child EANs", s.excludedChildEans],
        ["Total Available Qty", s.totalAvailableQty],
        ["Total Available PCS", s.totalAvailablePcs],
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
        "<tr><td>Found in bundle file</td><td>" + rc.foundInBundleFile + "</td></tr>" +
        "<tr><td>Matched in inventory</td><td>" + rc.matchedInInventory + "</td></tr>" +
        "<tr><td>Excluded because Available Qty ≤ 0 / invalid</td><td>" + rc.excludedQtyZeroOrInvalid + "</td></tr>" +
        "<tr><td>Not found in inventory</td><td>" + rc.notFoundInInventory + "</td></tr>" +
        "<tr><td>Invalid or missing Child EAN</td><td>" + rc.invalidOrMissingChildEan + "</td></tr>" +
        "<tr><td>Excluded by warehouse filter</td><td>" + rc.excludedByWarehouseFilter + "</td></tr>";
    }

    function renderAvailableTable(rows) {
      if (!rows.length) { availableTableWrap.innerHTML = '<p class="status">Không có dòng nào khả dụng (Available Qty &gt; 0) cho lựa chọn hiện tại.</p>'; return; }
      const cols = [
        ["bundleSku", "Bundle SKU"], ["warehouse", "Warehouse"], ["childEan", "Child EAN"], ["barcode", "Barcode"],
        ["partnerSku", "Partner SKU"], ["product", "Product"], ["category", "Category"], ["unit", "Unit"],
        ["conditionType", "Condition Type"], ["stockQty", "Stock Qty"], ["freezeQty", "Freeze Qty"], ["availableQty", "Available Qty"],
        ["pendingIn", "Pending In"], ["pendingOut", "Pending Out"], ["weightKg", "Weight (Kg)"], ["cbm", "CBM"],
        ["eanQty", "EAN Qty / PCS per Carton"], ["availablePcs", "Total Available PCS"],
        ["lastStockoutDate", "Last Stockout Date"], ["lastOutboundDate", "Last Outbound Date"], ["lastInboundDate", "Last Inbound Date"],
      ];
      let html = "<table class=\"data-table\"><thead><tr>" + cols.map(function (c) { return "<th>" + c[1] + "</th>"; }).join("") + "</tr></thead><tbody>";
      rows.forEach(function (r) {
        html += "<tr>" + cols.map(function (c) { return "<td>" + escapeHtml(r[c[0]] === null || r[c[0]] === undefined ? "" : r[c[0]]) + "</td>"; }).join("") + "</tr>";
      });
      html += "</tbody></table>";
      availableTableWrap.innerHTML = html;
    }

    function renderExcludedTable(rows) {
      excludedCountEl.textContent = String(rows.length);
      if (!rows.length) { excludedTableWrap.innerHTML = '<p class="status">Không có Child EAN nào bị loại.</p>'; return; }
      let html = '<table class="data-table"><thead><tr><th>Bundle SKU</th><th>Child EAN</th><th>Product</th><th>EAN Qty</th><th>Warehouse</th><th>Exclusion Reason</th></tr></thead><tbody>';
      rows.forEach(function (r) {
        html += "<tr><td>" + escapeHtml(r.bundleSku) + "</td><td>" + escapeHtml(r.childEan) + "</td><td>" + escapeHtml(r.product) +
          "</td><td>" + escapeHtml(r.eanQty === null || r.eanQty === undefined ? "" : r.eanQty) + "</td><td>" + escapeHtml(r.warehouse) +
          '</td><td><span class="rtoa-reason-badge">' + escapeHtml(EXCLUSION_LABELS[r.reason] || r.reason) + "</span></td></tr>";
      });
      html += "</tbody></table>";
      excludedTableWrap.innerHTML = html;
    }

    runBtn.addEventListener("click", async function () {
      if (!bundleParsed || !invParsed) return;
      if (!checkedBundles.size) { log("⚠️ Vui lòng chọn ít nhất 1 Bundle SKU."); return; }
      runBtn.disabled = true;
      logEl.textContent = "";
      log("[INFO] Đang chạy đối chiếu Bundle x Real-Time Inventory...");
      await tick();
      try {
        const result = runMatch({
          bundleRecords: bundleParsed.records,
          inventoryIndex: invIndex,
          selectedBundleSkus: checkedBundles,
          warehouseFilter: getWarehouseFilter(),
        });
        lastResult = result;
        log("[INFO] Available rows: " + result.available.length + " | Excluded rows: " + result.excluded.length);
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
      bundleParsed = null; invParsed = null; invIndex = null;
      bundleColOverride = {}; invColOverride = {}; checkedBundles = new Set(); lastResult = null;
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
    buildInventoryIndex: buildInventoryIndex,
    runMatch: runMatch,
    buildExportWorkbook: buildExportWorkbook,
    exportFilename: exportFilename,
    EXCLUSION_LABELS: EXCLUSION_LABELS,
    BUNDLE_COLUMN_DEFS: BUNDLE_COLUMN_DEFS,
    INVENTORY_COLUMN_DEFS: INVENTORY_COLUMN_DEFS,
    _internal: {
      normalizeHeaderText: normalizeHeaderText, normalizeEan: normalizeEan, toNumberOrNull: toNumberOrNull,
      detectColumns: detectColumns, fixSheetRange: fixSheetRange,
    },
  };

  if (typeof module !== "undefined" && module.exports) module.exports = api;
  return api;
})();
