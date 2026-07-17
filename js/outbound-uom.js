/* ==========================================================
 * outbound-uom.js — Outbound: Detect Outbound Type & Suggest Convert UOM
 * Port of "WPIC_KR_Outbound_Check_Convert_FIXED.ipynb" (Python/openpyxl)
 * to 100% client-side JS (SheetJS), same philosophy as the rest of the
 * portal: files never leave the browser.
 *
 * Inputs (2 Excel files, user-uploaded):
 *   1) Inventory — "Tồn kho theo thời gian thực" export.
 *      Required columns: Mã SKU đối tác, ĐVT, Khả dụng, Tên sản phẩm, Mã sản phẩm
 *   2) Order — Outbound Request export (sheet as exported by FOMS).
 *      Required columns: SKU Code, Requested Qty, Order Number
 *
 * Outputs (2 downloadable .xlsx, built client-side with SheetJS):
 *   1) Outbound_Type_Plan.xlsx — per order line: detect outbound type
 *      (CARTON vs PCS), quantities, stock detail, status.
 *   2) Suggest_Xa_Le.xlsx — only the lines flagged "cần bóc carton"
 *      (i.e. PCS lines where loose stock alone isn't enough), aggregated
 *      into "open N cartons of size X -> Y PCS released" suggestions.
 * ========================================================== */
(function () {
  "use strict";

  // ---- low-level parsing helpers (mirror clean_id / to_num / parse_uom in the notebook) ----
  function cleanId(v) {
    if (v === null || v === undefined) return "";
    if (typeof v === "number" && Number.isInteger(v)) return String(v);
    return String(v).trim();
  }

  function toNum(v) {
    if (v === null || v === undefined) return 0;
    const f = parseFloat(String(v).replace(/,/g, ""));
    return Number.isNaN(f) ? 0 : f;
  }

  function parseUom(u) {
    u = cleanId(u).toUpperCase().replace(/\s+/g, "");
    if (u === "PCS") return { kind: "PCS", size: null };
    const m = u.match(/^CARTON_(\d+)PCS$/);
    if (m) return { kind: "CARTON", size: parseInt(m[1], 10) };
    return { kind: "KHAC", size: null };
  }

  function headerIndex(headerRow) {
    const idx = {};
    (headerRow || []).forEach(function (h, i) { if (h !== null && h !== undefined) idx[String(h).trim()] = i; });
    return idx;
  }

  function requireCols(idx, required, label) {
    const missing = required.filter(function (h) { return !(h in idx); });
    if (missing.length) throw new Error(label + " thiếu cột: " + missing.join(", "));
  }

  // ---- Inventory ----
  function parseInventory(rows, log) {
    if (!rows.length) throw new Error("File tồn kho rỗng.");
    const idx = headerIndex(rows[0]);
    requireCols(idx, ["Mã SKU đối tác", "ĐVT", "Khả dụng", "Tên sản phẩm", "Mã sản phẩm"], "Inventory");
    const skuRows = new Map();
    const eanRows = new Map();
    let lineCount = 0;
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row) continue;
      const sku = cleanId(row[idx["Mã SKU đối tác"]]);
      if (!sku) continue;
      const ean = cleanId(row[idx["Mã sản phẩm"]]);
      const record = {
        sku: sku,
        tenSp: row[idx["Tên sản phẩm"]],
        ean: ean,
        dvt: row[idx["ĐVT"]],
        khaDung: toNum(row[idx["Khả dụng"]]),
      };
      if (!skuRows.has(sku)) skuRows.set(sku, []);
      skuRows.get(sku).push(record);
      if (ean) {
        if (!eanRows.has(ean)) eanRows.set(ean, []);
        eanRows.get(ean).push(record);
      }
      lineCount++;
    }
    log("[INFO] Tồn kho: " + skuRows.size + " SKU distinct; " + eanRows.size + " EAN distinct; " + lineCount + " dòng.");
    return { skuRows: skuRows, eanRows: eanRows };
  }

  // ---- Order ----
  function parseOrder(rows, log) {
    if (!rows.length) throw new Error("File order rỗng.");
    const idx = headerIndex(rows[0]);
    requireCols(idx, ["SKU Code", "Requested Qty", "Order Number"], "Order");
    const orderLines = [];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row) continue;
      const sku = cleanId(row[idx["SKU Code"]]);
      const qty = toNum(row[idx["Requested Qty"]]);
      if (!sku || qty <= 0) continue;
      orderLines.push({ orderNum: row[idx["Order Number"]], sku: sku, qty: qty });
    }
    log("[INFO] Đơn outbound: " + orderLines.length + " dòng SKU.");
    return orderLines;
  }

  // ---- Optional Bundle check (skip silently if no bundle file given) ----
  function checkBundle(orderLines, bundleRows, log) {
    if (!bundleRows || !bundleRows.length) return;
    const bundleSkus = new Set();
    for (let i = 1; i < bundleRows.length; i++) {
      const row = bundleRows[i];
      if (!row) continue;
      const sku = cleanId(row[0]);
      if (sku) bundleSkus.add(sku);
    }
    const orderSkus = new Set(orderLines.map(function (o) { return o.sku; }));
    const overlap = Array.from(orderSkus).filter(function (s) { return bundleSkus.has(s); });
    if (overlap.length) {
      throw new Error("Có " + overlap.length + " Bundle SKU trong order, cần explode riêng (chưa hỗ trợ trong tool này): " + overlap.slice(0, 20).join(", "));
    }
    log("[INFO] Đã đọc " + bundleSkus.size + " Bundle SKU. Order không có Bundle SKU — tiếp tục check tồn trực tiếp.");
  }

  // ---- Core: detect outbound type + split CARTON/PCS for 1 SKU ----
  // Priority: use biggest carton size first (whole cartons only), remainder -> loose
  // PCS if enough; if loose isn't enough for the remainder but total stock still
  // covers the requested qty, the WHOLE line collapses into a single PCS line
  // (this is the "cần bóc carton / xả lẻ" case) — mirrors split_carton_pcs() in
  // the source notebook exactly.
  function splitCartonPcs(qty, rows) {
    if (!rows.length) {
      return {
        lines: [{ uom: "-", qtyUom: 0, qtyPcs: 0, type: "-" }],
        status: "KHONG TIM THAY SKU",
        note: "SKU không tồn tại trong file tồn kho realtime",
        totalStock: 0, looseAvail: 0, cartonDetail: "(không có trong tồn kho)",
      };
    }

    let looseAvail = 0;
    const cartonOptions = []; // [size, avail]
    rows.forEach(function (row) {
      const u = parseUom(row.dvt);
      if (u.kind === "PCS") looseAvail += row.khaDung;
      else if (u.kind === "CARTON" && row.khaDung > 0) cartonOptions.push([u.size, row.khaDung]);
    });
    cartonOptions.sort(function (a, b) { return b[0] - a[0]; }); // biggest carton first
    const totalStock = looseAvail + cartonOptions.reduce(function (s, o) { return s + o[0] * o[1]; }, 0);

    let lines = [];
    let note = null;
    let rem = qty;
    cartonOptions.forEach(function (opt) {
      if (rem <= 0) return;
      const size = opt[0], avail = opt[1];
      const full = Math.floor(rem / size);
      const use = Math.min(full, avail);
      if (use > 0) {
        lines.push({ uom: "CARTON_" + size + "PCS", qtyUom: use, qtyPcs: use * size, type: "CARTON" });
        rem -= use * size;
      }
    });

    if (rem > 0) {
      if (looseAvail >= rem) {
        lines.push({ uom: "PCS", qtyUom: rem, qtyPcs: rem, type: "PCS" });
      } else if (totalStock >= qty) {
        lines = [{ uom: "PCS", qtyUom: qty, qtyPcs: qty, type: "PCS" }];
        note = "PCS lẻ khả dụng không đủ cho phần dư sau khi trừ carton nguyên -> xả lẻ toàn bộ SL yêu cầu thành UOM PCS (cần bóc carton)";
      } else {
        lines.push({ uom: "PCS", qtyUom: looseAvail, qtyPcs: looseAvail, type: "PCS" });
      }
    }

    let status;
    if (totalStock >= qty) status = "DU HANG";
    else if (totalStock > 0) status = "THIEU HANG - XUAT 1 PHAN";
    else status = "HET HANG";

    const cartonDetail = cartonOptions.length
      ? cartonOptions.map(function (o) { return "CARTON_" + o[0] + "PCS: " + o[1] + " carton khả dụng"; }).join("; ")
      : "(không có carton)";

    return { lines: lines, status: status, note: note, totalStock: totalStock, looseAvail: looseAvail, cartonDetail: cartonDetail };
  }

  function runPlan(orderLines, skuRows, eanRows) {
    return orderLines.map(function (ol) {
      let rows = skuRows.get(ol.sku) || [];
      let matchedByEan = false, fallbackEan = "";
      if (!rows.length && /^\d{13}$/.test(ol.sku) && eanRows.has(ol.sku)) {
        rows = eanRows.get(ol.sku);
        matchedByEan = true;
        fallbackEan = ol.sku;
      }
      const outputSku = matchedByEan ? rows[0].sku : ol.sku;
      const productName = rows.length ? rows[0].tenSp : "(Không có trong tồn kho)";
      const ean = rows.length ? rows[0].ean : "";
      const r = splitCartonPcs(ol.qty, rows);
      const fallbackNote = matchedByEan ? ("Fallback EAN: Order SKU " + ol.sku + " -> EAN/SKU inventory " + fallbackEan) : "";
      const finalNote = [fallbackNote, r.note].filter(Boolean).join("; ");
      return {
        orderNum: ol.orderNum, originalSku: ol.sku, sku: outputSku, ean: ean, productName: productName,
        requestedQty: ol.qty, looseAvail: r.looseAvail, cartonDetail: r.cartonDetail,
        totalStock: r.totalStock, status: r.status, lines: r.lines, note: finalNote,
        shortage: Math.max(ol.qty - r.totalStock, 0),
      };
    });
  }

  const STATUS_DISPLAY = {
    "DU HANG": "ĐỦ HÀNG",
    "THIEU HANG - XUAT 1 PHAN": "THIẾU HÀNG - XUẤT 1 PHẦN",
    "HET HANG": "HẾT HÀNG",
    "KHONG TIM THAY SKU": "KHÔNG TÌM THẤY SKU",
  };

  function flattenAndSort(results) {
    const flat = [];
    results.forEach(function (r) {
      r.lines.forEach(function (line) { flat.push(Object.assign({}, r, { line: line })); });
    });
    const typePriority = { CARTON: 0, PCS: 1, "-": 2 };
    flat.sort(function (a, b) {
      const pa = typePriority[a.line.type] === undefined ? 3 : typePriority[a.line.type];
      const pb = typePriority[b.line.type] === undefined ? 3 : typePriority[b.line.type];
      if (pa !== pb) return pa - pb;
      if (a.sku !== b.sku) return a.sku < b.sku ? -1 : 1;
      return a.line.uom < b.line.uom ? -1 : (a.line.uom > b.line.uom ? 1 : 0);
    });
    return flat;
  }

  // ---- "Suggest convert UOM" (xả lẻ): only PCS-type lines flagged as needing
  // cartons broken open. If a SKU has several carton sizes on hand, the
  // smallest one is chosen (minimises leftover loose stock created by opening).
  function parseCartonDetail(detail) {
    if (!detail || detail.indexOf("CARTON") === -1) return [];
    return detail.split(";").map(function (s) { return s.trim(); }).map(function (s) {
      const m = s.match(/CARTON_(\d+)PCS:\s*(\d+)\s*carton/);
      return m ? { size: parseInt(m[1], 10), avail: parseInt(m[2], 10) } : null;
    }).filter(Boolean);
  }

  function buildXaLeSuggestions(flat, log) {
    const out = [];
    flat.forEach(function (item) {
      const line = item.line;
      if (line.type !== "PCS") return; // CTN lines are already whole cartons — nothing to convert
      if (!item.note || item.note.indexOf("xả lẻ") === -1) return;
      const options = parseCartonDetail(item.cartonDetail);
      if (!options.length) { log("[WARNING] " + item.sku + ": được đánh dấu cần xả lẻ nhưng không đọc được carton chi tiết — bỏ qua."); return; }
      options.sort(function (a, b) { return a.size - b.size; });
      const chosen = options[0];
      const shortfall = line.qtyPcs - item.looseAvail;
      if (shortfall <= 0) return;
      const cartonsNeeded = Math.ceil(shortfall / chosen.size);
      const pcsReleased = cartonsNeeded * chosen.size;
      let note = "Đơn " + item.orderNum + ": cần " + line.qtyPcs + " PCS lẻ, tồn PCS lẻ hiện có " + item.looseAvail +
        ", thiếu " + shortfall + " -> mở " + cartonsNeeded + " carton " + chosen.size + "PCS/carton";
      if (options.length > 1) note += " (SKU có nhiều loại carton, đã chọn loại nhỏ nhất " + chosen.size + "PCS để giảm dư thừa)";
      if (cartonsNeeded > chosen.avail) note += " | CẢNH BÁO: chỉ có " + chosen.avail + " carton khả dụng, KHÔNG ĐỦ để mở " + cartonsNeeded + " carton";
      out.push({ sku: item.sku, ean: item.ean, uom: "CARTON_" + chosen.size + "PCS", cartonsNeeded: cartonsNeeded, pcsReleased: pcsReleased, note: note });
    });
    out.sort(function (a, b) { return a.sku < b.sku ? -1 : (a.sku > b.sku ? 1 : 0); });
    return out;
  }

  // ---- xlsx build helper (SheetJS, already loaded globally as `XLSX`) ----
  function aoaToXlsxBlob(aoa, sheetName) {
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    return new Blob([out], { type: "application/octet-stream" });
  }

  // Some exports (e.g. the realtime inventory file) declare a wrong/stale
  // "!ref" dimension (e.g. A1:R2) even though the sheet actually has
  // thousands of rows of real cell data. SheetJS's sheet_to_json trusts
  // "!ref" and silently truncates in that case. Recompute the true range
  // from the actual cell addresses before parsing — mirrors the notebook's
  // ws.reset_dimensions() workaround for the same file in openpyxl.
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
      ws["!ref"] = XLSX.utils.encode_range({ s: { r: Math.min(minR, declared ? declared.s.r : minR), c: Math.min(minC, declared ? declared.s.c : minC) }, e: { r: maxR, c: maxC } });
    }
  }

  function pickFirstSheetRows(workbook, preferredName) {
    const name = workbook.SheetNames.find(function (n) { return n.toLowerCase() === (preferredName || "").toLowerCase(); }) || workbook.SheetNames[0];
    const ws = workbook.Sheets[name];
    if (!ws) throw new Error('Sheet "' + name + '" not found');
    fixSheetRange(ws);
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    return { sheetName: name, rows: rows };
  }

  // ---- Public entry point ----
  async function generate(opts) {
    const log = opts.log || function () {};

    log("[INFO] Đang đọc file tồn kho: " + opts.inventoryFile.name);
    const invWb = await WOPUtils.readWorkbookFromFile(opts.inventoryFile);
    const invPick = pickFirstSheetRows(invWb, "Sheet1");
    const skuEan = parseInventory(invPick.rows, log);
    const skuRows = skuEan.skuRows, eanRows = skuEan.eanRows;

    log("[INFO] Đang đọc file order: " + opts.orderFile.name);
    const orderWb = await WOPUtils.readWorkbookFromFile(opts.orderFile);
    const orderPick = pickFirstSheetRows(orderWb, "to10916");
    const orderLines = parseOrder(orderPick.rows, log);

    if (opts.bundleFile) {
      log("[INFO] Đang đọc file bundle: " + opts.bundleFile.name);
      const bundleWb = await WOPUtils.readWorkbookFromFile(opts.bundleFile);
      const bundlePick = pickFirstSheetRows(bundleWb, "Data");
      checkBundle(orderLines, bundlePick.rows, log);
    }

    const results = runPlan(orderLines, skuRows, eanRows);
    const flat = flattenAndSort(results);
    log("[INFO] Tổng số dòng output: " + flat.length);

    const statusCounts = {};
    results.forEach(function (r) { statusCounts[r.status] = (statusCounts[r.status] || 0) + 1; });
    Object.keys(statusCounts).forEach(function (k) {
      log("[INFO]   " + (STATUS_DISPLAY[k] || k) + ": " + statusCounts[k]);
    });

    const planAoa = [[
      "STT", "Order Number", "EAN", "SKU", "Tên sản phẩm", "Requested Qty (PCS)", "Outbound Type", "Outbound UOM",
      "Outbound Qty (theo UOM)", "Outbound Qty quy đổi PCS", "Tồn khả dụng - PCS lẻ", "Tồn khả dụng - Carton (chi tiết)",
      "Tổng tồn khả dụng (quy đổi PCS)", "Trạng thái đáp ứng", "Thiếu (PCS)", "Ghi chú",
    ]];
    flat.forEach(function (item, i) {
      const line = item.line;
      planAoa.push([
        i + 1, item.orderNum, item.ean, item.sku, item.productName, line.qtyPcs, line.type, line.uom,
        line.qtyUom, line.qtyPcs, item.looseAvail, item.cartonDetail, item.totalStock,
        STATUS_DISPLAY[item.status] || item.status, item.shortage, item.note || "",
      ]);
    });

    const xaLe = buildXaLeSuggestions(flat, log);
    const totalCartons = xaLe.reduce(function (s, x) { return s + x.cartonsNeeded; }, 0);
    log("[INFO] Số SKU cần xả lẻ (mở carton): " + xaLe.length);
    log("[INFO] Tổng số carton cần mở (toàn đơn): " + totalCartons);

    const xaLeAoa = [["SKU", "EAN", "UOM", "Số CARTON cần xả lẻ", "Tổng PCS lấy ra từ xả lẻ", "Ghi chú"]];
    xaLe.forEach(function (x) { xaLeAoa.push([x.sku, x.ean, x.uom, x.cartonsNeeded, x.pcsReleased, x.note]); });
    xaLeAoa.push([]);
    xaLeAoa.push([]);
    xaLeAoa.push(["Tổng số SKU cần xả lẻ carton:", null, xaLe.length]);
    xaLeAoa.push(["Tổng số carton cần mở (toàn đơn):", null, totalCartons]);

    const planBlob = aoaToXlsxBlob(planAoa, "Outbound Plan");
    const xaLeBlob = aoaToXlsxBlob(xaLeAoa, "Suggest_Xa_Le");

    return {
      files: [
        { name: "Outbound_Type_Plan.xlsx", blob: planBlob, count: flat.length + " dòng" },
        { name: "Suggest_Xa_Le.xlsx", blob: xaLeBlob, count: xaLe.length + " SKU / " + totalCartons + " carton" },
      ],
      results: results, flat: flat, xaLe: xaLe,
    };
  }

  window.WOPOutboundUom = {
    generate: generate,
    // exported for unit testing (Node) — not used by the UI directly
    _internal: { cleanId: cleanId, toNum: toNum, parseUom: parseUom, splitCartonPcs: splitCartonPcs, buildXaLeSuggestions: buildXaLeSuggestions, flattenAndSort: flattenAndSort, runPlan: runPlan, parseInventory: parseInventory, parseOrder: parseOrder },
  };
})();
