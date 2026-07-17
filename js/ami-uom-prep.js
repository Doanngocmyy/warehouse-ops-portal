/* ==========================================================
 * ami-uom-prep.js — Pre-inbound: chuẩn bị file cập nhật UOM cho AMI WMS
 *
 * Port của load_work_table() trong script "AMI SKU RE-UPDATE — SMART FAST
 * MODE" (Python/Selenium, chạy local) sang JS thuần, 100% client-side —
 * cùng triết lý với phần còn lại của portal: file không rời khỏi trình
 * duyệt.
 *
 * QUAN TRỌNG: tool này CHỈ chuẩn bị & làm sạch file input. Nó KHÔNG tự
 * đăng nhập / cập nhật AMI WMS (ami.partner.vnfai.com) — việc đó vẫn cần
 * script Python (Selenium) chạy trên máy bạn, vì JS trên 1 trang web tĩnh
 * không thể (và không nên) điều khiển DOM của 1 origin khác, và càng
 * không nên chứa sẵn mật khẩu đăng nhập trong code public.
 *
 * Input: 1 file Excel bất kỳ có 1 cột chứa "Mã SKU đối tác" (hoặc biến
 * thể: SKU partner...) và 1 cột ĐVT/UOM/Qty (carton pack size). Tool tự
 * dò dòng header + cột, chấp nhận nhiều định dạng ĐVT: 120 / 120.0 /
 * "120 - CARTON_120PCS" / "CARTON_120PCS - 120" / "CARTON_120PCS".
 * (Giống hệt logic find_header_row / find_col_by_keywords /
 * extract_qty_from_input trong script Python gốc.)
 *
 * Output: 1 file .xlsx sạch, đúng 2 cột "Mã SKU đối tác" | "ĐVT" (đã
 * validate, bỏ dòng thiếu/lỗi, dedupe theo SKU+qty) — dùng trực tiếp làm
 * EXCEL_PATH cho script Python cập nhật AMI, không cần chỉnh sửa gì thêm.
 * ========================================================== */
(function () {
  "use strict";

  function normText(v) {
    if (v === null || v === undefined) return "";
    let s = String(v).trim().toLowerCase();
    s = s.replace(/\n/g, " ").replace(/\t/g, " ");
    s = s.replace(/\s+/g, " ");
    return s;
  }

  const HEADER_KEYS_SKU = [
    "mã sku đối tác", "ma sku doi tac", "sku đối tác", "sku doi tac",
    "partner sku", "sku partner",
  ];
  const HEADER_KEYS_UOM = ["đvt", "dvt", "uom", "qty", "quantity"];

  function findHeaderRow(rows, maxScan) {
    maxScan = Math.min(rows.length, maxScan || 40);
    for (let r = 0; r < maxScan; r++) {
      const joined = (rows[r] || []).map(normText).join(" | ");
      if (HEADER_KEYS_SKU.some(function (k) { return joined.indexOf(k) !== -1; })) return r;
    }
    for (let r = 0; r < maxScan; r++) {
      const joined = (rows[r] || []).map(normText).join(" | ");
      if (joined.indexOf("sku") !== -1 && HEADER_KEYS_UOM.some(function (k) { return joined.indexOf(k) !== -1; })) return r;
    }
    return null;
  }

  function findColByKeywords(headerRow, allKeywords, anyKeywords) {
    const norm = headerRow.map(normText);
    if (allKeywords && allKeywords.length) {
      for (let i = 0; i < norm.length; i++) {
        if (allKeywords.every(function (k) { return norm[i].indexOf(k) !== -1; })) return i;
      }
    }
    if (anyKeywords && anyKeywords.length) {
      for (let i = 0; i < norm.length; i++) {
        if (anyKeywords.some(function (k) { return norm[i].indexOf(k) !== -1; })) return i;
      }
    }
    return -1;
  }

  function extractQtyFromInput(v) {
    let s = String(v === null || v === undefined ? "" : v).trim().toUpperCase();
    s = s.replace(/\s+/g, " ");
    if (s === "" || s === "NAN") return null;

    const f = parseFloat(s.replace(/,/g, ""));
    if (!Number.isNaN(f) && Number.isInteger(f)) return f;

    let m = s.match(/^(\d+)\s*-\s*CARTON_(\d+)PCS$/);
    if (m && m[1] === m[2]) return parseInt(m[1], 10);

    m = s.match(/CARTON_(\d+)PCS\s*-\s*(\d+)/);
    if (m && m[1] === m[2]) return parseInt(m[1], 10);

    m = s.match(/CARTON_(\d+)PCS/);
    if (m) return parseInt(m[1], 10);

    const nums = s.match(/\d+/g);
    if (nums && nums.length) return parseInt(nums[0], 10);

    return null;
  }

  function cleanCell(v) {
    if (v === null || v === undefined) return "";
    let s = String(v).trim();
    if (s.endsWith(".0") && /^\d+$/.test(s.slice(0, -2))) s = s.slice(0, -2);
    return s;
  }

  function detectSkuCol(header) {
    let c = findColByKeywords(header, ["mã", "sku", "đối", "tác"]);
    if (c < 0) c = findColByKeywords(header, ["ma", "sku", "doi", "tac"]);
    if (c < 0) c = findColByKeywords(header, ["partner", "sku"]);
    if (c < 0) c = findColByKeywords(header, null, ["mã sku đối tác", "sku đối tác", "sku partner", "partner sku"]);
    if (c < 0) c = findColByKeywords(header, null, ["sku"]);
    return c;
  }

  function detectUomCol(header) {
    let c = findColByKeywords(header, null, ["đvt", "dvt", "uom", "đơn vị", "don vi"]);
    if (c < 0) c = findColByKeywords(header, null, ["qty", "quantity"]);
    return c;
  }

  // ---- Public entry point ----
  async function generate(opts) {
    const log = opts.log || function () {};
    log("[INFO] Đang đọc file: " + opts.file.name);
    const wb = await WOPUtils.readWorkbookFromFile(opts.file);
    const sheetName = wb.SheetNames[0];
    const rows = WOPUtils.sheetToRows(wb, sheetName);
    if (!rows.length) throw new Error("File rỗng.");

    const headerRowIdx = findHeaderRow(rows, 40);
    const header = headerRowIdx !== null ? rows[headerRowIdx] : rows[0];
    const dataStart = headerRowIdx !== null ? headerRowIdx + 1 : 1;
    log("[INFO] Detected header row: " + (headerRowIdx === null ? "(fallback: dòng đầu tiên)" : headerRowIdx + 1));

    const skuCol = detectSkuCol(header);
    const uomCol = detectUomCol(header);

    if (skuCol < 0) throw new Error("Không tìm được cột SKU partner / Mã SKU đối tác. Header đọc được: " + header.map(function (h) { return h === null || h === undefined ? "" : h; }).join(" | "));
    if (uomCol < 0) throw new Error("Không tìm được cột ĐVT / UOM / QTY. Header đọc được: " + header.map(function (h) { return h === null || h === undefined ? "" : h; }).join(" | "));

    log("[INFO] Detected SKU column: " + header[skuCol]);
    log("[INFO] Detected UOM column: " + header[uomCol]);

    const seen = new Set();
    const out = [];
    let totalScanned = 0, droppedEmpty = 0, droppedBadQty = 0, droppedDup = 0;

    for (let r = dataStart; r < rows.length; r++) {
      const row = rows[r];
      if (!row) continue;
      totalScanned++;
      const sku = cleanCell(row[skuCol]);
      const uomRaw = cleanCell(row[uomCol]);
      if (!sku || sku.toLowerCase() === "nan" || !uomRaw || uomRaw.toLowerCase() === "nan") { droppedEmpty++; continue; }
      const qty = extractQtyFromInput(uomRaw);
      if (qty === null || qty <= 0) { droppedBadQty++; continue; }
      const key = sku + "||" + qty;
      if (seen.has(key)) { droppedDup++; continue; }
      seen.add(key);
      out.push({ sku: sku, qty: qty, uomLabel: qty + " - CARTON_" + qty + "PCS" });
    }

    log("[INFO] Tổng dòng quét: " + totalScanned);
    if (droppedEmpty) log("[INFO] Bỏ qua (thiếu SKU hoặc ĐVT): " + droppedEmpty);
    if (droppedBadQty) log("[WARNING] Bỏ qua (không đọc được số lượng/carton từ ĐVT): " + droppedBadQty);
    if (droppedDup) log("[INFO] Bỏ trùng (SKU + qty giống dòng trước): " + droppedDup);
    log("[INFO] Số dòng hợp lệ sau khi làm sạch: " + out.length);

    if (!out.length) throw new Error("Không có dòng nào hợp lệ sau khi làm sạch — kiểm tra lại file nguồn.");

    const aoa = [["Mã SKU đối tác", "ĐVT"]];
    out.forEach(function (o) { aoa.push([o.sku, o.uomLabel]); });

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const wbOut = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wbOut, ws, "AMI_UOM_Update");
    const wbOutArr = XLSX.write(wbOut, { bookType: "xlsx", type: "array" });
    const blob = new Blob([wbOutArr], { type: "application/octet-stream" });

    return {
      files: [{ name: "AMI_UOM_Update_Input.xlsx", blob: blob, count: out.length + " SKU" }],
      rows: out,
    };
  }

  window.WOPAmiUomPrep = {
    generate: generate,
    // exported for unit testing (Node) — không dùng trực tiếp trong UI
    _internal: {
      normText: normText, findHeaderRow: findHeaderRow, findColByKeywords: findColByKeywords,
      extractQtyFromInput: extractQtyFromInput, cleanCell: cleanCell,
      detectSkuCol: detectSkuCol, detectUomCol: detectUomCol,
    },
  };
})();
