/* ==========================================================
 * asn-mix.js — port of "Label Automation - Mix Case.ipynb"
 * Sheet "Mix" fixed columns:
 *   C=SKU/EAN  H=Qty  I=group marker (1 or "MIX")
 *   J=Carton start  K=Carton end  P=Carton ID base
 * Output: 1 combined "carton summary" label per carton, listing
 * every SKU packed inside it (Carton ID = colP + "-" + totalQty).
 *
 * NEW: auto shrink font to fit standard 6x4in label up to
 * SHRINK_MAX_ROWS SKU rows; beyond that, auto-switch the page to
 * A4 at normal font size instead of shrinking further.
 * ========================================================== */
(function () {
  const U = WOPUtils;

  const DEFAULT_SHRINK_MAX_ROWS = 5;
  const MIN_FONT_SCALE = 0.55;
  const BASE_ROW_H = 30;
  const BASE_BAND_FIXED = 42 + 34 + 30 + 26 + 60; // header+carton+colHeader+footerHeader+footerValueMin (scale=1)
  const USABLE_H = WOPPdf.LABEL_H - 22; // minus margins*2

  function isMissing(v) { return U.isMissing(v); }

  function computeLayout(nRows, shrinkMaxRows) {
    const neededAtScale1 = BASE_BAND_FIXED + BASE_ROW_H * nRows;
    if (nRows <= shrinkMaxRows && neededAtScale1 <= USABLE_H) {
      return { pageW: WOPPdf.LABEL_W, pageH: WOPPdf.LABEL_H, fontScale: 1, rowHeight: BASE_ROW_H, mode: "normal" };
    }
    if (nRows <= shrinkMaxRows) {
      // shrink to fit standard label
      let scale = USABLE_H / neededAtScale1;
      scale = Math.max(scale, MIN_FONT_SCALE);
      return { pageW: WOPPdf.LABEL_W, pageH: WOPPdf.LABEL_H, fontScale: scale, rowHeight: BASE_ROW_H, mode: "shrink" };
    }
    // too many SKUs for the small label even after shrinking -> grow to A4
    const usableA4 = WOPPdf.A4_H - 22;
    const neededA4 = BASE_BAND_FIXED + BASE_ROW_H * nRows;
    let scale = 1;
    if (neededA4 > usableA4) scale = Math.max(usableA4 / neededA4, MIN_FONT_SCALE);
    return { pageW: WOPPdf.A4_W, pageH: WOPPdf.A4_H, fontScale: scale, rowHeight: BASE_ROW_H, mode: "grow_a4" };
  }

  async function generate({ rows, refNoOverride, shrinkMaxRows, log }) {
    if (!rows || rows.length < 6) throw new Error("Sheet 'Mix' trống hoặc thiếu dữ liệu.");
    shrinkMaxRows = shrinkMaxRows || DEFAULT_SHRINK_MAX_ROWS;
    const refCell = rows[0] ? rows[0][0] : null;
    const REF_NO = refNoOverride || (isMissing(refCell) ? "" : U.cleanText(refCell));
    log("[INFO] Ref No: " + REF_NO + " | Shrink threshold: " + shrinkMaxRows + " SKU rows");

    const skipped = [];
    const groups = [];
    let current = null;
    const startIdx = 4;

    function isGroupStart(row) { return U.groupMarker(row[8]) !== ""; }

    for (let i = startIdx; i < rows.length; i++) {
      const row = rows[i] || [];
      const excelRowNo = i + 1;
      const skuCell = row[2], qtyCell = row[7], markerI = row[8];
      const cartonFromCell = row[9], cartonToCell = row[10], cartonIdCell = row[15];

      if (isGroupStart(row)) {
        if (current) groups.push(current);
        current = { startRow: excelRowNo, cartonFromRaw: cartonFromCell, cartonToRaw: cartonToCell, cartonIdRaw: cartonIdCell, lines: [] };
      }

      if (!current) {
        if (!isMissing(skuCell) || !isMissing(qtyCell)) {
          skipped.push({ excel_row: excelRowNo, reason: "Orphan SKU row (no active group start above) — need I=1/MIX marker", sku_raw: skuCell, qty_raw: qtyCell });
        }
        continue;
      }
      if (isMissing(skuCell) && isMissing(qtyCell)) continue;
      if (isMissing(skuCell) || isMissing(qtyCell)) {
        skipped.push({ excel_row: excelRowNo, reason: "Inside group but missing SKU or QTY", sku_raw: skuCell, qty_raw: qtyCell });
        continue;
      }
      current.lines.push([excelRowNo, skuCell, qtyCell]);
    }
    if (current) groups.push(current);
    log("[INFO] Detected groups: " + groups.length);

    const { pdfDoc, font } = await WOPPdf.createDoc();
    let totalLabels = 0;
    const layoutCounts = { normal: 0, shrink: 0, grow_a4: 0 };

    for (const g of groups) {
      let cartonFrom, cartonTo;
      try {
        cartonFrom = U.parseIntSafe(g.cartonFromRaw, "Carton start (J)");
        cartonTo = U.parseIntSafe(g.cartonToRaw, "Carton end (K)");
      } catch (e) {
        skipped.push({ excel_row: g.startRow, reason: "Group skipped: cannot parse carton range J/K: " + e.message });
        continue;
      }
      if (cartonTo < cartonFrom) {
        skipped.push({ excel_row: g.startRow, reason: "Group skipped: invalid carton range (K < J)" });
        continue;
      }
      if (!g.lines.length) {
        skipped.push({ excel_row: g.startRow, reason: "Group skipped: no SKU lines found under this group" });
        continue;
      }
      const skuLines = [];
      for (const [rowNo, skuRaw, qtyRaw] of g.lines) {
        const sku = U.safeAscii(skuRaw);
        if (!sku) { skipped.push({ excel_row: rowNo, reason: "SKU becomes empty after cleaning", sku_raw: skuRaw }); continue; }
        let qty;
        try { qty = U.parseIntSafe(qtyRaw, "Qty (H)"); } catch (e) {
          skipped.push({ excel_row: rowNo, reason: "Cannot parse qty (H): " + e.message, sku_raw: skuRaw, qty_raw: qtyRaw });
          continue;
        }
        skuLines.push({ sku, qty });
      }
      if (!skuLines.length) {
        skipped.push({ excel_row: g.startRow, reason: "Group skipped: all SKU lines invalid after cleaning/parsing" });
        continue;
      }
      const totalQty = skuLines.reduce((s, r) => s + r.qty, 0);
      const ctnQtyVal = cartonTo - cartonFrom + 1;
      const rawCartonId = g.cartonIdRaw;
      if (isMissing(rawCartonId) || !U.cleanText(rawCartonId)) {
        skipped.push({ excel_row: g.startRow, reason: "Group skipped: missing Carton ID in column P at group start row (I=1/MIX)" });
        continue;
      }
      const cartonIdBase = U.cleanText(rawCartonId);
      const cartonIdGenerated = cartonIdBase + "-" + totalQty;

      const layout = computeLayout(skuLines.length, shrinkMaxRows);
      layoutCounts[layout.mode]++;

      for (let cartonNo = cartonFrom; cartonNo <= cartonTo; cartonNo++) {
        const page = pdfDoc.addPage([layout.pageW, layout.pageH]);
        WOPLabels.drawMixSummaryLabel(page, font, {
          pageW: layout.pageW, pageH: layout.pageH, fontScale: layout.fontScale, rowHeight: layout.rowHeight,
          cartonNo, refNo: REF_NO, rows: skuLines, cartonId: cartonIdGenerated, ctnQty: ctnQtyVal,
        });
        totalLabels++;
      }
    }

    log("[INFO] Layout used — normal: " + layoutCounts.normal + " | shrink: " + layoutCounts.shrink + " | grow-to-A4: " + layoutCounts.grow_a4);
    log("[INFO] Total labels: " + totalLabels + " | Skipped/warnings: " + skipped.length);
    if (!totalLabels) throw new Error("Không có label nào được tạo — kiểm tra lại marker I=1/MIX, range J/K, và Carton ID cột P.");

    const bytes = await pdfDoc.save();
    const files = [{ name: "ASN_Mix_Carton_Summary_Labels.pdf", blob: new Blob([bytes], { type: "application/pdf" }), count: totalLabels }];

    if (skipped.length) {
      const rowsCsv = [["excel_row", "reason", "sku_raw", "qty_raw"]];
      for (const s of skipped) rowsCsv.push([s.excel_row, s.reason, s.sku_raw || "", s.qty_raw || ""]);
      files.push({ name: "ASN_Mix_Skipped_Rows.csv", blob: new Blob([U.csvFromRows(rowsCsv)], { type: "text/csv;charset=utf-8" }), count: skipped.length });
    }

    return { files, stats: { groups: groups.length, labels: totalLabels, skipped: skipped.length, layoutCounts } };
  }

  window.WOPAsnMix = { generate: generate, computeLayout: computeLayout, DEFAULT_SHRINK_MAX_ROWS: DEFAULT_SHRINK_MAX_ROWS };
})();
