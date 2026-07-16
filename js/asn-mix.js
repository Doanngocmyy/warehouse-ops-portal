/* ==========================================================
 * asn-mix.js — port of "Label Automation - Mix Case.ipynb"
 * Sheet "Mix" fixed columns:
 *   C=SKU/EAN  H=Qty  I=group marker (1 or "MIX")
 *   J=Carton start  K=Carton end  P=Carton ID base
 * Output: 1 combined "carton summary" label per carton, listing
 * every SKU packed inside it (Carton ID = colP + "-" + totalQty).
 *
 * Auto shrink font to fit standard 6x4in label up to
 * SHRINK_MAX_ROWS SKU rows; beyond that, grow the label's HEIGHT
 * only as much as needed (width stays 6in, no A4 sheet switch) up to
 * MAX_LABEL_H, falling back to font-shrink only if still too tall.
 *
 * Output pages are sorted ascending by Carton No. (top -> bottom of
 * the PDF) across ALL groups, and a Carton No. sequence check logs
 * [WARNING] lines for any missing or duplicate/overlapping carton
 * numbers before export (port of "PDF in order.ipynb").
 * ========================================================== */
(function () {
  const U = WOPUtils;

  const DEFAULT_SHRINK_MAX_ROWS = 5;
  const MIN_FONT_SCALE = 0.55;
  const BASE_ROW_H = 30;
  const BASE_BAND_FIXED = 42 + 34 + 30 + 26 + 60; // header+carton+colHeader+footerHeader+footerValueMin (scale=1)
  const USABLE_H = WOPPdf.LABEL_H - 22; // minus margins*2
  const MAX_USABLE_H = WOPPdf.MAX_LABEL_H - 22;

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
    // too many SKUs for the standard label -> grow label HEIGHT only,
    // just enough to fit at full font size (no A4 switch, no shrink yet)
    if (neededAtScale1 <= MAX_USABLE_H) {
      const pageH = neededAtScale1 + 22;
      return { pageW: WOPPdf.LABEL_W, pageH: pageH, fontScale: 1, rowHeight: BASE_ROW_H, mode: "grow_label" };
    }
    // even the max label height isn't enough -> use max height + shrink font
    const scale = Math.max(MAX_USABLE_H / neededAtScale1, MIN_FONT_SCALE);
    return { pageW: WOPPdf.LABEL_W, pageH: WOPPdf.MAX_LABEL_H, fontScale: scale, rowHeight: BASE_ROW_H, mode: "grow_label_shrink" };
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

    let totalLabels = 0;
    const layoutCounts = { normal: 0, shrink: 0, grow_label: 0, grow_label_shrink: 0 };
    // Collect label draw-jobs first (instead of writing pages immediately)
    // so the whole document can be re-sorted by Carton No. before export.
    const pendingLabels = [];
    const cartonAssignments = [];

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
        cartonAssignments.push({ cartonNo: cartonNo, groupId: g.startRow });
        pendingLabels.push({
          cartonNo: cartonNo, pageW: layout.pageW, pageH: layout.pageH,
          fontScale: layout.fontScale, rowHeight: layout.rowHeight,
          rows: skuLines, cartonId: cartonIdGenerated, ctnQty: ctnQtyVal,
        });
        totalLabels++;
      }
    }

    log("[INFO] Layout used — normal: " + layoutCounts.normal + " | shrink: " + layoutCounts.shrink + " | grow-label: " + layoutCounts.grow_label + " | grow-label+shrink: " + layoutCounts.grow_label_shrink);
    log("[INFO] Total labels: " + totalLabels + " | Skipped/warnings: " + skipped.length);
    if (!totalLabels) throw new Error("Không có label nào được tạo — kiểm tra lại marker I=1/MIX, range J/K, và Carton ID cột P.");

    // Carton No. sequence check — warn on missing numbers, and on the same
    // physical carton number being claimed by two different groups/rows.
    U.checkCartonSequence(cartonAssignments, log, "Carton No.");

    // Sort ALL pages ascending by Carton No. (top -> bottom of the PDF),
    // regardless of which group/Excel row they came from.
    pendingLabels.sort((a, b) => a.cartonNo - b.cartonNo);

    const { pdfDoc, font } = await WOPPdf.createDoc();
    for (const item of pendingLabels) {
      const page = pdfDoc.addPage([item.pageW, item.pageH]);
      WOPLabels.drawMixSummaryLabel(page, font, {
        pageW: item.pageW, pageH: item.pageH, fontScale: item.fontScale, rowHeight: item.rowHeight,
        cartonNo: item.cartonNo, refNo: REF_NO, rows: item.rows, cartonId: item.cartonId, ctnQty: item.ctnQty,
      });
    }

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
