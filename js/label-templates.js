/* ==========================================================
 * label-templates.js — visual label drawing routines
 * Recreated to match real production samples (ASN_MIX_.pdf,
 * ASN_NO NEED.pdf, test mix label.pdf) at 432x288pt (6x4in).
 * ========================================================== */
(function () {
  const P = WOPPdf;
  const black = PDFLib.rgb(0, 0, 0);

  const MARGIN = 11;
  // column boundaries as fraction of interior width (matches sample proportions)
  const COL_A = 0.478; // EAN/#SKU column
  const COL_B = 0.132; // Qty column
  // COL_C = remainder (Barcode/QR column)

  function wrapLines(font, text, size, maxWidth, maxLines) {
    const words = String(text).split(" ").filter(Boolean);
    const lines = [];
    let cur = "";
    for (const w of words) {
      const test = cur ? cur + " " + w : w;
      if (font.widthOfTextAtSize(test, size) <= maxWidth || !cur) {
        cur = test;
      } else {
        lines.push(cur);
        cur = w;
      }
      if (lines.length >= maxLines - 1) break;
    }
    if (cur) lines.push(cur);
    // if still words left un-pushed (loop broke early), append remainder to last line
    return lines.slice(0, maxLines);
  }

  /* ---------- PER-EAN LABEL (Single + Mix-per-row) ---------- */
  // Matches ASN_MIX_.pdf / ASN_NO NEED.pdf exactly: TPLG header,
  // Carton No./Ref No. row, #EAN/Total Qty/Barcode header, value row.
  function drawPerEanLabel(page, font, opts) {
    const W = P.LABEL_W, H = P.LABEL_H;
    const iw = W - 2 * MARGIN, ih = H - 2 * MARGIN;
    const xA0 = MARGIN, xA1 = MARGIN + iw * COL_A;
    const xB1 = xA1 + iw * COL_B;
    const xC1 = MARGIN + iw;

    const bandHeader = 56, bandCarton = 56, bandColHeader = 46;
    const bandValue = ih - bandHeader - bandCarton - bandColHeader;

    const yTop = H - MARGIN;
    const yHeaderBot = yTop - bandHeader;
    const yCartonBot = yHeaderBot - bandCarton;
    const yColHeaderBot = yCartonBot - bandColHeader;
    const yValueBot = MARGIN;

    // outer border
    P.drawRect(page, MARGIN, MARGIN, iw, ih, 1.4);
    // header divider
    P.drawHLine(page, MARGIN, MARGIN + iw, yHeaderBot, 1.4);
    // TPLG
    P.drawLeftText(page, font, "TPLG", MARGIN + 12, (yTop + yHeaderBot) / 2, 30, black);

    // carton/ref row
    P.drawHLine(page, MARGIN, MARGIN + iw, yCartonBot, 1.4);
    P.drawVLine(page, xA1, yCartonBot, yHeaderBot, 1.2);
    P.drawVLine(page, xB1, yCartonBot, yHeaderBot, 1.2);
    const cartonMidY = (yHeaderBot + yCartonBot) / 2;
    const subDivX = xA0 + (xA1 - xA0) * 0.62;
    P.drawVLine(page, subDivX, yCartonBot, yHeaderBot, 1.0);
    P.drawCenteredText(page, font, "Carton No.", (xA0 + subDivX) / 2, cartonMidY, 13, black);
    P.drawCenteredText(page, font, String(opts.cartonNo), (subDivX + xA1) / 2, cartonMidY, 18, black);
    P.drawLeftText(page, font, "Ref No.", xA1 + 6, yHeaderBot - 12, 9, black);
    const refLines = wrapLines(font, opts.refNo || "", 7.5, xC1 - (xA1 + 6) - 4, 2);
    let ry = yHeaderBot - 26;
    for (const ln of refLines) {
      page.drawText(ln, { x: xA1 + 6, y: ry, size: 7.5, font: font, color: black });
      ry -= 10;
    }

    // column header row
    P.drawHLine(page, MARGIN, MARGIN + iw, yColHeaderBot, 1.4);
    P.drawVLine(page, xA1, yColHeaderBot, yCartonBot, 1.2);
    P.drawVLine(page, xB1, yColHeaderBot, yCartonBot, 1.2);
    const colHeaderMidY = (yCartonBot + yColHeaderBot) / 2;
    P.drawCenteredText(page, font, opts.itemLabel || "#EAN", (xA0 + xA1) / 2, colHeaderMidY, 13, black);
    P.drawCenteredText(page, font, "Total Qty", (xA1 + xB1) / 2, colHeaderMidY, 12, black);
    P.drawCenteredText(page, font, "Barcode", (xB1 + xC1) / 2, colHeaderMidY, 13, black);

    // value row
    P.drawVLine(page, xA1, yValueBot, yColHeaderBot, 1.2);
    P.drawVLine(page, xB1, yValueBot, yColHeaderBot, 1.2);
    const valueMidY = (yColHeaderBot + yValueBot) / 2;
    const eanSize = P.fitText(font, String(opts.ean), (xA1 - xA0) - 16, 22, 9);
    P.drawLeftText(page, font, String(opts.ean), xA0 + 8, valueMidY, eanSize, black);
    P.drawCenteredText(page, font, String(opts.qty), (xA1 + xB1) / 2, valueMidY, 20, black);

    const qrBoxSize = Math.min(xC1 - xB1 - 16, bandValue - 16);
    const qrX = xB1 + ((xC1 - xB1) - qrBoxSize) / 2;
    const qrY = valueMidY - qrBoxSize / 2;
    WOPUtils.drawQr(page, String(opts.qrText != null ? opts.qrText : opts.ean), qrX, qrY, qrBoxSize, black);
  }

  /* ---------- MIX CARTON SUMMARY LABEL (multi-SKU) ---------- */
  // Matches "test mix label.pdf": header, Carton No/Ref row (single line),
  // #SKU/Total Qty/Barcode header, N sku rows (barcode col = N/A),
  // Carton ID/Batch# + CTN Qty + QR footer row.
  // pageW/pageH allow the "grow to A4" mode; fontScale allows the "shrink" mode.
  function drawMixSummaryLabel(page, font, opts) {
    const W = opts.pageW || P.LABEL_W, H = opts.pageH || P.LABEL_H;
    const scale = opts.fontScale || 1;
    const iw = W - 2 * MARGIN, ih = H - 2 * MARGIN;
    const xA0 = MARGIN, xA1 = MARGIN + iw * COL_A;
    const xB1 = xA1 + iw * COL_B;
    const xC1 = MARGIN + iw;

    const bandHeader = 42 * scale;
    const bandCarton = 34 * scale;
    const bandColHeader = 30 * scale;
    const bandFooterHeader = 26 * scale;
    const rowH = (opts.rowHeight || 30) * scale;
    const nRows = opts.rows.length;
    const bandRows = rowH * nRows;
    const bandFooterValue = Math.max(ih - bandHeader - bandCarton - bandColHeader - bandRows - bandFooterHeader, 60);

    const yTop = H - MARGIN;
    const yHeaderBot = yTop - bandHeader;
    const yCartonBot = yHeaderBot - bandCarton;
    const yColHeaderBot = yCartonBot - bandColHeader;
    const yRowsBot = yColHeaderBot - bandRows;
    const yFooterHeaderBot = yRowsBot - bandFooterHeader;
    const yFooterValueBot = MARGIN;

    P.drawRect(page, MARGIN, MARGIN, iw, ih, 1.4);
    P.drawHLine(page, MARGIN, MARGIN + iw, yHeaderBot, 1.4);
    P.drawLeftText(page, font, "TPLG", MARGIN + 12, (yTop + yHeaderBot) / 2, 26 * scale, black);

    // carton/ref row (single line style, like test mix label.pdf)
    P.drawHLine(page, MARGIN, MARGIN + iw, yCartonBot, 1.4);
    P.drawVLine(page, xB1, yCartonBot, yHeaderBot, 1.2);
    const cartonMidY = (yHeaderBot + yCartonBot) / 2;
    P.drawCenteredText(page, font, "Carton No.", (xA0 + xB1) / 2, cartonMidY, 13 * scale, black);
    const refSize = 9 * scale;
    P.drawLeftText(page, font, "Ref No.  " + (opts.refNo || ""), xB1 + 8, cartonMidY, refSize, black);
    P.drawCenteredText(page, font, String(opts.cartonNo), (xA0 + xB1) * 0.72, cartonMidY, 16 * scale, black);

    // column header
    P.drawHLine(page, MARGIN, MARGIN + iw, yColHeaderBot, 1.4);
    P.drawVLine(page, xA1, yColHeaderBot, yCartonBot, 1.2);
    P.drawVLine(page, xB1, yColHeaderBot, yCartonBot, 1.2);
    const colHeadY = (yCartonBot + yColHeaderBot) / 2;
    P.drawCenteredText(page, font, "#SKU", (xA0 + xA1) / 2, colHeadY, 13 * scale, black);
    P.drawCenteredText(page, font, "Total Qty", (xA1 + xB1) / 2, colHeadY, 11 * scale, black);
    P.drawCenteredText(page, font, "Barcode", (xB1 + xC1) / 2, colHeadY, 12 * scale, black);

    // sku rows
    let rowTop = yColHeaderBot;
    for (let i = 0; i < nRows; i++) {
      const rowBot = rowTop - rowH;
      P.drawHLine(page, MARGIN, MARGIN + iw, rowBot, 1.0);
      P.drawVLine(page, xA1, rowBot, rowTop, 1.0);
      P.drawVLine(page, xB1, rowBot, rowTop, 1.0);
      const midY = (rowTop + rowBot) / 2;
      const sku = opts.rows[i].sku;
      const skuSize = P.fitText(font, sku, (xA1 - xA0) - 14, 13 * scale, 6);
      P.drawLeftText(page, font, sku, xA0 + 8, midY, skuSize, black);
      P.drawCenteredText(page, font, String(opts.rows[i].qty), (xA1 + xB1) / 2, midY, 13 * scale, black);
      P.drawCenteredText(page, font, "N/A", (xB1 + xC1) / 2, midY, 12 * scale, black);
      rowTop = rowBot;
    }

    // footer header
    P.drawHLine(page, MARGIN, MARGIN + iw, yFooterHeaderBot, 1.4);
    P.drawVLine(page, xA1, yFooterHeaderBot, yRowsBot, 1.2);
    P.drawVLine(page, xB1, yFooterHeaderBot, yRowsBot, 1.2);
    const fhMidY = (yRowsBot + yFooterHeaderBot) / 2;
    P.drawCenteredText(page, font, "Carton ID / Batch #", (xA0 + xA1) / 2, fhMidY, 12 * scale, black);
    P.drawCenteredText(page, font, "CTN Qty", (xA1 + xB1) / 2, fhMidY, 11 * scale, black);
    P.drawCenteredText(page, font, "Barcode", (xB1 + xC1) / 2, fhMidY, 12 * scale, black);

    // footer value
    P.drawVLine(page, xA1, yFooterValueBot, yFooterHeaderBot, 1.2);
    P.drawVLine(page, xB1, yFooterValueBot, yFooterHeaderBot, 1.2);
    const fvMidY = (yFooterHeaderBot + yFooterValueBot) / 2;
    const cidSize = P.fitText(font, opts.cartonId, (xA1 - xA0) - 14, 15 * scale, 7);
    P.drawCenteredText(page, font, opts.cartonId, (xA0 + xA1) / 2, fvMidY, cidSize, black);
    P.drawCenteredText(page, font, String(opts.ctnQty), (xA1 + xB1) / 2, fvMidY, 16 * scale, black);

    const qrBoxSize = Math.min(xC1 - xB1 - 16, yFooterHeaderBot - yFooterValueBot - 16);
    const qrX = xB1 + ((xC1 - xB1) - qrBoxSize) / 2;
    const qrY = fvMidY - qrBoxSize / 2;
    WOPUtils.drawQr(page, String(opts.cartonId), qrX, qrY, qrBoxSize, black);
  }

  window.WOPLabels = { drawPerEanLabel: drawPerEanLabel, drawMixSummaryLabel: drawMixSummaryLabel, wrapLines: wrapLines };
})();
