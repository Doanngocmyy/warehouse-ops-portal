/* ==========================================================
 * asn-single.js — port of "Single carton_Duplicate_New logic.ipynb"
 * Sheet "Single" fixed columns:
 *   B=SKU  C=EAN  H=Qty/box  I=Carton qty/label qty
 *   J=Carton start  K=Carton end  L=MIX marker
 * MIX marker: following blank-carton rows share the same carton
 * group, but every row still becomes its own per-EAN single label.
 * ========================================================== */
(function () {
  const U = WOPUtils;

  function isMissing(v) { return U.isMissing(v); }
  function hasCartonNo(row) { return !isMissing(row[9]) || !isMissing(row[10]); }
  function rowHasItemData(row) { return !isMissing(row[1]) || !isMissing(row[2]); }

  function buildRecords(row, excelRowNo, cartonStart, cartonEnd, labelQty, refNo, groupSort, skipped) {
    const records = [];
    const sku = U.safeAscii(row[1]);
    const ean = U.safeAscii(row[2]);
    if (!ean) {
      skipped.push({ excel_row: excelRowNo, reason: "EAN becomes empty after cleaning", sku_raw: row[1], ean_raw: row[2] });
      return records;
    }
    let qty, expanded;
    try {
      qty = U.parseIntSafe(row[7], "QTY/BOX (col H)");
      expanded = U.distributeInnerCartons(cartonStart, cartonEnd, labelQty);
    } catch (e) {
      skipped.push({ excel_row: excelRowNo, reason: "Invalid line: " + e.message, sku_raw: row[1], ean_raw: row[2] });
      return records;
    }
    expanded.forEach(function (cartonNo, idx) {
      records.push({ groupSort: groupSort, excelRow: excelRowNo, sku: sku, ean: ean, qty: qty, cartonNo: cartonNo, innerSeq: idx + 1 });
    });
    return records;
  }

  async function generate({ rows, refNoOverride, log }) {
    if (!rows || rows.length < 6) throw new Error("Sheet 'Single' trống hoặc thiếu dữ liệu.");
    const refCell = rows[0] ? rows[0][0] : null;
    const REF_NO = refNoOverride || (isMissing(refCell) ? "" : U.cleanText(refCell));
    log("[INFO] Ref No: " + REF_NO);

    const records = [];
    const skipped = [];
    let i = 5, normalGroups = 0, mixGroups = 0, rowsScanned = 0;

    while (i < rows.length) {
      const row = rows[i] || [];
      const excelRowNo = i + 1;
      rowsScanned++;

      if (!rowHasItemData(row)) { i++; continue; }

      if (U.mixMarker(row[11])) {
        const groupStartRow = excelRowNo;
        mixGroups++;
        let groupStart, groupEnd, groupLabelQty;
        try {
          groupStart = U.parseIntSafe(row[9], "MIX group carton start (col J)");
          groupEnd = isMissing(row[10]) ? groupStart : U.parseIntSafe(row[10], "MIX group carton end (col K)");
          groupLabelQty = isMissing(row[8]) ? (groupEnd - groupStart + 1) : U.parseIntSafe(row[8], "MIX group label qty (col I)");
        } catch (e) {
          skipped.push({ excel_row: groupStartRow, reason: "Invalid MIX group header: " + e.message });
          i++; continue;
        }
        const groupRows = [];
        let j = i;
        while (j < rows.length) {
          const r = rows[j] || [];
          if (j !== i) {
            if (U.mixMarker(r[11])) break;
            if (hasCartonNo(r)) break;
            if (!rowHasItemData(r)) break;
          }
          groupRows.push([j + 1, r]);
          j++;
        }
        for (const [lineRowNo, r] of groupRows) {
          const recs = buildRecords(r, lineRowNo, groupStart, groupEnd, groupLabelQty, REF_NO, groupStartRow, skipped);
          records.push(...recs);
        }
        i = j;
        continue;
      }

      try {
        if (isMissing(row[2])) throw new Error("EAN (col C) is missing");
        if (isMissing(row[7])) throw new Error("QTY/BOX (col H) is missing");
        if (isMissing(row[8])) throw new Error("CARTON QTY / LABEL QTY (col I) is missing");
        if (isMissing(row[9])) throw new Error("CARTON START (col J) is missing");
        const cartonStart = U.parseIntSafe(row[9], "CARTON START (col J)");
        const cartonEnd = isMissing(row[10]) ? cartonStart : U.parseIntSafe(row[10], "CARTON END (col K)");
        const labelQty = U.parseIntSafe(row[8], "CARTON QTY / LABEL QTY (col I)");
        normalGroups++;
        const recs = buildRecords(row, excelRowNo, cartonStart, cartonEnd, labelQty, REF_NO, excelRowNo, skipped);
        records.push(...recs);
      } catch (e) {
        skipped.push({ excel_row: excelRowNo, reason: "Skipped normal row: " + e.message, sku_raw: row[1], ean_raw: row[2] });
      }
      i++;
    }

    records.sort((a, b) => (a.groupSort - b.groupSort) || (a.cartonNo - b.cartonNo) || (a.excelRow - b.excelRow) || (a.innerSeq - b.innerSeq));

    log("[INFO] Rows scanned: " + rowsScanned + " | Normal: " + normalGroups + " | MIX groups: " + mixGroups);
    log("[INFO] Total labels: " + records.length + " | Skipped: " + skipped.length);
    if (!records.length) throw new Error("Không có label nào được tạo — kiểm tra lại dữ liệu Excel.");

    const { pdfDoc, font } = await WOPPdf.createDoc();
    for (const rec of records) {
      const page = pdfDoc.addPage([WOPPdf.LABEL_W, WOPPdf.LABEL_H]);
      WOPLabels.drawPerEanLabel(page, font, { cartonNo: rec.cartonNo, refNo: REF_NO, ean: rec.ean, qty: rec.qty, itemLabel: "#EAN" });
    }
    const bytes = await pdfDoc.save();
    const files = [{ name: "ASN_Single_Labels.pdf", blob: new Blob([bytes], { type: "application/pdf" }), count: records.length }];

    if (skipped.length) {
      const rowsCsv = [["excel_row", "reason", "sku_raw", "ean_raw"]];
      for (const s of skipped) rowsCsv.push([s.excel_row, s.reason, s.sku_raw || "", s.ean_raw || ""]);
      files.push({ name: "ASN_Single_Skipped_Rows.csv", blob: new Blob([U.csvFromRows(rowsCsv)], { type: "text/csv;charset=utf-8" }), count: skipped.length });
    }

    return { files, stats: { rowsScanned, normalGroups, mixGroups, labels: records.length, skipped: skipped.length } };
  }

  window.WOPAsnSingle = { generate: generate };
})();
