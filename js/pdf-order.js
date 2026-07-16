/* ==========================================================
 * pdf-order.js — merge one or more label PDFs (e.g. ASN Single +
 * ASN Mix output, or any label PDFs) into a single PDF, with every
 * page re-sorted ascending by Carton No. — read directly from each
 * page's text ("Carton No. <n>") via pdf.js — and the same
 * missing/duplicate sequence-check warnings used by the ASN
 * generators. Port of "PDF in order.ipynb".
 * ========================================================== */
(function () {
  const U = WOPUtils;
  const CARTON_REGEX = /Carton\s*No\.?\s*(\d+)/i;

  async function extractPageText(pdfjsDoc, pageNum) {
    const page = await pdfjsDoc.getPage(pageNum);
    const content = await page.getTextContent();
    return content.items.map(function (it) { return it.str; }).join(" ");
  }

  async function generate({ files, log }) {
    if (!files || !files.length) throw new Error("Chưa chọn file PDF nào.");
    log("[INFO] Số file PDF đầu vào: " + files.length);

    const srcPdfLibDocs = [];
    const pagesWithCarton = []; // { cartonNo, fileIdx, pageIdx, fileName }
    const noMatchPages = [];

    for (let fi = 0; fi < files.length; fi++) {
      const file = files[fi];
      const buf = await file.arrayBuffer();

      const pdfLibDoc = await PDFLib.PDFDocument.load(buf);
      srcPdfLibDocs.push(pdfLibDoc);

      const pdfjsDoc = await pdfjsLib.getDocument({ data: buf.slice(0) }).promise;
      const nPages = pdfjsDoc.numPages;
      log("[INFO] \"" + file.name + "\": " + nPages + " trang");

      for (let p = 1; p <= nPages; p++) {
        const text = await extractPageText(pdfjsDoc, p);
        const m = CARTON_REGEX.exec(text);
        if (!m) {
          noMatchPages.push({ fileName: file.name, page: p });
          continue;
        }
        pagesWithCarton.push({ cartonNo: parseInt(m[1], 10), fileIdx: fi, pageIdx: p - 1, fileName: file.name });
      }
    }

    if (noMatchPages.length) {
      log("[WARNING] Không tìm thấy \"Carton No.\" trên " + noMatchPages.length + " trang:");
      noMatchPages.slice(0, 30).forEach(function (x) {
        log("   - " + x.fileName + " — trang " + x.page);
      });
      if (noMatchPages.length > 30) log("   ... (+" + (noMatchPages.length - 30) + " more)");
    }

    log("[INFO] Tổng số trang có Carton No.: " + pagesWithCarton.length);
    if (!pagesWithCarton.length) throw new Error("Không tìm thấy Carton No. nào trong (các) PDF đã chọn.");

    // Carton No. sequence check — groupId = source file name, so a carton
    // number split across two different source files (or two different
    // spots in the same file) is flagged just like a duplicate.
    const assignments = pagesWithCarton.map(function (p) { return { cartonNo: p.cartonNo, groupId: p.fileName }; });
    U.checkCartonSequence(assignments, log, "Carton No.");

    // Sort ALL pages ascending by Carton No. (top -> bottom of the merged PDF).
    pagesWithCarton.sort(function (a, b) { return a.cartonNo - b.cartonNo; });

    const outDoc = await PDFLib.PDFDocument.create();
    for (const item of pagesWithCarton) {
      const [copied] = await outDoc.copyPages(srcPdfLibDocs[item.fileIdx], [item.pageIdx]);
      outDoc.addPage(copied);
    }

    const bytes = await outDoc.save();
    const outFiles = [{ name: "ASN_Labels_Sorted.pdf", blob: new Blob([bytes], { type: "application/pdf" }), count: pagesWithCarton.length }];

    if (noMatchPages.length) {
      const rowsCsv = [["file", "page"]];
      for (const x of noMatchPages) rowsCsv.push([x.fileName, x.page]);
      outFiles.push({ name: "PDF_In_Order_No_Carton_Match.csv", blob: new Blob([U.csvFromRows(rowsCsv)], { type: "text/csv;charset=utf-8" }), count: noMatchPages.length });
    }

    return { files: outFiles, stats: { inputFiles: files.length, totalPages: pagesWithCarton.length, noMatch: noMatchPages.length } };
  }

  window.WOPPdfOrder = { generate: generate };
})();
