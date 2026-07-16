/* ==========================================================
 * pdf-engine.js — shared pdf-lib label canvas helpers
 * Page size: 432 x 288 pt (6x4 in) for standard labels,
 * or A4 (595 x 842 pt) portrait for the "grow" mix summary layout.
 * All text uses built-in Helvetica-Bold (ASCII-only content).
 * ========================================================== */
(function () {
  const LABEL_W = 432, LABEL_H = 288;
  const A4_W = 595.28, A4_H = 841.89;

  async function createDoc() {
    const pdfDoc = await PDFLib.PDFDocument.create();
    const font = await pdfDoc.embedFont(PDFLib.StandardFonts.HelveticaBold);
    return { pdfDoc, font };
  }

  function fitText(font, text, maxWidth, startSize, minSize) {
    let size = startSize;
    while (size > minSize && font.widthOfTextAtSize(text, size) > maxWidth) size -= 0.5;
    return Math.max(size, minSize);
  }

  function drawCenteredText(page, font, text, cx, cy, size, color) {
    const w = font.widthOfTextAtSize(text, size);
    page.drawText(text, { x: cx - w / 2, y: cy - size * 0.35, size: size, font: font, color: color || PDFLib.rgb(0, 0, 0) });
  }

  function drawLeftText(page, font, text, x, cy, size, color) {
    page.drawText(text, { x: x, y: cy - size * 0.35, size: size, font: font, color: color || PDFLib.rgb(0, 0, 0) });
  }

  function drawRect(page, x, y, w, h, lineWidth) {
    page.drawRectangle({ x: x, y: y, width: w, height: h, borderColor: PDFLib.rgb(0, 0, 0), borderWidth: lineWidth || 1.2 });
  }

  function drawHLine(page, x1, x2, y, lineWidth) {
    page.drawLine({ start: { x: x1, y: y }, end: { x: x2, y: y }, thickness: lineWidth || 1.2, color: PDFLib.rgb(0, 0, 0) });
  }

  function drawVLine(page, x, y1, y2, lineWidth) {
    page.drawLine({ start: { x: x, y: y1 }, end: { x: x, y: y2 }, thickness: lineWidth || 1.2, color: PDFLib.rgb(0, 0, 0) });
  }

  window.WOPPdf = {
    LABEL_W: LABEL_W, LABEL_H: LABEL_H, A4_W: A4_W, A4_H: A4_H,
    createDoc: createDoc, fitText: fitText,
    drawCenteredText: drawCenteredText, drawLeftText: drawLeftText,
    drawRect: drawRect, drawHLine: drawHLine, drawVLine: drawVLine,
  };
})();
