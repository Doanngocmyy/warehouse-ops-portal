/* ==========================================================
 * utils.js — shared helpers for Warehouse Ops Portal
 * (Excel parsing, text cleaning, QR drawing, PDF label helpers)
 * ========================================================== */
(function () {
  const BIDI_AND_INVIS = [
    "​", "‌", "‍", "﻿", "­", " ",
    "‪", "‫", "‬", "‭", "‮",
    "⁦", "⁧", "⁨", "⁩",
  ];
  const DASH_VARIANTS = {
    "‐": "-", "‑": "-", "‒": "-", "–": "-",
    "—": "-", "−": "-", "﹣": "-", "－": "-",
  };

  function cleanText(s) {
    if (s === null || s === undefined) return "";
    s = String(s).replace(/\r/g, " ").replace(/\n/g, " ").replace(/\t/g, " ").trim();
    for (const k in DASH_VARIANTS) s = s.split(k).join(DASH_VARIANTS[k]);
    s = s.normalize("NFKC");
    for (const ch of BIDI_AND_INVIS) s = s.split(ch).join("");
    s = s.replace(/[\^~]/g, " ");
    return s.replace(/\s+/g, " ").trim();
  }

  function safeAscii(s) {
    s = cleanText(s).toUpperCase();
    s = s.replace(/[^A-Z0-9\-_/().\s]/g, "");
    return s.replace(/\s+/g, " ").trim();
  }

  function safeFilename(s, maxLen) {
    maxLen = maxLen || 160;
    s = safeAscii(s);
    s = s.replace(/[\\/:*?"<>|]+/g, "_");
    s = s.replace(/\s+/g, "_").replace(/^_+|_+$/g, "");
    return (s.slice(0, maxLen).replace(/_+$/, "") || "UNNAMED");
  }

  function isMissing(v) {
    if (v === null || v === undefined) return true;
    if (typeof v === "number" && Number.isNaN(v)) return true;
    if (typeof v === "string" && v.trim() === "") return true;
    return false;
  }

  function parseIntSafe(v, fieldName) {
    fieldName = fieldName || "";
    if (isMissing(v)) throw new Error(fieldName + " is missing");
    if (typeof v === "string") v = v.trim().replace(/,/g, "");
    const n = parseInt(Number(v), 10);
    if (Number.isNaN(n)) throw new Error(fieldName + " invalid number: " + v);
    return n;
  }

  function mixMarker(v) {
    return cleanText(v).toUpperCase() === "MIX";
  }

  function groupMarker(v) {
    if (isMissing(v)) return "";
    const s = cleanText(v).toUpperCase();
    const n = parseFloat(s);
    if (!Number.isNaN(n) && n === 1) return "1";
    if (s === "MIX") return "MIX";
    return "";
  }

  function distributeInnerCartons(cartonStart, cartonEnd, labelQty) {
    const outerCount = cartonEnd - cartonStart + 1;
    if (outerCount <= 0) throw new Error("Outer carton range is invalid");
    if (labelQty <= 0) throw new Error("Label quantity must be > 0");
    const base = Math.floor(labelQty / outerCount);
    const remainder = labelQty % outerCount;
    const result = [];
    for (let offset = 0; offset < outerCount; offset++) {
      const cartonNo = cartonStart + offset;
      const repeat = base + (offset < remainder ? 1 : 0);
      for (let k = 0; k < repeat; k++) result.push(cartonNo);
    }
    return result;
  }

  async function readWorkbookFromFile(file) {
    const buf = await file.arrayBuffer();
    return XLSX.read(buf, { type: "array", cellDates: false });
  }

  function sheetToRows(workbook, sheetName) {
    const ws = workbook.Sheets[sheetName];
    if (!ws) throw new Error('Sheet "' + sheetName + '" not found');
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 5000);
  }

  function csvFromRows(rows) {
    return rows.map(function (r) {
      return r.map(function (c) {
        return '"' + String(c === null || c === undefined ? "" : c).replace(/"/g, '""') + '"';
      }).join(",");
    }).join("\n");
  }

  // ---- QR code matrix (via qrcode-generator lib, global `qrcode`) ----
  function buildQrMatrix(text) {
    const qr = qrcode(0, "M");
    qr.addData(String(text));
    qr.make();
    const n = qr.getModuleCount();
    const modules = [];
    for (let r = 0; r < n; r++) {
      const row = [];
      for (let c = 0; c < n; c++) row.push(qr.isDark(r, c));
      modules.push(row);
    }
    return { n: n, modules: modules };
  }

  // Draw QR into a pdf-lib page inside box (x,y = bottom-left in PDF pts, size = box side length)
  function drawQr(page, text, x, y, size, darkColor) {
    const built = buildQrMatrix(text);
    const n = built.n, modules = built.modules;
    const cell = size / n;
    for (let r = 0; r < n; r++) {
      for (let c = 0; c < n; c++) {
        if (!modules[r][c]) continue;
        const px = x + c * cell;
        const py = y + (n - 1 - r) * cell;
        page.drawRectangle({ x: px, y: py, width: cell + 0.4, height: cell + 0.4, color: darkColor });
      }
    }
  }

  window.WOPUtils = {
    cleanText: cleanText, safeAscii: safeAscii, safeFilename: safeFilename,
    isMissing: isMissing, parseIntSafe: parseIntSafe,
    mixMarker: mixMarker, groupMarker: groupMarker, distributeInnerCartons: distributeInnerCartons,
    readWorkbookFromFile: readWorkbookFromFile, sheetToRows: sheetToRows,
    downloadBlob: downloadBlob, csvFromRows: csvFromRows, drawQr: drawQr,
  };
})();
