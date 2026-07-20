/* ==========================================================
 * app.js — nav wiring + ASN upload/generate orchestration
 * ========================================================== */
(function () {
  // ---- top nav ----
  document.querySelectorAll("nav.tabs button[data-tab]").forEach(function (btn) {
    btn.addEventListener("click", function () {
      document.querySelectorAll("nav.tabs button[data-tab]").forEach(b => b.classList.remove("active"));
      document.querySelectorAll(".section").forEach(s => s.classList.remove("active"));
      btn.classList.add("active");
      document.getElementById("section-" + btn.getAttribute("data-tab")).classList.add("active");
    });
  });
  // ---- pre-inbound sub tabs ----
  document.querySelectorAll(".subtabs button[data-sub]").forEach(function (btn) {
    btn.addEventListener("click", function () {
      const parent = btn.closest(".card");
      parent.querySelectorAll(".subtabs button[data-sub]").forEach(b => b.classList.remove("active"));
      parent.querySelectorAll(".subsection").forEach(s => s.classList.remove("active"));
      btn.classList.add("active");
      parent.querySelector("#sub-" + btn.getAttribute("data-sub")).classList.add("active");
    });
  });

  // ---- Pre-inbound: OCR ảnh (drag/drop, paste, file) ----
  if (window.WOPOcrScan) WOPOcrScan.init();

  // ---- flowchart interactivity ----
  const ioPanel = document.getElementById("io-detail");
  const exPanel = document.getElementById("ex-detail");
  WOPFlowcharts.initClickToDetail(document.getElementById("io-diagram"), ioPanel);
  WOPFlowcharts.initClickToDetail(document.getElementById("ex-diagram"), exPanel);

  document.getElementById("io-export").addEventListener("click", function () {
    const svgs = document.getElementById("io-diagram").querySelectorAll("svg");
    svgs.forEach(function (svg, idx) {
      WOPFlowcharts.exportSvgAsPng(svg, "SOP_Kho_" + (idx === 0 ? "Inbound" : "Outbound") + ".png");
    });
  });
  document.getElementById("ex-export").addEventListener("click", function () {
    const svgs = document.getElementById("ex-diagram").querySelectorAll("svg");
    svgs.forEach(function (svg, idx) {
      WOPFlowcharts.exportSvgAsPng(svg, "SOP_Express_Flow_" + (idx === 0 ? "LuongA" : "LuongB") + ".png");
    });
  });

  // ---- generic sheet-picker modal ----
  const modal = document.getElementById("sheetModalOverlay");
  const sheetSelect = document.getElementById("sheetSelect");
  let sheetModalResolve = null;
  document.getElementById("sheetModalConfirm").addEventListener("click", function () {
    modal.classList.remove("show");
    if (sheetModalResolve) sheetModalResolve(sheetSelect.value);
  });
  document.getElementById("sheetModalCancel").addEventListener("click", function () {
    modal.classList.remove("show");
    if (sheetModalResolve) sheetModalResolve(null);
  });
  function pickSheet(workbook, preferredName) {
    return new Promise(function (resolve) {
      if (workbook.SheetNames.length === 1) { resolve(workbook.SheetNames[0]); return; }
      sheetSelect.innerHTML = "";
      workbook.SheetNames.forEach(function (name) {
        const opt = document.createElement("option");
        opt.value = name; opt.textContent = name;
        if (name.toLowerCase() === (preferredName || "").toLowerCase()) opt.selected = true;
        sheetSelect.appendChild(opt);
      });
      sheetModalResolve = resolve;
      modal.classList.add("show");
    });
  }

  // ---- shared generator wiring ----
  function wireGenerator(cfg) {
    const fileInput = document.getElementById(cfg.fileId);
    const statusEl = document.getElementById(cfg.statusId);
    const genBtn = document.getElementById(cfg.genBtnId);
    const logEl = document.getElementById(cfg.logId);
    const resultsEl = document.getElementById(cfg.resultsId);
    let currentRows = null;

    fileInput.addEventListener("change", async function () {
      const file = fileInput.files[0];
      if (!file) return;
      statusEl.textContent = "Đang đọc file...";
      try {
        const wb = await WOPUtils.readWorkbookFromFile(file);
        const sheetName = await pickSheet(wb, cfg.preferredSheet);
        if (!sheetName) { statusEl.textContent = "Đã huỷ."; currentRows = null; genBtn.disabled = true; return; }
        currentRows = WOPUtils.sheetToRows(wb, sheetName);
        statusEl.textContent = "Đã nạp sheet \"" + sheetName + "\" — " + currentRows.length + " dòng. Sẵn sàng tạo label.";
        genBtn.disabled = false;
      } catch (e) {
        statusEl.textContent = "Lỗi đọc file: " + e.message;
        currentRows = null; genBtn.disabled = true;
      }
    });

    genBtn.addEventListener("click", async function () {
      if (!currentRows) return;
      genBtn.disabled = true;
      logEl.textContent = "";
      resultsEl.innerHTML = "";
      const log = function (msg) { logEl.textContent += msg + "\n"; logEl.scrollTop = logEl.scrollHeight; };
      try {
        const extra = cfg.getExtraOptions ? cfg.getExtraOptions() : {};
        const out = await cfg.generateFn(Object.assign({ rows: currentRows, log: log }, extra));
        log("✅ Hoàn tất.");
        out.files.forEach(function (f) {
          const a = document.createElement("a");
          a.className = "file-chip"; a.href = "#"; a.textContent = "⬇ " + f.name + " (" + f.count + ")";
          a.addEventListener("click", function (e) { e.preventDefault(); WOPUtils.downloadBlob(f.blob, f.name); });
          resultsEl.appendChild(a);
        });
      } catch (e) {
        logEl.textContent += "❌ Lỗi: " + e.message + "\n";
      } finally {
        genBtn.disabled = false;
      }
    });
  }

  wireGenerator({
    fileId: "singleFileInput", statusId: "singleStatus", genBtnId: "singleGenBtn",
    logId: "singleLog", resultsId: "singleResults", preferredSheet: "Single",
    generateFn: WOPAsnSingle.generate,
    getExtraOptions: function () { return { refNoOverride: document.getElementById("singleRefNo").value.trim() || null }; },
  });

  wireGenerator({
    fileId: "mixFileInput", statusId: "mixStatus", genBtnId: "mixGenBtn",
    logId: "mixLog", resultsId: "mixResults", preferredSheet: "Mix",
    generateFn: WOPAsnMix.generate,
    getExtraOptions: function () {
      return {
        refNoOverride: document.getElementById("mixRefNo").value.trim() || null,
        shrinkMaxRows: parseInt(document.getElementById("mixShrinkThreshold").value, 10) || WOPAsnMix.DEFAULT_SHRINK_MAX_ROWS,
      };
    },
  });

  // ---- PDF In Order (merge N label PDFs + sort by Carton No.) ----
  (function wirePdfOrder() {
    const fileInput = document.getElementById("pdfOrderFileInput");
    if (!fileInput) return;
    const statusEl = document.getElementById("pdfOrderStatus");
    const genBtn = document.getElementById("pdfOrderGenBtn");
    const logEl = document.getElementById("pdfOrderLog");
    const resultsEl = document.getElementById("pdfOrderResults");
    let currentFiles = null;

    fileInput.addEventListener("change", function () {
      currentFiles = (fileInput.files && fileInput.files.length) ? Array.from(fileInput.files) : null;
      if (!currentFiles) { statusEl.textContent = "Chưa chọn file."; genBtn.disabled = true; return; }
      statusEl.textContent = "Đã chọn " + currentFiles.length + " file PDF. Sẵn sàng gộp & sắp xếp.";
      genBtn.disabled = false;
    });

    genBtn.addEventListener("click", async function () {
      if (!currentFiles) return;
      genBtn.disabled = true;
      logEl.textContent = "";
      resultsEl.innerHTML = "";
      const log = function (msg) { logEl.textContent += msg + "\n"; logEl.scrollTop = logEl.scrollHeight; };
      try {
        const out = await WOPPdfOrder.generate({ files: currentFiles, log: log });
        log("✅ Hoàn tất.");
        out.files.forEach(function (f) {
          const a = document.createElement("a");
          a.className = "file-chip"; a.href = "#"; a.textContent = "⬇ " + f.name + " (" + f.count + ")";
          a.addEventListener("click", function (e) { e.preventDefault(); WOPUtils.downloadBlob(f.blob, f.name); });
          resultsEl.appendChild(a);
        });
      } catch (e) {
        logEl.textContent += "❌ Lỗi: " + e.message + "\n";
      } finally {
        genBtn.disabled = false;
      }
    });
  })();

  // ---- Outbound: Detect Outbound Type & Suggest Convert UOM ----
  (function wireOutboundUom() {
    const invInput = document.getElementById("outboundInventoryInput");
    if (!invInput) return;
    const orderInput = document.getElementById("outboundOrderInput");
    const bundleInput = document.getElementById("outboundBundleInput");
    const statusEl = document.getElementById("outboundUomStatus");
    const genBtn = document.getElementById("outboundUomGenBtn");
    const logEl = document.getElementById("outboundUomLog");
    const resultsEl = document.getElementById("outboundUomResults");

    function refreshReady() {
      const ready = invInput.files.length && orderInput.files.length;
      genBtn.disabled = !ready;
      if (ready) {
        let msg = "Đã chọn: " + invInput.files[0].name + " + " + orderInput.files[0].name;
        if (bundleInput.files.length) msg += " + " + bundleInput.files[0].name + " (bundle)";
        msg += ". Sẵn sàng chạy.";
        statusEl.textContent = msg;
      } else {
        statusEl.textContent = "Cần chọn cả 2 file: Tồn kho realtime + Order outbound.";
      }
    }
    invInput.addEventListener("change", refreshReady);
    orderInput.addEventListener("change", refreshReady);
    bundleInput.addEventListener("change", refreshReady);

    genBtn.addEventListener("click", async function () {
      if (!invInput.files.length || !orderInput.files.length) return;
      genBtn.disabled = true;
      logEl.textContent = "";
      resultsEl.innerHTML = "";
      const log = function (msg) { logEl.textContent += msg + "\n"; logEl.scrollTop = logEl.scrollHeight; };
      try {
        const out = await WOPOutboundUom.generate({
          inventoryFile: invInput.files[0],
          orderFile: orderInput.files[0],
          bundleFile: bundleInput.files.length ? bundleInput.files[0] : null,
          log: log,
        });
        log("✅ Hoàn tất.");
        out.files.forEach(function (f) {
          const a = document.createElement("a");
          a.className = "file-chip"; a.href = "#"; a.textContent = "⬇ " + f.name + " (" + f.count + ")";
          a.addEventListener("click", function (e) { e.preventDefault(); WOPUtils.downloadBlob(f.blob, f.name); });
          resultsEl.appendChild(a);
        });
      } catch (e) {
        logEl.textContent += "❌ Lỗi: " + e.message + "\n";
      } finally {
        genBtn.disabled = false;
      }
    });
  })();
})();
