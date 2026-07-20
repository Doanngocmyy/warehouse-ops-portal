/* ==========================================================
 * ocr-scan.js — Pre-inbound: OCR ảnh (drag & drop / paste / file)
 * Chạy 100% client-side bằng Tesseract.js — không upload ảnh ra ngoài.
 * ========================================================== */
window.WOPOcrScan = (function () {
  function init() {
    const dropzone = document.getElementById("ocrDropzone");
    if (!dropzone) return; // markup không tồn tại (trang khác) -> bỏ qua

    const fileInput = document.getElementById("ocrFileInput");
    const statusEl = document.getElementById("ocrStatus");
    const thumbsEl = document.getElementById("ocrThumbs");
    const runBtn = document.getElementById("ocrRunBtn");
    const clearBtn = document.getElementById("ocrClearBtn");
    const logEl = document.getElementById("ocrLog");
    const resultEl = document.getElementById("ocrResultText");
    const copyBtn = document.getElementById("ocrCopyBtn");
    const langSelect = document.getElementById("ocrLang");

    let images = []; // [{ file, url }]

    function log(msg) {
      logEl.textContent += msg + "\n";
      logEl.scrollTop = logEl.scrollHeight;
    }

    function refreshUi() {
      runBtn.disabled = images.length === 0;
      statusEl.textContent = images.length
        ? ("Đã có " + images.length + " ảnh — sẵn sàng chạy OCR.")
        : "Chưa có ảnh nào.";
      thumbsEl.innerHTML = "";
      images.forEach(function (img, idx) {
        const wrap = document.createElement("div");
        wrap.className = "ocr-thumb";
        const im = document.createElement("img");
        im.src = img.url;
        im.alt = "ảnh " + (idx + 1);
        const rm = document.createElement("button");
        rm.type = "button";
        rm.className = "ocr-thumb-remove";
        rm.title = "Xoá ảnh này";
        rm.textContent = "✕";
        rm.addEventListener("click", function (e) {
          e.stopPropagation();
          URL.revokeObjectURL(img.url);
          images.splice(idx, 1);
          refreshUi();
        });
        wrap.appendChild(im);
        wrap.appendChild(rm);
        thumbsEl.appendChild(wrap);
      });
    }

    function addFiles(fileList) {
      const files = Array.from(fileList || []).filter(function (f) {
        return f && f.type && f.type.indexOf("image/") === 0;
      });
      if (!files.length) return;
      files.forEach(function (f) {
        images.push({ file: f, url: URL.createObjectURL(f) });
      });
      refreshUi();
    }

    // ---- click khung để chọn file ----
    dropzone.addEventListener("click", function () {
      fileInput.click();
    });
    fileInput.addEventListener("change", function () {
      addFiles(fileInput.files);
      fileInput.value = "";
    });

    // ---- kéo-thả ----
    ["dragenter", "dragover"].forEach(function (evt) {
      dropzone.addEventListener(evt, function (e) {
        e.preventDefault();
        e.stopPropagation();
        dropzone.classList.add("dragover");
      });
    });
    ["dragleave", "drop"].forEach(function (evt) {
      dropzone.addEventListener(evt, function (e) {
        e.preventDefault();
        e.stopPropagation();
        dropzone.classList.remove("dragover");
      });
    });
    dropzone.addEventListener("drop", function (e) {
      if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length) {
        addFiles(e.dataTransfer.files);
      }
    });

    // ---- dán ảnh (Ctrl+V) ----
    function handlePaste(e) {
      const items = (e.clipboardData && e.clipboardData.items) || [];
      const files = [];
      for (let i = 0; i < items.length; i++) {
        if (items[i].type && items[i].type.indexOf("image/") === 0) {
          const f = items[i].getAsFile();
          if (f) files.push(f);
        }
      }
      if (files.length) {
        e.preventDefault();
        addFiles(files);
      }
    }
    dropzone.addEventListener("paste", handlePaste);
    // fallback: cho phép Ctrl+V ở bất kỳ đâu miễn tab OCR đang mở (không cần click vào khung trước)
    document.addEventListener("paste", function (e) {
      const sub = document.getElementById("sub-ocr-scan");
      if (sub && sub.classList.contains("active")) handlePaste(e);
    });

    clearBtn.addEventListener("click", function () {
      images.forEach(function (img) { URL.revokeObjectURL(img.url); });
      images = [];
      resultEl.value = "";
      copyBtn.disabled = true;
      logEl.textContent = "";
      refreshUi();
    });

    runBtn.addEventListener("click", async function () {
      if (!images.length) return;
      if (typeof Tesseract === "undefined") {
        log("❌ Không tải được thư viện OCR (Tesseract.js). Kiểm tra kết nối mạng rồi thử lại.");
        return;
      }
      runBtn.disabled = true;
      clearBtn.disabled = true;
      copyBtn.disabled = true;
      logEl.textContent = "";
      resultEl.value = "";
      const lang = langSelect.value || "vie+eng";
      let worker = null;
      try {
        log("⏳ Đang tải model ngôn ngữ (" + lang + ")... (lần đầu có thể mất chút thời gian)");
        worker = await Tesseract.createWorker(lang, 1, {
          logger: function (m) {
            if (m && m.status && typeof m.progress === "number") {
              log("… " + m.status + " (" + Math.round(m.progress * 100) + "%)");
            }
          },
        });
        const parts = [];
        for (let i = 0; i < images.length; i++) {
          log("🔍 Đang nhận diện ảnh " + (i + 1) + "/" + images.length + "...");
          const out = await worker.recognize(images[i].file);
          const text = ((out && out.data && out.data.text) || "").trim();
          parts.push(images.length > 1 ? ("--- Ảnh " + (i + 1) + " ---\n" + text) : text);
        }
        resultEl.value = parts.join("\n\n").trim();
        copyBtn.disabled = resultEl.value.length === 0;
        log("✅ Hoàn tất. Có thể sửa lại kết quả rồi bấm \"Copy toàn bộ\".");
      } catch (e) {
        log("❌ Lỗi OCR: " + e.message);
      } finally {
        if (worker) { try { await worker.terminate(); } catch (e2) { /* ignore */ } }
        runBtn.disabled = false;
        clearBtn.disabled = false;
      }
    });

    // ---- copy toàn bộ khối kết quả (giống nút copy code-block của ChatGPT) ----
    copyBtn.addEventListener("click", async function () {
      const text = resultEl.value;
      if (!text) return;
      try {
        await navigator.clipboard.writeText(text);
      } catch (e) {
        resultEl.select();
        resultEl.setSelectionRange(0, text.length);
        document.execCommand("copy");
      }
      const original = copyBtn.textContent;
      copyBtn.textContent = "✅ Đã copy!";
      setTimeout(function () { copyBtn.textContent = original; }, 1500);
    });

    resultEl.addEventListener("input", function () {
      copyBtn.disabled = resultEl.value.length === 0;
    });

    refreshUi();
  }

  return { init: init };
})();
