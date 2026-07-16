/* ==========================================================
 * flowcharts.js — interactivity (click-to-detail) + image export
 * for the two SOP diagrams on the Overview page.
 * ========================================================== */
(function () {
  const DETAILS = {
    "io-pre-inbound": "Pre-inbound: Factory (domestic VN hoặc cross-border) gửi Packing List (PL) trước cho KEC để chuẩn bị tạo ASN — bước này không có khách hàng tham gia trực tiếp.",
    "io-asn-creation": "KEC tạo ASN, sau đó upload Inbound Request (IR) lên hệ thống FOMS để lưu lại, đối chiếu khi hàng thực tế về kho.",
    "io-d1-factory": "D-1: Factory báo trước cho KEC lịch hàng sắp về (D-1), gửi kèm bộ chứng từ và thông tin tài xế. Domestic (VN) cần CI/PL/TKHQ(Barcode); Cross-border (CN) cần CI/PL/AN/EDO/BL.",
    "io-d1-kec": "KEC review chứng từ + thông tin tài xế nhận được từ Factory, kiểm tra đầy đủ trước khi forward.",
    "io-d1-forward": "KEC forward bộ chứng từ đã review sang WH để WH chuẩn bị thủ tục thông quan/nhập kho.",
    "io-d1-wh": "WH dựa trên chứng từ nhận được, nộp/submit bộ chứng từ cho nhân viên hải quan trước khi xe hàng tới.",
    "io-arrival": "Xe tải đến kho — bàn giao hàng hoá cho WH, bắt đầu tính SLA Inbound D+1.",
    "io-sla": "Inbound SLA: D+1. Gồm: Unloading → Counting → Checking → Put-away → Inventory Updated.",
    "io-done": "Inbound Done — hàng đã được xử lý xong và cập nhật tồn kho.",
    "ob-request": "Customer submit Outbound Request (OR) — yêu cầu xuất hàng ra khỏi kho.",
    "ob-kec-check": "KEC nhận OR: check tồn kho, check UOM để quyết định loại outbound (nguyên kiện/lẻ...), check có yêu cầu VAS không, và gợi ý convert UOM nếu cần.",
    "ob-wh-plan": "WH điền các thông tin còn thiếu vào OR: CDS (Customs Declaration Sheet) và Vendor.",
    "ob-kec-confirm": "KEC review lại OR đã được điền đầy đủ, gửi cho khách hàng TOPOLOGIE xác nhận; đồng thời upload lên hệ thống để vendor tiến hành pick-pack.",
    "ob-customer-confirm": "Customer xác nhận CO (Certificate of Origin) — không cần chờ bước này mới tiếp tục xử lý song song.",
    "ob-execution": "WH thực thi: Pick → VAS (nếu có) → Labeling → Pack → DO Preparation → nộp chứng từ cho nhân viên hải quan (họ giám sát + nhập dữ liệu import).",
    "ob-sla": "Customs SLA (thuộc WH, không phải Factory): D0 = gửi booking + driver info cùng lúc; D1 = Pickup. => ETD = D+2 (đối với AIR/SEA tính theo giờ closing/cut-off).",
    "ob-done": "Outbound Done.",
    "ex-a-topologie1": "TOPOLOGIE sends an Outbound Request (OR) to KEC — starts the Self-placed booking flow.",
    "ex-a-kec1": "KEC receives the OR and checks stock before forwarding it to the warehouse.",
    "ex-a-kec2": "KEC sends it to the WH, assigning CDS or VENDOR documentation per shipment.",
    "ex-a-wh1": "Bonded WH confirms CDS, receives the OR, and does pick & pack + dim weight.",
    "ex-a-wh2": "Bonded WH sends the DO (Delivery Order) to KEC once pick & pack is done.",
    "ex-a-kec3": "KEC uses the DO + shipment info to book freight via EXPRESS FWD (Alphatrans).",
    "ex-a-fwd": "EXPRESS FWD (Alphatrans) returns the AWB, Chargeable Weight (C.W), and rate quote.",
    "ex-a-kec4": "KEC sends the rate quote to TOPOLOGIE for confirmation.",
    "ex-a-topologie2": "TOPOLOGIE confirms the quote — approves shipping at the quoted rate.",
    "ex-a-kec5": "KEC gets driver info (books its own truck for the leg to the warehouse).",
    "ex-a-wh3": "Bonded WH preps the DO + customs declaration (COT 2pm), then hands off to the truck.",
    "ex-b-cnee1": "CNEE (the nominated consignee) books the AWB directly with the carrier — starts this flow.",
    "ex-b-cnee2": "CNEE selects remote pickup and sends KEC the control number to confirm the booking.",
    "ex-b-kec1": "KEC checks stock and sends the WH to fill in the matching CDS / VENDOR docs.",
    "ex-b-wh1": "Bonded WH does pick & pack and measures dim weight for the shipment.",
    "ex-b-wh2": "Bonded WH sends the DO to KEC once pick & pack is complete.",
    "ex-b-kec2": "KEC verifies the info and dim weight, and collects driver details.",
    "ex-b-kec3": "KEC sends the booking + driver info to Bonded WH to prepare handover.",
    "ex-b-wh3": "Bonded WH preps the DO + customs declaration (COT 2pm, ~1 day approval).",
    "ex-b-truck1": "Truck arrives; the driver submits documents directly at the customs counter.",
    "ex-b-truck2": "Truck waits for the customs seal cut, then departs — Flow B complete.",
  };

  function showDetail(panelEl, key) {
    if (!panelEl) return;
    panelEl.textContent = DETAILS[key] || "(Chưa có mô tả chi tiết cho bước này)";
  }

  function initClickToDetail(rootEl, panelEl) {
    rootEl.querySelectorAll("[data-key]").forEach(function (el) {
      el.classList.add("clickable");
      el.addEventListener("click", function () { showDetail(panelEl, el.getAttribute("data-key")); });
    });
  }

  async function exportSvgAsPng(svgEl, filename) {
    const xml = new XMLSerializer().serializeToString(svgEl);
    const svg64 = btoa(unescape(encodeURIComponent(xml)));
    const img = new Image();
    const scale = 2;
    await new Promise(function (resolve, reject) {
      img.onload = resolve; img.onerror = reject;
      img.src = "data:image/svg+xml;base64," + svg64;
    });
    const vb = svgEl.viewBox.baseVal;
    const w = (vb && vb.width) || svgEl.width.baseVal.value;
    const h = (vb && vb.height) || svgEl.height.baseVal.value;
    const canvas = document.createElement("canvas");
    canvas.width = w * scale; canvas.height = h * scale;
    const ctx = canvas.getContext("2d");
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
    canvas.toBlob(function (blob) { WOPUtils.downloadBlob(blob, filename); }, "image/png");
  }

  async function exportHtmlAsPng(containerEl, filename) {
    const canvas = await html2canvas(containerEl, { backgroundColor: "#ffffff", scale: 2 });
    canvas.toBlob(function (blob) { WOPUtils.downloadBlob(blob, filename); }, "image/png");
  }

  window.WOPFlowcharts = { initClickToDetail, exportSvgAsPng, exportHtmlAsPng, showDetail };
})();
