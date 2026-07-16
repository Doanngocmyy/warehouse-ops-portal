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
    "ex-a-topologie1": "TOPOLOGIE gửi Outbound Request (OR) cho KEC — mở đầu luồng Self-placed booking.",
    "ex-a-kec1": "KEC nhận OR, kiểm tra tồn kho trước khi gửi tiếp cho kho phân loại chứng từ.",
    "ex-a-kec2": "KEC gửi kho: phân chia loại chứng từ CDS / VENDOR tuỳ theo shipment.",
    "ex-a-wh1": "Bonded WH phản hồi CDS, nhận OR, thực hiện pick & pack, đo dim weight.",
    "ex-a-wh2": "Bonded WH gửi DO (Delivery Order) cho KEC sau khi pick & pack xong.",
    "ex-a-kec3": "KEC dùng DO + thông tin lô hàng để book cước qua EXPRESS FWD (Alphatrans).",
    "ex-a-fwd": "EXPRESS FWD (Alphatrans) gửi lại AWB, Chargeable Weight (C.W) và báo giá cước.",
    "ex-a-kec4": "KEC gửi báo giá cước cho TOPOLOGIE để xác nhận.",
    "ex-a-topologie2": "TOPOLOGIE confirm báo giá — đồng ý xuất hàng theo cước đã báo.",
    "ex-a-kec5": "KEC lấy thông tin tài xế (KEC tự book tài xế nội bộ cho chặng vận chuyển tới kho).",
    "ex-a-wh3": "Bonded WH làm DO + khai hải quan (COT 14:00) rồi bàn giao cho truck lấy hàng.",
    "ex-b-cnee1": "CNEE (bên nhận chỉ định) tự đặt AWB trực tiếp với hãng chuyển phát — mở đầu luồng CNEE-nominated booking.",
    "ex-b-cnee2": "CNEE chọn remote pickup và gửi control number cho KEC để KEC biết lô hàng đã được đặt.",
    "ex-b-kec1": "KEC check tồn kho, gửi kho điền chứng từ CDS / VENDOR tương ứng.",
    "ex-b-wh1": "Bonded WH pick & pack, đo dim weight cho lô hàng.",
    "ex-b-wh2": "Bonded WH gửi DO cho KEC sau khi hoàn tất pick & pack.",
    "ex-b-kec2": "KEC check lại thông tin, dim weight, đồng thời lấy thông tin tài xế.",
    "ex-b-kec3": "KEC gửi booking + thông tin tài xế cho Bonded WH để chuẩn bị giao hàng.",
    "ex-b-wh3": "Bonded WH làm DO + khai hải quan (COT 14:00, duyệt khoảng 1 ngày).",
    "ex-b-truck1": "Truck đến kho, tài xế tự nộp chứng từ tại quầy hải quan (đặc thù luồng CNEE-nominated).",
    "ex-b-truck2": "Chờ hải quan cắt seal, truck rời kho — hoàn tất luồng B.",
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
