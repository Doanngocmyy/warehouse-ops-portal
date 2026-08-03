/* ==========================================================
 * tests/outbound-availability.test.js
 * Plain Node.js assert-based tests (no framework/deps) for the
 * Bundle-SKU-centric matching logic in js/outbound-availability.js.
 * Run with: node tests/outbound-availability.test.js
 * ========================================================== */
"use strict";
const assert = require("assert");
const path = require("path");

global.window = global.window || {};
const mod = require(path.join(__dirname, "..", "js", "outbound-availability.js"));

let passed = 0, failed = 0;
function test(name, fn) {
  try {
    fn();
    passed++;
    console.log("  ok  - " + name);
  } catch (e) {
    failed++;
    console.log("FAIL  - " + name);
    console.log("        " + (e && e.message ? e.message : e));
  }
}

// ---- helpers ----
// Build a "bundles" array the same shape parseBundleFile() groups into:
// { bundleSku, bundleName, children:[{childSku, childProduct, unitCode, unitName, qty, availabilityRatio}], totalChildPcs }
function bundle(sku, name, children) {
  const totalPcs = children.reduce(function (s, c) { return s + (typeof c.qty === "number" ? c.qty : 0); }, 0);
  return { bundleSku: sku, bundleName: name, children: children, totalChildPcs: totalPcs };
}
function child(sku, product, qty) {
  return { childSku: sku, childProduct: product, unitCode: "PCS", unitName: "PCS", qty: qty, availabilityRatio: 100 };
}
function invRow(warehouse, invProductCode, availableQty, extra) {
  return Object.assign({
    warehouse: warehouse, invProductCode: invProductCode, partnerSku: "",
    product: "Inv " + invProductCode, category: "Accessories", unit: "MIX", conditionType: "Mới",
    stockQty: availableQty, freezeQty: 0, availableQtyRaw: availableQty,
    availableQty: mod._internal.toNumberOrNull(availableQty),
    pendingIn: 0, pendingOut: 0, weightKg: 1, cbm: 0.01,
    lastStockoutDate: "", lastOutboundDate: "", lastInboundDate: "",
  }, extra || {});
}
function selKeys(bundles) {
  return new Set(bundles.map(function (b) { return mod._internal.normalizeBundleSkuKey(b.bundleSku); }));
}

console.log("\n=== Real-Time Outbound Availability — Bundle-SKU matching engine tests ===\n");

// ---- Test 1: one Bundle SKU, multiple child rows, Available Qty = 1 ----
test("Test 1: Bundle with 3 child rows, Available Qty=1 -> 1 available carton, all 3 children retained", function () {
  const bundles = [bundle("TP-BUNDLE-001", "Bundle 001", [child("A", "Child A", 10), child("B", "Child B", 20), child("C", "Child C", 5)])];
  const idx = mod.buildBundleKeyIndex([invRow("SG2", "TP-BUNDLE-001", 1)], "invProductCode");
  const result = mod.runMatch({ bundles: bundles, bundleKeyIndex: idx, selectedBundleSkuKeys: selKeys(bundles) });
  assert.strictEqual(result.available.length, 1);
  assert.strictEqual(result.available[0].childCount, 3);
  assert.strictEqual(result.available[0].children.length, 3);
  assert.strictEqual(result.available[0].totalChildPcs, 35);
  assert.strictEqual(result.summary.availableBundleCartons, 1);
});

// ---- Test 2: Bundle SKU absent from inventory ----
test("Test 2: Bundle SKU absent from inventory -> entire bundle excluded", function () {
  const bundles = [bundle("TP-BUNDLE-002", "Bundle 002", [child("A", "A", 1), child("B", "B", 2), child("C", "C", 3), child("D", "D", 4), child("E", "E", 5)])];
  const idx = mod.buildBundleKeyIndex([], "invProductCode"); // empty inventory
  const result = mod.runMatch({ bundles: bundles, bundleKeyIndex: idx, selectedBundleSkuKeys: selKeys(bundles) });
  assert.strictEqual(result.available.length, 0);
  assert.strictEqual(result.excluded.length, 1);
  assert.strictEqual(result.excluded[0].reason, "BUNDLE_SKU_NOT_FOUND_IN_INVENTORY");
  assert.strictEqual(result.excluded[0].childCount, 5); // none of its 5 child lines appear available
});

// ---- Test 3: Bundle SKU exists, Available Qty = 0 ----
test("Test 3: Bundle SKU exists with Available Qty=0 -> entire bundle excluded (AVAILABLE_QTY_ZERO)", function () {
  const bundles = [bundle("TP-BUNDLE-003", "Bundle 003", [child("A", "A", 1)])];
  const idx = mod.buildBundleKeyIndex([invRow("SG2", "TP-BUNDLE-003", 0)], "invProductCode");
  const result = mod.runMatch({ bundles: bundles, bundleKeyIndex: idx, selectedBundleSkuKeys: selKeys(bundles) });
  assert.strictEqual(result.available.length, 0);
  assert.strictEqual(result.excluded[0].reason, "AVAILABLE_QTY_ZERO");
});

// ---- Test 4: Bundle SKU exists, Available Qty > 0 ----
test("Test 4: Bundle SKU exists with Available Qty>0 -> entire bundle included", function () {
  const bundles = [bundle("TP-BUNDLE-004", "Bundle 004", [child("A", "A", 1)])];
  const idx = mod.buildBundleKeyIndex([invRow("SG2", "TP-BUNDLE-004", 7)], "invProductCode");
  const result = mod.runMatch({ bundles: bundles, bundleKeyIndex: idx, selectedBundleSkuKeys: selKeys(bundles) });
  assert.strictEqual(result.available.length, 1);
  assert.strictEqual(result.available[0].availableQty, 7);
});

// ---- Test 5: same Child SKU under two different Bundle SKUs ----
test("Test 5: same Child SKU under 2 different Bundle SKUs -> two separate cartons", function () {
  const bundles = [
    bundle("TP-A", "Bundle A", [child("SHARED-CHILD", "Shared", 5)]),
    bundle("TP-B", "Bundle B", [child("SHARED-CHILD", "Shared", 5)]),
  ];
  const idx = mod.buildBundleKeyIndex([invRow("SG2", "TP-A", 2), invRow("SG2", "TP-B", 3)], "invProductCode");
  const result = mod.runMatch({ bundles: bundles, bundleKeyIndex: idx, selectedBundleSkuKeys: selKeys(bundles) });
  assert.strictEqual(result.available.length, 2);
  assert.strictEqual(result.summary.availableBundleCartons, 2);
  const skus = result.available.map(function (r) { return r.bundleSku; }).sort();
  assert.deepStrictEqual(skus, ["TP-A", "TP-B"]);
});

// ---- Test 6: multiple child rows must not inflate the available bundle-carton count ----
test("Test 6: 1 bundle with 10 child rows + Available Qty=1 -> still exactly 1 available carton", function () {
  const children = [];
  for (let i = 0; i < 10; i++) children.push(child("C" + i, "Child " + i, i + 1));
  const bundles = [bundle("TP-MANY-CHILDREN", "Many Children", children)];
  const idx = mod.buildBundleKeyIndex([invRow("SG2", "TP-MANY-CHILDREN", 1)], "invProductCode");
  const result = mod.runMatch({ bundles: bundles, bundleKeyIndex: idx, selectedBundleSkuKeys: selKeys(bundles) });
  assert.strictEqual(result.available.length, 1);
  assert.strictEqual(result.summary.availableBundleCartons, 1);
  assert.strictEqual(result.available[0].childCount, 10);
});

// ---- Test 7: Total Child PCS must equal sum of SL for the available bundle ----
test("Test 7: Total Child PCS = sum(SL) for available bundle", function () {
  const bundles = [bundle("TP-PCS-SUM", "PCS Sum", [child("A", "A", 100), child("B", "B", 10), child("C", "C", 10), child("D", "D", 20)])];
  const idx = mod.buildBundleKeyIndex([invRow("SG2", "TP-PCS-SUM", 1)], "invProductCode");
  const result = mod.runMatch({ bundles: bundles, bundleKeyIndex: idx, selectedBundleSkuKeys: selKeys(bundles) });
  assert.strictEqual(result.available[0].totalChildPcs, 140);
  assert.strictEqual(result.summary.totalChildPcsInAvailableBundles, 140);
});

// ---- Test 8: inventory matching must use Bundle SKU, not Child SKU ----
test("Test 8: matching uses Bundle SKU only — a Child SKU present in inventory must NOT cause a false match", function () {
  const bundles = [bundle("TP-NOT-IN-INV", "Not In Inv", [child("4895227934032", "Child EAN", 100)])];
  // Inventory only has the CHILD EAN as its own row, never the Bundle SKU itself.
  const idx = mod.buildBundleKeyIndex([invRow("SG2", "4895227934032", 50)], "invProductCode");
  const result = mod.runMatch({ bundles: bundles, bundleKeyIndex: idx, selectedBundleSkuKeys: selKeys(bundles) });
  assert.strictEqual(result.available.length, 0);
  assert.strictEqual(result.excluded[0].reason, "BUNDLE_SKU_NOT_FOUND_IN_INVENTORY");
});

// ---- Test 9: Bundle SKU may match against either Mã sản phẩm or Mã SKU đối tác ----
test("Test 9: auto-detection picks whichever inventory column actually contains the Bundle SKUs", function () {
  const bundles = [bundle("TP-X1", "X1", [child("A", "A", 1)]), bundle("TP-X2", "X2", [child("B", "B", 1)])];
  const invRecords = [
    { warehouse: "SG2", invProductCode: "9999999999999", partnerSku: "TP-X1", availableQty: 5, product: "", category: "", unit: "", conditionType: "", stockQty: 5, freezeQty: 0, pendingIn: 0, pendingOut: 0, weightKg: 0, cbm: 0, lastStockoutDate: "", lastOutboundDate: "", lastInboundDate: "" },
    { warehouse: "SG2", invProductCode: "8888888888888", partnerSku: "TP-X2", availableQty: 3, product: "", category: "", unit: "", conditionType: "", stockQty: 3, freezeQty: 0, pendingIn: 0, pendingOut: 0, weightKg: 0, cbm: 0, lastStockoutDate: "", lastOutboundDate: "", lastInboundDate: "" },
  ];
  const bundleKeys = selKeys(bundles);
  const detection = mod.detectBundleSkuKeySource(bundleKeys, invRecords, null);
  assert.strictEqual(detection.field, "partnerSku"); // Bundle SKUs live under partnerSku here, not invProductCode
  assert.strictEqual(detection.overlap, 2);
  const result = mod.runMatch({ bundles: bundles, bundleKeyIndex: detection.index, selectedBundleSkuKeys: bundleKeys });
  assert.strictEqual(result.available.length, 2);
});

// ---- Test 10: no overlap -> clear warning, not silent empty processing ----
test("Test 10: zero overlap on both candidate columns is detectable (not silently swallowed)", function () {
  const bundles = [bundle("TP-NOWHERE", "Nowhere", [child("A", "A", 1)])];
  const invRecords = [{ warehouse: "SG2", invProductCode: "1111111111111", partnerSku: "SOMETHING-ELSE", availableQty: 5, product: "", category: "", unit: "", conditionType: "", stockQty: 5, freezeQty: 0, pendingIn: 0, pendingOut: 0, weightKg: 0, cbm: 0, lastStockoutDate: "", lastOutboundDate: "", lastInboundDate: "" }];
  const detection = mod.detectBundleSkuKeySource(selKeys(bundles), invRecords, null);
  assert.strictEqual(detection.overlap, 0);
  assert.strictEqual(detection.tried.length, 2);
  assert.ok(detection.tried.every(function (t) { return t.overlap === 0; }));
});

// ---- Test 11: warehouse filter excludes the entire bundle inventory row + child details ----
test("Test 11: warehouse filter excludes the whole bundle-carton row (children go with it)", function () {
  const bundles = [bundle("TP-WH", "WH Bundle", [child("A", "A", 1), child("B", "B", 2)])];
  const idx = mod.buildBundleKeyIndex([invRow("WH-North", "TP-WH", 4), invRow("WH-South", "TP-WH", 6)], "invProductCode");
  const result = mod.runMatch({ bundles: bundles, bundleKeyIndex: idx, selectedBundleSkuKeys: selKeys(bundles), warehouseFilter: new Set(["WH-South"]) });
  assert.strictEqual(result.available.length, 1);
  assert.strictEqual(result.available[0].warehouse, "WH-South");
  assert.strictEqual(result.available[0].childCount, 2); // children travel with the carton
  assert.strictEqual(result.excluded.length, 1);
  assert.strictEqual(result.excluded[0].reason, "WAREHOUSE_FILTERED_OUT");
  assert.strictEqual(result.excluded[0].childCount, 2);
});

// ---- Test 12: EAN / SKU values remain text in Excel export ----
test("Test 12: buildExportWorkbook keeps Bundle SKU / Child SKU as text", function () {
  global.XLSX = {
    utils: {
      aoa_to_sheet: function (aoa) { return { __aoa: aoa }; },
      book_new: function () { return { sheets: {} }; },
      book_append_sheet: function (wb, ws, name) { wb.sheets[name] = ws; },
    },
    write: function () { return new Uint8Array([1, 2, 3]); },
  };
  global.Blob = function (parts, opts) { this.parts = parts; this.opts = opts; this.size = 3; this.type = opts.type; };
  const fakeResult = {
    available: [{
      bundleSku: "04895227934032", bundleName: "B", warehouse: "SG2", invProductCode: "04895227934032", partnerSku: "TP-01",
      invProductName: "P", category: "C", unit: "MIX", conditionType: "Mới", stockQty: 1, freezeQty: 0, availableQty: 1,
      pendingIn: 0, pendingOut: 0, weightKg: 1, cbm: 0.01, lastStockoutDate: "", lastInboundDate: "", lastOutboundDate: "",
      childCount: 1, totalChildPcs: 10, children: [{ childSku: "04895227934099", childProduct: "Child", unitCode: "PCS", unitName: "PCS", qty: 10, availabilityRatio: 100 }],
    }],
    excluded: [],
    summary: { totalUniqueBundleCartonsInFile: 1, availableBundleCartons: 1, unavailableBundleCartons: 0, totalChildSkuLinesInAvailableBundles: 1, totalChildPcsInAvailableBundles: 10 },
    reconciliation: { uniqueBundleSkusFound: 1, matchedInInventory: 1, availableQtyGt0: 1, availableQtyLte0: 0, notFoundInInventory: 0, invalidOrMissingIdentifiers: 0, excludedByWarehouseFilter: 0 },
  };
  mod.buildExportWorkbook(fakeResult); // must not throw
  assert.strictEqual(typeof fakeResult.available[0].bundleSku, "string");
  assert.strictEqual(fakeResult.available[0].bundleSku, "04895227934032"); // leading zero preserved
  assert.strictEqual(typeof fakeResult.available[0].children[0].childSku, "string");
  delete global.XLSX; delete global.Blob;
});

// ---- extra: Bundle SKU normalization (case + whitespace + trailing .0) ----
test("Extra: normalizeBundleSkuKey is case-insensitive and whitespace/'.0'-safe", function () {
  const a = mod._internal.normalizeBundleSkuKey("TP-251223-09");
  const b = mod._internal.normalizeBundleSkuKey("tp-251223-09");
  const c = mod._internal.normalizeBundleSkuKey(" TP-251223-09 ");
  const d = mod._internal.normalizeBundleSkuKey("TP-251223-09.0");
  assert.strictEqual(a, b);
  assert.strictEqual(a, c);
  assert.strictEqual(a, d);
});

// ---- extra: column detection against the real uploaded header rows ----
test("Extra: detectColumns resolves real Bundle file headers (Bundle SKU / Child SKU / SL)", function () {
  const header = ["Bundle SKU (Barcode)", "Tên Bundle", "Child SKU (Barcode)", "Tên sản phẩm con", "Mã ĐVT", "Tên ĐVT", "SL", "Tỷ Lệ Khả Dụng (%)"];
  const det = mod._internal.detectColumns(header, mod.BUNDLE_COLUMN_DEFS);
  assert.strictEqual(det.unresolvedRequired.length, 0);
  assert.strictEqual(det.mapping.bundleSku, 0);
  assert.strictEqual(det.mapping.childSku, 2);
  assert.strictEqual(det.mapping.unitCode, 4);
  assert.strictEqual(det.mapping.unitName, 5);
  assert.strictEqual(det.mapping.qty, 6);
  assert.strictEqual(det.mapping.availabilityRatio, 7);
});

test("Extra: detectColumns resolves the REAL uploaded inventory export headers", function () {
  const header = ["#", "Kho", "Mã sản phẩm", "Mã SKU đối tác", "Tên sản phẩm", "Danh mục", "ĐVT", "Tình trạng hàng hoá",
    "Tồn kho", "Phong tỏa", "Khả dụng", "Đi đường", "Chờ nhập", "Chờ xuất", "Khối lượng (Kg)", "CBM",
    "Ngày hết tồn gần nhất", "Ngày nhập kho gần nhất", "Ngày xuất kho gần nhất"];
  const det = mod._internal.detectColumns(header, mod.INVENTORY_COLUMN_DEFS);
  assert.strictEqual(det.unresolvedRequired.length, 0);
  assert.strictEqual(det.mapping.invProductCode, 2);
  assert.strictEqual(det.mapping.partnerSku, 3);
  assert.strictEqual(det.mapping.freezeQty, 9);
  assert.strictEqual(det.mapping.availableQty, 10);
  assert.strictEqual(det.mapping.lastStockoutDate, 16);
  assert.strictEqual(det.mapping.lastInboundDate, 17);
  assert.strictEqual(det.mapping.lastOutboundDate, 18);
});

test("Extra: parseBundleFile groups multiple child rows under one Bundle SKU (real reference example)", function () {
  const rows = [
    ["Bundle SKU (Barcode)", "Tên Bundle", "Child SKU (Barcode)", "Tên sản phẩm con", "Mã ĐVT", "Tên ĐVT", "SL", "Tỷ Lệ Khả Dụng (%)"],
    ["TP-251223-03-140", "Bundle X", "4895227934032", "6.0mm Rope Strap", "PCS", "PCS", 100, 100],
    ["TP-251223-03-140", "Bundle X", "4894961070310", "Bungee Strap", "PCS", "PCS", 10, 100],
    ["TP-251223-03-140", "Bundle X", "4895227950230", "Bungee Wrist Strap", "PCS", "PCS", 10, 100],
    ["TP-251223-03-140", "Bundle X", "4894961034947", "Bungee Wrist Strap", "PCS", "PCS", 20, 100],
  ];
  const parsed = mod.parseBundleFile(rows);
  assert.strictEqual(parsed.childRows.length, 4);
  assert.strictEqual(parsed.bundles.length, 1); // grouped into ONE bundle carton, not 4
  assert.strictEqual(parsed.bundles[0].bundleSku, "TP-251223-03-140");
  assert.strictEqual(parsed.bundles[0].children.length, 4);
  assert.strictEqual(parsed.bundles[0].totalChildPcs, 140);
});

console.log("\n=== " + passed + " passed, " + failed + " failed ===\n");
process.exit(failed ? 1 : 0);
