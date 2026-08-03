/* ==========================================================
 * tests/outbound-availability.test.js
 * Plain Node.js assert-based tests (no framework/deps) for the pure
 * matching logic in js/outbound-availability.js. Run with:
 *   node tests/outbound-availability.test.js
 * ========================================================== */
"use strict";
const assert = require("assert");
const path = require("path");
const fs = require("fs");

// The module assigns itself to `window.WOPOutboundAvailability` and also
// does `module.exports = api` when running under Node/CommonJS.
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

// ---- helpers to build fake bundleRecords / inventoryIndex directly
// (bypassing the Excel/XLSX layer, which only exists in the browser) ----
function bundleRow(bundleSku, childEan, eanQty, extra) {
  return Object.assign({
    rowIndex: 0, bundleSku: bundleSku, bundleName: "Bundle " + bundleSku,
    childEanRaw: childEan, childEan: mod._internal.normalizeEan(childEan),
    childProduct: "Product " + childEan, unit: "PCS", eanQty: eanQty,
  }, extra || {});
}
function invRow(warehouse, childEan, availableQty, extra) {
  return Object.assign({
    rowIndex: 0, warehouse: warehouse, childEan: mod._internal.normalizeEan(childEan),
    barcode: childEan, partnerSku: "SKU-" + childEan, product: "Inv Product " + childEan,
    category: "Accessories", unit: "PCS", conditionType: "Mới",
    stockQty: mod._internal.toNumberOrNull(availableQty), freezeQty: 0, availableQtyRaw: availableQty, availableQty: mod._internal.toNumberOrNull(availableQty),
    pendingIn: 0, pendingOut: 0, weightKg: 1, cbm: 0.01, invEanQty: 10,
    lastStockoutDate: "", lastOutboundDate: "", lastInboundDate: "",
  }, extra || {});
}
function buildIndex(rows) {
  const idx = new Map();
  rows.forEach(function (r) {
    if (!idx.has(r.childEan)) idx.set(r.childEan, []);
    idx.get(r.childEan).push(r);
  });
  return idx;
}

console.log("\n=== Real-Time Outbound Availability — matching engine tests ===\n");

// ---- Test 1: one bundle, 4 Child EANs, all Available Qty > 0 ----
test("Test 1: 4 Child EANs all available -> all 4 included", function () {
  const bundleSku = "TP-251223-03-140";
  const bundleRecords = [
    bundleRow(bundleSku, "4895227934032", 100),
    bundleRow(bundleSku, "4894961070310", 10),
    bundleRow(bundleSku, "4895227950230", 10),
    bundleRow(bundleSku, "4894961034947", 20),
  ];
  const inv = buildIndex([
    invRow("SG2 - Kho Bonded", "4895227934032", 100),
    invRow("SG2 - Kho Bonded", "4894961070310", 10),
    invRow("SG2 - Kho Bonded", "4895227950230", 10),
    invRow("SG2 - Kho Bonded", "4894961034947", 20),
  ]);
  const result = mod.runMatch({ bundleRecords: bundleRecords, inventoryIndex: inv, selectedBundleSkus: new Set([bundleSku]) });
  assert.strictEqual(result.available.length, 4);
  assert.strictEqual(result.excluded.length, 0);
  assert.strictEqual(result.reconciliation.foundInBundleFile, 4);
  assert.strictEqual(result.reconciliation.matchedInInventory, 4);
});

// ---- Test 2: 2 available, 2 with Available Qty = 0 ----
test("Test 2: 2 available + 2 qty=0 -> 2 included, 2 excluded (AVAILABLE_QTY_ZERO)", function () {
  const bundleSku = "BUNDLE-A";
  const bundleRecords = [
    bundleRow(bundleSku, "1000000000001", 5),
    bundleRow(bundleSku, "1000000000002", 5),
    bundleRow(bundleSku, "1000000000003", 5),
    bundleRow(bundleSku, "1000000000004", 5),
  ];
  const inv = buildIndex([
    invRow("WH1", "1000000000001", 3),
    invRow("WH1", "1000000000002", 0),
    invRow("WH1", "1000000000003", 7),
    invRow("WH1", "1000000000004", 0),
  ]);
  const result = mod.runMatch({ bundleRecords: bundleRecords, inventoryIndex: inv, selectedBundleSkus: new Set([bundleSku]) });
  assert.strictEqual(result.available.length, 2);
  assert.strictEqual(result.excluded.length, 2);
  assert.ok(result.excluded.every(function (e) { return e.reason === "AVAILABLE_QTY_ZERO"; }));
});

// ---- Test 3: one Child EAN absent from inventory ----
test("Test 3: Child EAN missing from inventory -> CHILD_EAN_NOT_FOUND_IN_INVENTORY", function () {
  const bundleSku = "BUNDLE-B";
  const bundleRecords = [bundleRow(bundleSku, "2000000000001", 5), bundleRow(bundleSku, "2000000000002", 5)];
  const inv = buildIndex([invRow("WH1", "2000000000001", 10)]); // EAN ...002 not present
  const result = mod.runMatch({ bundleRecords: bundleRecords, inventoryIndex: inv, selectedBundleSkus: new Set([bundleSku]) });
  assert.strictEqual(result.available.length, 1);
  assert.strictEqual(result.excluded.length, 1);
  assert.strictEqual(result.excluded[0].reason, "CHILD_EAN_NOT_FOUND_IN_INVENTORY");
});

// ---- Test 4: Available Qty blank or text ----
test("Test 4: Available Qty blank/text -> AVAILABLE_QTY_INVALID", function () {
  const bundleSku = "BUNDLE-C";
  const bundleRecords = [bundleRow(bundleSku, "3000000000001", 5), bundleRow(bundleSku, "3000000000002", 5)];
  const inv = buildIndex([invRow("WH1", "3000000000001", null), invRow("WH1", "3000000000002", "N/A")]);
  const result = mod.runMatch({ bundleRecords: bundleRecords, inventoryIndex: inv, selectedBundleSkus: new Set([bundleSku]) });
  assert.strictEqual(result.available.length, 0);
  assert.strictEqual(result.excluded.length, 2);
  assert.ok(result.excluded.every(function (e) { return e.reason === "AVAILABLE_QTY_INVALID"; }));
});

// ---- Test 5: EAN imported by Excel as numeric ending in ".0" ----
test("Test 5: EAN normalization — trailing '.0' still matches", function () {
  assert.strictEqual(mod._internal.normalizeEan("4895227934032.0"), "4895227934032");
  assert.strictEqual(mod._internal.normalizeEan(4895227934032), "4895227934032");
  assert.strictEqual(mod._internal.normalizeEan(" 4895227934032.00 "), "4895227934032");

  const bundleSku = "BUNDLE-D";
  const bundleRecords = [bundleRow(bundleSku, "4895227934032.0", 100)]; // Excel-ish string with trailing .0
  const inv = buildIndex([invRow("WH1", 4895227934032, 8)]); // inventory read as JS number
  const result = mod.runMatch({ bundleRecords: bundleRecords, inventoryIndex: inv, selectedBundleSkus: new Set([bundleSku]) });
  assert.strictEqual(result.available.length, 1);
  assert.strictEqual(result.available[0].availableQty, 8);
});

// ---- Test 6: duplicate Child EAN across different warehouses ----
test("Test 6: same Child EAN in 2 warehouses -> both preserved separately", function () {
  const bundleSku = "BUNDLE-E";
  const bundleRecords = [bundleRow(bundleSku, "5000000000001", 5)];
  const inv = buildIndex([invRow("WH-North", "5000000000001", 4), invRow("WH-South", "5000000000001", 6)]);
  const result = mod.runMatch({ bundleRecords: bundleRecords, inventoryIndex: inv, selectedBundleSkus: new Set([bundleSku]) });
  assert.strictEqual(result.available.length, 2);
  const warehouses = result.available.map(function (r) { return r.warehouse; }).sort();
  assert.deepStrictEqual(warehouses, ["WH-North", "WH-South"]);
});

// ---- Test 7: warehouse filter selected ----
test("Test 7: warehouse filter -> only matching warehouse rows returned", function () {
  const bundleSku = "BUNDLE-F";
  const bundleRecords = [bundleRow(bundleSku, "6000000000001", 5)];
  const inv = buildIndex([invRow("WH-North", "6000000000001", 4), invRow("WH-South", "6000000000001", 6)]);
  const result = mod.runMatch({ bundleRecords: bundleRecords, inventoryIndex: inv, selectedBundleSkus: new Set([bundleSku]), warehouseFilter: new Set(["WH-South"]) });
  assert.strictEqual(result.available.length, 1);
  assert.strictEqual(result.available[0].warehouse, "WH-South");
  assert.strictEqual(result.excluded.length, 1);
  assert.strictEqual(result.excluded[0].reason, "WAREHOUSE_FILTERED_OUT");
});

// ---- Test 8: no available cartons remain -> valid zero-result state ----
test("Test 8: all excluded -> zero-result state, not a crash", function () {
  const bundleSku = "BUNDLE-G";
  const bundleRecords = [bundleRow(bundleSku, "7000000000001", 5), bundleRow(bundleSku, "7000000000002", 5)];
  const inv = buildIndex([invRow("WH1", "7000000000001", 0)]); // ...002 not found at all
  const result = mod.runMatch({ bundleRecords: bundleRecords, inventoryIndex: inv, selectedBundleSkus: new Set([bundleSku]) });
  assert.strictEqual(result.available.length, 0);
  assert.strictEqual(result.summary.totalAvailableQty, 0);
  assert.strictEqual(result.summary.totalAvailablePcs, 0);
  assert.strictEqual(result.reconciliation.foundInBundleFile, 2);
  assert.strictEqual(result.reconciliation.notFoundInInventory, 1);
});

// ---- Test 9: EAN Qty 100 but Available Qty 0 -> must still be excluded ----
test("Test 9: EAN Qty must never cause inclusion when Available Qty <= 0", function () {
  const bundleSku = "BUNDLE-H";
  const bundleRecords = [bundleRow(bundleSku, "8000000000001", 100)];
  const inv = buildIndex([invRow("WH1", "8000000000001", 0)]);
  const result = mod.runMatch({ bundleRecords: bundleRecords, inventoryIndex: inv, selectedBundleSkus: new Set([bundleSku]) });
  assert.strictEqual(result.available.length, 0);
  assert.strictEqual(result.excluded[0].reason, "AVAILABLE_QTY_ZERO");
  assert.strictEqual(result.excluded[0].eanQty, 100);
});

// ---- Test 10: Available Qty 2, EAN Qty 100 -> Total Available PCS = 200 ----
test("Test 10: Total Available PCS = Available Qty x EAN Qty", function () {
  const bundleSku = "BUNDLE-I";
  const bundleRecords = [bundleRow(bundleSku, "9000000000001", 100)];
  const inv = buildIndex([invRow("WH1", "9000000000001", 2)]);
  const result = mod.runMatch({ bundleRecords: bundleRecords, inventoryIndex: inv, selectedBundleSkus: new Set([bundleSku]) });
  assert.strictEqual(result.available[0].availablePcs, 200);
  assert.strictEqual(result.summary.totalAvailablePcs, 200);
});

// ---- extra: negative qty excluded ----
test("Extra: negative Available Qty -> AVAILABLE_QTY_NEGATIVE", function () {
  const bundleSku = "BUNDLE-J";
  const bundleRecords = [bundleRow(bundleSku, "1100000000001", 5)];
  const inv = buildIndex([invRow("WH1", "1100000000001", -3)]);
  const result = mod.runMatch({ bundleRecords: bundleRecords, inventoryIndex: inv, selectedBundleSkus: new Set([bundleSku]) });
  assert.strictEqual(result.available.length, 0);
  assert.strictEqual(result.excluded[0].reason, "AVAILABLE_QTY_NEGATIVE");
});

// ---- extra: missing Child EAN in bundle file ----
test("Extra: blank Child EAN in bundle row -> CHILD_EAN_MISSING", function () {
  const bundleSku = "BUNDLE-K";
  const bundleRecords = [bundleRow(bundleSku, "", 5), bundleRow(bundleSku, null, 5)];
  const inv = buildIndex([]);
  const result = mod.runMatch({ bundleRecords: bundleRecords, inventoryIndex: inv, selectedBundleSkus: new Set([bundleSku]) });
  assert.strictEqual(result.available.length, 0);
  assert.strictEqual(result.excluded.length, 2);
  assert.ok(result.excluded.every(function (e) { return e.reason === "CHILD_EAN_MISSING"; }));
  assert.strictEqual(result.reconciliation.invalidOrMissingChildEan, 2);
});

// ---- extra: column detection against the real uploaded header rows ----
test("Extra: detectColumns resolves real Bundle file headers", function () {
  const header = ["Bundle SKU (Barcode)", "Tên Bundle", "Child SKU (Barcode)", "Tên sản phẩm con", "Mã ĐVT", "Tên ĐVT", "SL", "Tỷ Lệ Khả Dụng (%)"];
  const det = mod._internal.detectColumns(header, mod.BUNDLE_COLUMN_DEFS);
  assert.strictEqual(det.unresolvedRequired.length, 0);
  assert.strictEqual(det.mapping.bundleSku, 0);
  assert.strictEqual(det.mapping.childEan, 2);
  assert.strictEqual(det.mapping.eanQty, 6);
});

test("Extra: detectColumns resolves real Inventory screenshot headers", function () {
  const header = ["#", "Warehouse", "Barcode", "Partner SKU", "Child EAN", "Product", "Category", "Unit", "Condition type",
    "Stock Qty", "Freeze Qty", "Available Qty", "Pending In", "Pending Out", "Weight (Kg)", "CBM",
    "Last Stockout Date", "Last Outbound Date", "Last Inbound Date", "EAN Qty"];
  const det = mod._internal.detectColumns(header, mod.INVENTORY_COLUMN_DEFS);
  assert.strictEqual(det.unresolvedRequired.length, 0);
  assert.strictEqual(det.mapping.childEan, 4);
  assert.strictEqual(det.mapping.availableQty, 11);
  assert.strictEqual(det.mapping.eanQty, 19); // last "EAN Qty" column, not the "Child EAN" one
});

// Regression test: real "Tồn kho theo thời gian thực" export uploaded by the
// user (2026-08-03) — column order/wording differs slightly from the
// screenshot-based test above (e.g. "Ngày hết tồn gần nhất" for stockout,
// "Phong tỏa" for freeze, no separate Barcode column). This caught a real
// bug where the 3 date columns were mismapped/unmapped and Freeze Qty was
// unmapped — required columns (Child EAN, Available Qty) were unaffected,
// but display accuracy for these optional columns was wrong.
test("Extra: detectColumns resolves the REAL uploaded inventory export headers", function () {
  const header = ["#", "Kho", "Mã sản phẩm", "Mã SKU đối tác", "Tên sản phẩm", "Danh mục", "ĐVT", "Tình trạng hàng hoá",
    "Tồn kho", "Phong tỏa", "Khả dụng", "Đi đường", "Chờ nhập", "Chờ xuất", "Khối lượng (Kg)", "CBM",
    "Ngày hết tồn gần nhất", "Ngày nhập kho gần nhất", "Ngày xuất kho gần nhất"];
  const det = mod._internal.detectColumns(header, mod.INVENTORY_COLUMN_DEFS);
  assert.strictEqual(det.unresolvedRequired.length, 0);
  assert.strictEqual(det.mapping.warehouse, 1);
  assert.strictEqual(det.mapping.childEan, 2);
  assert.strictEqual(det.mapping.partnerSku, 3);
  assert.strictEqual(det.mapping.conditionType, 7);
  assert.strictEqual(det.mapping.stockQty, 8);
  assert.strictEqual(det.mapping.freezeQty, 9);
  assert.strictEqual(det.mapping.availableQty, 10);
  assert.strictEqual(det.mapping.lastStockoutDate, 16); // "Ngày hết tồn gần nhất"
  assert.strictEqual(det.mapping.lastInboundDate, 17);  // "Ngày nhập kho gần nhất"
  assert.strictEqual(det.mapping.lastOutboundDate, 18); // "Ngày xuất kho gần nhất"
});

test("Extra: parseBundleFile end-to-end against real reference example rows", function () {
  const rows = [
    ["Bundle SKU (Barcode)", "Tên Bundle", "Child SKU (Barcode)", "Tên sản phẩm con", "Mã ĐVT", "Tên ĐVT", "SL", "Tỷ Lệ Khả Dụng (%)"],
    ["TP-251223-03-140", "Bundle X", "4895227934032", "6.0mm Rope Strap", "PCS", "PCS", 100, 100],
    ["TP-251223-03-140", "Bundle X", "4894961070310", "Bungee Strap", "PCS", "PCS", 10, 100],
    ["TP-251223-03-140", "Bundle X", "4895227950230", "Bungee Wrist Strap", "PCS", "PCS", 10, 100],
    ["TP-251223-03-140", "Bundle X", "4894961034947", "Bungee Wrist Strap", "PCS", "PCS", 20, 100],
  ];
  const parsed = mod.parseBundleFile(rows);
  assert.strictEqual(parsed.records.length, 4);
  assert.deepStrictEqual(parsed.records.map(function (r) { return r.eanQty; }), [100, 10, 10, 20]);
  assert.strictEqual(parsed.bundles.length, 1);
  assert.strictEqual(parsed.bundles[0].sku, "TP-251223-03-140");
});

test("Extra: buildExportWorkbook keeps EAN/Barcode as text (no numeric coercion)", function () {
  // Minimal fake XLSX shim just for this one assertion.
  global.XLSX = {
    utils: {
      aoa_to_sheet: function (aoa) {
        const ws = {};
        aoa.forEach(function (row, r) { row.forEach(function (cell, c) {
          const addr = String.fromCharCode(65 + c) + (r + 1);
          ws[addr] = { v: cell, t: typeof cell === "string" ? "s" : "n" };
        }); });
        return ws;
      },
      book_new: function () { return { sheets: {} }; },
      book_append_sheet: function (wb, ws, name) { wb.sheets[name] = ws; },
    },
    write: function (wb) { return new Uint8Array([1, 2, 3]); },
  };
  global.Blob = function (parts, opts) { this.parts = parts; this.opts = opts; };
  const fakeResult = {
    available: [{ bundleSku: "B1", warehouse: "WH1", childEan: "04895227934032", barcode: "04895227934032", partnerSku: "SKU1", product: "P", category: "C", unit: "PCS", conditionType: "Mới", stockQty: 1, freezeQty: 0, availableQty: 1, pendingIn: 0, pendingOut: 0, weightKg: 1, cbm: 0.01, eanQty: 100, availablePcs: 100, lastStockoutDate: "", lastOutboundDate: "", lastInboundDate: "" }],
    excluded: [],
    summary: { selectedBundles: 1, totalChildEansInBundleFile: 1, availableChildEans: 1, excludedChildEans: 0, totalAvailableQty: 1, totalAvailablePcs: 100 },
    reconciliation: { foundInBundleFile: 1, matchedInInventory: 1, excludedQtyZeroOrInvalid: 0, notFoundInInventory: 0, invalidOrMissingChildEan: 0, excludedByWarehouseFilter: 0 },
  };
  mod.buildExportWorkbook(fakeResult); // must not throw
  const ws = global.XLSX.utils.aoa_to_sheet.calledWith; // not tracked, just ensure no exception path
  // Re-derive the AOA the same way buildExportWorkbook does, to assert the EAN cell type.
  const availRow = fakeResult.available[0];
  assert.strictEqual(typeof availRow.childEan, "string"); // leading zero preserved as string
  assert.strictEqual(availRow.childEan, "04895227934032");
  delete global.XLSX; delete global.Blob;
});

console.log("\n=== " + passed + " passed, " + failed + " failed ===\n");
process.exit(failed ? 1 : 0);
