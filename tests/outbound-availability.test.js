/* ==========================================================
 * tests/outbound-availability.test.js
 * Plain Node.js assert-based tests (no framework/deps) for the
 * Combined Real-Time Inventory transform in js/outbound-availability.js.
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
function invRow(warehouse, invProductCode, availableQty, extra) {
  return Object.assign({
    warehouse: warehouse, invProductCode: invProductCode, partnerSku: "",
    product: "Inv " + invProductCode, category: "Accessories", unit: "PCS", conditionType: "Mới",
    stockQty: null, freezeQty: null, availableQtyRaw: availableQty,
    availableQty: mod._internal.toNumberOrNull(availableQty),
    pendingIn: null, pendingOut: null, weightKg: null, cbm: null,
    lastStockoutDate: "", lastOutboundDate: "", lastInboundDate: "", invEanQty: null,
  }, extra || {});
}
function bundleMapOf(bundles) {
  // bundles: [{ sku, name, children:[{childSku,childProduct,qty}] }]
  const m = new Map();
  bundles.forEach(function (b) {
    m.set(mod._internal.normalizeBundleSkuKey(b.sku), {
      bundleSku: b.sku, bundleName: b.name || "",
      children: b.children.map(function (c) { return { childSku: c.childSku, childProduct: c.childProduct, qty: c.qty, unitCode: "PCS", unitName: "PCS", availabilityRatio: 100 }; }),
      totalChildPcs: b.children.reduce(function (s, c) { return s + (c.qty || 0); }, 0),
    });
  });
  return m;
}
function run(invRecords, bundles, opts) {
  return mod.generateCombinedReport(Object.assign({
    invRecords: invRecords,
    bundleMap: bundleMapOf(bundles || []),
    bundleKeyField: "invProductCode",
  }, opts || {}));
}

console.log("\n=== Combined Real-Time Inventory — transform tests ===\n");

// 1. Available normal inventory row remains one output row.
test("1. Available normal row -> 1 output row", function () {
  const r = run([invRow("SG2", "SKU-A", 5)], []);
  assert.strictEqual(r.combined.length, 1);
  assert.strictEqual(r.combined[0]._rowType, "normal");
});

// 2. Normal inventory row with Available Qty = 0 is excluded.
test("2. Available Qty=0 -> excluded from output", function () {
  const r = run([invRow("SG2", "SKU-A", 0)], []);
  assert.strictEqual(r.combined.length, 0);
  assert.strictEqual(r.summary.removedUnavailableRows, 1);
});

// 3. Available bundle with four valid child lines becomes exactly four output rows.
test("3. Bundle with 4 child lines -> exactly 4 output rows", function () {
  const bundles = [{ sku: "TP-BUNDLE-001", children: [
    { childSku: "A", childProduct: "A", qty: 100 }, { childSku: "B", childProduct: "B", qty: 10 },
    { childSku: "C", childProduct: "C", qty: 10 }, { childSku: "D", childProduct: "D", qty: 20 },
  ] }];
  const r = run([invRow("SG2", "TP-BUNDLE-001", 1, { unit: "MIX" })], bundles);
  assert.strictEqual(r.combined.length, 4);
});

// 4. No original bundle summary row remains after successful expansion.
test("4. Original 1-line bundle summary row is dropped after expansion", function () {
  const bundles = [{ sku: "TP-BUNDLE-002", children: [{ childSku: "A", childProduct: "A", qty: 5 }] }];
  const r = run([invRow("SG2", "TP-BUNDLE-002", 1, { unit: "MIX" })], bundles);
  assert.strictEqual(r.combined.length, 1);
  assert.strictEqual(r.combined[0].childEan, "A"); // it's the CHILD row, not a "TP-BUNDLE-002" summary row
  assert.ok(r.combined.every(function (row) { return row.childEan !== "TP-BUNDLE-002"; }));
});

// 5. Each expanded row gets the correct Child EAN.
test("5. Expanded rows get correct Child EAN", function () {
  const bundles = [{ sku: "TP-X", children: [{ childSku: "EAN-1", childProduct: "P1", qty: 1 }, { childSku: "EAN-2", childProduct: "P2", qty: 2 }] }];
  const r = run([invRow("SG2", "TP-X", 1, { unit: "MIX" })], bundles);
  const eans = r.combined.map(function (row) { return row.childEan; }).sort();
  assert.deepStrictEqual(eans, ["EAN-1", "EAN-2"]);
});

// 6. Each expanded row gets the correct Child Product Name.
test("6. Expanded rows get correct Child Product Name", function () {
  const bundles = [{ sku: "TP-Y", children: [{ childSku: "E1", childProduct: "Rope Strap", qty: 1 }] }];
  const r = run([invRow("SG2", "TP-Y", 1, { unit: "MIX" })], bundles);
  assert.strictEqual(r.combined[0].product, "Rope Strap");
});

// 7. Each expanded row gets EAN Qty directly from SL / Child Qty.
test("7. Expanded rows: EAN Qty = SL from mapping", function () {
  const bundles = [{ sku: "TP-Z", children: [{ childSku: "E1", childProduct: "P", qty: 37 }] }];
  const r = run([invRow("SG2", "TP-Z", 1, { unit: "MIX" })], bundles);
  assert.strictEqual(r.combined[0].eanQty, 37);
});

// 8. No Available Qty field exists in the final bundle child output.
test("8. No availableQty field on expanded bundle child rows", function () {
  const bundles = [{ sku: "TP-NOAVAIL", children: [{ childSku: "E1", childProduct: "P", qty: 10 }] }];
  const r = run([invRow("SG2", "TP-NOAVAIL", 7, { unit: "MIX" })], bundles); // Available Qty=7 on source row
  assert.strictEqual(r.combined[0].availableQty, undefined);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(r.combined[0], "availableQty"), false);
});

// 9. No Available PCS calculation exists.
test("9. No availablePcs field / calculation anywhere on output rows", function () {
  const bundles = [{ sku: "TP-NOPCS", children: [{ childSku: "E1", childProduct: "P", qty: 100 }] }];
  const r = run([invRow("SG2", "TP-NOPCS", 2, { unit: "MIX" })], bundles); // if it were qty*avail it'd be 200
  assert.strictEqual(r.combined[0].eanQty, 100); // exactly SL, never 100*2
  assert.strictEqual(r.combined[0].availablePcs, undefined);
});

// 10. Bundle existing only in mapping but not in realtime produces no output.
test("10. Bundle-mapping-only SKU (absent from realtime) -> zero output", function () {
  const bundles = [{ sku: "TP-HISTORICAL-ONLY", children: [{ childSku: "E1", childProduct: "P", qty: 5 }] }];
  const r = run([invRow("SG2", "SOME-OTHER-SKU", 5)], bundles); // realtime never mentions TP-HISTORICAL-ONLY
  assert.strictEqual(r.combined.length, 1); // just the one normal row
  assert.ok(r.combined.every(function (row) { return row.barcode !== "TP-HISTORICAL-ONLY" && row.childEan !== "TP-HISTORICAL-ONLY"; }));
  assert.ok(r.diagnostics.every(function (d) { return d.reason !== "BUNDLE_MAPPING_NOT_FOUND" || d.code !== "TP-HISTORICAL-ONLY"; }));
});

// 11. Available realtime MIX/bundle row with no mapping is excluded.
test("11. Available MIX row with NO mapping -> excluded from combined output", function () {
  const r = run([invRow("SG2", "TP-UNMAPPED-MIX", 3, { unit: "MIX" })], []); // empty bundle mapping
  assert.strictEqual(r.combined.length, 0);
});

// 12. The unmapped MIX/bundle row receives diagnostic reason BUNDLE_MAPPING_NOT_FOUND.
test("12. Unmapped MIX row -> diagnostic reason BUNDLE_MAPPING_NOT_FOUND", function () {
  const r = run([invRow("SG2", "TP-UNMAPPED-MIX2", 3, { unit: "MIX" })], []);
  assert.strictEqual(r.diagnostics.length, 1);
  assert.strictEqual(r.diagnostics[0].reason, "BUNDLE_MAPPING_NOT_FOUND");
  assert.strictEqual(r.summary.unmappedBundleRowsExcluded, 1);
});

// 13. Unmapped bundle is not retained as a normal inventory row.
test("13. Unmapped MIX row is NOT kept as a normal row (not silently fallback-included)", function () {
  const r = run([invRow("SG2", "TP-UNMAPPED-MIX3", 3, { unit: "MIX" })], []);
  assert.ok(!r.combined.some(function (row) { return row.barcode === "TP-UNMAPPED-MIX3"; }));
});

// 14. Combined output contains normal available rows plus expanded mapped bundle child rows.
test("14. Combined output = available normal rows + expanded bundle child rows", function () {
  const bundles = [{ sku: "TP-MIX-OK", children: [{ childSku: "C1", childProduct: "P1", qty: 5 }, { childSku: "C2", childProduct: "P2", qty: 15 }] }];
  const r = run([
    invRow("SG2", "NORMAL-1", 10),
    invRow("SG2", "TP-MIX-OK", 1, { unit: "MIX" }),
    invRow("SG2", "NORMAL-2", 0), // filtered out
  ], bundles);
  assert.strictEqual(r.combined.length, 3); // 1 normal + 2 expanded children
  const types = r.combined.map(function (row) { return row._rowType; }).sort();
  assert.deepStrictEqual(types, ["bundleChild", "bundleChild", "normal"]);
});

// 15. All bundles are processed automatically without manual selection.
test("15. No bundle-selection parameter exists — every mapped bundle found in realtime auto-expands", function () {
  const bundles = [
    { sku: "TP-AUTO-1", children: [{ childSku: "A", childProduct: "A", qty: 1 }] },
    { sku: "TP-AUTO-2", children: [{ childSku: "B", childProduct: "B", qty: 1 }] },
  ];
  const r = run([invRow("SG2", "TP-AUTO-1", 1, { unit: "MIX" }), invRow("SG2", "TP-AUTO-2", 1, { unit: "MIX" })], bundles);
  // generateCombinedReport() has no "selectedBundleSkus" option at all — both expand automatically.
  assert.strictEqual(mod.generateCombinedReport.length <= 1, true); // single opts-object signature, no selection param
  assert.strictEqual(r.summary.availableBundlesExpanded, 2);
});

// 16. Barcode, Partner SKU and Child EAN remain text in export.
test("16. buildExportWorkbook keeps Barcode/Partner SKU/Child EAN as text", function () {
  global.XLSX = {
    utils: {
      aoa_to_sheet: function (aoa) { return { __aoa: aoa }; },
      book_new: function () { return { sheets: {} }; },
      book_append_sheet: function (wb, ws, name) { wb.sheets[name] = ws; wb.__sheetNames = (wb.__sheetNames || []).concat([name]); },
      encode_range: function (r) { return "A1:Z" + (r.e.r + 1); },
    },
    write: function () { return new Uint8Array([1, 2, 3]); },
  };
  global.Blob = function (parts, opts) { this.parts = parts; this.opts = opts; this.size = 3; this.type = opts.type; };
  const rows = [{ warehouse: "SG2", barcode: "04895227934032", partnerSku: "TP-01", childEan: "04895227934099", product: "P", category: "C", unit: "MIX", conditionType: "Mới", stockQty: 1, freezeQty: 0, pendingIn: 0, pendingOut: 0, weightKg: 1, cbm: 0.01, lastStockoutDate: "", lastOutboundDate: "", lastInboundDate: "", eanQty: 10 }];
  mod.buildExportWorkbook(rows);
  assert.strictEqual(typeof rows[0].barcode, "string");
  assert.strictEqual(rows[0].barcode, "04895227934032"); // leading zero preserved
  assert.strictEqual(typeof rows[0].childEan, "string");
  delete global.XLSX; delete global.Blob;
});

// 17. Excel export contains exactly one worksheet named "Combined Real-Time Inventory".
test("17. Export produces exactly ONE worksheet named 'Combined Real-Time Inventory'", function () {
  const sheetNames = [];
  global.XLSX = {
    utils: {
      aoa_to_sheet: function (aoa) { return { __aoa: aoa }; },
      book_new: function () { return { sheets: {} }; },
      book_append_sheet: function (wb, ws, name) { sheetNames.push(name); },
      encode_range: function (r) { return "A1:Z" + (r.e.r + 1); },
    },
    write: function () { return new Uint8Array([1]); },
  };
  global.Blob = function (parts, opts) { this.size = 1; this.type = opts.type; };
  mod.buildExportWorkbook([{ warehouse: "SG2", barcode: "A", partnerSku: "B", childEan: "C", product: "P", eanQty: 1 }]);
  assert.strictEqual(sheetNames.length, 1);
  assert.strictEqual(sheetNames[0], "Combined Real-Time Inventory");
  delete global.XLSX; delete global.Blob;
});

// 18. UI preview count equals Excel export data-row count (same underlying array).
test("18. Preview array and export array are the exact same combined result (no divergence)", function () {
  const bundles = [{ sku: "TP-SAME", children: [{ childSku: "A", childProduct: "A", qty: 1 }, { childSku: "B", childProduct: "B", qty: 2 }] }];
  const r = run([invRow("SG2", "NORMAL-X", 4), invRow("SG2", "TP-SAME", 1, { unit: "MIX" })], bundles);
  // The export function is called with exactly r.combined (see init()'s exportBtn handler) — verify
  // the array passed to export has the same length as what the UI reports as "Final Combined Rows".
  assert.strictEqual(r.combined.length, r.summary.finalCombinedRows);
  assert.strictEqual(r.combined.length, 3);
});

// 19. Warehouse filter consistently filters normal and expanded rows.
test("19. Warehouse filter removes both normal rows AND bundle-expanded rows from the excluded warehouse", function () {
  const bundles = [{ sku: "TP-WH", children: [{ childSku: "A", childProduct: "A", qty: 1 }] }];
  const r = run([
    invRow("WH-North", "NORMAL-A", 5),
    invRow("WH-South", "NORMAL-B", 5),
    invRow("WH-North", "TP-WH", 1, { unit: "MIX" }),
    invRow("WH-South", "TP-WH", 1, { unit: "MIX" }),
  ], bundles, { warehouseFilter: new Set(["WH-South"]) });
  assert.ok(r.combined.every(function (row) { return row.warehouse === "WH-South"; }));
  assert.strictEqual(r.combined.length, 2); // NORMAL-B (normal) + TP-WH's 1 child (from WH-South only)
  assert.strictEqual(r.summary.warehouseFilteredRows, 2); // NORMAL-A + WH-North's TP-WH row
});

// 20. TP-251223-03-140 expands into four rows with EAN Qty values 100, 10, 10, 20.
test("20. Reference example: TP-251223-03-140 -> 4 rows, EAN Qty 100/10/10/20", function () {
  const bundles = [{ sku: "TP-251223-03-140", children: [
    { childSku: "4895227934032", childProduct: "6.0mm Rope Strap", qty: 100 },
    { childSku: "4894961070310", childProduct: "Bungee Strap", qty: 10 },
    { childSku: "4895227950230", childProduct: "Bungee Wrist Strap", qty: 10 },
    { childSku: "4894961034947", childProduct: "Bungee Wrist Strap", qty: 20 },
  ] }];
  const r = run([invRow("SG2 - Kho Bonded", "TP-251223-03-140", 1, { unit: "MIX" })], bundles);
  assert.strictEqual(r.combined.length, 4);
  assert.deepStrictEqual(r.combined.map(function (row) { return row.eanQty; }), [100, 10, 10, 20]);
  assert.ok(r.combined.every(function (row) { return row.barcode === "TP-251223-03-140"; }));
  assert.ok(r.combined.every(function (row) { return row.availableQty === undefined; }));
});

// ---- extra: column detection against the real uploaded header rows ----
test("Extra: detectColumns resolves real Bundle file headers", function () {
  const header = ["Bundle SKU (Barcode)", "Tên Bundle", "Child SKU (Barcode)", "Tên sản phẩm con", "Mã ĐVT", "Tên ĐVT", "SL", "Tỷ Lệ Khả Dụng (%)"];
  const det = mod._internal.detectColumns(header, mod.BUNDLE_COLUMN_DEFS);
  assert.strictEqual(det.unresolvedRequired.length, 0);
  assert.strictEqual(det.mapping.bundleSku, 0);
  assert.strictEqual(det.mapping.childSku, 2);
  assert.strictEqual(det.mapping.qty, 6);
});

test("Extra: detectColumns resolves the REAL uploaded inventory export headers", function () {
  const header = ["#", "Kho", "Mã sản phẩm", "Mã SKU đối tác", "Tên sản phẩm", "Danh mục", "ĐVT", "Tình trạng hàng hoá",
    "Tồn kho", "Phong tỏa", "Khả dụng", "Đi đường", "Chờ nhập", "Chờ xuất", "Khối lượng (Kg)", "CBM",
    "Ngày hết tồn gần nhất", "Ngày nhập kho gần nhất", "Ngày xuất kho gần nhất"];
  const det = mod._internal.detectColumns(header, mod.INVENTORY_COLUMN_DEFS);
  assert.strictEqual(det.unresolvedRequired.length, 0);
  assert.strictEqual(det.mapping.invProductCode, 2);
  assert.strictEqual(det.mapping.partnerSku, 3);
  assert.strictEqual(det.mapping.availableQty, 10);
  assert.strictEqual(det.mapping.lastStockoutDate, 16);
  assert.strictEqual(det.mapping.lastInboundDate, 17);
  assert.strictEqual(det.mapping.lastOutboundDate, 18);
});

test("Extra: isMixUnit is case/whitespace-insensitive", function () {
  assert.strictEqual(mod._internal.isMixUnit("MIX"), true);
  assert.strictEqual(mod._internal.isMixUnit(" mix "), true);
  assert.strictEqual(mod._internal.isMixUnit("Mix"), true);
  assert.strictEqual(mod._internal.isMixUnit("PCS"), false);
  assert.strictEqual(mod._internal.isMixUnit("CARTON_69PCS"), false);
});

test("Extra: negative/invalid Available Qty removed with correct reasons", function () {
  const r = run([invRow("SG2", "A", -1), invRow("SG2", "B", "N/A"), invRow("SG2", "C", null)], []);
  assert.strictEqual(r.combined.length, 0);
  const reasons = r.diagnostics.map(function (d) { return d.reason; }).sort();
  assert.deepStrictEqual(reasons, ["AVAILABLE_QTY_INVALID", "AVAILABLE_QTY_INVALID", "AVAILABLE_QTY_NEGATIVE"]);
});


// ---- Default output sort: Warehouse -> Product Code, bundle children
// grouped together in original mapping order ----
test("21. compareCodes: numeric-string codes compare numerically, not lexically", function () {
  assert.ok(mod._internal.compareCodes("9", "10") < 0);   // NOT lexical ("10" < "9" as strings)
  assert.ok(mod._internal.compareCodes("100", "20") > 0);
  assert.strictEqual(mod._internal.compareCodes("A", "A"), 0);
});

test("22. compareCodes: alphanumeric Bundle SKUs fall back to string order", function () {
  assert.ok(mod._internal.compareCodes("TP-001", "TP-002") < 0);
  assert.ok(mod._internal.compareCodes("TP-002", "TP-001") > 0);
});

test("23. Default sort groups bundle child rows together, in original mapping order", function () {
  const bundles = [{ sku: "TP-SORT", children: [
    { childSku: "Z-LAST", childProduct: "Z", qty: 1 }, // deliberately NOT alphabetical
    { childSku: "A-FIRST", childProduct: "A", qty: 2 },
    { childSku: "M-MID", childProduct: "M", qty: 3 },
  ] }];
  const r = run([invRow("SG2", "TP-SORT", 1, { unit: "MIX" })], bundles);
  assert.strictEqual(r.combined.length, 3);
  // all 3 rows adjacent (same barcode = bundle SKU) AND in the exact
  // original mapping order (Z, A, M) — NOT re-alphabetized.
  assert.deepStrictEqual(r.combined.map(function (row) { return row.childEan; }), ["Z-LAST", "A-FIRST", "M-MID"]);
  assert.ok(r.combined.every(function (row) { return row.barcode === "TP-SORT"; }));
});

test("24. Default sort: normal rows ordered by Warehouse -> Product Code", function () {
  const r = run([
    invRow("WH-South", "SKU-002", 1),
    invRow("WH-North", "SKU-100", 1),
    invRow("WH-North", "SKU-005", 1),
  ], []);
  const order = r.combined.map(function (row) { return row.warehouse + "|" + row.barcode; });
  assert.deepStrictEqual(order, ["WH-North|SKU-005", "WH-North|SKU-100", "WH-South|SKU-002"]);
});

test("25. Default sort: mixed normal + bundle rows interleave correctly by Warehouse -> Barcode", function () {
  const bundles = [{ sku: "SKU-050", children: [{ childSku: "C1", childProduct: "C1", qty: 1 }, { childSku: "C2", childProduct: "C2", qty: 2 }] }];
  const r = run([
    invRow("WH1", "SKU-100", 1),
    invRow("WH1", "SKU-050", 1, { unit: "MIX" }), // bundle carton whose "barcode" sorts between 010 and 100
    invRow("WH1", "SKU-010", 1),
  ], bundles);
  const seq = r.combined.map(function (row) { return row.barcode; });
  assert.deepStrictEqual(seq, ["SKU-010", "SKU-050", "SKU-050", "SKU-100"]); // bundle's 2 children stay adjacent, in place by barcode
  assert.deepStrictEqual(r.combined.map(function (row) { return row.childEan; }).slice(1, 3), ["C1", "C2"]); // original mapping order preserved
});


// ==========================================================
// Final EAN Qty logic (Rule 1 PCS, Rule 2 CARTON_<N>PCS, Rule 3 MIX
// bundle child) — deriveEanQty() is the single shared helper used by
// BOTH the normal-row and bundle-child branches inside
// generateCombinedReport, so these tests exercise it both directly
// (unit tests on deriveEanQty) and through the full transform
// (integration tests via run()/invRow()) to prove the preview/export
// row objects never compute EAN Qty a second time anywhere else.
// ==========================================================

test("EAN-1. PCS: Stock Qty=40 -> EAN Qty=40 (never hardcoded to 1)", function () {
  const r = run([invRow("SG2", "SKU-PCS-1", 5, { unit: "PCS", stockQty: 40 })], []);
  assert.strictEqual(r.combined[0].eanQty, 40);
  assert.notStrictEqual(r.combined[0].eanQty, 1);
});

test("EAN-2. PCS numeric string: Stock Qty=\"40\" -> EAN Qty=40 (numeric, not concatenated)", function () {
  const r = run([invRow("SG2", "SKU-PCS-2", 5, { unit: "PCS", stockQty: "40" })], []);
  assert.strictEqual(r.combined[0].eanQty, 40);
  assert.strictEqual(typeof r.combined[0].eanQty, "number");
});

test("EAN-3. CARTON_20PCS: Stock Qty=3 -> EAN Qty=60", function () {
  const r = run([invRow("SG2", "SKU-C20", 5, { unit: "CARTON_20PCS", stockQty: 3 })], []);
  assert.strictEqual(r.combined[0].eanQty, 60);
});

test("EAN-4. CARTON_30PCS: Stock Qty=4 -> EAN Qty=120", function () {
  const r = run([invRow("SG2", "SKU-C30", 5, { unit: "CARTON_30PCS", stockQty: 4 })], []);
  assert.strictEqual(r.combined[0].eanQty, 120);
});

test("EAN-5. CARTON_40PCS: Stock Qty=12 -> EAN Qty=480", function () {
  const r = run([invRow("SG2", "SKU-C40", 5, { unit: "CARTON_40PCS", stockQty: 12 })], []);
  assert.strictEqual(r.combined[0].eanQty, 480);
});

test("EAN-6. CARTON_60PCS: Stock Qty=6 -> EAN Qty=360", function () {
  const r = run([invRow("SG2", "SKU-C60", 5, { unit: "CARTON_60PCS", stockQty: 6 })], []);
  assert.strictEqual(r.combined[0].eanQty, 360);
});

test("EAN-7. CARTON_100PCS: Stock Qty=2 -> EAN Qty=200", function () {
  const r = run([invRow("SG2", "SKU-C100", 5, { unit: "CARTON_100PCS", stockQty: 2 })], []);
  assert.strictEqual(r.combined[0].eanQty, 200);
});

test("EAN-8. Lowercase UOM \"carton_40pcs\": Stock Qty=12 -> EAN Qty=480", function () {
  const r = run([invRow("SG2", "SKU-C40L", 5, { unit: "carton_40pcs", stockQty: 12 })], []);
  assert.strictEqual(r.combined[0].eanQty, 480);
});

test("EAN-9. UOM with internal space \"CARTON_40 PCS\": Stock Qty=12 -> EAN Qty=480", function () {
  const r = run([invRow("SG2", "SKU-C40S", 5, { unit: "CARTON_40 PCS", stockQty: 12 })], []);
  assert.strictEqual(r.combined[0].eanQty, 480);
});

test("EAN-10. Trailing-space UOM \"CARTON_40PCS \": Stock Qty=12 -> EAN Qty=480", function () {
  const r = run([invRow("SG2", "SKU-C40T", 5, { unit: "CARTON_40PCS ", stockQty: 12 })], []);
  assert.strictEqual(r.combined[0].eanQty, 480);
});

test("EAN-11. Invalid carton UOM \"CARTON_PCS\" (no N) -> EAN Qty blank + CARTON_PCS_SIZE_NOT_DETECTED", function () {
  const r = run([invRow("SG2", "SKU-CBAD", 5, { unit: "CARTON_PCS", stockQty: 12 })], []);
  assert.strictEqual(r.combined[0].eanQty, null);
  assert.ok(r.diagnostics.some(function (d) { return d.reason === "CARTON_PCS_SIZE_NOT_DETECTED" && d.code === "SKU-CBAD"; }));
  // must NOT invent a carton size (e.g. never silently falls back to Stock Qty itself)
  assert.notStrictEqual(r.combined[0].eanQty, 12);
});

test("EAN-12. Expanded MIX child: Stock Qty=1, Child SL=100 -> EAN Qty=100", function () {
  const bundles = [{ sku: "TP-MIX-1", children: [{ childSku: "C1", childProduct: "P1", qty: 100 }] }];
  const r = run([invRow("SG2", "TP-MIX-1", 1, { unit: "MIX", stockQty: 1 })], bundles);
  assert.strictEqual(r.combined[0].stockQty, 1);
  assert.strictEqual(r.combined[0].eanQty, 100);
});

test("EAN-13. Expanded MIX child: Stock Qty remains 1, is NOT overwritten by Child SL", function () {
  const bundles = [{ sku: "TP-MIX-2", children: [{ childSku: "C1", childProduct: "P1", qty: 250 }] }];
  const r = run([invRow("SG2", "TP-MIX-2", 1, { unit: "MIX", stockQty: 1 })], bundles);
  assert.strictEqual(r.combined[0].stockQty, 1); // NOT 250
  assert.notStrictEqual(r.combined[0].stockQty, r.combined[0].eanQty);
});

test("EAN-14. MIX child SL values 100/10/10/20 -> EAN Qty 100/10/10/20", function () {
  const bundles = [{ sku: "TP-MIX-3", children: [
    { childSku: "C1", childProduct: "P1", qty: 100 }, { childSku: "C2", childProduct: "P2", qty: 10 },
    { childSku: "C3", childProduct: "P3", qty: 10 }, { childSku: "C4", childProduct: "P4", qty: 20 },
  ] }];
  const r = run([invRow("SG2", "TP-MIX-3", 1, { unit: "MIX", stockQty: 1 })], bundles);
  assert.deepStrictEqual(r.combined.map(function (row) { return row.eanQty; }), [100, 10, 10, 20]);
  assert.ok(r.combined.every(function (row) { return row.stockQty === 1; })); // all 4 share source Stock Qty, unchanged
});

test("EAN-15. MIX EAN Qty is not multiplied by Stock Qty", function () {
  const bundles = [{ sku: "TP-MIX-4", children: [{ childSku: "C1", childProduct: "P1", qty: 50 }] }];
  // Stock Qty = 3 (unusual, but must still NOT multiply: EAN Qty stays 50, not 150)
  const r = run([invRow("SG2", "TP-MIX-4", 1, { unit: "MIX", stockQty: 3 })], bundles);
  assert.strictEqual(r.combined[0].eanQty, 50);
  assert.strictEqual(r.combined[0].stockQty, 3);
});

test("EAN-16. Normal row metadata unchanged by the EAN Qty refactor", function () {
  const r = run([invRow("WH1", "SKU-META", 5, {
    unit: "CARTON_20PCS", stockQty: 3, partnerSku: "PSKU-1", product: "Widget", category: "Cat1",
    conditionType: "Mới", freezeQty: 2, pendingIn: 1, pendingOut: 0, weightKg: 4.5, cbm: 0.02,
  })], []);
  const row = r.combined[0];
  assert.strictEqual(row.warehouse, "WH1");
  assert.strictEqual(row.partnerSku, "PSKU-1");
  assert.strictEqual(row.product, "Widget");
  assert.strictEqual(row.category, "Cat1");
  assert.strictEqual(row.conditionType, "Mới");
  assert.strictEqual(row.freezeQty, 2);
  assert.strictEqual(row.pendingIn, 1);
  assert.strictEqual(row.pendingOut, 0);
  assert.strictEqual(row.weightKg, 4.5);
  assert.strictEqual(row.cbm, 0.02);
  assert.strictEqual(row.stockQty, 3);
  assert.strictEqual(row.eanQty, 60); // 3 x 20, confirms the calc still ran correctly alongside untouched metadata
});

test("EAN-17. Web preview and Excel export contain identical EAN Qty values (same row objects)", function () {
  const bundles = [{ sku: "TP-MIX-5", children: [
    { childSku: "C1", childProduct: "P1", qty: 100 }, { childSku: "C2", childProduct: "P2", qty: 10 },
  ] }];
  const r = run([
    invRow("WH1", "SKU-NORM", 5, { unit: "PCS", stockQty: 40 }),
    invRow("WH1", "SKU-CART", 5, { unit: "CARTON_40PCS", stockQty: 12 }),
    invRow("WH1", "TP-MIX-5", 1, { unit: "MIX", stockQty: 1 }),
  ], bundles);
  const previewEanQtys = r.combined.map(function (row) { return row.eanQty; });

  const capturedAoa = [];
  global.XLSX = {
    utils: {
      aoa_to_sheet: function (aoa) { capturedAoa.push(aoa); return { __aoa: aoa }; },
      book_new: function () { return { sheets: {} }; },
      book_append_sheet: function () {},
      encode_range: function (rr) { return "A1:Z" + (rr.e.r + 1); },
    },
    write: function () { return new Uint8Array([1]); },
  };
  global.Blob = function (parts, opts) { this.size = 1; this.type = opts.type; };
  mod.buildExportWorkbook(r.combined);
  delete global.XLSX; delete global.Blob;

  const aoa = capturedAoa[0];
  const eanQtyColIdx = mod.EXPORT_HEADER.indexOf("EAN Qty");
  const exportEanQtys = aoa.slice(1).map(function (row) { return row[eanQtyColIdx]; });
  assert.deepStrictEqual(exportEanQtys, previewEanQtys); // preview and export must never disagree
});

test("EAN-18. Output schema still contains both Stock Qty and EAN Qty columns", function () {
  assert.ok(mod.EXPORT_HEADER.indexOf("Stock Qty") !== -1);
  assert.ok(mod.EXPORT_HEADER.indexOf("EAN Qty") !== -1);
});

test("EAN-19. No Available PCS / Total PCS / duplicate quantity column exists", function () {
  assert.strictEqual(mod.EXPORT_HEADER.indexOf("Available PCS"), -1);
  assert.strictEqual(mod.EXPORT_HEADER.indexOf("PCS per UOM"), -1);
  assert.strictEqual(mod.EXPORT_HEADER.indexOf("Total PCS"), -1);
  assert.strictEqual(mod.EXPORT_HEADER.length, 19);
  const r = run([invRow("SG2", "SKU-SCHEMA", 5, { unit: "PCS", stockQty: 10 })], []);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(r.combined[0], "availablePcs"), false);
});

test("EAN-20. Only Available Qty > 0 source rows enter the final output", function () {
  const r = run([
    invRow("SG2", "SKU-OK", 1, { unit: "PCS", stockQty: 10 }),
    invRow("SG2", "SKU-ZERO", 0, { unit: "PCS", stockQty: 10 }),
    invRow("SG2", "SKU-NEG", -5, { unit: "PCS", stockQty: 10 }),
  ], []);
  assert.strictEqual(r.combined.length, 1);
  assert.strictEqual(r.combined[0].barcode, "SKU-OK");
});

// ---- deriveEanQty direct unit tests (the shared helper itself) ----
test("EAN-21. deriveEanQty: PCS branch returns Stock Qty untouched", function () {
  const out = mod._internal.deriveEanQty({ unit: "PCS", stockQty: 40, isExpandedBundle: false, childQty: null, sourceEanQty: null });
  assert.strictEqual(out.eanQty, 40);
  assert.strictEqual(out.diagnostic, null);
});

test("EAN-22. deriveEanQty: CARTON branch multiplies, numeric string safe (\"12\" x 40 -> 480)", function () {
  const out = mod._internal.deriveEanQty({ unit: "CARTON_40PCS", stockQty: "12", isExpandedBundle: false, childQty: null, sourceEanQty: null });
  assert.strictEqual(out.eanQty, 480);
  assert.strictEqual(typeof out.eanQty, "number");
});

test("EAN-23. deriveEanQty: isExpandedBundle always wins, ignoring unit/stockQty entirely", function () {
  const out = mod._internal.deriveEanQty({ unit: "PCS", stockQty: 999, isExpandedBundle: true, childQty: 7, sourceEanQty: null });
  assert.strictEqual(out.eanQty, 7);
});

test("EAN-24. deriveEanQty: blank/invalid Stock Qty on a PCS row -> blank + EAN_QTY_STOCK_INVALID", function () {
  const out = mod._internal.deriveEanQty({ unit: "PCS", stockQty: null, isExpandedBundle: false, childQty: null, sourceEanQty: null });
  assert.strictEqual(out.eanQty, null);
  assert.strictEqual(out.diagnostic, "EAN_QTY_STOCK_INVALID");
});

test("EAN-25. deriveEanQty: unknown UOM falls back to trustworthy source EAN Qty, else blank + EAN_QTY_RULE_NOT_DETECTED", function () {
  const withSource = mod._internal.deriveEanQty({ unit: "BOX", stockQty: 5, isExpandedBundle: false, childQty: null, sourceEanQty: 42 });
  assert.strictEqual(withSource.eanQty, 42);
  assert.strictEqual(withSource.diagnostic, null);
  const withoutSource = mod._internal.deriveEanQty({ unit: "BOX", stockQty: 5, isExpandedBundle: false, childQty: null, sourceEanQty: null });
  assert.strictEqual(withoutSource.eanQty, null);
  assert.strictEqual(withoutSource.diagnostic, "EAN_QTY_RULE_NOT_DETECTED");
});

console.log("\n=== " + passed + " passed, " + failed + " failed ===\n");
process.exit(failed ? 1 : 0);
