# -*- coding: utf-8 -*-
"""Generate the sanitized SYNTHETIC packing-list PDF fixtures used by
test_pl_ocr_core.py. Run with `python3 generate.py` from this directory to
regenerate SYN-1001-WarehouseAlpha-CN.pdf / SYN-1002-WarehouseBeta-CN.pdf /
SYN-DIM.xlsx in place. Every code, barcode, and address below is fabricated
-- see the 2026-08-03 security incident note in test_pl_ocr_core.py for why
this replaced a real-file fixture set that should never have been
committed to a public repo. Requires: pip install reportlab openpyxl."""
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from pathlib import Path

OUT = Path(__file__).resolve().parent
OUT.mkdir(parents=True, exist_ok=True)
styles = getSampleStyleSheet()
GRID = TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black)])
HDR_ROW = ["#", "Barcode", "Ma san pham", "Ten hang hoa", "DVT", "Tinh trang So luong"]
HDR_ROW_8COL = ["#", "Barcode", "Ma san pham", "Ten hang hoa", "DVT", "Tinh trang", "So luong", ""]

def noise_block():
    return [
        Paragraph("8/2/26, 10:15 AM about:blank", styles["Normal"]),
        Paragraph("FAKE-WAREHOUSE (synthetic test fixture -- not a real facility)", styles["Normal"]),
        Paragraph("Dia chi: 123 Test Street, Test District, Test City, Testland (fabricated)", styles["Normal"]),
        Paragraph("Nguoi lien he: Synthetic Tester - 0000000000 (fabricated)", styles["Normal"]),
        Spacer(1, 6),
    ]

# ── File A: package spans 2 pages, page1 has merged condition+qty (6-col)
#    including a SKU split across a newline ("-BOX") and a SKU containing a
#    U+FFFE artifact in place of '-'; page2 continuation repeats the header
#    and uses the 8-col layout (condition + quantity in separate cells).
def build_file_a():
    doc = SimpleDocTemplate(str(OUT / "SYN-1001-WarehouseAlpha-CN.pdf"), pagesize=A4)
    elems = []
    elems += noise_block()
    data = [
        ["Ma kien hang: PGKECSYN10010001 1/1", "", "", "", "", ""],
        HDR_ROW,
        ["1", "1112223330001", "TP-SYN-A-001-\nBOX", "Widget A", "PCS", "Moi 10"],
        ["2", "1112223330002", "TP-SYN-A-002", "Widget B", "PCS", "Moi 5"],
        ["3", "1112223330003", "TP-SYN-A-003", "Widget C", "PCS", "Moi 7"],
    ]
    t = Table(data)
    t.setStyle(GRID)
    elems.append(t)
    elems.append(Spacer(1, 20))
    # page break so package continues on page 2
    from reportlab.platypus import PageBreak
    elems.append(PageBreak())
    elems.append(Paragraph("# Barcode Ma san pham Ten hang hoa DVT Tinh trang So luong", styles["Normal"]))
    data2 = [
        ["Ma kien hang: PGKECSYN10010001", "1/1", "", "", "", "", "", ""],
        HDR_ROW_8COL,
        ["4", "1112223330004", "TP-SYN-A-004", "Widget D", "PCS", "Moi", "8", ""],
        ["5", "1112223330005", "TP-SYN-A-005", "Widget E", "PCS", "Moi", "6", ""],
        ["Tong cong", "", "", "", "", "", "36", ""],
    ]
    t2 = Table(data2)
    t2.setStyle(GRID)
    elems.append(t2)
    doc.build(elems)

# ── File B: ONE page, TWO packages (spec bug #1: "mot trang co nhieu
#    package"), both using the merged condition+qty (6-col) shape.
def build_file_b():
    doc = SimpleDocTemplate(str(OUT / "SYN-1002-WarehouseBeta-CN.pdf"), pagesize=A4)
    elems = []
    elems += noise_block()
    data1 = [
        ["Ma kien hang: PGKECSYN10020001 1/2", "", "", "", "", ""],
        HDR_ROW,
        ["1", "1112223330011", "TP-SYN-B-011", "Gadget A", "PCS", "Moi 9"],
        ["2", "1112223330012", "TP-SYN-B-012", "Gadget B", "PCS", "Moi 4"],
        ["Tong cong", "", "", "", "", "13"],
    ]
    t1 = Table(data1)
    t1.setStyle(GRID)
    elems.append(t1)
    elems.append(Spacer(1, 16))
    data2 = [
        ["Ma kien hang: PGKECSYN10020002 2/2", "", "", "", "", ""],
        HDR_ROW,
        ["1", "1112223330021", "TP-SYN-B-021", "Gadget C", "PCS", "Moi 11"],
        ["2", "1112223330022", "TP-SYN-B-022", "Gadget D", "PCS", "Moi 3"],
        ["Tong cong", "", "", "", "", "14"],
    ]
    t2 = Table(data2)
    t2.setStyle(GRID)
    elems.append(t2)
    doc.build(elems)

def build_dim():
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Lo Hang (QRcode o giua)", "Ma Kien (Barcode o tren hoac duoi)",
               "Dai (cm)", "Rong (cm)", "Cao (cm)", "Nang (kg) -> Nhap theo can", "CBM"])
    rows = [
        ("SYN-1001-WarehouseAlpha-CN", "PGKECSYN10010001", 50, 40, 30, 12.0, 0.06),
        ("SYN-1002-WarehouseBeta-CN",  "PGKECSYN10020001", 45, 35, 25, 8.0, 0.039),
        ("SYN-1002-WarehouseBeta-CN",  "PGKECSYN10020002", 45, 35, 25, 9.0, 0.039),
    ]
    for r in rows:
        ws.append(list(r))
    wb.save(str(OUT / "SYN-DIM.xlsx"))


build_file_a()
build_file_b()
build_dim()
print("built:", sorted(p.name for p in OUT.glob("*.pdf")) + ["SYN-DIM.xlsx"])
