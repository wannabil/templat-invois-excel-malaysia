"""Build a free BM Excel template: Sebut Harga -> Pesanan Belian (rekod PO pelanggan) -> Nota Penghantaran -> Invois -> Resit, with a document register.
Demo data only (fictional company and customer)."""
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation

out = sys.argv[1] if len(sys.argv) > 1 else "templat-invois-sebut-harga.xlsx"
wb = Workbook()
B = Font(bold=True)
H = Font(bold=True, size=16)
FILL = PatternFill("solid", fgColor="1F2937")
WH = Font(bold=True, color="FFFFFF")
thin = Side(style="thin", color="D1D5DB")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
RM = '"RM"#,##0.00'

# ---- Tetapan (settings) ----
s = wb.active
s.title = "Tetapan"
rows = [
    ("Tetapan syarikat (isi sekali, semua dokumen guna maklumat ini)", None),
    (None, None),
    ("Nama syarikat", "Kedai Contoh Enterprise"),
    ("No. SSM", "202501000000 (contoh)"),
    ("Alamat", "No. 1, Jalan Contoh, 47800 Petaling Jaya, Selangor"),
    ("Telefon", "03-0000 0000"),
    ("E-mel", "akaun@contoh.my"),
    ("Bank & no. akaun", "Maybank 0000 0000 0000 (contoh)"),
    ("Kadar cukai (0 jika tidak berdaftar SST)", 0),
    ("Terma bayaran (hari)", 14),
]
for r, (k, v) in enumerate(rows, 1):
    s.cell(r, 1, k).font = H if r == 1 else B
    if v is not None:
        s.cell(r, 2, v)
s["B9"].number_format = "0%"
s.column_dimensions["A"].width = 42
s.column_dimensions["B"].width = 55
s["A12"] = "Nota: ini templat percuma, bukan nasihat cukai atau perakaunan. Untuk e-Invois LHDN, rujuk hasil.gov.my."
s["A12"].font = Font(italic=True, color="6B7280")

# ---- Daftar (register) ----
d = wb.create_sheet("Daftar")
hdr = ["No. dokumen", "Jenis", "Tarikh", "Pelanggan", "Rujukan (dokumen asal)", "Jumlah (RM)", "Status", "Tarikh bayar"]
for c, h in enumerate(hdr, 1):
    x = d.cell(1, c, h); x.font = WH; x.fill = FILL
demo = [
    ("SH-2026-001", "Sebut Harga", "2026-11-02", "Syarikat Pelanggan Sdn Bhd", "", "='Sebut Harga'!F22", "Diterima", ""),
    ("PO-PLG-7781", "PO Pelanggan", "2026-11-04", "Syarikat Pelanggan Sdn Bhd", "SH-2026-001", "='Sebut Harga'!F22", "Diterima", ""),
    ("DO-2026-001", "Nota Penghantaran", "2026-11-06", "Syarikat Pelanggan Sdn Bhd", "PO-PLG-7781", "", "Dihantar", ""),
    ("INV-2026-001", "Invois", "2026-11-06", "Syarikat Pelanggan Sdn Bhd", "DO-2026-001", "=Invois!F22", "Belum bayar", ""),
    ("RS-2026-001", "Resit", "2026-11-15", "Syarikat Pelanggan Sdn Bhd", "INV-2026-001", "=Resit!C12", "Dibayar", "2026-11-15"),
]
for r, row in enumerate(demo, 2):
    for c, v in enumerate(row, 1):
        d.cell(r, c, v)
    d.cell(r, 6).number_format = RM
dv = DataValidation(type="list", formula1='"Sebut Harga,PO Pelanggan,Nota Penghantaran,Invois,Resit"', allow_blank=True)
d.add_data_validation(dv); dv.add("B2:B500")
dv2 = DataValidation(type="list", formula1='"Draf,Dihantar,Diterima,Belum bayar,Dibayar,Batal"', allow_blank=True)
d.add_data_validation(dv2); dv2.add("G2:G500")
for col, w in zip("ABCDEFGH", (16, 20, 12, 30, 24, 14, 14, 14)):
    d.column_dimensions[col].width = w
d["J1"] = "Invois belum bayar (RM)"; d["J1"].font = B
d["J2"] = '=SUMIFS(F:F,B:B,"Invois",G:G,"Belum bayar")'; d["J2"].number_format = RM
d["J4"] = "Nombor seterusnya: guna format JENIS-TAHUN-NNN (cth INV-2026-002). Satu siri bagi setiap jenis, jangan guna semula nombor."
d.freeze_panes = "A2"


def doc_sheet(name, title, number, ref_label, ref, extra_note):
    w = wb.create_sheet(name)
    w["A1"] = "=Tetapan!B3"; w["A1"].font = H
    w["A2"] = '=Tetapan!B5'
    w["A3"] = '="SSM "&Tetapan!B4'
    w["A4"] = '=Tetapan!B6&"  |  "&Tetapan!B7'
    w["E1"] = title; w["E1"].font = Font(bold=True, size=18)
    meta = [("No.", number), ("Tarikh", "2026-11-06"), (ref_label, ref)]
    for i, (k, v) in enumerate(meta, 2):
        w.cell(i, 5, k).font = B; w.cell(i, 6, v)
    w["A6"] = "Kepada:"; w["A6"].font = B
    w["A7"] = "Syarikat Pelanggan Sdn Bhd"
    w["A8"] = "Lot 2, Jalan Pelanggan, 40000 Shah Alam, Selangor"
    w["A9"] = "u.p.: Puan Contoh, 012-000 0000"
    cols = ["#", "Perkara", "Kuantiti", "Unit", "Harga seunit", "Jumlah"]
    for c, h in enumerate(cols, 1):
        x = w.cell(11, c, h); x.font = WH; x.fill = FILL
    items = [("Kotak kek 8 inci (100 unit)", 2, "pek", 45), ("Pelekat logo custom", 500, "keping", 0.3), ("Penghantaran Lembah Klang", 1, "trip", 25)]
    for i in range(8):
        r = 12 + i
        w.cell(r, 1, i + 1)
        if i < len(items):
            for c, v in enumerate(items[i], 2):
                w.cell(r, c, v)
        w.cell(r, 6, f'=IF(C{r}="","",C{r}*E{r})')
        w.cell(r, 5).number_format = RM; w.cell(r, 6).number_format = RM
        for c in range(1, 7):
            w.cell(r, c).border = BOX
    w["E20"] = "Jumlah kecil"; w["F20"] = "=SUM(F12:F19)"
    w["E21"] = '="Cukai "&TEXT(Tetapan!B9,"0%")'; w["F21"] = "=F20*Tetapan!B9"
    w["E22"] = "JUMLAH"; w["F22"] = "=F20+F21"
    for r in (20, 21, 22):
        w.cell(r, 5).font = B; w.cell(r, 6).number_format = RM
    w["F22"].font = Font(bold=True, size=13)
    w["A24"] = extra_note
    w["A25"] = '="Bayaran ke: "&Tetapan!B8'
    w.column_dimensions["A"].width = 6; w.column_dimensions["B"].width = 40
    for col, wd in zip("CDEF", (10, 10, 16, 27)):
        w.column_dimensions[col].width = wd
    w.sheet_properties.pageSetUpPr.fitToPage = True
    w.page_setup.fitToWidth = 1
    w.page_setup.fitToHeight = 1
    w.print_area = "A1:F25"
    return w


doc_sheet("Sebut Harga", "SEBUT HARGA", "SH-2026-001", "Sah sehingga", "2026-12-02", "Harga sah 30 hari. Sila keluarkan PO dengan merujuk nombor sebut harga ini.")
dsheet = doc_sheet("Nota Penghantaran", "NOTA PENGHANTARAN", "DO-2026-001", "Rujukan PO", "PO-PLG-7781", "Diterima dalam keadaan baik oleh: ____________________  Tarikh: __________  Cop syarikat:")
for r in range(12, 20):  # DO shows quantities, not prices
    dsheet.cell(r, 5).value = None; dsheet.cell(r, 6).value = None
for r in (20, 21, 22):
    dsheet.cell(r, 5).value = None; dsheet.cell(r, 6).value = None
dsheet["E11"] = "Catatan"; dsheet["F11"] = "Tandatangan"
doc_sheet("Invois", "INVOIS", "INV-2026-001", "Rujukan DO / PO", "DO-2026-001 / PO-PLG-7781", '="Terma: bayar dalam "&Tetapan!B10&" hari dari tarikh invois."')
wb["Invois"]["E5"] = "Tarikh akhir"; wb["Invois"]["E5"].font = B
wb["Invois"]["F5"] = "=DATEVALUE(F3)+Tetapan!B10"; wb["Invois"]["F5"].number_format = "yyyy-mm-dd"

# ---- Resit ----
rs = wb.create_sheet("Resit")
rs["A1"] = "=Tetapan!B3"; rs["A1"].font = H
rs["E1"] = "RESIT"; rs["E1"].font = Font(bold=True, size=18)
for i, (k, v) in enumerate([("No. resit", "RS-2026-001"), ("Tarikh", "2026-11-15")], 2):
    rs.cell(i, 5, k).font = B; rs.cell(i, 6, v)
lines = [("Diterima daripada", "Syarikat Pelanggan Sdn Bhd"), ("Untuk invois", "INV-2026-001"), ("Cara bayaran", "FPX / pindahan bank"), ("Rujukan transaksi", "(no. rujukan bank)")]
for i, (k, v) in enumerate(lines, 7):
    rs.cell(i, 1, k).font = B; rs.cell(i, 3, v)
rs["A12"] = "Jumlah diterima"; rs["A12"].font = B
rs["C12"] = "=Invois!F22"; rs["C12"].number_format = RM; rs["C12"].font = Font(bold=True, size=13)
rs.column_dimensions["A"].width = 20; rs.column_dimensions["C"].width = 32

# ---- Cara guna ----
g = wb.create_sheet("Cara Guna", 0)
steps = [
    "Templat Invois & Sebut Harga (Excel) - percuma",
    "",
    "1. Isi tab Tetapan sekali sahaja (nama syarikat, SSM, bank, cukai, terma bayaran).",
    "2. Pelanggan tanya harga -> isi tab Sebut Harga, simpan sebagai PDF, hantar.",
    "3. Pelanggan setuju -> mereka hantar PO. Rekod nombor PO mereka dalam tab Daftar.",
    "4. Barang dihantar -> isi Nota Penghantaran (tiada harga), minta pelanggan tandatangan.",
    "5. Keluarkan Invois, rujuk nombor DO dan PO.",
    "6. Duit masuk -> keluarkan Resit, tukar status invois kepada Dibayar dalam Daftar.",
    "",
    "Tab Daftar tunjuk semua dokumen dan jumlah invois yang belum dibayar.",
    "Tanda Excel dah tak cukup: taip semula data yang sama di banyak tab, nombor invois bertindih,",
    "susah nampak siapa belum bayar, ramai staf edit fail yang sama.",
    "",
    "Panduan penuh & bila patut tukar ke sistem: https://miraclemikhael.com/services/custom-business-systems",
    "Templat ini bukan nasihat cukai/perakaunan. Untuk e-Invois LHDN, rujuk https://www.hasil.gov.my",
]
for r, t in enumerate(steps, 1):
    g.cell(r, 1, t)
g["A1"].font = H
g.column_dimensions["A"].width = 110
for ws in wb.worksheets:
    if not ws.sheet_properties.pageSetUpPr.fitToPage:
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
    ws.page_setup.orientation = "landscape" if ws.title == "Daftar" else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
wb.active = 0
wb.save(out)
print("saved", out)
