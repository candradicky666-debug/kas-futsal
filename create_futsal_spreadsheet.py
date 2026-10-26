import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

wb = openpyxl.Workbook()

# ==============================================================================
# SHEET 1: REKAP KAS FUTSAL (PRO TABLE DESIGN)
# ==============================================================================
ws = wb.active
ws.title = "Rekap Kas Futsal"
ws.views.sheetView[0].showGridLines = True
ws.sheet_properties.tabColor = "107C41"

# Freeze header rows so table scrolls like a pro sheet
ws.freeze_panes = "A10"

# Palette
C_HEADER_FILL = "107C41"    # Professional Excel Dark Emerald Green
C_HEADER_TEXT = "FFFFFF"
C_ZEBRA = "F8FAF9"          # Clean subtle row striping
C_WHITE = "FFFFFF"
C_TOTAL_BG = "E2E8F0"       # Professional Slate Gray
C_TEXT_DARK = "0F172A"      # Crisp Slate 900
C_TEXT_MUTED = "475569"
C_BORDER_GRID = "CBD5E1"    # Subtle Excel Grid
C_BORDER_DARK = "475569"

# Professional Fonts
font_main_title = Font(name="Segoe UI", size=16, bold=True, color="107C41")
font_subtitle = Font(name="Segoe UI", size=10, italic=True, color=C_TEXT_MUTED)
font_kpi_lbl = Font(name="Segoe UI", size=9, bold=True, color=C_TEXT_MUTED)
font_tbl_header = Font(name="Segoe UI", size=11, bold=True, color=C_HEADER_TEXT)
font_row_no = Font(name="Segoe UI", size=10, bold=True, color="64748B")
font_data = Font(name="Segoe UI", size=10, color=C_TEXT_DARK)
font_data_bold = Font(name="Segoe UI", size=10, bold=True, color=C_TEXT_DARK)
font_inc = Font(name="Segoe UI", size=10, bold=True, color="15803D")  # Green bold
font_exp = Font(name="Segoe UI", size=10, bold=True, color="B91C1C")  # Red bold
font_bal = Font(name="Segoe UI", size=10, bold=True, color="1D4ED8")  # Blue bold

font_tot_label = Font(name="Segoe UI", size=11, bold=True, color="0F172A")
font_tot_val = Font(name="Segoe UI", size=11, bold=True, color="0F172A")

# Alignments
align_center = Alignment(horizontal="center", vertical="center")
align_left = Alignment(horizontal="left", vertical="center")
align_right = Alignment(horizontal="right", vertical="center")
align_header = Alignment(horizontal="center", vertical="center", wrap_text=True)

# Borders
side_thin = Side(style="thin", color=C_BORDER_GRID)
side_med = Side(style="medium", color=C_BORDER_DARK)
side_double = Side(style="double", color="0F172A")

border_cell = Border(left=side_thin, right=side_thin, top=side_thin, bottom=side_thin)
border_header = Border(left=side_thin, right=side_thin, top=side_med, bottom=side_med)
border_total = Border(left=side_thin, right=side_thin, top=side_med, bottom=side_double)

# Indonesian Currency format with Rp
CURRENCY_FORMAT = '"Rp "#,##0'

# Title Header
ws["B2"] = "BUKU KAS & REKAPITULASI KEUANGAN FUTSAL"
ws["B2"].font = font_main_title
ws["B3"] = "Laporan Resmi Kas Tim: Pemasukan, Pengeluaran, Sisa Kas Berjalan & Total Akumulasi"
ws["B3"].font = font_subtitle
ws.row_dimensions[2].height = 24
ws.row_dimensions[3].height = 18

# ------------------------------------------------------------------------------
# KPI Summary Cards (B5:G7)
# ------------------------------------------------------------------------------
def build_card(start_c, end_c, title, formula, bg_color, num_color):
    ws.merge_cells(start_row=5, start_column=start_c, end_row=5, end_column=end_c)
    lbl = ws.cell(row=5, column=start_c, value=title)
    lbl.font = font_kpi_lbl
    lbl.alignment = align_center

    ws.merge_cells(start_row=6, start_column=start_c, end_row=7, end_column=end_c)
    val = ws.cell(row=6, column=start_c, value=formula)
    val.font = Font(name="Segoe UI", size=15, bold=True, color=num_color)
    val.alignment = align_center
    val.number_format = CURRENCY_FORMAT

    for r in range(5, 8):
        for c in range(start_c, end_c + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
            cell.border = border_cell

build_card(2, 3, "TOTAL PEMASUKAN", "=SUM(E10:E34)", "DCFCE7", "15803D")  # Light green card
build_card(4, 5, "TOTAL PENGELUARAN", "=SUM(F10:F34)", "FEE2E2", "B91C1C") # Light red card
build_card(6, 7, "SISA KAS (SALDO AKHIR)", "=B6-D6", "DBEAFE", "1D4ED8")   # Light blue card

ws.row_dimensions[5].height = 20
ws.row_dimensions[6].height = 18
ws.row_dimensions[7].height = 18
ws.row_dimensions[8].height = 14

# ------------------------------------------------------------------------------
# Table Header Row (Row 9) - BOLD HEADERS
# ------------------------------------------------------------------------------
columns = [
    ("No", 7),
    ("Tanggal", 15),
    ("Keterangan Transaksi", 38),
    ("Kategori", 20),
    ("Pemasukan (Rp)", 20),
    ("Pengeluaran (Rp)", 20),
    ("Sisa Kas (Rp)", 20),
]

ws.row_dimensions[9].height = 30
for idx, (col_name, col_width) in enumerate(columns, start=1):
    c = ws.cell(row=9, column=idx, value=col_name)
    c.font = font_tbl_header
    c.fill = PatternFill(start_color=C_HEADER_FILL, end_color=C_HEADER_FILL, fill_type="solid")
    c.alignment = align_header
    c.border = border_header
    ws.column_dimensions[get_column_letter(idx)].width = col_width

# Category Data Validation dropdown
cat_dv = DataValidation(
    type="list",
    formula1='"Iuran Kas,Sewa Lapangan,Konsumsi,Perlengkapan,Turnamen,Uang Sparing,Kas Masuk Lain,Lain-lain"',
    allow_blank=True
)
cat_dv.error = "Pilih kategori dari daftar yang tersedia"
cat_dv.errorTitle = "Kategori Tidak Valid"
ws.add_data_validation(cat_dv)

sample_transactions = [
    (1, "2026-10-01", "Saldo Awal Kas Tim", "Kas Masuk Lain", 500000, 0),
    (2, "2026-10-04", "Iuran Kas Mingguan Pemain (Pekan 1)", "Iuran Kas", 300000, 0),
    (3, "2026-10-04", "Sewa Lapangan Futsal (2 Jam)", "Sewa Lapangan", 0, 350000),
    (4, "2026-10-04", "Air Mineral 1 Karton + Es Batu", "Konsumsi", 0, 45000),
    (5, "2026-10-11", "Iuran Kas Mingguan Pemain (Pekan 2)", "Iuran Kas", 280000, 0),
    (6, "2026-10-11", "Sewa Lapangan Futsal (2 Jam)", "Sewa Lapangan", 0, 350000),
    (7, "2026-10-15", "Beli Bola Futsal Specs Original (1 pcs)", "Perlengkapan", 0, 240000),
    (8, "2026-10-18", "Iuran Kas Mingguan Pemain (Pekan 3)", "Iuran Kas", 300000, 0),
    (9, "2026-10-18", "Sewa Lapangan Futsal (2 Jam)", "Sewa Lapangan", 0, 350000),
    (10, "2026-10-18", "Uang Pendaftaran Turnamen Antar Tim", "Turnamen", 0, 200000),
    (11, "2026-10-25", "Hadiah Juara 3 Turnamen Futsal", "Kas Masuk Lain", 500000, 0),
    (12, "2026-10-25", "Iuran Kas Mingguan Pemain (Pekan 4)", "Iuran Kas", 320000, 0),
]

start_row = 10
total_entry_rows = 25  # rows 10 to 34
end_data_row = start_row + total_entry_rows - 1 # 34

for i in range(total_entry_rows):
    curr_row = start_row + i
    ws.row_dimensions[curr_row].height = 22
    is_sample = i < len(sample_transactions)

    cat_dv.add(ws.cell(row=curr_row, column=4))
    row_bg = C_ZEBRA if (i % 2 == 1) else C_WHITE

    if is_sample:
        no, tgl, desc, kat, pem, peng = sample_transactions[i]
        ws.cell(row=curr_row, column=1, value=no)
        ws.cell(row=curr_row, column=2, value=tgl)
        ws.cell(row=curr_row, column=3, value=desc)
        ws.cell(row=curr_row, column=4, value=kat)
        ws.cell(row=curr_row, column=5, value=pem)
        ws.cell(row=curr_row, column=6, value=peng)
    else:
        ws.cell(row=curr_row, column=1, value=f'=IF(C{curr_row}<>""; ROW()-9; "")'.replace(";", ","))
        ws.cell(row=curr_row, column=2, value="")
        ws.cell(row=curr_row, column=3, value="")
        ws.cell(row=curr_row, column=4, value="")
        ws.cell(row=curr_row, column=5, value=0)
        ws.cell(row=curr_row, column=6, value=0)

    # Formula Running Balance
    ws.cell(row=curr_row, column=7, value=f'=IF(AND(C{curr_row}="", E{curr_row}=0, F{curr_row}=0), "", SUM(E$10:E{curr_row}) - SUM(F$10:F{curr_row}))')

    # Formatting & Styles
    c1 = ws.cell(row=curr_row, column=1)
    c1.alignment = align_center
    c1.font = font_row_no

    c2 = ws.cell(row=curr_row, column=2)
    c2.alignment = align_center
    c2.font = font_data

    c3 = ws.cell(row=curr_row, column=3)
    c3.alignment = align_left
    c3.font = font_data

    c4 = ws.cell(row=curr_row, column=4)
    c4.alignment = align_center
    c4.font = font_data

    # Kolom Pemasukan (Rp) - BOLD & GREEN
    c5 = ws.cell(row=curr_row, column=5)
    c5.alignment = align_right
    c5.font = font_inc
    c5.number_format = CURRENCY_FORMAT

    # Kolom Pengeluaran (Rp) - BOLD & RED
    c6 = ws.cell(row=curr_row, column=6)
    c6.alignment = align_right
    c6.font = font_exp
    c6.number_format = CURRENCY_FORMAT

    # Kolom Sisa Kas (Rp) - BOLD & BLUE
    c7 = ws.cell(row=curr_row, column=7)
    c7.alignment = align_right
    c7.font = font_bal
    c7.number_format = CURRENCY_FORMAT

    for col in range(1, 8):
        cell = ws.cell(row=curr_row, column=col)
        cell.fill = PatternFill(start_color=row_bg, end_color=row_bg, fill_type="solid")
        cell.border = border_cell

# ------------------------------------------------------------------------------
# Total Row (Row 35) - BOLD & DOUBLE UNDERLINE
# ------------------------------------------------------------------------------
tot_row = end_data_row + 1 # 35
ws.row_dimensions[tot_row].height = 28

ws.merge_cells(start_row=tot_row, start_column=1, end_row=tot_row, end_column=4)
lbl_tot = ws.cell(row=tot_row, column=1, value="TOTAL KESELURUHAN")
lbl_tot.font = font_tot_label
lbl_tot.alignment = Alignment(horizontal="right", vertical="center")

c_tot_pem = ws.cell(row=tot_row, column=5, value=f"=SUM(E{start_row}:E{end_data_row})")
c_tot_peng = ws.cell(row=tot_row, column=6, value=f"=SUM(F{start_row}:F{end_data_row})")
c_tot_sisa = ws.cell(row=tot_row, column=7, value=f"=E{tot_row}-F{tot_row}")

for c in range(1, 8):
    cell = ws.cell(row=tot_row, column=c)
    cell.font = font_tot_val
    cell.fill = PatternFill(start_color=C_TOTAL_BG, end_color=C_TOTAL_BG, fill_type="solid")
    cell.border = border_total
    if c in [5, 6, 7]:
        cell.number_format = CURRENCY_FORMAT
        cell.alignment = align_right
        if c == 5:
            cell.font = Font(name="Segoe UI", size=11, bold=True, color="15803D")
        elif c == 6:
            cell.font = Font(name="Segoe UI", size=11, bold=True, color="B91C1C")
        elif c == 7:
            cell.font = Font(name="Segoe UI", size=11, bold=True, color="1D4ED8")

# ==============================================================================
# SHEET 2: IURAN PEMAIN
# ==============================================================================
ws2 = wb.create_sheet(title="Iuran Pemain")
ws2.views.sheetView[0].showGridLines = True
ws2.sheet_properties.tabColor = "2563EB"
ws2.freeze_panes = "A6"

ws2["B2"] = "CATATAN IURAN KAS PEMAIN FUTSAL"
ws2["B2"].font = font_main_title
ws2["B3"] = "Monitoring Pembayaran Iuran Mingguan Per Anggota (Tarif: Rp 20.000 / Pekan)"
ws2["B3"].font = font_subtitle
ws2.row_dimensions[2].height = 24
ws2.row_dimensions[3].height = 18

headers_iuran = [
    ("No", 7),
    ("Nama Pemain", 26),
    ("Posisi / No. Punggung", 22),
    ("Minggu 1 (Rp)", 16),
    ("Minggu 2 (Rp)", 16),
    ("Minggu 3 (Rp)", 16),
    ("Minggu 4 (Rp)", 16),
    ("Total Iuran (Rp)", 18),
    ("Status", 16)
]

ws2.row_dimensions[5].height = 30
for idx, (h_name, h_width) in enumerate(headers_iuran, start=1):
    c = ws2.cell(row=5, column=idx, value=h_name)
    c.font = font_tbl_header
    c.fill = PatternFill(start_color=C_HEADER_FILL, end_color=C_HEADER_FILL, fill_type="solid")
    c.alignment = align_header
    c.border = border_header
    ws2.column_dimensions[get_column_letter(idx)].width = h_width

players = [
    (1, "Rizky (Kapten)", "Flank #7", 20000, 20000, 20000, 20000),
    (2, "Budi Santoso", "Anchor #4", 20000, 20000, 20000, 0),
    (3, "Andi Pratama", "Pivot #9", 20000, 20000, 20000, 20000),
    (4, "Fajar Nugraha", "Kiper #1", 20000, 20000, 20000, 20000),
    (5, "Dimas Setiawan", "Flank #11", 20000, 0, 20000, 0),
    (6, "Eko Wahyudi", "Anchor #5", 20000, 20000, 20000, 20000),
    (7, "Hendra Wijaya", "Pivot #10", 20000, 20000, 20000, 20000),
    (8, "Bayu Aditya", "Kiper #12", 20000, 20000, 0, 20000),
    (9, "Yoga Pratama", "Flank #8", 20000, 20000, 20000, 20000),
    (10, "Gilang Ramadhan", "Anchor #3", 20000, 20000, 20000, 20000),
    (11, "Doni Saputra", "Flank #14", 20000, 20000, 20000, 20000),
    (12, "Taufik Hidayat", "Pivot #17", 0, 20000, 20000, 20000),
]

iuran_start = 6
for idx, p in enumerate(players):
    r = iuran_start + idx
    ws2.row_dimensions[r].height = 22
    row_bg = C_ZEBRA if (idx % 2 == 1) else C_WHITE

    ws2.cell(row=r, column=1, value=p[0]).alignment = align_center
    ws2.cell(row=r, column=2, value=p[1]).alignment = align_left
    ws2.cell(row=r, column=3, value=p[2]).alignment = align_center

    for w_i in range(4):
        c_w = ws2.cell(row=r, column=4 + w_i, value=p[3 + w_i])
        c_w.alignment = align_right
        c_w.number_format = CURRENCY_FORMAT
        c_w.font = font_data

    # Total per player - BOLD
    c_tot = ws2.cell(row=r, column=8, value=f"=SUM(D{r}:G{r})")
    c_tot.alignment = align_right
    c_tot.font = font_data_bold
    c_tot.number_format = CURRENCY_FORMAT

    # Status Lunas/Belum
    c_stat = ws2.cell(row=r, column=9, value=f'=IF(H{r}>=80000, "Lunas", "Kurang Rp " & TEXT(80000-H{r}, "#,##0"))')
    c_stat.alignment = align_center
    c_stat.font = font_data_bold

    for c in range(1, 10):
        cell = ws2.cell(row=r, column=c)
        cell.fill = PatternFill(start_color=row_bg, end_color=row_bg, fill_type="solid")
        cell.border = border_cell

# Total row for Iuran
iuran_tot_r = iuran_start + len(players)
ws2.row_dimensions[iuran_tot_r].height = 28
ws2.merge_cells(start_row=iuran_tot_r, start_column=1, end_row=iuran_tot_r, end_column=3)
c_iuran_lbl = ws2.cell(row=iuran_tot_r, column=1, value="TOTAL IURAN TERKUMPUL")
c_iuran_lbl.font = font_tot_label
c_iuran_lbl.alignment = Alignment(horizontal="right", vertical="center")

for w in range(4):
    col = 4 + w
    col_let = get_column_letter(col)
    c_sub = ws2.cell(row=iuran_tot_r, column=col, value=f"=SUM({col_let}{iuran_start}:{col_let}{iuran_tot_r-1})")
    c_sub.font = font_tot_val
    c_sub.alignment = align_right
    c_sub.number_format = CURRENCY_FORMAT

c_tot_all = ws2.cell(row=iuran_tot_r, column=8, value=f"=SUM(H{iuran_start}:H{iuran_tot_r-1})")
c_tot_all.font = font_tot_val
c_tot_all.alignment = align_right
c_tot_all.number_format = CURRENCY_FORMAT

ws2.cell(row=iuran_tot_r, column=9, value="")

for c in range(1, 10):
    cell = ws2.cell(row=iuran_tot_r, column=c)
    cell.fill = PatternFill(start_color=C_TOTAL_BG, end_color=C_TOTAL_BG, fill_type="solid")
    cell.border = border_total

output_file = "Rekap_Uang_Kas_Futsal.xlsx"
wb.save(output_file)
print(f"Spreadsheet Excel Pro berhasil dibuat: {output_file}")
