"""Tạo file Excel mẫu giống sổ theo dõi giao hàng để thử ứng dụng."""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

OUT = Path(__file__).resolve().parent.parent / "samples" / "So_giao_hang_mau.xlsx"

HEADERS = [
    "Khách hàng", "NV", "Ba vị 100g", "PROBI 6.5ML", "SC VNM",
    "SB VNM", "Sữa chua", "ADM", "Men sống", "Ghi chú",
]
ROWS = [
    ["HKD VĂN AN",            "Thiện", 0,   12,  190, 0,  0,   0,  0,  "trưa 11h"],
    ["CÔNG TY CÁP BẮC",       "Thiện", 0,   0,   0,   24, 0,   0,  0,  "trong ngày"],
    ["MN HOA MI TT",          "Thiện", 40,  20,  120, 0,  60,  30, 12, "thứ 2"],
    ["MN HÀ HỘI",             "Thiện", 0,   10,  720, 0,  0,   0,  0,  "1q HĐ"],
    ["MN LÊ LỢI",             "Thiện", 25,  8,   260, 0,  0,   24, 7,  ""],
    ["MN QUANG LĂNG",         "Thiện", 0,   0,   183, 12, 0,   45, 0,  "trong ngày"],
]


def main() -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Giao hàng"

    ws["A1"] = "SỔ THEO DÕI GIAO HÀNG"
    ws["A1"].font = Font(bold=True, size=14)
    ws.merge_cells("A1:J1")
    ws["A1"].alignment = Alignment(horizontal="center")
    ws["A2"] = "Tháng 01/2026"
    ws.merge_cells("A2:J2")
    ws["A2"].alignment = Alignment(horizontal="center")

    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    head_fill = PatternFill("solid", fgColor="DDE5F0")

    header_row = 4
    for c, h in enumerate(HEADERS, start=1):
        cell = ws.cell(row=header_row, column=c, value=h)
        cell.font = Font(bold=True)
        cell.fill = head_fill
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for r, row in enumerate(ROWS, start=header_row + 1):
        for c, v in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = border
            if c >= 3:
                cell.alignment = Alignment(horizontal="center")

    widths = [22, 10, 12, 13, 10, 10, 11, 9, 12, 18]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = w

    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    return OUT


if __name__ == "__main__":
    print("Đã tạo:", main())
