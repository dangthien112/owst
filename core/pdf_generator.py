"""
Sinh phiếu giao hàng PDF bám theo mẫu PGH của công ty Bảo Hùng.

Khổ A5 ngang (mặc định) 1 liên/trang: header công ty (logo góc trái,
tên + địa chỉ căn giữa), tiêu đề PHIẾU GIAO HÀNG, ngày + số phiếu,
khối thông tin khách hàng (căn lề trái, đỏ/đen giống mẫu),
bảng 7 cột co giãn vừa bề rộng trang, dòng Cộng, số tiền bằng chữ,
số chứng từ gốc kèm theo, chữ ký 2 bên.
"""
from __future__ import annotations

import datetime as dt
import os
import re
from pathlib import Path

from reportlab.lib.pagesizes import A4, A5, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

from core.text_utils import doc_so_tien, fmt_qty, money, money_dec


def get_pagesize(page_type: str = "A5_NGANG"):
    """
    page_type có thể là:
    - 'A5_NGANG': Khổ A5 nằm ngang (A5 landscape) — 1 phiếu/trang đúng mẫu PGH.
    - 'A5'      : Khổ A5 dọc.
    - 'A4'      : Khổ A4.
    """
    if page_type == "A5_NGANG":
        return landscape(A5)
    elif page_type == "A5":
        return A5
    return A4


# Mặc định A5
_DEFAULT_PAGESIZE = A5
PAGE_W, PAGE_H = _DEFAULT_PAGESIZE

# ───────────────────────────── Fonts ─────────────────────────────────────
_FONT_DIR = Path(__file__).resolve().parent.parent / "fonts"
_FONT = "TimesNewRoman"
_FONT_B = "TimesNewRoman-Bold"
_FONT_I = "TimesNewRoman-Italic"
_FONT_BI = "TimesNewRoman-BoldItalic"


def _register_fonts() -> None:
    """Đăng ký font Unicode (bắt buộc để hiển thị đúng tiếng Việt)."""
    if _FONT in pdfmetrics.getRegisteredFontNames():
        return

    # Thử tìm Times New Roman trên Linux (msttcorefonts)
    ttf_path = "/usr/share/fonts/truetype/msttcorefonts/Times_New_Roman.ttf"
    ttf_b = "/usr/share/fonts/truetype/msttcorefonts/Times_New_Roman_Bold.ttf"
    ttf_i = "/usr/share/fonts/truetype/msttcorefonts/Times_New_Roman_Italic.ttf"
    ttf_bi = "/usr/share/fonts/truetype/msttcorefonts/Times_New_Roman_Bold_Italic.ttf"

    if os.path.exists(ttf_path):
        pdfmetrics.registerFont(TTFont(_FONT, ttf_path))
    else:
        pdfmetrics.registerFont(TTFont(_FONT, "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf"))

    if os.path.exists(ttf_b):
        pdfmetrics.registerFont(TTFont(_FONT_B, ttf_b))
    else:
        pdfmetrics.registerFont(TTFont(_FONT_B, "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf"))

    # Nghiêng / đậm nghiêng (nếu không có thì dùng lại font thường)
    if os.path.exists(ttf_i):
        pdfmetrics.registerFont(TTFont(_FONT_I, ttf_i))
    else:
        pdfmetrics.registerFont(TTFont(_FONT_I, ttf_path if os.path.exists(ttf_path)
                                       else "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf"))

    if os.path.exists(ttf_bi):
        pdfmetrics.registerFont(TTFont(_FONT_BI, ttf_bi))
    else:
        pdfmetrics.registerFont(TTFont(_FONT_BI, ttf_b if os.path.exists(ttf_b)
                                       else "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf"))


def wrap_text(text: str, font: str, size: float, max_w: float) -> list[str]:
    """Ngắt dòng theo bề rộng thực của font."""
    text = (text or "").replace("\n", " ").strip()
    if not text:
        return [""]
    words = text.split()
    lines: list[str] = []
    cur = ""
    for w in words:
        probe = f"{cur} {w}".strip()
        if pdfmetrics.stringWidth(probe, font, size) <= max_w:
            cur = probe
        else:
            if cur:
                lines.append(cur)
            # từ quá dài → cắt theo ký tự
            while pdfmetrics.stringWidth(w, font, size) > max_w:
                cut = len(w)
                while cut > 1 and pdfmetrics.stringWidth(w[:cut], font, size) > max_w:
                    cut -= 1
                lines.append(w[:cut])
                w = w[cut:]
            cur = w
    if cur:
        lines.append(cur)
    return lines or [""]


# ─────────────────────────── Bố cục ──────────────────────────────────────
_PROJECT_LOGO = str(Path(__file__).resolve().parent.parent / "logo.png")

MARGIN = 18
GAP = 14
CUT_Y = PAGE_H / 2 - GAP / 2  # đường cắt

# Tỷ lệ cột theo mẫu PGH — quy đổi theo bề rộng thực của trang (tổng = 1.0)
COL_RATIOS = [
    ("STT", 0.06),
    ("Tên hàng", 0.34),
    ("Đơn vị", 0.09),
    ("Số lượng", 0.10),
    ("Đơn giá\n(Bao gồm thuế)", 0.14),
    ("Thành tiền", 0.16),
    ("HSD/EXP", 0.11),
]
COLS: list[tuple[str, float]] = []


def _set_cols(page_width: float) -> None:
    """Co giãn cột cho vừa đúng bề rộng nội dung trang (không lệch trái)."""
    global COLS
    w = page_width - 2 * MARGIN
    COLS = [(label, w * ratio) for label, ratio in COL_RATIOS]


C_DARK = (0.11, 0.14, 0.20)
C_LINE = (0.35, 0.36, 0.40)
C_HEAD = (0.925, 0.933, 0.945)
C_SOFT = (0.60, 0.62, 0.68)
C_RED = (0.80, 0.10, 0.10)  # màu đỏ khối thông tin khách hàng theo mẫu

S_TITLE = 11.5
S_COMPANY = 9.0
S_SMALL = 7.0
S_BODY = 7.6
S_TABLE = 7.6


class Cfg:
    def __init__(self, d: dict | None = None):
        d = d or {}
        self.ten_cong_ty = d.get("ten_cong_ty", "CÔNG TY TNHH TM & DV")
        self.dia_chi = d.get("dia_chi", "")
        self.mst = d.get("mst", "")
        self.dien_thoai = d.get("dien_thoai", "")
        self.email = d.get("email", "")
        self.logo = d.get("logo", "")
        self.tieu_de = d.get("tieu_de", "PHIẾU GIAO HÀNG")
        self.so_lien = int(d.get("so_lien", 1) or 1)
        self.nhan_lien = d.get("nhan_lien", ["", ""])
        self.hien_gia = bool(d.get("hien_don_gia", True))
        self.nguoi_giao = d.get("nguoi_giao", "")
        self.ghi_chu_footer = d.get("ghi_chu_footer", "")
        self.mau_so_phieu = d.get("mau_so_phieu", "")
        self.so_phieu_dau = d.get("so_phieu_dau", "")
        self.so_bat_dau = int(d.get("so_bat_dau", 1) or 1)
        self.delivery_date = d.get("delivery_date", "").strip()


def _date_text(order: dict) -> str:
    """Chuỗi ngày dạng 'Ngày dd Tháng mm Năm yyyy' giống mẫu PGH.
    Nếu ngày bị bỏ trống trong đơn → trả '....' để in dấu chấm."""
    d = (order.get("date") or "").strip()
    if not d or d == "....":
        import datetime as _dt
        return f"Ngày..... Tháng .... Năm {_dt.date.today().year}"
    parts = str(d).split("/")
    if len(parts) == 3:
        return f"Ngày {parts[0]} Tháng {parts[1]} Năm {parts[2]}"
    return f"Ngày {d}"


def _find_logo(cfg: Cfg) -> str:
    """Logo ưu tiên theo cấu hình, nếu không có thì dùng logo.png gốc dự án."""
    if cfg.logo and os.path.exists(cfg.logo):
        return cfg.logo
    return _PROJECT_LOGO if os.path.exists(_PROJECT_LOGO) else ""


def _draw_header(c: canvas.Canvas, x: float, y: float, w: float, cfg: Cfg,
                 order: dict, compact: bool = False, date_txt: str = "") -> float:
    """Vẽ phần đầu phiếu bám theo mẫu PGH: logo góc trái, tên công ty & địa chỉ
    căn giữa, tiêu đề, ngày + số phiếu, khối thông tin khách hàng."""
    # Logo góc trên bên trái
    logo_path = _find_logo(cfg)
    if logo_path:
        try:
            logo_h = 30
            c.drawImage(logo_path, x, y - logo_h, width=logo_h, height=logo_h,
                        preserveAspectRatio=True, mask="auto")
        except Exception:
            pass

    cx = x + w / 2

    # Tên công ty (đậm, HOA) & địa chỉ — căn giữa
    c.setFillColorRGB(*C_DARK)
    c.setFont(_FONT_B, S_COMPANY)
    c.drawCentredString(cx, y, cfg.ten_cong_ty.upper())
    y -= 13
    c.setFont(_FONT, S_SMALL)
    c.setFillColorRGB(0.28, 0.30, 0.35)
    c.drawCentredString(cx, y, cfg.dia_chi)
    y -= 22

    # Tiêu đề PHIẾU GIAO HÀNG
    c.setFont(_FONT_B, 16)
    c.setFillColorRGB(*C_DARK)
    c.drawCentredString(cx, y, cfg.tieu_de.upper())
    y -= 18

    # Ngày (nghiêng) + Số phiếu — căn giữa
    raw_date = order.get("date") or ""
    date_txt = _date_text(order)
    code = order.get("code") or ""
    c.setFont(_FONT_I, 9.5)
    c.setFillColorRGB(*C_DARK)
    c.drawCentredString(cx, y, date_txt)
    y -= 15
    c.setFont(_FONT_B, 10)
    c.drawCentredString(cx, y, f"SỐ: {code}")
    y -= 20

    # Thông tin đơn vị — nhãn thường đen, giá trị ĐẬM đỏ (theo mẫu chuẩn Bảo Hùng)
    cust = order.get("customer") or ""
    addr = order.get("address") or ""
    note = order.get("note") or ""

    # Dòng 1: Tên đơn vị
    c.setFont(_FONT, S_BODY)
    c.setFillColorRGB(*C_DARK)
    label_1 = "Tên đơn vị: "
    w1 = pdfmetrics.stringWidth(label_1, _FONT, S_BODY)
    c.drawString(x, y, label_1)
    c.setFont(_FONT_B, S_BODY)
    c.setFillColorRGB(*C_RED)
    c.drawString(x + w1, y, cust)
    y -= 14

    # Dòng 2: Người đặt hàng
    c.setFont(_FONT, S_BODY)
    c.setFillColorRGB(*C_DARK)
    c.drawString(x, y, "Người đặt hàng: ")
    y -= 14

    # Dòng 3: Địa chỉ
    c.setFont(_FONT, S_BODY)
    c.setFillColorRGB(*C_DARK)
    label_3 = "Địa chỉ: "
    w3 = pdfmetrics.stringWidth(label_3, _FONT, S_BODY)
    c.drawString(x, y, label_3)
    c.setFont(_FONT_B, S_BODY)
    c.setFillColorRGB(*C_RED)
    c.drawString(x + w3, y, addr if addr else "")
    y -= 14

    # Dòng 4: Diễn giải
    c.setFont(_FONT, S_BODY)
    c.setFillColorRGB(*C_DARK)
    label_4 = "Diễn giải: Bán hàng "
    w4 = pdfmetrics.stringWidth(label_4, _FONT, S_BODY)
    c.drawString(x, y, label_4)
    c.setFont(_FONT_B, S_BODY)
    c.setFillColorRGB(*C_RED)
    c.drawString(x + w4, y, cust)
    y -= 14

    # Dòng 5: Lưu ý
    c.setFont(_FONT, S_BODY)
    c.setFillColorRGB(*C_DARK)
    c.drawString(x, y, f"Lưu ý: {note}")
    y -= 16

    return y


def _draw_table_header(c: canvas.Canvas, x: float, y: float) -> float:
    h = 24  # header bảng 2 dòng cho 'Đơn giá (Bao gồm thuế)'
    total_w = sum(wd for _, wd in COLS)
    c.setFillColorRGB(1, 1, 1)
    c.rect(x, y - h, total_w, h, stroke=0, fill=1)
    c.setStrokeColorRGB(*C_LINE)
    c.setLineWidth(0.6)
    c.rect(x, y - h, total_w, h, stroke=1, fill=0)
    c.setFillColorRGB(*C_DARK)
    c.setFont(_FONT_B, S_TABLE - 0.5)
    cx = x
    for i, (label, cw) in enumerate(COLS):
        if "\n" in label:
            lines = label.split("\n")
            c.drawCentredString(cx + cw / 2, y - 9, lines[0])
            c.drawCentredString(cx + cw / 2, y - 19, lines[1])
        else:
            c.drawCentredString(cx + cw / 2, y - h / 2 - 3, label)
        cx += cw
        if i < len(COLS) - 1:
            c.line(cx, y, cx, y - h)
    return y - h


def _row_height(item: dict, w_name: float) -> float:
    label = item["ten"]
    n = len(wrap_text(label, _FONT, S_TABLE, w_name - 8))
    return max(13.5, n * 9.3 + 4.2)


def _draw_row(c: canvas.Canvas, x: float, y: float, h: float, idx: int,
              item: dict, show_price: bool) -> None:
    total_w = sum(wd for _, wd in COLS)
    c.setStrokeColorRGB(*C_LINE)
    c.setLineWidth(0.5)
    c.rect(x, y - h, total_w, h, stroke=1, fill=0)

    c.setFont(_FONT, S_TABLE)
    c.setFillColorRGB(*C_DARK)
    cx = x
    w_name = COLS[1][1]
    label = item["ten"]
    lines = wrap_text(label, _FONT, S_TABLE, w_name - 8)

    # STT
    c.drawCentredString(cx + COLS[0][1] / 2, y - h + 4.6, str(idx))
    cx += COLS[0][1]
    # tên hàng (nhiều dòng, canh trên)
    ty = y - 9.2
    for line in lines:
        c.drawString(cx + 4, ty, line)
        ty -= 9.3
    cx += COLS[1][1]
    # ĐVT
    c.drawCentredString(cx + COLS[2][1] / 2, y - h + 4.6, (item.get("dvt") or "")[:10])
    cx += COLS[2][1]
    # SL (căn phải)
    c.drawRightString(cx + COLS[3][1] - 4, y - h + 4.6, fmt_qty(item["qty"]))
    cx += COLS[3][1]
    # đơn giá (định dạng 1.234.567,00 giống mẫu)
    c.drawRightString(cx + COLS[4][1] - 4, y - h + 4.6,
                      money_dec(item["gia"]) if show_price else "")
    cx += COLS[4][1]
    # thành tiền
    c.drawRightString(cx + COLS[5][1] - 4, y - h + 4.6,
                      money(item["thanh_tien"]) if show_price else "")
    cx += COLS[5][1]
    # HSD/EXP
    c.drawCentredString(cx + COLS[6][1] / 2, y - h + 4.6, "")

    # đường dọc
    lx = x
    for i, (_, cw) in enumerate(COLS[:-1]):
        lx += cw
        c.line(lx, y, lx, y - h)


def _draw_footer_block(c: canvas.Canvas, x: float, y: float, w: float,
                       order: dict, cfg: Cfg, lien: str, extra: float = 0.0,
                       date_txt: str = "") -> float:
    """Tổng cộng + số tiền bằng chữ + số chứng từ gốc + chữ ký. Trả về y dưới cùng."""
    total_w = sum(wd for _, wd in COLS)

    # dòng tổng cộng
    h = 18
    c.setStrokeColorRGB(*C_LINE)
    c.setLineWidth(0.6)
    c.rect(x, y - h, total_w, h, stroke=1, fill=0)

    # Cột 0=STT, 1=Tên, 2=ĐVT, 3=SL, 4=Đơn giá — tw là vị trí bắt đầu ô Thành tiền
    tw = sum(wd for _, wd in COLS[:5])
    c.line(x + tw, y, x + tw, y - h)

    c.setFont(_FONT_B, 9.5)
    c.setFillColorRGB(*C_DARK)
    c.drawString(x + 2, y - h + 5.5, "Cộng")
    c.drawRightString(x + total_w - (COLS[6][1] + 4), y - h + 5.5,
                      money(order.get("total", 0)))
    y -= h

    # số tiền bằng chữ (nhãn đen, giá trị nghiêng)
    words = doc_so_tien(order.get("total", 0))
    y -= 8
    c.setFont(_FONT, S_SMALL + 1.0)
    c.setFillColorRGB(*C_DARK)
    label = "Số tiền bằng chữ: "
    c.drawString(x + 2, y, label)
    c.setFont(_FONT_BI, S_SMALL + 1.0)
    c.drawString(x + 2 + pdfmetrics.stringWidth(label, _FONT_BI, S_SMALL + 1.0), y, words)
    y -= 14

    # số chứng từ gốc kèm theo
    c.setFont(_FONT, S_SMALL + 1.0)
    c.setFillColorRGB(*C_DARK)
    c.drawString(x + 2, y, "Số chứng từ gốc kèm theo: ...")
    y -= 16

    # chữ ký: 2 cột, cột phải có dòng ngày phía trên (giống mẫu)
    sign_h = 56
    y -= 2 + max(0.0, extra)
    col_w = w / 2
    c.setFont(_FONT_I, 9)
    c.setFillColorRGB(*C_DARK)
    if date_txt:
        c.drawCentredString(x + col_w + col_w / 2, y, date_txt)
        y -= 13

    c.setFont(_FONT_B, S_BODY + 0.5)
    c.drawCentredString(x + col_w / 2, y - 10, "Người mua hàng")
    c.drawCentredString(x + col_w + col_w / 2, y - 10, "Người giao hàng")
    y -= 26

    c.setFont(_FONT_I, S_SMALL)
    c.setFillColorRGB(0.35, 0.37, 0.42)
    c.drawCentredString(x + col_w / 2, y, "(Ký, họ tên)")
    c.drawCentredString(x + col_w + col_w / 2, y, "(Ký, họ tên)")
    y -= sign_h

    return y - 10


def _measure_copy(order: dict, cfg: Cfg, w: float, show_price: bool,
                  items: list[dict] | None = None, cont: bool = False) -> float:
    """Đo chính xác chiều cao 1 liên bằng cách vẽ thử ra canvas rỗng."""
    from io import BytesIO

    _register_fonts()
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=_DEFAULT_PAGESIZE)
    top = PAGE_H - MARGIN
    y = _render_copy(c, MARGIN, top, w, order, cfg, "", show_price, items=items, cont=cont)
    y = _draw_footer_block(c, MARGIN, y, w, order, cfg, "", date_txt=_date_text(order))
    c.showPage()
    return top - y


# giữ tên cũ cho tương thích
measure_copies = _measure_copy


def _render_copy(c: canvas.Canvas, x: float, y_top: float, w: float, order: dict,
                 cfg: Cfg, lien: str, show_price: bool, items: list[dict] | None = None,
                 cont: bool = False) -> float:
    """Vẽ 1 liên tại (x, y_top). Trả về y sau khi vẽ xong."""
    y = y_top
    date_txt = _date_text(order)

    if cont:
        c.setFont(_FONT_B, 8.5)
        c.setFillColorRGB(*C_DARK)
        c.drawString(x, y - 9, f"{cfg.tieu_de} (tiếp theo) — {order.get('customer','')} — Số: {order.get('code','')}")
        y -= 13
        c.setStrokeColorRGB(*C_LINE)
        c.line(x, y, x + w, y)
        y -= 4
    else:
        y = _draw_header(c, x, y, w, cfg, order, date_txt=date_txt)

    y = _draw_table_header(c, x, y)
    for i, item in enumerate(items if items is not None else order.get("items", []), start=1):
        h = _row_height(item, COLS[1][1])
        _draw_row(c, x, y, h, i, item, show_price)
        y -= h
    return y


def build_order_pdf(order: dict, out_path: str, cfg: Cfg, show_price: bool = True,
                    so_lien: int | None = None, page_type: str = "A5_DOC") -> str:
    """Sinh 1 file PDF cho 1 đơn. Trả về đường dẫn file."""
    _register_fonts()
    so_lien = int(so_lien if so_lien is not None else cfg.so_lien)
    so_lien = max(1, min(3, so_lien))
    cfg.so_lien = so_lien

    pagesize = get_pagesize(page_type)
    pw, ph = pagesize
    _set_cols(pw)

    out_path = str(out_path)
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)

    w = pw - 2 * MARGIN

    c = canvas.Canvas(out_path, pagesize=pagesize)
    c.setTitle(f"{cfg.tieu_de} {order.get('code','')}")
    c.setAuthor(cfg.ten_cong_ty)
    c.setSubject(order.get("customer", ""))

    # Căn giữa toàn bộ phần vẽ
    render_x = (pw - w) / 2
    date_txt = _date_text(order)

    # Render các liên (so_lien = 1 → chỉ 1 trang duy nhất, không lặp)
    for lien_i in range(so_lien):
        y_after = _render_copy(c, render_x, ph - MARGIN, w, order, cfg, "", show_price)
        # Luôn vẽ footer: dòng Cộng, tiền bằng chữ, chứng từ gốc, chữ ký
        _draw_footer_block(c, render_x, y_after, w, order, cfg, "", date_txt=date_txt)
        c.showPage()
        if so_lien == 1:
            break

    c.save()
    return out_path


def _draw_cut_marks(c: canvas.Canvas) -> None:
    """Đường cắt giữa 2 liên + nhãn kéo cắt (tắt theo yêu cầu — mẫu không có)."""
    # c.saveState()
    # c.setStrokeColorRGB(*C_SOFT)
    # c.setLineWidth(0.4)
    # c.setDash(2, 2)
    # c.line(MARGIN, CUT_Y, PAGE_W - MARGIN, CUT_Y)
    # c.setDash()
    # c.setFont(_FONT, 6.2)
    # c.setFillColorRGB(0.42, 0.45, 0.50)
    # c.drawCentredString(PAGE_W / 2, CUT_Y + 6, "— — — — — —  CẮT THEO ĐƯỜNG NÀY  — — — — — —")
    # c.restoreState()


def safe_filename(text: str, max_len: int = 80) -> str:
    """Bỏ ký tự cấm trong tên file Windows."""
    s = re.sub(r'[\\/:*?"<>|\r\n\t]+', " ", str(text or "")).strip()
    s = re.sub(r"\s+", " ", s)
    return s[:max_len].strip() or "don_hang"


def build_all(orders: list[dict], out_dir: str, cfg: Cfg, show_price: bool = True,
              so_lien: int | None = None, one_file: bool = False,
              progress=None, page_type: str = "A5_DOC") -> list[str]:
    """Sinh PDF cho toàn bộ danh sách đơn. Trả về list đường dẫn đã tạo."""
    os.makedirs(out_dir, exist_ok=True)
    paths: list[str] = []
    total = len(orders)
    so_lien = int(so_lien if so_lien is not None else cfg.so_lien)
    so_lien = max(1, min(3, so_lien))

    from core.pdf_generator import _render_copy, _draw_footer_block, _register_fonts, MARGIN

    pagesize = get_pagesize(page_type)
    pw, ph = pagesize
    _set_cols(pw)

    # Nhóm theo nhân viên giao hàng nếu có
    if one_file:
        from collections import defaultdict
        groups = defaultdict(list)
        for o in orders:
            groups[(o.get("nguoi_giao") or "").strip() or "Chung"].append(o)

        from reportlab.pdfgen import canvas as _cv
        _register_fonts()
        render_w = pw - 2 * MARGIN
        render_x = (pw - render_w) / 2

        for nv, group in groups.items():
            fname = f"Phieu_giao_hang_NV_{safe_filename(nv)}.pdf"
            target = os.path.join(out_dir, fname)
            c = _cv.Canvas(target, pagesize=pagesize)

            for idx, o in enumerate(group, start=1):
                date_txt = _date_text(o)
                for k in range(so_lien):
                    y = _render_copy(c, render_x, ph - MARGIN, render_w, o, cfg, "", show_price)
                    # Luôn vẽ footer cho mỗi phiếu
                    _draw_footer_block(c, render_x, y, render_w, o, cfg, "", date_txt=date_txt)
                    c.showPage()
                    if so_lien == 1:
                        break
                if progress:
                    progress(idx, len(group), o)
            c.save()
            paths.append(target)
        return paths

    for idx, o in enumerate(orders, start=1):
        name = f"{safe_filename(o.get('code') or idx)}_{safe_filename(o.get('customer'))}.pdf"
        p = os.path.join(out_dir, name)
        n = 1
        while os.path.exists(p) and not p.endswith(f"__{n}.pdf"):
            p = os.path.join(out_dir, f"{safe_filename(o.get('code') or idx)}_{safe_filename(o.get('customer'))}__{n}.pdf")
            n += 1
        build_order_pdf(o, p, cfg, show_price=show_price, so_lien=so_lien, page_type=page_type)
        paths.append(p)
        if progress:
            progress(idx, total, o)
    return paths


if __name__ == "__main__":
    demo = {
        "code": "BH32734",
        "date": dt.date.today().strftime("%d/%m/%Y"),
        "customer": "TRƯỜNG MẦM NON HỒNG VÂN 2 – phân hiệu HÀ HỘI",
        "address": "Số 106 đường Phú Cốc, thôn Khê Hồi, Xã Hồng Vân, TP Hà Nội",
        "phone": "",
        "items": [
            {"ma": "SC", "ten": "Sữa chua ăn VNM có đường 60g", "dvt": "Hộp",
             "qty": 720, "gia": 3942, "thanh_tien": 2838240},
        ],
        "total": 2838240,
    }
    cfg = Cfg({
        "ten_cong_ty": "CÔNG TY TNHH THƯƠNG MẠI VÀ DỊCH VỤ ĐẦU TƯ BẢO HÙNG",
        "dia_chi": "Số 12, Ngách 291/61 Lê Trọng Tấn, Phường Dương Nội, Thành phố Hà Nội, Việt Nam",
        "mst": "01 0004448",
        "dien_thoai": "0385 155 956",
    })
    print(build_order_pdf(demo, "/home/thien/order2pdf/samples/demo.pdf", cfg, page_type="A5_NGANG"))
