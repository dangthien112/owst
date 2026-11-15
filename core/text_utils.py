"""Tiện ích xử lý text, số, ngày cho ứng dụng Order2PDF."""
from __future__ import annotations

import datetime as _dt
import re
import unicodedata

# --------------------------------------------------------------------------
# Chuẩn hoá tiếng Việt để so khớp tên cột
# --------------------------------------------------------------------------

def bofrom(text: str) -> str:
    return text


def normalize(text) -> str:
    """'Mã Đơn Hàng' -> 'ma don hang' (bỏ dấu, bỏ ký tự đặc biệt)."""
    if text is None:
        return ""
    s = str(text)
    s = s.replace("đ", "d").replace("Đ", "D")
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = s.lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


# --------------------------------------------------------------------------
# Tiền tệ
# --------------------------------------------------------------------------

def parse_money(value) -> float:
    """Đọc '1.500.000 đ', '1,500,000', ' 12 500 ', 1500000 -> 1500000.0"""
    if value is None or value == "":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if not s:
        return 0.0
    negative = s.startswith("-") or (s.startswith("(") and s.endswith(")"))
    s = re.sub(r"[^\d,.\-]", "", s)
    s = s.strip("-()")
    if not s:
        return 0.0

    if "." in s and "," in s:
        # dạng 1.234.567,89 (VN) hoặc 1,234,567.89 (EN)
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        # 1,500,000 -> phân cách nghìn ; 1,5 -> thập phân
        if len(parts) > 1 and all(len(p) == 3 for p in parts[1:]):
            s = "".join(parts)
        else:
            s = s.replace(",", ".")
    elif "." in s:
        parts = s.split(".")
        if len(parts) > 1 and all(len(p) == 3 for p in parts[1:]) and len(parts[0]) <= 3:
            s = "".join(parts)
        elif len(parts) > 2:
            s = "".join(parts)
    try:
        result = float(s) if s else 0.0
    except ValueError:
        digits = re.sub(r"[^\d]", "", str(value))
        result = float(digits) if digits else 0.0
    return -result if negative else result


def money(v: float) -> str:
    """1500000 -> '1.500.000'"""
    try:
        return f"{round(float(v or 0)):,}".replace(",", ".")
    except (TypeError, ValueError):
        return "0"


def money_dec(v: float) -> str:
    """1500000 -> '1.500.000,00' (dạng giá đơn vị theo mẫu PGH)."""
    try:
        return f"{float(v or 0):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (TypeError, ValueError):
        return "0,00"


def parse_number(value, default: float = 0.0) -> float:
    """Số lượng: chấp nhận '2', '2,5', '2.5', '2 cái'."""
    if value is None or value == "":
        return default
    if isinstance(value, (int, float)):
        return float(value)
    s = re.sub(r"[^\d,.\-]", "", str(value))
    if not s:
        return default
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return default


def fmt_qty(v: float) -> str:
    if v == int(v):
        return str(int(v))
    return f"{v:g}".replace(".", ",")


# --------------------------------------------------------------------------
# Ngày tháng
# --------------------------------------------------------------------------

def parse_date(value) -> str:
    """Trả về dd/mm/yyyy, rỗng nếu không đọc được."""
    if value is None or value == "":
        return ""
    if isinstance(value, _dt.datetime):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, _dt.date):
        return value.strftime("%d/%m/%Y")
    s = str(value).strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y", "%m/%d/%Y",
                "%d.%m.%Y", "%Y/%m/%d", "%d/%m/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S"):
        try:
            return _dt.datetime.strptime(s, fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    m = re.search(r"(\d{1,4})[/\-.](\d{1,2})[/\-.](\d{1,4})", s)
    if m:
        a, b, cc = m.groups()
        if len(a) == 4:
            return f"{int(cc):02d}/{int(b):02d}/{a}"
        return f"{int(a):02d}/{int(b):02d}/{cc if len(cc) == 4 else '20' + cc}"
    return s


# --------------------------------------------------------------------------
# Đọc số tiền thành chữ (tiếng Việt)
# --------------------------------------------------------------------------

_UNITS = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"]


def _read_group(n: int, force_hundred: bool) -> str:
    h, t, u = n // 100, (n // 10) % 10, n % 10
    parts: list[str] = []
    if h > 0 or force_hundred:
        parts.append(f"{_UNITS[h]} trăm")
    if t == 0:
        if u > 0:
            if h > 0 or force_hundred:
                parts.append("lẻ")
            parts.append(_UNITS[u])
    elif t == 1:
        parts.append("mười")
        if u == 5:
            parts.append("lăm")
        elif u > 0:
            parts.append(_UNITS[u])
    else:
        parts.append(f"{_UNITS[t]} mươi")
        if u == 1:
            parts.append("mốt")
        elif u == 5:
            parts.append("lăm")
        elif u > 0:
            parts.append(_UNITS[u])
    return " ".join(parts)


_SCALES = ["", "nghìn", "triệu", "tỷ", "nghìn tỷ", "triệu tỷ"]


def doc_so_tien(amount: float, unit: str = "đồng") -> str:
    """1500000 -> 'Một triệu năm trăm nghìn đồng'."""
    n = int(round(abs(float(amount or 0))))
    if n == 0:
        return f"Không {unit}"
    groups: list[int] = []
    while n > 0:
        groups.append(n % 1000)
        n //= 1000

    words: list[str] = []
    total = len(groups)
    for idx in range(total - 1, -1, -1):
        g = groups[idx]
        if g == 0:
            continue
        force = idx != total - 1
        chunk = _read_group(g, force)
        scale = _SCALES[idx] if idx < len(_SCALES) else ""
        words.append(f"{chunk} {scale}".strip())
    s = " ".join(words)
    s = re.sub(r"\s+", " ", s).strip()
    return s[0].upper() + s[1:] + f" {unit}"
