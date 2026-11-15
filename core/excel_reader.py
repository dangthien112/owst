"""
Đọc file Excel đơn hàng → danh sách đơn.

Hỗ trợ 2 kiểu bố cục phổ biến:

A. BẢNG NGANG (giống sổ theo dõi giao hàng)
     | Khách hàng     | NV    | SC VNM | PROBI | ADM | ...
     | HKD VĂN AN     | Thiện | 190    | 12    | 99  | ...
   → mỗi dòng = 1 đơn, mỗi cột sản phẩm = số lượng

B. BẢNG DÀI (mỗi dòng 1 mặt hàng)
     | Mã đơn | Khách hàng | Sản phẩm | ĐVT | Số lượng | Đơn giá |
   → các dòng cùng mã đơn được gộp thành 1 đơn
"""
from __future__ import annotations

import csv
import io
from pathlib import Path

from core.text_utils import normalize, parse_date, parse_money, parse_number
from core.pdf_generator import Cfg

# ── Từ khoá nhận diện cột ────────────────────────────────────────────────
ADDRESS_KEYS = ["dia chi", "address", "noi giao", "diem giao"]
PHONE_KEYS = ["dien thoai", "sdt", "phone", "tel"]
CUSTOMER_KEYS = [
    "khach hang", "ten khach", "ten khach hang", "khach", "ho ten",
    "nguoi mua", "nguoi nhan", "dai ly", "cua hang", "ho kinh doanh", "hkd",
    "ten don vi", "don vi mua", "ten truong", "customer", "client", "buyer",
]
ORDER_ID_KEYS = ["ma don", "ma don hang", "so don", "so phieu", "ma phieu", "order", "order id", "ma hd", "so hd"]
PRODUCT_KEYS = ["san pham", "ten hang", "ten san pham", "mat hang", "hang hoa", "product", "item", "ten sp"]
QTY_KEYS = ["so luong", "sl", "qty", "quantity", "so luong giao"]
UNIT_KEYS = ["dvt", "don vi", "don vi tinh", "unit"]
PRICE_KEYS = ["don gia", "gia ban", "don gia ban", "price", "unit price"]
AMOUNT_KEYS = ["thanh tien", "tong tien", "so tien", "amount", "total"]
NOTE_KEYS = ["ghi chu", "note", "gio giao", "thoi gian", "ca giao"]
NGUOI_GIAO_KEYS = ["nguoi giao", "nguoi giao hang", "nhan vien", "nv", "staff"]
DATE_KEYS = ["ngay", "ngay dat", "ngay giao", "date"]

TOTAL_ROW_WORDS = ["tong", "cong", "tong cong", "tong so", "total", "sum"]


def _hits(header_norm: str, keys: list[str]) -> bool:
    """Header có chứa từ khoá không (khớp theo từ với key ngắn, tránh 'sl' khớp 'sale')."""
    if not header_norm:
        return False
    tokens = header_norm.split()
    for k in keys:
        if header_norm == k:
            return True
        if len(k) <= 3:
            if k in tokens:
                return True
        elif k in header_norm:
            return True
    return False


# ── Đọc dữ liệu thô ──────────────────────────────────────────────────────

def load_grid(path: str, sheet: str | int | None = None) -> tuple[list[list], str]:
    """Trả về (grid 2 chiều, tên sheet). Hỗ trợ .xlsx/.xlsm/.csv."""
    p = Path(path)
    suffix = p.suffix.lower()

    if suffix == ".csv":
        raw = p.read_bytes()
        text = None
        for enc in ("utf-8-sig", "utf-8", "cp1258", "latin-1"):
            try:
                text = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            text = raw.decode("utf-8", errors="replace")
        sample = text[:4096]
        try:
            delim = csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
        except csv.Error:
            delim = "," if sample.count(",") >= sample.count(";") else ";"
        return [list(r) for r in csv.reader(io.StringIO(text), delimiter=delim)], p.stem

    if suffix not in (".xlsx", ".xlsm"):
        raise ValueError(
            f"Không hỗ trợ định dạng {suffix or '(không rõ)'}. "
            "Hãy lưu file sang .xlsx (Excel 2007 trở lên) hoặc .csv."
        )

    from openpyxl import load_workbook

    wb = load_workbook(path, data_only=True, read_only=True)
    if sheet is None:
        ws = wb.worksheets[0]
    elif isinstance(sheet, int):
        ws = wb.worksheets[sheet]
    else:
        ws = wb[sheet]
    grid = [list(r) for r in ws.iter_rows(values_only=True)]
    name = ws.title
    wb.close()
    return grid, name


def find_header_row(grid: list[list]) -> int:
    """Tìm dòng tiêu đề (trong 25 dòng đầu): nhiều ô text + có từ khoá nhất."""
    best_idx, best_score = 0, -1
    for idx, row in enumerate(grid[:25]):
        texts = [c for c in row if isinstance(c, str) and c.strip()]
        if len(texts) < 2:
            continue
        norms = [normalize(c) for c in texts]
        kw = sum(
            1 for n in norms
            if _hits(n, CUSTOMER_KEYS + PRODUCT_KEYS + QTY_KEYS + ORDER_ID_KEYS + DATE_KEYS + NOTE_KEYS)
        )
        score = kw * 3 + len(texts)
        if score > best_score:
            best_idx, best_score = idx, score
    return best_idx


def _detect_columns(headers: list) -> dict:
    cols = {"customer": None, "order_id": None, "product": None, "qty": None,
            "unit": None, "price": None, "amount": None, "date": None,
            "address": None, "phone": None, "nguoi_giao": None, "notes": []}
    taken: set[int] = set()

    def take(key: str, keys: list[str]) -> None:
        for i, h in enumerate(headers):
            if i in taken:
                continue
            n = normalize(h)
            if n and _hits(n, keys):
                cols[key] = i
                taken.add(i)
                return

    take("order_id", ORDER_ID_KEYS)
    take("customer", CUSTOMER_KEYS)
    take("nguoi_giao", NGUOI_GIAO_KEYS)
    take("product", PRODUCT_KEYS)
    take("qty", QTY_KEYS)
    take("unit", UNIT_KEYS)
    take("price", PRICE_KEYS)
    take("amount", AMOUNT_KEYS)
    take("date", DATE_KEYS)
    take("phone", PHONE_KEYS)
    take("address", ADDRESS_KEYS)

    for i, h in enumerate(headers):
        if i in taken:
            continue
        n = normalize(h)
        if n and _hits(n, NOTE_KEYS):
            cols["notes"].append(i)
    return cols


def _match_product_header(header, products: list[dict], mapping: dict | None = None) -> dict | None:
    """Khớp tiêu đề cột với 1 sản phẩm trong bảng giá.
    Ưu tiên: khớp theo mapping, khớp nguyên văn, mã hàng là 1 từ trong tiêu đề, chứa tên sản phẩm.
    """
    n = normalize(header)
    if not n:
        return None
    
    # Ưu tiên 1: khớp từ mapping cấu hình
    if mapping and n in mapping:
        pdf_name = mapping[n]
        # Tìm trong danh sách sản phẩm theo tên đã ánh xạ
        for p in products:
            if normalize(p.get("ten", "")) == normalize(pdf_name):
                return p
        # Nếu không tìm thấy sản phẩm cụ thể, trả về đối tượng ảo
        return {"ma": n, "ten": pdf_name, "dvt": "", "gia": 0}

    n_tokens = set(n.split())
    best, best_score = None, 0
    for p in products:
        score = 0
        for token in (normalize(p.get("ma", "")), normalize(p.get("ten", "")),
                      normalize(p.get("ten_ngan", ""))):
            if not token:
                continue
            if token == n:
                score = max(score, 100 + len(token))
            elif len(token) >= 2 and token in n_tokens:
                score = max(score, 60 + len(token))
            elif len(token) >= 4 and token in n:
                score = max(score, 30 + len(token))
            elif len(n) >= 4 and n in token:
                score = max(score, 20 + len(n))
        if score > best_score:
            best, best_score = p, score
    return best


def _cell(row: list, idx: int | None):
    if idx is None or idx >= len(row):
        return None
    v = row[idx]
    return None if v in (None, "") else v


def _generate_code(seq: int, cfg: Cfg) -> str:
    # Lấy số bắt đầu từ cfg
    start = getattr(cfg, "so_bat_dau", 1)
    # Lấy tiền tố từ so_phieu_dau (giữ logic cũ) hoặc so_phieu_dau trực tiếp
    # Nếu so_phieu_dau là "BH00001", tiền tố là "BH", số là 1
    import re
    base = getattr(cfg, "so_phieu_dau", "")
    prefix = ""
    start_val = start
    digits = 5
    
    m = re.search(r"(\d+)$", base)
    if m:
        prefix = base[:m.start()]
        digits = len(m.group(1))
    
    current = start_val + seq - 1
    return f"{prefix}{current:0{digits}d}"

def read_orders(path: str, products: list[dict], cfg: Cfg | None = None, sheet: str | int | None = None) -> list[dict]:
    """
    Trả về list đơn hàng:
      {code, date, customer, phone, address, note,
       items: [{ma, ten, dvt, qty, gia, thanh_tien}],
       subtotal, discount, total, sheet}
    """
    grid, sheet_name = load_grid(path, sheet)
    grid = [r for r in grid if any(c not in (None, "") for c in r)]
    if not grid:
        return []

    hdr = find_header_row(grid)
    headers = grid[hdr]
    cols = _detect_columns(headers)
    body = grid[hdr + 1:]

    # Kiểu B: bảng dài (có cột tên sản phẩm + số lượng)
    if cols["product"] is not None and cols["qty"] is not None:
        return _read_long(body, cols, products, sheet_name, cfg)

    # Kiểu A: bảng ngang, mỗi cột là 1 sản phẩm
    reserved = {v for v in cols.values() if isinstance(v, int)} | set(cols["notes"])
    prod_cols: dict[int, dict] = {}
    mapping = cfg.mapping if cfg and hasattr(cfg, "mapping") else {}
    for i, h in enumerate(headers):
        if i in reserved:
            continue
        p = _match_product_header(h, products, mapping)
        if p is not None:
            prod_cols[i] = p
    if not prod_cols:
        raise ValueError(
            "Không nhận ra cột sản phẩm nào trong file Excel.\n\n"
            "Cách xử lý: mở tab 'Bảng giá' và đặt Mã hàng đúng như tiêu đề cột "
            "(ví dụ cột ghi 'SC VNM' thì mã hàng nên là 'SC' hoặc 'SC VNM'), "
            "hoặc đổi tiêu đề cột trong Excel cho khớp."
        )

    orders: list[dict] = []
    seq = 0
    for row in body:
        if not any(c not in (None, "") for c in row):
            continue
        # bỏ dòng tổng cộng
        row_text = " ".join(normalize(c) for c in row if isinstance(c, str))
        if any(row_text.strip().startswith(w) for w in TOTAL_ROW_WORDS) and all(
            parse_number(_cell(row, ci)) == 0 for ci in prod_cols
        ):
            continue

        customer = str(_cell(row, cols["customer"]) or "").strip()
        items = []
        for ci, p in prod_cols.items():
            qty = parse_number(_cell(row, ci))
            if qty <= 0:
                continue
            gia = parse_money(p.get("gia", 0))
            items.append({
                "ma": p.get("ma", ""),
                "ten": p.get("ten", ""),
                "dvt": p.get("dvt", ""),
                "qty": qty,
                "gia": gia,
                "thanh_tien": round(qty * gia),
            })
        if not items:
            continue
        seq += 1
        code_val = str(_cell(row, cols["order_id"]) or "").strip()
        if not code_val and cfg is not None:
            code_val = _generate_code(seq, cfg)
        else:
            code_val = code_val or str(seq)

        note_parts = [str(_cell(row, ci)).strip() for ci in cols["notes"] if _cell(row, ci) is not None]
        total = round(sum(i["thanh_tien"] for i in items))
        orders.append({
            "code": code_val,
            "date": parse_date(_cell(row, cols["date"])),
            "customer": customer or f"Khách {seq}",
            "nguoi_giao": str(_cell(row, cols["nguoi_giao"]) or "").strip(),
            "phone": str(_cell(row, cols["phone"]) or ""),
            "address": str(_cell(row, cols["address"]) or ""),
            "note": " | ".join(note_parts),
            "items": items,
            "subtotal": total,
            "discount": 0,
            "total": total,
            "sheet": sheet_name,
        })
    return orders


def _read_long(body, cols, products, sheet_name, cfg: Cfg | None = None) -> list[dict]:
    """Bảng dài: gộp các dòng theo mã đơn (hoặc theo khách nếu không có mã đơn)."""
    groups: dict[str, dict] = {}
    counter = 0
    for row in body:
        prod_name = _cell(row, cols["product"])
        qty = parse_number(_cell(row, cols["qty"]))
        if not prod_name or qty <= 0:
            continue
        p = _match_product_header(prod_name, products, cfg.mapping if cfg and hasattr(cfg, "mapping") else {}) or {}
        excel_gia = parse_money(_cell(row, cols["price"])) if cols["price"] is not None else 0
        gia = excel_gia or parse_money(p.get("gia", 0))

        code = str(_cell(row, cols["order_id"]) or "").strip()
        customer = str(_cell(row, cols["customer"]) or "").strip()
        key = code or f"__{customer}_{len(groups)}"
        if key not in groups:
            counter += 1
            if not code and cfg is not None:
                assigned_code = _generate_code(counter, cfg)
            else:
                assigned_code = code or str(counter)

            groups[key] = {
                "code": assigned_code,
                "date": parse_date(_cell(row, cols["date"])),
                "customer": customer or f"Khách {counter}",
                "nguoi_giao": str(_cell(row, cols["nguoi_giao"]) or "").strip(),
                "phone": str(_cell(row, cols["phone"]) or ""),
                "address": str(_cell(row, cols["address"]) or ""),
                "note": "",
                "items": [],
                "sheet": sheet_name,
            }
        groups[key]["items"].append({
            "ma": p.get("ma", ""),
            "ten": p.get("ten") or str(prod_name).strip(),
            "dvt": p.get("dvt") or str(_cell(row, cols["unit"]) or "").strip(),
            "qty": qty,
            "gia": gia,
            "thanh_tien": round(qty * gia),
        })

    out = list(groups.values())
    for o in out:
        o["subtotal"] = round(sum(i["thanh_tien"] for i in o["items"]))
        o["discount"] = 0
        o["total"] = o["subtotal"]
    return out


if __name__ == "__main__":
    import json
    import sys

    if len(sys.argv) < 2:
        print("Dùng: python -m core.excel_reader <file.xlsx>")
        raise SystemExit(1)
    pp = Path(__file__).resolve().parent.parent / "products.json"
    prods = json.loads(pp.read_text(encoding="utf-8"))["products"]
    res = read_orders(sys.argv[1], prods)
    print(f"{len(res)} đơn")
    for o in res[:5]:
        print(o["customer"], "|", o["note"], "| tổng", o["total"],
              "|", [(i["ten"], i["qty"], i["gia"]) for i in o["items"]])
