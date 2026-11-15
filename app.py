"""
Order2PDF — chuyển file Excel đơn hàng thành phiếu giao hàng PDF.

Chạy có giao diện:      python app.py
Chạy không giao diện:   python app.py --cli <file.xlsx> --out <thu_muc> [--one-file]
"""
from __future__ import annotations

import json
import os
import queue
import shutil
import subprocess
import sys
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# đảm bảo import được core/ khi chạy từ thư mục khác
ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.pdf_generator import Cfg, build_all, build_order_pdf
from core.excel_reader import read_orders
from core.text_utils import money


# ───────────────────────── Đường dẫn khi đóng gói ────────────────────────
def app_dir() -> Path:
    """Thư mục chứa app (khi build exe là thư mục chứa .exe)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return ROOT


def res_dir() -> Path:
    """Thư mục chứa tài nguyên đi kèm (fonts...) khi build exe."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return ROOT


PRODUCTS_FILE = app_dir() / "products.json"
CONFIG_FILE = app_dir() / "config.json"

DEFAULT_PRODUCTS = {"products": []}
DEFAULT_CONFIG = {
    "ten_cong_ty": "CÔNG TY CỦA BẠN",
    "dia_chi": "",
    "mst": "",
    "dien_thoai": "",
    "email": "",
    "logo": "",
    "tieu_de": "PHIẾU GIAO HÀNG",
    "so_lien": 2,
    "nhan_lien": ["", ""],
    "hien_don_gia": True,
    "nguoi_giao": "",
    "ghi_chu_footer": "",
    "tien_to_so_phieu": "BH",
    "so_bat_dau": 1,
    "dung_ngay_hom_nay": True,
    "delivery_date": "",
}


def load_json(path: Path, default: dict) -> dict:
    """Đọc JSON; nếu chưa có thì tạo từ bản mẫu đi kèm (hoặc default)."""
    if path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            pass
    bundled = res_dir() / path.name
    if bundled.exists() and bundled != path:
        try:
            data = json.loads(bundled.read_text(encoding="utf-8"))
            path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
            return data
        except Exception:
            pass
    return dict(default)


def save_json(path: Path, data) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def open_folder(path: str) -> None:
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


# ─────────────────────────── Ứng dụng ────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Order2PDF — Excel → Phiếu giao hàng PDF")
        self.geometry("1060x720")
        self.minsize(900, 620)

        self.cfg_data = load_json(CONFIG_FILE, DEFAULT_CONFIG)
        for k, v in DEFAULT_CONFIG.items():
            self.cfg_data.setdefault(k, v)
        self.products_data = load_json(PRODUCTS_FILE, DEFAULT_PRODUCTS)
        self.products_data.setdefault("products", [])

        self.excel_path: str | None = None
        self.orders: list[dict] = []
        self.out_dir = tk.StringVar(value=str(Path.home() / "PhieuGiaoHang"))
        self.one_file = tk.BooleanVar(value=False)
        self.show_price = tk.BooleanVar(value=bool(self.cfg_data.get("hien_don_gia", True)))
        self.status = tk.StringVar(value="Sẵn sàng")
        self.log_q: queue.Queue = queue.Queue()
        self.worker: threading.Thread | None = None

        self._build_style()
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(120, self._drain_queue)

    # ── giao diện ────────────────────────────────────────────────────────
    def _build_style(self):
        st = ttk.Style(self)
        try:
            st.theme_use("clam")
        except tk.TclError:
            pass
        st.configure(".", font=("Segoe UI", 10))
        st.configure("Title.TLabel", font=("Segoe UI Semibold", 15))
        st.configure("Sub.TLabel", font=("Segoe UI", 9), foreground="#666")
        st.configure("Accent.TButton", font=("Segoe UI Semibold", 10), padding=(14, 7))
        st.configure("TNotebook.Tab", padding=(16, 8), font=("Segoe UI", 10))
        st.configure("Treeview", rowheight=24)
        st.configure("Treeview.Heading", font=("Segoe UI Semibold", 9))

    def _build_ui(self):
        head = ttk.Frame(self, padding=(16, 14, 16, 6))
        head.pack(fill="x")
        ttk.Label(head, text="Chuyển Excel thành phiếu giao hàng PDF", style="Title.TLabel").pack(anchor="w")
        ttk.Label(head, text="Mỗi dòng trong Excel = 1 phiếu. Đơn giá lấy từ tab Bảng giá.",
                  style="Sub.TLabel").pack(anchor="w", pady=(2, 0))

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=14, pady=8)
        self.nb = nb
        self.tab_main = ttk.Frame(nb, padding=12)
        self.tab_price = ttk.Frame(nb, padding=12)
        self.tab_customer = ttk.Frame(nb, padding=12)
        self.tab_company = ttk.Frame(nb, padding=12)
        nb.add(self.tab_main, text="Xuất phiếu")
        nb.add(self.tab_price, text="Bảng giá mặt hàng")
        nb.add(self.tab_customer, text="Thông tin khách hàng")
        nb.add(self.tab_company, text="Thông tin công ty")
        nb.bind("<<NotebookTabChanged>>", self._on_tab_change)

        self._build_main_tab()
        self._build_price_tab()
        self._build_customer_tab()
        self._build_company_tab()

        bar = ttk.Frame(self, padding=(16, 4, 16, 10))
        bar.pack(fill="x")
        ttk.Label(bar, textvariable=self.status, style="Sub.TLabel").pack(side="left")

    # ── tab 1 ────────────────────────────────────────────────────────────
    def _build_main_tab(self):
        f = self.tab_main

        box = ttk.LabelFrame(f, text=" 1. File Excel đơn hàng ", padding=10)
        box.pack(fill="x")
        row = ttk.Frame(box)
        row.pack(fill="x")
        self.lbl_file = ttk.Label(row, text="Chưa chọn file — bấm 'Chọn file Excel…'",
                                  foreground="#888")
        self.lbl_file.pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="Chọn file Excel…", command=self.pick_file, width=20).pack(side="right")
        ttk.Button(row, text="Đọc lại", command=self.reload_orders, width=10).pack(side="right", padx=(0, 6))

        box2 = ttk.LabelFrame(f, text=" 2. Cấu hình in ", padding=10)
        box2.pack(fill="x", pady=(10, 0))
        row2 = ttk.Frame(box2)
        row2.pack(fill="x")
        ttk.Label(row2, text="Tiền tố số phiếu:").pack(side="left", padx=(0, 6))
        self.prefix_var = tk.StringVar(value=str(self.cfg_data.get("tien_to_so_phieu", "BH")))
        ttk.Entry(row2, textvariable=self.prefix_var, width=10).pack(side="left")
        ttk.Label(row2, text="Số bắt đầu:").pack(side="left", padx=(12, 6))
        self.start_var = tk.StringVar(value=str(self.cfg_data.get("so_bat_dau", 1)))
        ttk.Entry(row2, textvariable=self.start_var, width=10).pack(side="left")

        row3 = ttk.Frame(box2)
        row3.pack(fill="x", pady=(10, 0))
        ttk.Label(row3, text="Nơi lưu PDF:").pack(side="left", padx=(0, 6))
        ttk.Entry(row3, textvariable=self.out_dir).pack(side="left", fill="x", expand=True)
        ttk.Button(row3, text="Mở", width=6,
                   command=lambda: open_folder(self.out_dir.get())).pack(side="left", padx=(6, 0))
        ttk.Button(row3, text="Chọn…", width=10, command=self.pick_outdir).pack(side="right")

        self.page_size = tk.StringVar(value="A5")

        opts = ttk.Frame(box2)
        opts.pack(fill="x", pady=(10, 0))
        ttk.Checkbutton(opts, text="Hiện đơn giá", variable=self.show_price).pack(side="left")
        ttk.Checkbutton(opts, text="Gộp đơn theo người giao vào 1 file PDF",
                        variable=self.one_file).pack(side="left", padx=(18, 0))
        ttk.Label(opts, text="Khổ giấy:").pack(side="left", padx=(18, 6))
        ttk.Combobox(opts, textvariable=self.page_size, values=["A4", "A5", "A5_NGANG"], 
                     width=10, state="readonly").pack(side="left")

        self.delivery_date_var = tk.StringVar(value="")
        self.delivery_date_var.trace_add("write", self._on_date_change)
        date_row = ttk.Frame(box2)
        date_row.pack(fill="x", pady=(10, 0))
        ttk.Label(date_row, text="Ngày giao hàng (nếu rỗng → ....):").pack(side="left", padx=(0, 6))
        ttk.Entry(date_row, textvariable=self.delivery_date_var, width=20).pack(side="left")
        ttk.Label(date_row, text="Định dạng dd/mm/yyyy", style="Sub.TLabel").pack(side="left", padx=(6, 0))
        self.blank_date = tk.BooleanVar(value=False)
        ttk.Checkbutton(date_row, text="Để trống để viết tay", variable=self.blank_date,
                        command=self._on_date_change).pack(side="left", padx=(12, 0))

        box3 = ttk.LabelFrame(f, text=" 3. Kiểm tra danh sách rồi xuất ", padding=10)
        box3.pack(fill="both", expand=True, pady=(10, 0))

        table = ttk.Frame(box3)
        table.pack(fill="both", expand=True)
        cols = ("customer", "date", "nguoi_giao", "note")
        self.tree = ttk.Treeview(table, columns=cols, show="headings", height=8)
        self.tree.tag_configure("odd", background="#F5F7FA")
        for cid, label, w, anchor in (
            ("customer", "Khách hàng", 300, "w"),
            ("date", "Ngày giao", 110, "center"),
            ("nguoi_giao", "Người giao", 100, "center"),
            ("note", "Ghi chú", 320, "w"),
        ):
            self.tree.heading(cid, text=label, anchor=anchor)
            self.tree.column(cid, width=w, minwidth=70, anchor=anchor,
                             stretch=(cid in ("customer", "note")))
        vs = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vs.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vs.pack(side="left", fill="y")

        actions = ttk.Frame(box3)
        actions.pack(fill="x", pady=(10, 0))
        self.pb = ttk.Progressbar(actions, mode="determinate")
        self.pb.pack(side="left", fill="x", expand=True, padx=(0, 12))
        ttk.Button(actions, text="Mở thư mục", width=14,
                   command=lambda: open_folder(self.out_dir.get())).pack(side="left", padx=(0, 8))
        self.btn_export = ttk.Button(actions, text="XUẤT PDF", style="Accent.TButton",
                                     width=16, command=self.export_pdf)
        self.btn_export.pack(side="left")
        self.btn_read = self.btn_export  # giữ tương thích

    # ── tab 2 ────────────────────────────────────────────────────────────
    def _build_price_tab(self):
        f = self.tab_price
        ttk.Label(f, text="Mỗi mặt hàng: mã (khớp tiêu đề cột trong Excel) · tên in trên phiếu · ĐVT · đơn giá.",
                  style="Sub.TLabel").pack(anchor="w", pady=(0, 8))

        wrap = ttk.Frame(f)
        wrap.pack(fill="both", expand=True)
        cols = ("ma", "ten_ngan", "ten", "dvt", "gia")
        self.price_tree = ttk.Treeview(wrap, columns=cols, show="headings", height=14)
        self.price_tree.tag_configure("odd", background="#F5F7FA")
        for cid, label, w, anchor in (
            ("ma", "Mã hàng", 75, "w"),
            ("ten_ngan", "Tên/Mã trên Excel", 140, "w"),
            ("ten", "Tên in trên phiếu", 280, "w"),
            ("dvt", "ĐVT", 60, "center"),
            ("gia", "Đơn giá (đ)", 100, "e"),
        ):
            self.price_tree.heading(cid, text=label, anchor=anchor)
            self.price_tree.column(cid, width=w, minwidth=50, anchor=anchor,
                                   stretch=(cid == "ten"))
        vs = ttk.Scrollbar(wrap, orient="vertical", command=self.price_tree.yview)
        self.price_tree.configure(yscrollcommand=vs.set)
        self.price_tree.pack(side="left", fill="both", expand=True)
        vs.pack(side="left", fill="y")
        self.price_tree.bind("<Double-1>", lambda e: self.edit_product())

        btns = ttk.Frame(f)
        btns.pack(fill="x", pady=(10, 0))
        ttk.Button(btns, text="Thêm mặt hàng", command=self.add_product).pack(side="left")
        ttk.Button(btns, text="Sửa mặt hàng", command=self.edit_product).pack(side="left", padx=6)
        ttk.Button(btns, text="Xoá", command=self.del_product).pack(side="left")
        ttk.Button(btns, text="Nạp từ Excel...", command=self.import_products_excel).pack(side="left", padx=(18, 0))
        ttk.Button(btns, text="Xuất Excel", command=self.export_products_excel).pack(side="left", padx=(6, 0))
        ttk.Button(btns, text="Lưu bảng giá", style="Accent.TButton", command=self.save_products).pack(side="right")

        self._fill_price_tree()

    def _build_customer_tab(self):
        f = self.tab_customer
        ttk.Label(f, text="Quản lý thông tin khách nhận hàng và ánh xạ tên đơn vị với mã khách trên file Excel.",
                  style="Sub.TLabel").pack(anchor="w", pady=(0, 8))

        wrap = ttk.Frame(f)
        wrap.pack(fill="both", expand=True)
        cols = ("excel_code", "ten_don_vi", "nguoi_dat", "dia_chi", "dien_giai")
        self.cust_tree = ttk.Treeview(wrap, columns=cols, show="headings", height=14)
        self.cust_tree.tag_configure("odd", background="#F5F7FA")
        for cid, label, w, anchor in (
            ("excel_code", "Mã/Tên trên Excel", 150, "w"),
            ("ten_don_vi", "Tên đơn vị (In trên phiếu)", 200, "w"),
            ("nguoi_dat", "Người đặt hàng", 120, "w"),
            ("dia_chi", "Địa chỉ giao hàng", 250, "w"),
            ("dien_giai", "Diễn giải", 120, "w"),
        ):
            self.cust_tree.heading(cid, text=label, anchor=anchor)
            self.cust_tree.column(cid, width=w, minwidth=60, anchor=anchor, stretch=(cid in ("dia_chi", "ten_don_vi")))
        vs = ttk.Scrollbar(wrap, orient="vertical", command=self.cust_tree.yview)
        self.cust_tree.configure(yscrollcommand=vs.set)
        self.cust_tree.pack(side="left", fill="both", expand=True)
        vs.pack(side="left", fill="y")
        self.cust_tree.bind("<Double-1>", lambda e: self.edit_customer())

        btns = ttk.Frame(f)
        btns.pack(fill="x", pady=(10, 0))
        ttk.Button(btns, text="Thêm khách hàng", command=self.add_customer).pack(side="left")
        ttk.Button(btns, text="Sửa khách hàng", command=self.edit_customer).pack(side="left", padx=6)
        ttk.Button(btns, text="Xoá", command=self.del_customer).pack(side="left")
        ttk.Button(btns, text="Nạp từ Excel...", command=self.import_customers_excel).pack(side="left", padx=(18, 0))
        ttk.Button(btns, text="Xuất Excel", command=self.export_customers_excel).pack(side="left", padx=(6, 0))
        ttk.Button(btns, text="Lưu thông tin", style="Accent.TButton", command=self.save_customers).pack(side="right")

        self._fill_customer_tree()

    def _fill_customer_tree(self):
        self.cust_tree.delete(*self.cust_tree.get_children())
        customers = self.cfg_data.get("customers", [])
        for i, c in enumerate(customers):
            self.cust_tree.insert("", "end", values=(
                c.get("excel_code", ""),
                c.get("ten_don_vi", ""),
                c.get("nguoi_dat", ""),
                c.get("dia_chi", ""),
                c.get("dien_giai", ""),
            ), tags=("odd",) if i % 2 else ())

    def add_customer(self):
        d = self._customer_dialog("Thêm khách hàng", {})
        if d:
            self.cfg_data.setdefault("customers", []).append(d)
            self._fill_customer_tree()
            save_json(CONFIG_FILE, self.cfg_data)

    def edit_customer(self):
        sel = self.cust_tree.selection()
        if not sel:
            messagebox.showinfo("Chọn dòng", "Hãy chọn một khách hàng để sửa.")
            return
        idx = self.cust_tree.index(sel[0])
        customers = self.cfg_data.setdefault("customers", [])
        if idx < len(customers):
            d = self._customer_dialog("Sửa khách hàng", customers[idx])
            if d:
                customers[idx] = d
                self._fill_customer_tree()
                save_json(CONFIG_FILE, self.cfg_data)

    def del_customer(self):
        sel = self.cust_tree.selection()
        if not sel:
            messagebox.showinfo("Chọn dòng", "Hãy chọn một khách hàng để xoá.")
            return
        if messagebox.askyesno("Xác nhận", "Xoá khách hàng đã chọn?"):
            idx = self.cust_tree.index(sel[0])
            customers = self.cfg_data.get("customers", [])
            if idx < len(customers):
                customers.pop(idx)
                self._fill_customer_tree()
                save_json(CONFIG_FILE, self.cfg_data)

    def save_customers(self):
        """Lưu thay đổi danh sách khách hàng (dùng khi muốn lưu thủ công)."""
        try:
            save_json(CONFIG_FILE, self.cfg_data)
            self.status.set("Đã lưu thông tin khách hàng")
            messagebox.showinfo("Đã lưu", "Thông tin khách hàng đã được lưu vào config.json.")
        except Exception as e:
            messagebox.showerror("Không lưu được", str(e))

    def _customer_dialog(self, title: str, init: dict) -> dict | None:
        dlg = tk.Toplevel(self)
        dlg.title(title)
        dlg.geometry("480x300")
        dlg.transient(self)
        dlg.grab_set()

        form = ttk.Frame(dlg, padding=16)
        form.pack(fill="both", expand=True)

        fields = [
            ("excel_code", "Mã/Tên trên Excel:", init.get("excel_code", "")),
            ("ten_don_vi", "Tên đơn vị (In trên phiếu):", init.get("ten_don_vi", "")),
            ("nguoi_dat", "Người đặt hàng:", init.get("nguoi_dat", "")),
            ("dia_chi", "Địa chỉ giao hàng:", init.get("dia_chi", "")),
            ("dien_giai", "Diễn giải:", init.get("dien_giai", "")),
        ]
        vars_map = {}
        for i, (k, lbl, val) in enumerate(fields):
            ttk.Label(form, text=lbl).grid(row=i, column=0, sticky="w", pady=6, padx=(0, 10))
            v = tk.StringVar(value=val)
            ttk.Entry(form, textvariable=v, width=35).grid(row=i, column=1, sticky="we", pady=6)
            vars_map[k] = v
        form.columnconfigure(1, weight=1)

        res = {}
        def on_ok():
            for k, v in vars_map.items():
                res[k] = v.get().strip()
            dlg.destroy()

        btns = ttk.Frame(dlg, padding=(16, 0, 16, 16))
        btns.pack(fill="x")
        ttk.Button(btns, text="Lưu", style="Accent.TButton", command=on_ok).pack(side="right")
        ttk.Button(btns, text="Huỷ", command=dlg.destroy).pack(side="right", padx=8)

        self.wait_window(dlg)
        return res if res else None

    def import_products_excel(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.csv")])
        if not p: return
        import openpyxl
        try:
            wb = openpyxl.load_workbook(p, data_only=True)
            ws = wb.active
            new_prods = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row[0]: continue
                new_prods.append({"ma": str(row[0]), "ten_ngan": str(row[1] or ""), "ten": str(row[2] or ""), "dvt": str(row[3] or ""), "gia": float(row[4] or 0)})
            self.products_data["products"] = new_prods
            self._fill_price_tree()
            messagebox.showinfo("Thành công", f"Đã nạp {len(new_prods)} mặt hàng.")
        except Exception as e:
            messagebox.showerror("Lỗi", str(e))

    def import_customers_excel(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.csv")])
        if not p: return
        import openpyxl
        try:
            wb = openpyxl.load_workbook(p, data_only=True)
            ws = wb.active
            new_custs = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row[0]: continue
                new_custs.append({"excel_code": str(row[0]), "ten_don_vi": str(row[1] or ""), "nguoi_dat": str(row[2] or ""), "dia_chi": str(row[3] or ""), "dien_giai": str(row[4] or "")})
            self.cfg_data["customers"] = new_custs
            self._fill_customer_tree()
            messagebox.showinfo("Thành công", f"Đã nạp {len(new_custs)} khách hàng.")
        except Exception as e:
            messagebox.showerror("Lỗi", str(e))

    def export_products_excel(self):
        p = filedialog.asksaveasfilename(defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="bang_gia_mau.xlsx", title="Xuất bảng giá ra Excel")
        if not p: return
        import openpyxl
        from openpyxl.styles import Font
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Bang gia"
        ws.append(["Mã hàng", "Tên/Mã trên Excel", "Tên in trên phiếu", "ĐVT", "Đơn giá"])
        for cell in ws[1]:
            cell.font = Font(bold=True)
        for pr in self.products_data["products"]:
            ws.append([pr.get("ma", ""), pr.get("ten_ngan", ""), pr.get("ten", ""),
                       pr.get("dvt", ""), pr.get("gia", 0)])
        for col, wdt in zip("ABCDE", (14, 30, 40, 12, 16)):
            ws.column_dimensions[col].width = wdt
        ws.freeze_panes = "A2"
        wb.save(p)
        self.status.set(f"Đã xuất bảng giá vào {p}")
        messagebox.showinfo("Xuất xong", f"Đã xuất {len(self.products_data['products'])} mặt hàng vào:\n{p}")

    def export_customers_excel(self):
        p = filedialog.asksaveasfilename(defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="khach_hang_mau.xlsx", title="Xuất danh sách khách hàng ra Excel")
        if not p: return
        import openpyxl
        from openpyxl.styles import Font
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Khach hang"
        ws.append(["Mã/Tên trên Excel", "Tên đơn vị", "Người đặt hàng", "Địa chỉ giao hàng", "Diễn giải"])
        for cell in ws[1]:
            cell.font = Font(bold=True)
        for c in self.cfg_data.get("customers", []):
            ws.append([c.get("excel_code", ""), c.get("ten_don_vi", ""), c.get("nguoi_dat", ""),
                       c.get("dia_chi", ""), c.get("dien_giai", "")])
        for col, wdt in zip("ABCDE", (26, 36, 20, 40, 24)):
            ws.column_dimensions[col].width = wdt
        ws.freeze_panes = "A2"
        wb.save(p)
        self.status.set(f"Đã xuất danh sách khách hàng vào {p}")
        messagebox.showinfo("Xuất xong",
                            f"Đã xuất {len(self.cfg_data.get('customers', []))} khách hàng vào:\n{p}")

    # ── tab 4 ────────────────────────────────────────────────────────────
    def _build_company_tab(self):
        f = self.tab_company
        form = ttk.Frame(f)
        form.pack(fill="x")
        self.comp_vars = {}
        fields = [
            ("ten_cong_ty", "Tên công ty", 70),
            ("dia_chi", "Địa chỉ", 70),
            ("tieu_de", "Tiêu đề phiếu", 40),
            ("nguoi_giao", "Người giao hàng (in sẵn)", 40),
        ]
        for i, (key, label, width) in enumerate(fields):
            r, col = i // 2, (i % 2) * 2
            ttk.Label(form, text=label).grid(row=r, column=col, sticky="w", padx=(0, 8), pady=6)
            var = tk.StringVar(value=str(self.cfg_data.get(key, "")))
            ttk.Entry(form, textvariable=var, width=width).grid(
                row=r, column=col + 1, sticky="we", padx=(0, 24), pady=6)
            self.comp_vars[key] = var
        form.columnconfigure(1, weight=1)
        form.columnconfigure(3, weight=1)

        row = ttk.Frame(f)
        row.pack(fill="x", pady=(10, 0))
        ttk.Label(row, text="Logo (PNG/JPG):").pack(side="left")
        self.logo_var = tk.StringVar(value=self.cfg_data.get("logo", ""))
        ttk.Entry(row, textvariable=self.logo_var).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(row, text="Chọn ảnh…", command=self.pick_logo).pack(side="left")
        ttk.Button(row, text="Bỏ logo", command=lambda: self.logo_var.set("")).pack(side="left", padx=(6, 0))

        row2 = ttk.Frame(f)
        row2.pack(fill="x", pady=(10, 0))
        ttk.Label(row2, text="Số liên mỗi tờ").pack(side="left")
        self.so_lien_var = tk.StringVar(value="1")
        ttk.Spinbox(row2, from_=1, to=2, width=5, textvariable=self.so_lien_var).pack(side="left", padx=(8, 6))
        ttk.Label(row2, text="(2 = in 2 liên giống nhau trên 1 tờ A4)", style="Sub.TLabel").pack(side="left")

        row3 = ttk.Frame(f)
        row3.pack(fill="x", pady=(12, 0))
        ttk.Label(row3, text="Ghi chú chân phiếu").pack(side="left")
        ghi_chu = tk.StringVar(value=str(self.cfg_data.get("ghi_chu_footer", "")))
        ttk.Entry(row3, textvariable=ghi_chu).pack(side="left", fill="x", expand=True, padx=(8, 0))
        self.comp_vars["ghi_chu_footer"] = ghi_chu

        ttk.Button(f, text="Lưu thông tin công ty", style="Accent.TButton",
                   command=self.save_company).pack(anchor="w", pady=(16, 0))

    def _on_tab_change(self, _e=None):
        idx = self.nb.index(self.nb.select()) if self.nb.select() else 0
        if idx == 1:
            n = len(self.products_data["products"])
            zero = sum(1 for p in self.products_data["products"] if not p.get("gia"))
            msg = f"{n} mặt hàng trong bảng giá"
            if zero:
                msg += f" · {zero} mặt hàng chưa có đơn giá"
            self.status.set(msg + " · nhấp đúp 1 dòng để sửa")
        elif idx == 2:
            self.status.set("Thông tin in trên đầu phiếu — sửa xong bấm 'Lưu thông tin công ty'")
        else:
            self.status.set(getattr(self, "_order_status", "Sẵn sàng"))

    # ── hành động: chọn file ─────────────────────────────────────────────
    def pick_file(self):
        p = filedialog.askopenfilename(
            title="Chọn file Excel đơn hàng",
            filetypes=[("Excel / CSV", "*.xlsx *.xlsm *.csv"), ("Tất cả", "*.*")],
        )
        if p:
            self.load_excel(p)

    def load_excel(self, path: str):
        self.excel_path = path
        self.lbl_file.configure(text=path, foreground="#111")
        self.reload_orders()

    def reload_orders(self):
        if not self.excel_path:
            messagebox.showinfo("Chưa có file", "Hãy chọn file Excel trước.")
            return
        try:
            self.orders = read_orders(self.excel_path, self.products_data["products"], Cfg(self.cfg_data))
        except Exception as e:
            messagebox.showerror("Không đọc được file",
                                 f"{e}\n\n{''.join(traceback.format_exception_only(type(e), e)).strip()}")
            return

        self._assign_codes()
        self._apply_customers()
        self.tree.delete(*self.tree.get_children())
        for i, o in enumerate(self.orders):
            self.tree.insert("", "end", values=(
                o["customer"], o.get("date") or "", o.get("nguoi_giao") or "",
                o.get("note", ""),
            ), tags=("odd",) if i % 2 else ())
        total = sum(o["total"] for o in self.orders)
        self._order_status = (f"{len(self.orders)} đơn · tổng {money(total)} đ"
                              f" · nguồn: {Path(self.excel_path).name}")
        self.status.set(self._order_status)

        no_price = [i["ten"] for o in self.orders for i in o["items"] if not i["gia"]]
        if no_price:
            uniq = sorted(set(no_price))
            messagebox.showwarning(
                "Có mặt hàng chưa có đơn giá",
                "Các mặt hàng sau chưa có đơn giá (sẽ in 0 đ):\n\n• " + "\n• ".join(uniq[:12])
                + "\n\nVào tab 'Bảng giá mặt hàng' để nhập giá.",
            )

    def _apply_customers(self):
        """Thay thế thông tin đơn theo danh sách khách ở tab 3 (khớp theo Mã/Tên trên Excel)."""
        import unicodedata

        def norm(s):
            return unicodedata.normalize("NFKC", str(s or "")).strip().lower()

        for o in self.orders:
            for c in self.cfg_data.get("customers", []):
                code = norm(c.get("excel_code"))
                if code and code == norm(o.get("customer")):
                    if c.get("ten_don_vi"):
                        o["customer"] = c["ten_don_vi"]
                    if c.get("dia_chi"):
                        o["address"] = c["dia_chi"]
                    if c.get("dien_giai"):
                        o["note"] = (o.get("note", "") + " | " + c["dien_giai"]).strip(" |")
                    break

    def _assign_codes(self):
        """Gán số phiếu: giữ mã từ Excel nếu có, không thì đánh số theo tiền tố từ giao diện."""
        prefix = str(self.prefix_var.get() if hasattr(self, "prefix_var") else self.cfg_data.get("tien_to_so_phieu", "")).strip()
        try:
            n = int(self.start_var.get() if hasattr(self, "start_var") else self.cfg_data.get("so_bat_dau", 1))
        except (TypeError, ValueError):
            n = 1
        used = {str(o.get("code") or "") for o in self.orders}
        if hasattr(self, "delivery_date_var"):
            self._delivery_date = self.delivery_date_var.get().strip()
        else:
            self._delivery_date = str(self.cfg_data.get("delivery_date", "") or "").strip()
        for o in self.orders:
            code = str(o.get("code") or "").strip()
            if not code or code.isdigit():
                while f"{prefix}{n:05d}" in used:
                    n += 1
                o["code"] = f"{prefix}{n:05d}"
                used.add(o["code"])
                n += 1
            if getattr(self, "blank_date", None) and self.blank_date.get():
                o["date"] = ""
            elif self._delivery_date:
                o["date"] = self._delivery_date
            elif self.cfg_data.get("dung_ngay_hom_nay", True) and not o.get("date"):
                import datetime as _dt
                o["date"] = _dt.date.today().strftime("%d/%m/%Y")

    def _on_date_change(self, *_):
        """Nhập/sửa ngày giao → cập nhật ngay vào các đơn đã nạp (không cần nạp lại file)."""
        if not self.orders:
            return
        d = self.delivery_date_var.get().strip()
        import datetime as _dt
        today = _dt.date.today().strftime("%d/%m/%Y")
        for o in self.orders:
            if getattr(self, "blank_date", None) and self.blank_date.get():
                o["date"] = ""
            elif d:
                o["date"] = d
            elif self.cfg_data.get("dung_ngay_hom_nay", True):
                o["date"] = today
        for iid, o in zip(self.tree.get_children(), self.orders):
            self.tree.set(iid, "date", o.get("date") or "")

    def pick_outdir(self):
        p = filedialog.askdirectory(title="Chọn thư mục lưu PDF")
        if p:
            self.out_dir.set(p)

    # ── xuất PDF ─────────────────────────────────────────────────────────
    def export_pdf(self):
        if not self.orders:
            messagebox.showinfo("Chưa có dữ liệu", "Hãy chọn file Excel và kiểm tra danh sách trước.")
            return
        if self.worker and self.worker.is_alive():
            return
        out = self.out_dir.get().strip()
        if not out:
            messagebox.showinfo("Thiếu thư mục", "Hãy chọn thư mục lưu PDF.")
            return
        
        # Cập nhật cfg từ UI trước khi chạy
        self.cfg_data.update(self._collect_cfg())
        cfg = Cfg(self.cfg_data)
        
        page_type = self.page_size.get()
        orders = self.orders
        one_file = self.one_file.get()
        show_price = self.show_price.get()
        
        self.btn_export.configure(state="disabled")
        self.pb.configure(maximum=max(1, len(orders)), value=0)
        self.status.set("Đang xuất PDF…")
        
        def work():
            try:
                def prog(i, n, o):
                    self.log_q.put(("progress", (i, n, o)))
                from core.pdf_generator import build_all
                paths = build_all(orders, out, cfg, show_price=show_price,
                                  one_file=one_file, progress=prog, page_type=page_type)
                self.log_q.put(("done", paths))
            except Exception as e:
                self.log_q.put(("error", f"{e}\n{traceback.format_exc()}"))
                
        self.worker = threading.Thread(target=work, daemon=True)
        self.worker.start()

    def _drain_queue(self):
        try:
            while True:
                kind, payload = self.log_q.get_nowait()
                if kind == "progress":
                    i, n, o = payload
                    self.pb.configure(value=i)
                    self.status.set(f"Đang xuất {i}/{n}: {o.get('customer','')}")
                elif kind == "done":
                    self.btn_export.configure(state="normal")
                    self.pb.configure(value=self.pb["maximum"])
                    paths = payload
                    self._order_status = f"Xong: {len(paths)} file PDF"
                    self.status.set(self._order_status)
                    if messagebox.askyesno("Hoàn tất",
                                           f"Đã tạo {len(paths)} file PDF trong:\n{self.out_dir.get()}\n\nMở thư mục?"):
                        open_folder(self.out_dir.get())
                elif kind == "error":
                    self.btn_export.configure(state="normal")
                    self.status.set("Lỗi khi xuất PDF")
                    messagebox.showerror("Lỗi", payload[:1200])
        except queue.Empty:
            pass
        self.after(120, self._drain_queue)

    # ── bảng giá ─────────────────────────────────────────────────────────
    def _fill_mapping_tree(self):
        self.map_tree.delete(*self.map_tree.get_children())
        mapping = self.cfg_data.get("mapping", {})
        for i, (k, v) in enumerate(mapping.items()):
            self.map_tree.insert("", "end", iid=str(i), values=(k, v),
                                tags=("odd",) if i % 2 else ())

    def add_mapping(self):
        dlg = tk.Toplevel(self)
        dlg.title("Thêm ánh xạ")
        dlg.transient(self)
        dlg.grab_set()
        
        ttk.Label(dlg, text="Mã (Excel):").grid(row=0, column=0, padx=10, pady=5)
        ent1 = ttk.Entry(dlg, width=30)
        ent1.grid(row=0, column=1, padx=10, pady=5)
        
        ttk.Label(dlg, text="Tên (PDF):").grid(row=1, column=0, padx=10, pady=5)
        ent2 = ttk.Entry(dlg, width=30)
        ent2.grid(row=1, column=1, padx=10, pady=5)
        
        def save():
            k, v = ent1.get().strip(), ent2.get().strip()
            if not k: return
            if "mapping" not in self.cfg_data: self.cfg_data["mapping"] = {}
            self.cfg_data["mapping"][k] = v
            self._fill_mapping_tree()
            dlg.destroy()
            
        ttk.Button(dlg, text="Lưu", command=save).grid(row=2, column=0, columnspan=2, pady=10)
        dlg.wait_window()

    def del_mapping(self):
        sel = self.map_tree.selection()
        if not sel: return
        i = int(sel[0])
        items = list(self.cfg_data.get("mapping", {}).items())
        k = items[i][0]
        if messagebox.askyesno("Xoá", f"Xoá ánh xạ '{k}'?"):
            del self.cfg_data["mapping"][k]
            self._fill_mapping_tree()

    def save_mapping(self):
        try:
            save_json(CONFIG_FILE, self.cfg_data)
            self.status.set("Đã lưu ánh xạ")
        except Exception as e:
            messagebox.showerror("Lỗi", str(e))

    def _fill_price_tree(self):
        self.price_tree.delete(*self.price_tree.get_children())
        for i, p in enumerate(self.products_data["products"]):
            self.price_tree.insert("", "end", iid=str(i), values=(
                p.get("ma", ""), p.get("ten_ngan", ""), p.get("ten", ""), p.get("dvt", ""), money(p.get("gia", 0))),
                tags=("odd",) if i % 2 else ())

    def _product_dialog(self, title: str, initial: dict) -> dict | None:
        dlg = tk.Toplevel(self)
        dlg.title(title)
        dlg.transient(self)
        dlg.resizable(False, False)
        dlg.grab_set()
        vars_ = {}
        rows = [("ma", "Mã hàng (nội bộ)"), ("ten_ngan", "Từ khoá trên Excel"),
                ("ten", "Tên in trên phiếu"), ("dvt", "Đơn vị tính"), ("gia", "Đơn giá (đồng)")]
        for i, (k, lab) in enumerate(rows):
            ttk.Label(dlg, text=lab).grid(row=i, column=0, sticky="w", padx=14, pady=7)
            v = tk.StringVar(value=str(initial.get(k, "")))
            ttk.Entry(dlg, textvariable=v, width=42).grid(row=i, column=1, padx=14, pady=7)
            vars_[k] = v
        result: dict = {}

        def ok():
            ma = vars_["ma"].get().strip()
            ten_ngan = vars_["ten_ngan"].get().strip()
            ten = vars_["ten"].get().strip()
            if not ten:
                messagebox.showwarning("Thiếu tên", "Tên mặt hàng không được để trống.", parent=dlg)
                return
            gia_raw = vars_["gia"].get().replace(".", "").replace(",", "").strip()
            try:
                gia = float(gia_raw or 0)
            except ValueError:
                messagebox.showwarning("Sai đơn giá", "Đơn giá phải là số.", parent=dlg)
                return
            result.update({"ma": ma, "ten_ngan": ten_ngan, "ten": ten, "dvt": vars_["dvt"].get().strip(), "gia": gia})
            dlg.destroy()

        bb = ttk.Frame(dlg)
        bb.grid(row=len(rows), column=0, columnspan=2, sticky="e", padx=14, pady=(6, 14))
        ttk.Button(bb, text="Huỷ", command=dlg.destroy).pack(side="right")
        ttk.Button(bb, text="Đồng ý", style="Accent.TButton", command=ok).pack(side="right", padx=6)
        dlg.wait_window()
        return result or None

    def add_product(self):
        d = self._product_dialog("Thêm mặt hàng", {"dvt": "Hộp", "gia": 0})
        if d:
            self.products_data["products"].append(d)
            self._fill_price_tree()
            self.save_products(silent=True)

    def edit_product(self):
        sel = self.price_tree.selection()
        if not sel:
            messagebox.showinfo("Chưa chọn", "Chọn 1 dòng trong bảng giá để sửa.")
            return
        i = int(sel[0])
        d = self._product_dialog("Sửa mặt hàng", self.products_data["products"][i])
        if d:
            self.products_data["products"][i] = d
            self._fill_price_tree()
            self.save_products(silent=True)
            if self.excel_path:
                self.reload_orders()

    def del_product(self):
        sel = self.price_tree.selection()
        if not sel:
            return
        i = int(sel[0])
        name = self.products_data["products"][i].get("ten", "")
        if messagebox.askyesno("Xoá mặt hàng", f"Xoá '{name}' khỏi bảng giá?"):
            self.products_data["products"].pop(i)
            self._fill_price_tree()
            self.save_products(silent=True)

    def save_products(self, silent: bool = False):
        try:
            save_json(PRODUCTS_FILE, self.products_data)
            self.status.set(f"Đã lưu bảng giá: {len(self.products_data['products'])} mặt hàng")
            if self.excel_path and not silent:
                self.reload_orders()
        except Exception as e:
            messagebox.showerror("Không lưu được", str(e))

    # ── thông tin công ty ────────────────────────────────────────────────
    def pick_logo(self):
        p = filedialog.askopenfilename(title="Chọn logo",
                                       filetypes=[("Ảnh", "*.png *.jpg *.jpeg"), ("Tất cả", "*.*")])
        if p:
            self.logo_var.set(p)

    def _collect_cfg(self) -> dict:
        d = dict(self.cfg_data)
        for k, var in self.comp_vars.items():
            d[k] = var.get().strip()
        d["logo"] = self.logo_var.get().strip()
        try:
            d["so_lien"] = max(1, min(3, int(self.so_lien_var.get())))
        except (TypeError, ValueError):
            d["so_lien"] = 2
        d["hien_don_gia"] = self.show_price.get()
        if hasattr(self, "prefix_var"):
            d["tien_to_so_phieu"] = self.prefix_var.get().strip()
        if hasattr(self, "start_var"):
            try:
                d["so_bat_dau"] = int(self.start_var.get().strip() or 1)
            except (TypeError, ValueError):
                d["so_bat_dau"] = 1
        d["delivery_date"] = self.delivery_date_var.get().strip()
        return d

    def save_company(self):
        self.cfg_data = self._collect_cfg()
        try:
            save_json(CONFIG_FILE, self.cfg_data)
        except Exception as e:
            messagebox.showerror("Không lưu được", str(e))
            return
        self.status.set("Đã lưu thông tin công ty")
        messagebox.showinfo("Đã lưu", "Thông tin công ty đã được lưu vào config.json.")

    def _on_close(self):
        if self.worker and self.worker.is_alive():
            if not messagebox.askyesno("Đang xuất PDF", "Đang xuất PDF, thoát vẫn?"):
                return
        self.destroy()


# ─────────────────────────── CLI ─────────────────────────────────────────
def run_cli(argv: list[str]) -> int:
    import argparse

    ap = argparse.ArgumentParser(description="Excel → PDF phiếu giao hàng")
    ap.add_argument("input", help="file Excel (.xlsx/.csv)")
    ap.add_argument("--out", default="output", help="thư mục lưu PDF")
    ap.add_argument("--one-file", action="store_true", help="gộp tất cả vào 1 PDF")
    ap.add_argument("--no-price", action="store_true", help="không in đơn giá")
    ap.add_argument("--page-type", default="A5_DOC", help="A4, A5, hoặc A5_NGANG")
    args = ap.parse_args(argv)

    cfg_data = load_json(CONFIG_FILE, DEFAULT_CONFIG)
    cfg = Cfg(cfg_data)
    products = load_json(PRODUCTS_FILE, DEFAULT_PRODUCTS).get("products", [])
    orders = read_orders(args.input, products, cfg)

    prefix = str(cfg_data.get("tien_to_so_phieu", "") or "")
    try:
        n = int(cfg_data.get("so_bat_dau", 1))
    except (TypeError, ValueError):
        n = 1
    used = set()
    for o in orders:
        if not str(o.get("code") or "").strip():
            while f"{prefix}{n:05d}" in used:
                n += 1
            o["code"] = f"{prefix}{n:05d}"
            used.add(o["code"])
            n += 1
        if not o.get("date"):
            import datetime as _dt
            o["date"] = _dt.date.today().strftime("%d/%m/%Y")

    paths = build_all(orders, args.out, cfg, show_price=not args.no_price, one_file=args.one_file, page_type=args.page_type)
    print(f"Đã tạo {len(paths)} file cho {len(orders)} đơn trong '{args.out}'")
    for p in paths:
        print("  -", p)
    return 0


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--cli":
        raise SystemExit(run_cli(sys.argv[2:]))
    App().mainloop()
