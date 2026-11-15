# Order2PDF — Excel → Phiếu giao hàng PDF

Chuyển bảng đơn hàng trong Excel thành **phiếu giao hàng PDF**, mỗi tờ A4 in **2 liên giống nhau** có đường cắt ở giữa, đơn giá lấy từ **bảng giá riêng sửa được trong app**.

Đúng mẫu phiếu của CÔNG TY TNHH THƯƠNG MẠI VÀ DỊCH VỤ ĐẦU TƯ BẢO HÙNG: header công ty, số phiếu, ngày, người mua, bảng STT/Tên hàng/ĐVT/SL/Đơn giá/Thành tiền, tổng cộng, số tiền bằng chữ tiếng Việt, 2 chữ ký.

---

## 1. Dùng ngay (chưa cần đóng gói)

```bash
python -m venv .venv
# Windows:  .venv\Scripts\activate
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Ba tab:

1. **Xuất phiếu** — chọn file Excel → kiểm tra danh sách → `XUẤT PDF`.
2. **Bảng giá mặt hàng** — sửa đơn giá từng mặt hàng (nhấp đúp 1 dòng), thêm/xoá. Lưu vào `products.json`.
3. **Thông tin công ty** — tên, địa chỉ, MST, điện thoại, logo, tiêu đề phiếu, tiền tố số phiếu, số liên mỗi tờ. Lưu vào `config.json`.

Chạy không cần giao diện:

```bash
python app.py --cli duong/dan/file.xlsx --out output
python app.py --cli file.xlsx --out output --one-file   # gộp tất cả vào 1 PDF
```

## 2. Đóng gói thành .exe cho Windows

Trên máy Windows có Python 3.11+:

```bat
build.bat
```

Kết quả: `dist\Order2PDF.exe` — file đơn, copy đi đâu cũng chạy, không cần cài Python.
Lần chạy đầu, app tự tạo `products.json` và `config.json` nằm cạnh file .exe để sửa.

> `.exe` phải build **trên Windows** (PyInstaller không build chéo từ Linux/Mac được).

## 3. File Excel đầu vào

App tự nhận diện tiêu đề cột tiếng Việt (có dấu hay không dấu đều được) và hỗ trợ **2 kiểu bảng**:

**A. Bảng ngang** — mỗi dòng 1 khách, mỗi cột là 1 mặt hàng (giống sổ theo dõi giao hàng):

| Khách hàng | SC VNM | PROBI 6.5ML | ADM | Ghi chú |
|---|---|---|---|---|
| MN HÀ HỘI | 720 | 10 | | 1q HĐ |
| HKD VĂN AN | 190 | 12 | 99 | trưa 11h |

**B. Bảng dài** — mỗi dòng 1 mặt hàng, gộp theo mã đơn:

| Số phiếu | Khách hàng | Tên hàng | ĐVT | Số lượng |
|---|---|---|---|---|
| BH900 | MN LÊ LỢI | SC VNM | Hộp | 100 |
| BH900 | MN LÊ LỢI | Probi 6.5ml | Chai | 20 |

Nhận diện được: `Khách hàng / HKD / Nhà trường / Đơn vị mua`, `Số phiếu / Mã đơn / Mã HĐ`, `Sản phẩm / Tên hàng / Mặt hàng`, `Số lượng / SL / Số lượng giao`, `Đơn giá / Giá`, `Thành tiền / Thành tiền (đ)`, `Ghi chú / NV / Ngày / Địa chỉ / Điện thoại`.

Định dạng đọc được: `.xlsx`, `.xlsm`, `.csv` (UTF-8). **Không đọc `.xls` cũ** — mở bằng Excel rồi *Save As → .xlsx*.

## 4. Bảng giá — phần quan trọng nhất

`products.json`: mỗi mặt hàng gồm `ma` (khớp tiêu đề cột trong Excel), `ten` (in trên phiếu), `dvt`, `gia`.

```json
{ "ma": "SC", "ten": "Sữa chua ăn VNM có đường 60g", "dvt": "Hộp", "gia": 3942 }
```

Cách khớp cột: mã hàng xuất hiện trong tiêu đề cột là khớp mạnh nhất. Ví dụ cột `SC VNM` → mặt hàng mã `SC`; cột `Ba vị 100g` → khớp theo tên `Sữa chua Ba Vị 100g`.
Cột nào không khớp mặt hàng nào sẽ bị bỏ qua — app cảnh báo khi có mặt hàng chưa có đơn giá (in 0 đ).

Thành tiền tự tính = số lượng × đơn giá; tổng cộng và **số tiền bằng chữ** tự sinh.

## 5. Cấu trúc mã nguồn

```
app.py                  giao diện tkinter + chế độ dòng lệnh
core/text_utils.py      bỏ dấu tiếng Việt, đọc số/tiền/ngày, đọc số tiền thành chữ
core/excel_reader.py    đọc Excel/CSV, nhận diện cột, gộp đơn
core/pdf_generator.py   vẽ phiếu A4: 2 liên + đường cắt, đo chiều cao thật
products.json           bảng giá (sửa trong app)
config.json             thông tin công ty (sửa trong app)
fonts/                  DejaVuSans + Bold (bắt buộc để in đúng dấu tiếng Việt)
tools/make_sample.py    tạo Excel mẫu
tools/make_icon.py      tạo assets/icon.ico
```

## 6. Đã kiểm tra

- 6 đơn mẫu → 6 PDF, mỗi file 1 trang 2 liên, nội dung không chạm đường cắt (liên trên kết thúc ~27pt trước đường cắt, liên dưới còn 14pt lề đáy).
- Chế độ gộp: 6 đơn → 1 PDF 6 trang.
- Số tiền bằng chữ đúng: 2.838.240 → "Hai triệu tám trăm ba mươi tám nghìn hai trăm bốn mươi đồng".
- Đọc được cả bảng ngang (xlsx/csv) và bảng dài (gộp đơn theo mã).

## 7. Còn phải làm

- Nhập **đơn giá thật** cho 7 mặt hàng trong tab *Bảng giá* (hiện chỉ SC = 3.942 đ lấy từ phiếu mẫu, các mã khác là giá tạm).
- Đối chiếu với 1 file Excel thật của công ty để chắc chắn tên cột khớp.
