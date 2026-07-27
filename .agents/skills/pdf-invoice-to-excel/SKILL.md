---
name: pdf-invoice-to-excel
description: Tự động đọc và trích xuất thông tin hóa đơn PDF (kể cả PDF dạng hình ảnh/scan bằng OCR) và ghi dữ liệu vào file Excel có sẵn. Tự động kiểm tra trùng lặp để bỏ qua hóa đơn đã có, căn giữa số liệu, định dạng phân chia hàng nghìn cho các cột tiền và tự động mở rộng vùng Excel Table (ListObject). Hãy sử dụng skill này bất cứ khi nào người dùng muốn đọc/nhập hóa đơn PDF vào Excel, tự động hóa lưu hóa đơn, hoặc tạo script quét hóa đơn PDF.
---

# Skill: Trích Xuất Hóa Đơn PDF và Lưu vào File Excel (PDF Invoice to Excel)

Skill này cung cấp quy trình và mã mẫu để tự động đọc thông tin từ các tệp hóa đơn PDF (dạng văn bản hoặc ảnh scan/OCR) trong một thư mục chỉ định và ghi dữ liệu tương ứng vào các cột trong file Excel có sẵn.

---

## 🚀 Các Tính Năng Chính
1. **Đọc PDF thông minh (Hỗ trợ OCR)**: Sử dụng `PyMuPDF (fitz)` kết hợp `rapidocr_onnxruntime` để đọc được cả hóa đơn điện tử dạng text và hóa đơn chụp/scan dạng hình ảnh.
2. **Tự động lọc trùng lặp**: Đọc danh sách số hóa đơn đã có trong cột khóa chính (thường là Cột A - `Số hóa đơn tt`), nếu hóa đơn PDF nào đã có trong Excel thì tự động bỏ qua.
3. **Chuẩn hóa Định dạng Cell (Căn giữa & Phân chia hàng nghìn)**:
   - Căn giữa toàn bộ dòng dữ liệu (`Alignment(horizontal='center', vertical='center')`).
   - Định dạng phân chia hàng nghìn cho các cột Tiền, Thuế, Tổng tiền (`#,##0`).
   - Định dạng 3 chữ số thập phân cho Cột Số lượng (`#,##0.000`).
4. **Mở rộng vùng Excel Table tự động**: Tự động nhận diện nếu dữ liệu Excel nằm trong đối tượng Bảng chuẩn (**Excel Formatted Table - `ListObject`**) và tự động cập nhật phạm vi bảng (`tbl.ref`) để tất cả dòng mới đều nằm trọn trong Bảng.
5. **Tích hợp File chạy nhanh 1-Click (`run.bat`)**: Tạo script batch để người dùng chỉ cần nhấp đôi chuột là tự động chạy quy trình quét hóa đơn.

---

## 📋 Yêu Cầu Thư Viện Python
Skill này sử dụng các thư viện Python sau:
```bash
pip install pymupdf rapidocr-onnxruntime openpyxl
```

---

## 🛠️ Quy Trình Thực Hiện Từng Bước

### Bước 1: Trích Xuất Dữ Liệu Hóa Đơn PDF (PyMuPDF + RapidOCR)
1. Mở file PDF và chuyển trang hóa đơn thành hình ảnh PNG (độ phân giải 200 DPI).
2. Chạy model RapidOCR để trích xuất văn bản từ hình ảnh:
   - **Số hóa đơn**: Tìm theo regex `r'(?:s[ốseó]|Số)[:\s]*(\d{5,8})'` hoặc lấy từ tên file.
   - **Ngày hóa đơn**: Tìm ngày theo regex `r'(\d{1,2})/(\d{1,2})/(\d{4})'` hoặc `r'Ngay (\d+) thang (\d+) nam (\d+)'`. Chuyển về định dạng chuẩn `dd/mm/yyyy`.
   - **Số lượng (Lít/Số mặt hàng)**: Đọc từ dòng mặt hàng hoặc số lượng.
   - **Tiền trước thuế, Tiền thuế GTGT, Tổng tiền thanh toán**: Đọc các mốc tiền tương ứng.

### Bước 2: Đọc File Excel & Lọc Trùng
1. Mở workbook Excel bằng `openpyxl.load_workbook(EXCEL_PATH)`.
2. Đọc toàn bộ giá trị cột `Số hóa đơn tt` (Cột A, bắt đầu từ dòng 2) đưa vào tập hợp `existing_invoices`.
3. Khi duyệt qua danh sách các file PDF trong thư mục chỉ định:
   - Nếu `inv_num in existing_invoices`: Bỏ qua (không ghi trùng).
   - Nếu chưa có: Thêm vào danh sách `new_invoices`.

### Bước 3: Ghi Dữ Liệu, Định Dạng Cell & Mở Rộng Excel Table (`ListObject`)
1. Xác định dòng trống tiếp theo `next_row`.
2. Ghi từng trường dữ liệu vào các cột tương ứng:
   - Cột A (1): Số hóa đơn tt
   - Cột B (2): Ngày (`dd/mm/yyyy`)
   - Cột C (3): Số lượng
   - Cột D (4): Tiền trước thuế
   - Cột E (5): Thuế GTGT
   - Cột F (6): Tổng tiền thanh toán
3. **Định dạng căn giữa & phân chia hàng nghìn**:
   ```python
   from openpyxl.styles import Alignment

   CENTER_ALIGN = Alignment(horizontal='center', vertical='center')
   FMT_CURRENCY = '_-* #,##0_-;\\-* #,##0_-;_-* "-"_-;_-@_-'
   FMT_QTY = '#,##0.000'

   def format_cell_row(sheet, row_idx):
       for c in range(1, 7):
           sheet.cell(row=row_idx, column=c).alignment = CENTER_ALIGN
       sheet.cell(row=row_idx, column=3).number_format = FMT_QTY
       sheet.cell(row=row_idx, column=4).number_format = FMT_CURRENCY
       sheet.cell(row=row_idx, column=5).number_format = FMT_CURRENCY
       sheet.cell(row=row_idx, column=6).number_format = FMT_CURRENCY
   ```
4. **Mở rộng Excel Table**:
   ```python
   def update_table_ranges(sheet):
       last_row = sheet.max_row
       for tbl_name in list(sheet.tables):
           tbl = sheet.tables[tbl_name]
           if hasattr(tbl, 'ref') and ":" in tbl.ref:
               start_cell, end_cell = tbl.ref.split(":")
               col_letter = re.match(r'([A-Za-z]+)(\d+)', end_cell).group(1)
               tbl.ref = f"{start_cell}:{col_letter}{last_row}"
   ```
5. Lưu workbook bằng `wb.save(EXCEL_PATH)`.

---

## 📁 Mã Nguồn Mẫu Sản Phẩm (`scripts/process_invoices.py`)
Mã nguồn tham khảo hoàn chỉnh nằm trong thư mục `scripts/process_invoices.py` của Skill này.
