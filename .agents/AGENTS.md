# QUY TẮC PHÁT TRIỂN DỰ ÁN & ĐẢM BẢO TÍNH KHẢ DI CHUYỂN (PORTABILITY RULES)

## 1. Tính Khả Di Chuyển Giữa Các Máy Tính (Cross-Machine Portability)
Mọi chương trình, kịch bản (script), hoặc ứng dụng web khi phát triển phải đảm bảo **chạy 100% không lỗi** khi được copy/chuyển sang bất kỳ máy tính hoặc ổ đĩa nào khác (`C:`, `D:`, `E:`, USB, v.v.).

### A. Xử lý đường dẫn động (Dynamic Path Handling):
- **TUYỆT ĐỐI KHÔNG** dùng đường dẫn tuyệt đối cứng (hardcoded path) như `d:\AI AGI...`, `C:\Users\...`.
- **LUÔN LUÔN** dùng đường dẫn động tương đối xác định từ vị trí chứa script:
  - In Python: `BASE_DIR = os.path.dirname(os.path.abspath(__file__))`
  - In Batch (`run.bat`): Sử dụng `%~dp0` để tham chiếu thư mục gốc của file `.bat` (ví dụ: `python "%~dp0process_invoices.py"`).
  - In Node.js: Sử dụng `path.join(__dirname, ...)` hoặc `import.meta.url`.

### B. Tự động kiểm tra & Cài đặt môi trường trong `run.bat`:
- Mọi chương trình khởi chạy bằng `run.bat` phải tích hợp quy trình kiểm tra tự động:
  1. **Kiểm tra Runtime (Python/Node.js)**: Kiểm tra xem máy tính đã cài đặt Python/Node.js và thêm vào PATH chưa. Nếu chưa có, đưa ra hướng dẫn rõ ràng.
  2. **Tạo file phụ thuộc (`requirements.txt` / `package.json`)**: Luôn duy trì file danh sách phụ thuộc đi kèm dự án.
  3. **Tự động cài đặt thư viện thiếu**: Trước khi chạy code chính, `run.bat` tự kiểm tra thư viện thiếu và chạy `python -m pip install -r "%~dp0requirements.txt"` để tự khắc phục trên máy mới.

### C. Ứng dụng Web App / Localhost:
- Các ứng dụng web app chạy trên localhost khi chép qua máy tính khác phải tự khởi động dịch vụ, mở port và có thể thực thi ngay qua `run.bat`.

### D. Tối ưu hóa & Phòng ngừa lỗi runtime:
- **Kiểm tra tồn tại**: Luôn dùng `os.path.exists()` kiểm tra file/thư mục đầu vào trước khi đọc/ghi.
- **Xử lý xung đột file**: Bắt lỗi `PermissionError` (khi file Excel đang mở trên máy người dùng) và đưa ra thông báo nhắc đóng file thay vì làm crash chương trình.
- **Chuẩn hóa UTF-8**: Thêm `chcp 65001 > nul` trong Batch Script và `sys.stdout.reconfigure(encoding='utf-8')` trong Python để hiển thị tiếng Việt chính xác trên mọi hệ điều hành.
