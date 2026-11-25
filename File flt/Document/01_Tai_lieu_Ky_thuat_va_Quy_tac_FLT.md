# Tài liệu Kỹ thuật, Định dạng & Quy tắc chuẩn hóa công cụ tạo file `.flt`

**Thư mục chính:** `D:\AI_AntiGravity skill\File flt`  
**Ngày cập nhật:** 25/11/2026

---

## 1. Kiến trúc & Thành phần thư mục

| Thành phần | Vai trò |
| --- | --- |
| `Tao_File_FLT.bat` | File command chạy chính (`chcp 65001`, gọi `python -X utf8 "%~dp0generate_flt.py"` và `pause`). |
| `generate_flt.py` | Script Python đọc dữ liệu Excel, gom nhóm, chuẩn hóa tên file và xuất file nhị phân `.flt`. |
| `result\` | Thư mục chứa toàn bộ file `.flt` đầu ra (tự động tạo nếu chưa có). |
| `cong K_A.flt` | File `.flt` mẫu dùng để đối chiếu cấu trúc nhị phân. |

**Yêu cầu môi trường:** Python 3 và thư viện `openpyxl` (`pip install openpyxl`).

---

## 2. Quy tắc đọc & chuẩn hóa dữ liệu Excel

### 2.1. Phạm vi file và sheet
- **File đầu vào:** Quét tất cả file `.xlsx` và `.xlsm` nằm cùng thư mục với `generate_flt.py`; tự động bỏ qua file tạm Excel có tiền tố `~$`.
- **Sheet được đọc:** Chỉ đọc các sheet đang hiển thị (`ws.sheet_state == "visible"`), không phân biệt tên sheet (ví dụ `New`, `New DW`, `Old DW`, `transfer`, `Transfer & rehire`...). Bỏ qua toàn bộ sheet ẩn (`hidden` / `veryHidden`).
- **Dòng tiêu đề:** Đọc tiêu đề tại dòng 1 (`Row 1`), chuẩn hóa bằng cách gộp khoảng trắng thừa và chuyển về chữ thường (`casefold`).

### 2.2. Nhận diện cột tự động
- **Cột mã nhân viên (`IT_COL_NAMES` — ưu tiên tìm trước):**
  - `IT Code`, `IT. No`, `IT No`, `IT_No`, `IT_Code`, `ITNo`, `ITCode`.
- **Cột dự phòng (`IT_FALLBACK_COL_NAMES`):**
  - `No.`, `No` — chỉ dùng khi sheet **không** có cột nào thuộc nhóm `IT_COL_NAMES` (tránh lấy nhầm cột số thứ tự `No.` trên các sheet đã có sẵn cột `IT Code`).
- **Cột cổng/máy (`GATE_COL_NAMES`):**
  - `Cổng được phép quẹt vân tay`, `Máy được phép quẹt bằng vân tay`, `Máy được phép quẹt vân tay`, `Machine No.`, `Machine No`, `Machine_No`.

### 2.3. Chuẩn hóa giá trị & lọc dữ liệu rác
- Bỏ qua dòng thiếu mã nhân viên hoặc thiếu tên cổng/máy.
- Bỏ qua các dòng placeholder dạng `Column1`, `Column11`... (từ PowerQuery/bảng mẫu).
- Nếu mã nhân viên được Excel lưu dạng số thực nguyên (ví dụ `59291.0`), tự động chuyển về số nguyên dạng chuỗi (`"59291"`).

---

## 3. Quy tắc gom nhóm & đặt tên file `.flt`

### 3.1. Gom nhóm theo tập cổng/máy (`normalize_gate_key`)
- Chuẩn hóa chữ hoa/thường và khoảng trắng thừa.
- Tách tiền tố đầu dòng (`cổng` hoặc `máy`, nếu có), sau đó tách danh sách theo dấu phẩy `,`, loại bỏ tiền tố lặp ở từng phần tử con và sắp xếp theo thứ tự bảng chữ cái.
- **Hệ quả:**
  - Không phân biệt thứ tự: `Cổng A, K` và `Cổng K, A` được gom vào **1 file duy nhất**.
  - Không phân biệt hoa/thường: `Cổng F1, nhà PK hoa chậu` và `Cổng F1, Nhà PK hoa chậu` được gom vào **1 file duy nhất**.
- **Loại trùng:** Trong cùng 1 nhóm cổng/máy, các `IT Code` trùng lặp chỉ được giữ lại 1 lần (giữ nguyên thứ tự xuất hiện đầu tiên).

### 3.2. Chuẩn hóa tên file đầu ra (`sanitize_filename`)
1. Bỏ dấu tiếng Việt (`đ` → `d`, `Đ` → `D`, loại bỏ toàn bộ dấu thanh/mũ bằng Unicode NFD).
2. Chuyển dấu phẩy và khoảng trắng xung quanh dấu phẩy (`\s*,\s*`) thành dấu gạch dưới `_`.
3. Thay thế các ký tự đặc biệt còn lại (ngoài `A-Z`, `a-z`, `0-9`, khoảng trắng, `_`, `-`) bằng `_`, gộp `_` liên tiếp và cắt khoảng trắng/dấu chấm ở hai đầu.
4. Ví dụ thực tế:
   - `Máy 73, 79, 120, 163` → `May 73_79_120_163.flt`
   - `Cổng F1, nhà PK hoa chậu` → `Cong F1_nha PK hoa chau.flt`
   - `4,5,6,7,8` → `4_5_6_7_8.flt`

---

## 4. Định dạng nhị phân file `.flt` (`build_flt_bytes`)

File `.flt` tuân thủ định dạng nhị phân **DevExpress `TcxFilterCriteria`** (khớp byte-for-byte với file mẫu `cong K_A.flt`):

1. **Header (9 byte):**
   - `FF FF FF FF 01` (5 byte cố định)
   - `Count` (4 byte unsigned int, little-endian `<I`): tổng số `IT Code` trong file.
2. **Mỗi bản ghi `IT Code` (mã hóa UTF-8, độ dài `n <= 255` byte):**
   - `00 06` (2 byte cố định)
   - `n` (4 byte unsigned int, little-endian `<I`)
   - Chuỗi `IT Code` (`n` byte)
   - `01 00 00 00 06` (5 byte cố định)
   - `n` (1 byte unsigned char `B`)
   - Chuỗi `IT Code` (`n` byte)

---

## 5. Vận hành & Xử lý lỗi thường gặp

- **Chạy hàng ngày:** Đặt file `.xlsx`/`.xlsm` vào `D:\AI_AntiGravity skill\File flt` → đóng Excel đang mở → nhấp đúp `Tao_File_FLT.bat` → lấy file trong `result\`.
- **Lỗi `PermissionError`:** File Excel đang được mở bởi Excel (đóng file trước khi chạy).
- **Báo `"Khong co du lieu hop le"`:** Kiểm tra file có đúng đuôi `.xlsx`/`.xlsm` và nằm đúng thư mục chứa `Tao_File_FLT.bat` hay không; kiểm tra sheet có đang bị ẩn hoặc tiêu đề cột có nằm ở dòng 1 không.
