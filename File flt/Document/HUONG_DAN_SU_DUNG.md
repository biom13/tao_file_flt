# Hướng dẫn sử dụng công cụ tạo file `.flt`

Công cụ này giúp tự động đọc toàn bộ file Excel (`.xlsx`, `.xlsm`) trong thư mục `D:\AntiGravity skill\File flt`, gom các mã nhân viên (`IT Code` / `IT. No`) theo từng cổng/máy được phép quẹt vân tay và xuất ra các file `.flt` đúng chuẩn định dạng để import vào hệ thống.

---

## 1. Cấu trúc thư mục

| Thành phần | Vai trò |
| --- | --- |
| `Tao_File_FLT.bat` | File chạy chính — chỉ cần nhấp đúp để thực thi |
| `generate_flt.py` | Script Python xử lý dữ liệu Excel và sinh file `.flt` |
| `result\` | Thư mục chứa toàn bộ file `.flt` được tạo ra |
| `cong K_A.flt` | File `.flt` mẫu dùng để đối chiếu định dạng |
| `HUONG_DAN_SU_DUNG.md` | Tài liệu hướng dẫn này |

---

## 2. Cách sử dụng

1. Chép các file Excel cần xử lý vào thư mục `D:\AntiGravity skill\File flt`.
2. Nhấp đúp vào file `Tao_File_FLT.bat`.
3. Đợi cửa sổ dòng lệnh báo `Hoan tat! Da tao ... file .flt.` rồi nhấn phím bất kỳ để đóng.
4. Mở thư mục `result\` để lấy các file `.flt` vừa tạo và import vào hệ thống.

> [!NOTE]
> Nếu trong `result\` đã có file `.flt` trùng tên từ lần chạy trước, công cụ sẽ tự động ghi đè bằng dữ liệu mới nhất. Nếu muốn làm sạch hoàn toàn trước khi chạy đợt mới, bạn có thể xóa các file cũ trong `result\`.

---

## 3. Quy tắc đọc dữ liệu từ Excel

- **Phạm vi file**: Quét tất cả file `.xlsx` và `.xlsm` nằm cùng thư mục với `generate_flt.py` (tự động bỏ qua file tạm của Excel có dạng `~$...`).
- **Phạm vi sheet**: Chỉ đọc các **sheet đang hiển thị (`visible`)** (không phân biệt tên sheet — ví dụ `New`, `New DW`, `Old DW`, `transfer`, `Transfer & rehire`... đều được đọc tự động). Các sheet đang bị ẩn (`hidden` / `veryHidden`) sẽ được bỏ qua để tránh lấy nhầm dữ liệu cũ hoặc dữ liệu nháp.
- **Nhận diện cột tự động** (không phân biệt chữ hoa/thường hay khoảng trắng thừa):
  - **Cột mã nhân viên**: `IT Code`, `IT. No`, `IT No`, `IT_No`, `IT_Code`, `ITNo`, `ITCode` (ưu tiên trước; nếu sheet không có các cột trên thì mới lấy cột `No.` / `No`).
  - **Cột cổng/máy quẹt vân tay**: `Cổng được phép quẹt vân tay`, `Máy được phép quẹt bằng vân tay`, `Máy được phép quẹt vân tay`, `Machine No.`, `Machine No`, `Machine_No`.
- **Lọc dữ liệu rác**: Tự động bỏ qua các dòng trống, dòng thiếu mã nhân viên hoặc thiếu tên cổng/máy, và các dòng tiêu đề phụ kiểu `Column1`, `Column11`...

---

## 4. Quy tắc gom nhóm và đặt tên file `.flt`

- **Gom nhóm không phân biệt thứ tự & hoa/thường**:
  - Các giá trị cổng/máy phân tách bằng dấu phẩy `,` sẽ được tách ra, chuẩn hóa chữ hoa/thường và sắp xếp lại để so sánh.
  - Ví dụ: `Cổng A, K` và `Cổng K, A` (hoặc `Cổng F1, nhà PK hoa chậu` và `Cổng F1, Nhà PK hoa chậu`) được xem là **cùng 1 nhóm** và gộp chung vào 1 file `.flt`.
  - Các mã `IT Code` trùng lặp trong cùng 1 nhóm cổng sẽ tự động được loại bỏ (chỉ giữ 1 lần).
- **Chuẩn hóa tên file đầu ra**:
  - Tên file `.flt` được đặt theo tên cổng/máy, **bỏ dấu tiếng Việt** và **loại bỏ ký tự đặc biệt**.
  - Dấu phẩy `,` (và khoảng trắng quanh dấu phẩy) được chuyển thành dấu gạch dưới `_`.
  - Ví dụ:
    - `Máy 73, 79, 120, 163` → `May 73_79_120_163.flt`
    - `Cổng F1, nhà PK hoa chậu` → `Cong F1_nha PK hoa chau.flt`
    - `Cổng A, K` → `Cong A_K.flt`

---

## 5. Định dạng file `.flt` đầu ra

File `.flt` được ghi theo định dạng nhị phân **DevExpress `TcxFilterCriteria`** khớp chính xác với file mẫu `cong K_A.flt`:
- Header: `FF FF FF FF 01` + số lượng `IT Code` (4 byte little-endian).
- Mỗi bản ghi `IT Code`: gồm độ dài + chuỗi `IT Code` theo đúng cấu trúc của `TcxFilterCriteriaItem` (giới hạn tối đa 255 byte cho mỗi `IT Code`).

---

## 6. Yêu cầu môi trường & xử lý lỗi thường gặp

- **Yêu cầu**: Máy cần cài sẵn **Python 3** và thư viện **`openpyxl`**.
  - Kiểm tra nhanh bằng PowerShell/CMD: `python -c "import openpyxl; print('OK')"`
  - Nếu thiếu `openpyxl`, cài bằng lệnh: `pip install openpyxl`
- **Lỗi `PermissionError` khi đang mở file Excel**:
  - Đóng các file Excel đang mở trong thư mục trước khi chạy `Tao_File_FLT.bat`.
- **Không thấy dữ liệu từ một sheet nào đó**:
  - Kiểm tra xem sheet đó có đang bị ẩn (Hidden) hay không, và dòng đầu tiên (Row 1) của sheet có đúng tên cột nằm trong danh sách ở Mục 3 hay không.
