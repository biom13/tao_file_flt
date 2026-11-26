# Hướng Dẫn Sử Dụng Công Cụ Tạo File Import Mẫu 7009 (`Tao_File_Import.bat`)

## 1. Mục đích
Công cụ `Tao_File_Import.bat` tự động chuyển đổi hàng loạt các tệp Excel nguồn (`.xlsx`) thành các tệp import nhân viên theo đúng **mẫu 7009 (`7009_Employee_Template.xlsx`)**, lưu kết quả vào thư mục con `result\` và giữ nguyên tuyệt đối các tệp nguồn.

---

## 2. Cấu trúc thư mục công cụ
```text
D:\AI_AntiGravity skill\tao-file-import\
├── Tao_File_Import.bat                     # File chạy chính (double-click để chạy)
├── resources\
│   └── 7009_Employee_Template.xlsx         # File mẫu 7009 chuẩn (SHA-256 cố định)
├── scripts\
│   └── tao_file_import.py                  # Script Python xử lý nghiệp vụ & kiểm tra
├── Documents\                              # Tài liệu hướng dẫn & tổng hợp yêu cầu
└── result\                                 # Thư mục con tự động sinh ra chứa file kết quả
```

---

## 3. Cách sử dụng hàng ngày

### Cách 1: Chạy ngay trong thư mục gốc
1. Chép các tệp Excel nguồn (`.xlsx`) vào `D:\AI_AntiGravity skill\tao-file-import\`.
2. Nhấp đúp chuột (double-click) vào **`Tao_File_Import.bat`**.
3. Chờ màn hình dòng lệnh chạy xong và xem báo cáo từng file:
   - `✅ Đã tạo: result\<tên_file>.xlsx (N nhân viên)` — Thành công.
   - `❌ LỖI: ...` — File nguồn có lỗi dữ liệu hoặc thiếu cột; hệ thống chỉ rõ tên sheet, số dòng, nguyên nhân và vẫn tiếp tục xử lý các file còn lại.
4. Mở thư mục con **`result\`** để lấy các tệp import đã hoàn chỉnh.

### Cách 2: Chép riêng `Tao_File_Import.bat` sang thư mục làm việc khác
- Bạn có thể chép **duy nhất file `Tao_File_Import.bat`** vào bất kỳ thư mục nào đang chứa các file Excel (ví dụ `D:\finger_temp\`).
- Khi nhấp đúp chạy, file `.bat` sẽ tự động gọi script và mẫu 7009 tại `D:\AI_AntiGravity skill\tao-file-import\`, quét toàn bộ file `.xlsx` trong thư mục hiện tại và lưu kết quả vào thư mục con **`result\`** ngay tại nơi đặt file `.bat`.

### Cách 3: Chạy dòng lệnh (tùy chọn cho 1 file lẻ)
```powershell
python -X utf8 "D:\AI_AntiGravity skill\tao-file-import\scripts\tao_file_import.py" "D:\duong_dan\file_nguon.xlsx"
```

---

## 4. Quy trình xử lý tự động bên trong
1. **Kiểm tra toàn vẹn mẫu 7009**: Xác thực mã SHA-256 của `7009_Employee_Template.xlsx` và cấu trúc 7 sheet chuẩn (`EMPLOYEES`, `SetCodes`, `Locations`, `Departments`, `Sections`, `Groups`, `Levels`).
2. **Quét sheet hiển thị**:
   - Bỏ qua hoàn toàn các sheet `hidden` và `veryHidden`.
   - Nhận diện nghiệp vụ (`Action`) dựa theo tên sheet hiển thị.
   - Giữ lại các sheet nguồn hiển thị khác (không phải sheet nghiệp vụ và không phải `EMPLOYEES`) ở cuối file kết quả.
3. **Dò dòng tiêu đề & đọc dữ liệu thông minh**:
   - Tự động dò dòng tiêu đề từ dòng 1 đến dòng 5 (hỗ trợ cả file có dòng tiêu đề nằm ở dòng 2).
   - Tự động bỏ qua cụm cột phụ bên phải thuộc phần `"Phụ trách quản lý / chấm công"` hoặc cột `Code` lặp lại.
   - Đọc trực tiếp các dòng có dữ liệu thực tế (xử lý nhanh ngay cả khi sheet bị định dạng thừa tới hơn 1 triệu dòng trống).
4. **Lọc & chuẩn hóa dữ liệu**:
   - Bỏ qua các dòng có `Remark` là `Đổi tên`.
   - Ánh xạ tên cột linh hoạt, quy đổi `Position` ➔ `Level`, chuẩn hóa `Location` khớp duy nhất theo danh mục `Locations`.`Code` của mẫu.
   - Áp dụng quy tắc riêng 8 cột cho `RESIGNED`.
5. **Chia lô & kiểm tra toàn diện**:
   - Kiểm tra trùng `Employee Code` và `ITCode` trên toàn bộ file nguồn.
   - Nếu `<= 30` nhân viên: xuất 1 file `result\<tên_gốc>.xlsx`.
   - Nếu `> 30` nhân viên: tự động tách tối đa 30 nhân viên/file với hậu tố `_01`, `_02`, ...
   - Xóa sạch dữ liệu/công thức thừa dưới dòng nhân viên cuối, gỡ named range/link lỗi (`Staff`, `Worker`) tương thích Excel 2021, đóng băng công thức ngoài trong sheet được giữ lại.
   - Chỉ ghi file kết quả chính thức sau khi tất cả các phần đều vượt qua bước kiểm tra hậu lưu (đối khớp font chữ, màu, viền, chiều cao dòng, dropdown Action, số lượng dòng).

---

## 5. Xử lý các thông báo lỗi thường gặp

| Thông báo lỗi | Nguyên nhân | Cách khắc phục |
| :--- | :--- | :--- |
| `Không tìm thấy sheet nghiệp vụ nào đang hiển thị` | File nguồn không có sheet nào đang hiện khớp bảng tên sheet nghiệp vụ. | Kiểm tra lại tên tab sheet hoặc bỏ ẩn (unhide) sheet cần import. |
| `Sheet '...' dòng X thiếu ITCode` (hoặc `Code`, `Name`, `Effected Date`) | Dòng dữ liệu bị bỏ trống trường bắt buộc. | Mở file nguồn điền đủ thông tin cho dòng báo lỗi rồi chạy lại `.bat`. |
| `Trùng Employee Code / ITCode trên toàn bộ nguồn` | Một mã nhân viên hoặc mã IT xuất hiện nhiều hơn 1 lần. | Kiểm tra và xóa/sửa bản ghi bị trùng trong file nguồn. |
| `Position không hợp lệ` | Cột chức vụ không thuộc `Daily Worker`, `DW`, `Staff`, `Worker`. | Chuẩn hóa lại giá trị chức vụ tại dòng báo lỗi. |
| `Location trống hoặc không khớp danh mục` | Địa điểm không khớp duy nhất với cột `Code` trong sheet `Locations` của mẫu 7009. | Kiểm tra lại chính tả địa điểm theo sheet `Locations` của mẫu. |
| `Công thức tham chiếu ngoài tại ... chưa có giá trị đã lưu` | Sheet giữ lại có công thức trỏ ra file Excel bên ngoài nhưng chưa lưu giá trị cache. | Mở file nguồn bằng Excel để tính lại, bấm Save rồi chạy lại `.bat`. |

---

## 6. Yêu cầu môi trường
- Windows với **Python 3.10+** đã cài trong `PATH`.
- Thư viện Python: `openpyxl` (`pip install openpyxl`).
- Không cần kết nối mạng hay mở AI chat khi vận hành hàng ngày.
