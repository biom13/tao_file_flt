# Tổng Hợp Yêu Cầu & Quy Tắc Nghiệp Vụ Tạo File Import 7009

Tài liệu này tổng hợp đầy đủ toàn bộ yêu cầu ban đầu và các quy tắc mở rộng đã được thống nhất qua quá trình chạy thử trên các tệp dữ liệu thực tế từ nhiều nguồn khác nhau.

---

## 1. Hình thức vận hành & Đầu ra
- **Hình thức chạy**: Đóng gói thành tệp **`.bat`** (`Tao_File_Import.bat`) gọi script Python cố định (`scripts/tao_file_import.py`).
- **Xử lý hàng loạt theo thư mục**: Khi chạy `Tao_File_Import.bat`, tự động quét tất cả các tệp `.xlsx` nằm trong cùng thư mục chứa tệp `.bat` (bỏ qua tệp tạm `~$*.xlsx` và tệp mẫu `7009_Employee_Template.xlsx`).
- **Thư mục kết quả**: Toàn bộ tệp import sinh ra được lưu vào thư mục con **`result\`** bên trong thư mục chứa tệp `.bat`.
- **Bảo toàn tệp nguồn**: Không bao giờ chỉnh sửa hay ghi đè tệp Excel nguồn.
- **Xử lý độc lập từng tệp**: Nếu một tệp nguồn gặp lỗi kiểm tra, hệ thống báo chi tiết lỗi của tệp đó và vẫn tiếp tục xử lý các tệp còn lại trong thư mục.

---

## 2. File mẫu & Định dạng (Mẫu 7009)
- Sử dụng đúng tệp mẫu gốc **`7009_Employee_Template.xlsx`** đóng gói trong `resources\` (mã kiểm tra toàn vẹn SHA-256: `3242d0ac486b87b48f71d553f9317c0fc9834dc350a6c44f6167ffa1cebe1a5a`).
- Tuyệt đối không tự dựng workbook từ đầu hay áp định dạng mặc định.
- Bảo toàn 100% **font chữ, cỡ chữ, màu sắc, đường viền, căn lề, định dạng ngày Excel, chiều cao hàng, độ rộng cột, vùng tiêu đề (dòng 1–5) và dropdown cột Action** trên sheet `EMPLOYEES`.
- Giữ nguyên 6 sheet tham chiếu ngay sau `EMPLOYEES` theo đúng thứ tự: `SetCodes`, `Locations`, `Departments`, `Sections`, `Groups`, `Levels`.
- Xóa sạch nội dung/công thức ở mọi dòng thừa dưới dòng nhân viên cuối cùng (và cắt gọn dòng trống thừa); xác nhận lại sau khi lưu không còn dữ liệu thừa ngoài 17 cột import.
- Tương thích Excel 2021: tự động gỡ các named range/external link lỗi của mẫu (`Staff`, `Worker`).

---

## 3. Quy tắc nhận diện Sheet & Action

### 3.1. Sheet bị bỏ qua & Sheet giữ lại
- **Sheet `hidden` và `veryHidden`**: Bỏ qua hoàn toàn, kể cả khi tên sheet khớp với một nghiệp vụ.
- **Sheet nguồn hiển thị không phải nghiệp vụ**: Sao chép vào cuối file kết quả (ngoại trừ sheet tên `EMPLOYEES`). Nếu trùng tên với 6 sheet tham chiếu của mẫu ➔ dừng và báo lỗi.
- **Công thức trong sheet giữ lại**: Giữ nguyên công thức nội bộ; công thức tham chiếu tệp ngoài/add-in được thay bằng giá trị cache đã lưu (nếu thiếu giá trị cache ➔ dừng và báo rõ sheet, ô lỗi).

### 3.2. Bảng ánh xạ Tên Sheet ➔ Action
Tên sheet được so khớp không phân biệt hoa thường, chuẩn hóa khoảng trắng thừa và dấu gạch dưới (`_`).

| Action đầu ra | Các tên sheet nguồn được chấp nhận |
| :--- | :--- |
| **`CREATE-NEW`** | `new`, `New Employee`, `New DW` |
| **`TRANSFER`** | `transfer`, `Transfer & rehire`, `transfer&rehire`, `transfer&rehire&đổi tên`, `transfer & rehire & đổi tên`, `Old DW` |
| **`RESIGNED`** | `resigned`, `resign`, `quit`, và **mọi tên sheet có chứa chữ `resign` hoặc `nghỉ việc` / `nghi viec`** (không phân biệt hoa thường, dấu tiếng Việt) |
| **`INFO-CHANGED`** | `info-changed`, `info changed`, `info_changed`, `info-change`, `info change`, `info_change` |

---

## 4. Quy tắc nhận diện Dòng tiêu đề & Ánh xạ Cột nguồn

### 4.1. Nhận diện dòng tiêu đề & loại trừ cột phụ
- Tự động dò dòng tiêu đề trong phạm vi **dòng 1 đến dòng 5** (hỗ trợ cả các tệp có dòng tiêu đề nằm ở dòng 2 do có tiêu đề nhóm ở dòng 1).
- **Bỏ qua phần `"Phụ trách quản lý / chấm công"`**: Tự động dừng đọc cột khi gặp nhóm cột có tiêu đề phía trên chứa `"Phụ trách..."` hoặc gặp cột `Code` lặp lại lần thứ 2 bên phải, tránh lấy nhầm thông tin của người quản lý.

### 4.2. Tên cột tương đương (áp dụng cho TẤT CẢ các sheet)
So khớp không phân biệt hoa thường và khoảng trắng thừa; nếu tồn tại cả hai cách viết thì ưu tiên tên cột chuẩn.

| Cột chuẩn | Các tên cột tương đương được chấp nhận |
| :--- | :--- |
| `IT. No` | `IT Code`, `No.` (ưu tiên `IT. No` và `IT Code` nếu có cùng lúc) |
| `Name` | `Full name`, `Names` |
| `Departments` | `Department` |
| `Group` | `Group/Actual` (và `Group / Actual`) |
| `Setcode` | `Set code` |
| `Position` | `Level` |
| `Effected Date` | `Effective date`, `Transfer/rehire date` (và `Transfer / rehire date`) |

---

## 5. Quy tắc lọc dòng & Loại trừ bản ghi
- **Điều kiện đưa vào kiểm tra**: Dòng có ít nhất 1 trong 3 trường `IT. No`, `Code`, `Name`. Các dòng trống hoặc chỉ chứa công thức phụ ở cột khác bị bỏ qua.
- **Điều kiện bắt buộc khi xuất**: Cả 3 trường `IT. No`, `Code`, `Name` đều phải có giá trị.
- **Loại trừ theo `Remark`**: Nếu cột `Remark` có giá trị là **`Đổi tên`** (không phân biệt hoa thường, dấu tiếng Việt hay khoảng trắng thừa) ➔ **bỏ qua hoàn toàn**, không thêm bản ghi đó vào tệp import.

---

## 6. Quy tắc xác định Ngày hiệu lực (`Effected Date`)
So khớp tên cột ngày không phân biệt hoa thường; lấy giá trị từ cột đầu tiên có dữ liệu theo thứ tự ưu tiên:

| Action | Thứ tự ưu tiên lấy cột ngày | Khi thiếu ngày |
| :--- | :--- | :--- |
| **`CREATE-NEW`** | `Ngày lấy vân tay` ➔ `Starting Date` ➔ `Effected Date` / `Effective date` ➔ `Transfer/rehire date` | **Dừng & báo lỗi** (không lấy ngày hôm nay) |
| **`TRANSFER`** | `Transfer From` ➔ `Transfer/rehire date` ➔ `Starting Date` ➔ `Effected Date` / `Effective date` ➔ `Ngày lấy vân tay` | **Dừng & báo lỗi** (không lấy ngày hôm nay) |
| **`RESIGNED`** | `Stopping date` ➔ `Quit Date` ➔ `Effected Date` / `Effective date` ➔ `Transfer/rehire date` | **Dừng & báo lỗi** (không lấy ngày hôm nay) |
| **`INFO-CHANGED`** | `Effected Date` / `Effective date` ➔ `Transfer/rehire date` | Riêng action này được **tự động dùng ngày hiện hành** khi chạy |

---

## 7. Quy tắc điền dữ liệu vào sheet `EMPLOYEES`

| Cột trong `EMPLOYEES` | Quy tắc điền dữ liệu |
| :--- | :--- |
| `Action` | Theo tên sheet nguồn (`CREATE-NEW`, `TRANSFER`, `RESIGNED`, `INFO-CHANGED`) |
| `Effected Date` | Theo bảng ưu tiên ngày ở Mục 6 |
| `Employee Code` | Từ cột `Code` |
| `FullName` | Từ cột `Name` (hoặc `Full name`, `Names`) |
| `ITCode` | Từ cột `IT. No` (hoặc `IT Code`) |
| `SetCode` | Từ cột `Setcode` (hoặc `Set code`) |
| `Level` | Quy đổi từ `Position` (hoặc `Level`): `Daily Worker` / `DW` ➔ `DW`; `staff` ➔ `Staff`; `worker` ➔ `Worker`. Giá trị khác ➔ dừng báo lỗi. |
| `Location` | Đối chiếu khớp duy nhất với cột `Code` trong sheet `Locations` của mẫu (bỏ qua dấu tiếng Việt, hoa thường, khoảng trắng; ví dụ `Hà Nội`, `HaNoi`, `HA NOI` ➔ `Ha Noi`). Dùng đúng cách viết trong mẫu. |
| `Department`, `Section`, `Group` | Từ `Departments` (hoặc `Department`), `Section`, `Group` |
| `Gender` | Giữ giá trị nguồn nếu có; thiếu cột hoặc trống thì để trống |
| `Bio-Code`, `Note`, `ID Card No.`, `Email`, `Phone Number` | Để trống |

### Quy tắc riêng cho nghiệp vụ `RESIGNED`
- Chỉ điền đúng **8 cột**: `Action`, `Effected Date`, `Employee Code`, `FullName`, `ITCode`, `SetCode`, `Level`, `Location`.
- `SetCode` và `Location` luôn cố định là **`Resignation`**.
- Toàn bộ các cột còn lại để trống (kể cả khi nguồn có sẵn phòng ban hoặc giới tính).
- Sheet nguồn chỉ cần các cột `IT. No`, `Code`, `Name`, `Position` và một cột ngày nghỉ hợp lệ.

---

## 8. Quy tắc chia tệp & Kiểm tra trước khi xuất
- **Giới hạn bản ghi**: Tối đa **30 nhân viên/tệp**. Nếu `> 30` nhân viên, chia theo đúng thứ tự sheet và thứ tự dòng trong nguồn thành các tệp hậu tố `_01`, `_02`, ...
- **Các lỗi bắt buộc dừng xuất**:
  1. Trùng `Employee Code` hoặc `ITCode` trên toàn bộ nguồn (kể cả khác sheet).
  2. Thiếu trường bắt buộc (`IT. No`, `Code`, `Name`, ngày hiệu lực, `Position`, `Location`, `Department` đối với nhóm không phải `RESIGNED`).
  3. `Position` không hợp lệ hoặc `Location` trống / không khớp duy nhất.
  4. Sheet nguồn hiển thị trùng tên với sheet tham chiếu của mẫu.
  5. Công thức ngoài trong sheet giữ lại không có giá trị đã lưu.
  6. Sai lệch định dạng, dropdown, thứ tự sheet hoặc còn dữ liệu thừa sau khi lưu.
