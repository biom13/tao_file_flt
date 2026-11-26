---
name: tao-file-import
description: >-
  Chuyển danh sách nhân viên từ một tệp Excel nguồn thành một hoặc nhiều tệp import theo mẫu 7009.
  Sử dụng mẫu 7009_Employee_Template.xlsx có sẵn, tự động kiểm tra lỗi (validation), điền dữ liệu
  vào sheet EMPLOYEES và chia nhỏ file nếu có trên 30 nhân viên. Kích hoạt khi người dùng yêu cầu
  tạo file import, xuất file 7009, hoặc xử lý danh sách nhân viên import.
---

# Skill: Tạo File Import Nhân Viên Mẫu 7009 (`tao-file-import`)

Skill này tự động hóa 100% quy trình chuyển đổi danh sách nhân viên từ tệp Excel nguồn sang các tệp import theo định dạng biểu mẫu 7009.

## Tài nguyên & Script
- **Script xử lý chính**: `scripts/tao_file_import.py` (hoặc `D:\AntiGravity skill\tao-file-import\scripts\tao_file_import.py`)
- **File mẫu 7009**: `resources/7009_Employee_Template.xlsx` (hoặc `D:\AntiGravity skill\tao-file-import\resources\7009_Employee_Template.xlsx`)

---

## Cách thực thi (Execution Runbook)

Khi người dùng yêu cầu chạy skill hoặc cung cấp đường dẫn file nguồn Excel:

1. **Chạy lệnh Python**:
   ```powershell
   python "D:\AntiGravity skill\tao-file-import\scripts\tao_file_import.py" "<duong_dan_file_nguon.xlsx>" ["<thu_muc_xuat>"]
   ```
   *(Nếu không truyền thư mục xuất, file kết quả sẽ được tạo ngay tại cùng thư mục với file nguồn).*

2. **Xử lý kết quả**:
   - **Thành công**: Thông báo cho người dùng danh sách các file `_01.xlsx`, `_02.xlsx`... đã được tạo cùng số lượng nhân viên trong từng file.
   - **Thất bại / Có lỗi validation**: Script sẽ liệt kê chính xác sheet, dòng và nguyên nhân lỗi (trùng mã Code, ITCode, sai chức vụ, trống địa điểm...). Trích xuất thông tin lỗi rõ ràng để người dùng chỉnh sửa file nguồn.

---

## Tóm tắt quy tắc nghiệp vụ cốt lõi

### 1. Phân loại Action & Thứ tự lấy ngày hiệu lực (Effected Date)
| Tên sheet nguồn | Action | Thứ tự ưu tiên lấy Effected Date |
| :--- | :--- | :--- |
| `new` | `CREATE-NEW` | `Ngày lấy vân tay` → `Starting Date` / `Starting date` → `Effected Date` |
| `transfer`, `Transfer & rehire`, `Old DW` | `TRANSFER` | `Transfer From` → `Transfer/rehire date` → `Starting Date` / `Starting date` → `Effected Date` |
| `resigned`, `resign`, `quit` | `RESIGNED` | `Stopping date` → `Quit Date` / `Quit date` → `Effected Date` |
| `info-changed`, `info changed`, `info_changed`, `info-change`, `info change`, `info_change` | `INFO-CHANGED` | `Effected Date` (nếu thiếu cột hoặc ô trống, tự động dùng ngày chạy skill) |

### 2. Từ đồng nghĩa cột (Sheet `Old DW`)
- Chấp nhận: `IT Code` → `IT. No`, `Full name` → `Name`, `Department` → `Departments`, `Level` → `Position`, `Set code` → `Setcode`.
- Nếu có cả hai, luôn ưu tiên tên cột chuẩn.

### 3. Quy tắc điền & chuyển đổi dữ liệu
- **Dòng hợp lệ**: Dòng có ít nhất 1 trong 3 trường `IT. No`, `Code`, `Name`. Khi xuất bắt buộc phải có đủ cả 3 trường.
- **Level**: Chuyển đổi từ cột `Position`:
  - `Daily Worker` hoặc `DW` → `DW`
  - `staff` → `Staff`
  - `worker` → `Worker`
  - Giá trị khác → Báo lỗi validation.
- **Location**: So khớp không phân biệt hoa thường, bỏ dấu và khoảng trắng với cột `Code` trong sheet `Locations` của mẫu. Sử dụng đúng tên chuẩn trong mẫu khi xuất.
- **Riêng RESIGNED**: Chỉ điền đúng 8 cột:
  `Action` (RESIGNED), `Effected Date`, `Employee Code`, `FullName`, `ITCode`, `SetCode` (Resignation), `Level`, `Location` (Resignation). Các cột khác để trống hoàn toàn.

### 4. Quy tắc chia tệp & bảo tồn
- Tối đa **30 nhân viên/tệp**. Tự động gắn hậu tố `_01`, `_02`...
- Bỏ qua hoàn toàn các sheet ẩn (`hidden`, `veryHidden`).
- Các sheet nguồn hiển thị khác ngoài nhóm nghiệp vụ được sao chép vào cuối file xuất (ngoại trừ sheet `EMPLOYEES` của nguồn).
- Công thức ngoài/add-in được thay thế bằng giá trị đã lưu trong cache Excel.
