# Tổng hợp Yêu cầu Người dùng — Công cụ tạo file `.flt`

**Thư mục làm việc chính:** `D:\AI_AntiGravity skill\File flt`  
**Ngày cập nhật:** 25/11/2026

---

## Bảng tổng hợp yêu cầu đã tinh gọn

| Hạng mục | Yêu cầu đã chốt | Trạng thái |
| --- | --- | --- |
| **1. Thư mục chính & file chạy** | Đặt bộ công cụ (`Tao_File_FLT.bat`, `generate_flt.py`) ngay trong `D:\AI_AntiGravity skill\File flt`; mỗi lần chạy tự động quét hết các file Excel trong thư mục này. | Đã hoàn thành |
| **2. Thư mục lưu kết quả** | Toàn bộ file `.flt` sinh ra được lưu vào thư mục con `D:\AI_AntiGravity skill\File flt\result\` cho gọn gàng, dễ tìm. | Đã hoàn thành |
| **3. Định dạng file `.flt`** | File `.flt` xuất ra phải đúng chuẩn cấu trúc nhị phân của file mẫu `cong K_A.flt` để import vào hệ thống không bị báo lỗi. | Đã hoàn thành |
| **4. Quy đổi tên cột tương đương** | - Cột mã nhân viên: `IT Code`, `IT. No` là như nhau; bổ sung thêm cột `No.` (dùng khi sheet không có cột `IT Code` / `IT. No`).<br>- Cột cổng/máy: `Machine No.`, `Cổng được phép quẹt vân tay` (và `Máy được phép quẹt bằng vân tay`) là như nhau. | Đã hoàn thành |
| **5. Phạm vi sheet** | Chỉ đọc các sheet đang hiển thị (`New`, `New DW`, `Old DW`, `transfer`...), tuyệt đối bỏ qua các sheet đang bị ẩn để không sinh nhầm dữ liệu cũ/nháp. | Đã hoàn thành |
| **6. Gom nhóm theo cổng/máy** | Gom các `IT Code` có cùng cổng/máy vào chung 1 file `.flt`; **không phân biệt thứ tự** (ví dụ `Cổng A, K` và `Cổng K, A` gom chung 1 file) và không phân biệt hoa/thường. | Đã hoàn thành |
| **7. Chuẩn hóa tên file đầu ra** | Đặt tên file theo tên cổng/máy, **không dấu tiếng Việt và không chứa ký tự đặc biệt** (ví dụ: `May 71_79_120_163.flt`). | Đã hoàn thành |
| **8. Tài liệu lưu trữ** | Xuất đầy đủ tài liệu Kỹ thuật/Quy tắc và tài liệu Tổng hợp Yêu cầu ở cả 2 định dạng `.docx` và `.md`. | Đã hoàn thành |
