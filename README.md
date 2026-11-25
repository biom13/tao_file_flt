# Excel add-in: Tạo file FLT

Add-in cho Excel 2021 (Windows). Thay cho `Tao_File_FLT.bat` + `generate_flt.py`: đọc dữ liệu ngay trong workbook đang mở, gom mã nhân viên theo cổng/máy và tạo các file `.flt`. Toàn bộ xử lý chạy trên máy của bạn, dữ liệu Excel không được gửi đi đâu.

## Cấu trúc

| File | Vai trò |
| --- | --- |
| `manifest.xml` | Khai báo add-in cho Excel (đây là file bạn cài hoặc chia sẻ) |
| `taskpane.html`, `taskpane.css`, `taskpane.js` | Giao diện bên phải Excel |
| `flt-core.js` | Logic nhận diện cột, gom nhóm, đặt tên file, ghi nhị phân `.flt`, tạo `.zip` |
| `assets/` | Icon |
| `Day_len_GitHub.bat` / `.ps1` | Đẩy code lên GitHub và bật Pages bằng dòng lệnh (dùng GitHub CLI) |
| `Day_len_GitHub_khong_gh.bat` / `.ps1` | Như trên nhưng chỉ cần git, không cần GitHub CLI |
| `Cau_hinh_manifest.bat` | Điền username và tên repo GitHub vào `manifest.xml` (chỉ cần khi đổi repo) |

## Cách nhanh (không cần thao tác trên web GitHub)

`manifest.xml` đã điền sẵn cho repo `biom13/tao_file_flt`. Nhấp đúp **`Day_len_GitHub.bat`**: script tự cài git và GitHub CLI nếu thiếu, đăng nhập GitHub (một lần), khởi tạo git, commit, push lên nhánh `main` và bật GitHub Pages. Nếu không muốn cài GitHub CLI, dùng **`Day_len_GitHub_khong_gh.bat`** (chỉ cần git; muốn bật Pages tự động thì dán một token, không thì bật tay theo hướng dẫn trong script). Xong thì làm tiếp **Bước 3** (cài vào Excel). Các Bước 1 và 2 bên dưới chỉ cần khi bạn làm tay hoặc đổi sang repo khác.

## Bước 1. Đưa code lên GitHub Pages

1. Tạo repository **Public** trên GitHub (ví dụ `flt-addin`). GitHub Pages miễn phí chỉ dùng được với repo public, nên không đưa dữ liệu thật hay thông tin nhạy cảm vào repo.
2. Upload toàn bộ file trong thư mục này (kể cả `.nojekyll` và thư mục `assets`) lên nhánh `main`.
3. Vào **Settings → Pages → Build and deployment**, chọn **Deploy from a branch**, nhánh `main`, thư mục `/ (root)`, bấm Save.
4. Chờ 1–2 phút, mở thử `https://<username>.github.io/<repo>/taskpane.html`. Thấy trang "Tạo file FLT" là thành công (trang sẽ báo chỉ chạy trong Excel, đó là bình thường).

## Bước 2. Điền địa chỉ vào manifest

Chạy `Cau_hinh_manifest.bat`, nhập GitHub username và tên repo. Hoặc mở `manifest.xml` bằng Notepad và thay `YOUR-USERNAME`, `YOUR-REPO`. Sau đó upload lại `manifest.xml` lên GitHub nếu bạn đã upload bản cũ.

## Bước 3. Cài vào Excel 2021

Excel 2021 dùng thư mục chia sẻ làm "Trusted Add-in Catalog":

1. Tạo thư mục, ví dụ `C:\OfficeAddins`, chép `manifest.xml` vào.
2. Chuột phải thư mục → **Properties → Sharing → Share…** để chia sẻ thư mục, ghi lại đường dẫn mạng (dạng `\\TEN-MAY\OfficeAddins`).
3. Mở Excel → **File → Options → Trust Center → Trust Center Settings… → Trusted Add-in Catalogs**. Dán đường dẫn mạng vào **Catalog Url**, bấm **Add catalog**, tick **Show in Menu**, bấm OK và khởi động lại Excel.
4. Vào **Insert → My Add-ins → SHARED FOLDER**, chọn **Tạo file FLT**, bấm **Add**.
5. Từ lần sau, nút **Tạo file FLT** nằm ở thẻ **Home**.

## Chia sẻ cho người khác

- Đặt `manifest.xml` vào một thư mục mạng dùng chung của công ty, rồi mỗi người thêm đường dẫn đó làm Trusted Add-in Catalog (bước 3). Cập nhật code trên GitHub là mọi người dùng bản mới, không phải cài lại.
- Hoặc gửi file `manifest.xml` cho từng người để họ làm bước 3 với thư mục trên máy mình.
- Máy người dùng cần có Internet (để tải trang từ GitHub Pages và `office.js` của Microsoft).

## Cách dùng

1. Mở workbook có dữ liệu (dòng 1 là tiêu đề cột).
2. Home → **Tạo file FLT** → **Quét workbook**.
3. Xem danh sách file và số mã, rồi chọn:
   - **Lưu vào thư mục result**: chọn một thư mục, add-in tạo thư mục con `result` và ghi các file `.flt` vào đó (nếu Excel hỗ trợ).
   - **Tải file .zip**: tải một file zip chứa toàn bộ `.flt`, giải nén ra là dùng.
   - **Tải từng file**: tải riêng từng file.

## Quy tắc xử lý (giống tài liệu kỹ thuật)

- Chỉ đọc sheet đang hiển thị, bỏ qua sheet ẩn; tiêu đề ở dòng 1.
- Cột mã nhân viên: `IT Code`, `IT. No`, `IT No`, `IT_No`, `IT_Code`, `ITNo`, `ITCode`; chỉ khi không có mới dùng `No.` / `No`.
- Cột cổng/máy: `Cổng được phép quẹt vân tay`, `Máy được phép quẹt bằng vân tay`, `Máy được phép quẹt vân tay`, `Machine No.`, `Machine No`, `Machine_No`.
- Bỏ dòng thiếu mã hoặc thiếu cổng/máy, bỏ dòng `Column1`, `Column11`…; mã dạng `59291.0` đổi thành `59291`.
- Gom nhóm không phân biệt thứ tự và hoa/thường (`Cổng A, K` = `Cổng K, A`); tiền tố là một phần của nhóm (`Cổng A` và `Máy A` là hai file khác nhau); mã trùng trong cùng nhóm chỉ giữ một lần.
- Tên file: bỏ dấu tiếng Việt, dấu phẩy thành `_`, ký tự đặc biệt thành `_`, cắt khoảng trắng/`_`/`.` ở hai đầu; nếu rỗng thì đặt `unknown_gate` (ví dụ `Máy 73, 79, 120, 163` → `May 73_79_120_163.flt`).
- Định dạng nhị phân `.flt` theo mục 4 của tài liệu (header 9 byte, mỗi mã một bản ghi).

## Đã đối chiếu với bản Python

Logic được port từ `generate_flt.py`. Đã chạy thử cùng một file Excel (16 nhóm, gồm sheet ẩn, sheet dùng cột `No.` dự phòng, mã `59291.0`, tên có dấu, tiền tố lặp, dấu phẩy cuối...) qua cả hai bản: tên file và nội dung từng byte giống nhau. File `.flt` tạo ra cũng khớp từng byte với file mẫu `Cong_A_K.flt`.

## Khác biệt so với bản .bat

- Add-in đọc **workbook đang mở**, không tự quét nhiều file `.xlsx` trong một thư mục như `generate_flt.py`. Muốn gộp nhiều file, mở từng file rồi quét, hoặc chép các sheet vào cùng một workbook.
- Không còn lỗi `PermissionError` vì không cần đóng Excel.
- Nếu hai nhóm khác nhau cho ra cùng một tên file sau khi bỏ dấu, bản Python ghi đè lẫn nhau, còn add-in tự đổi tên (thêm `_2`) và hiện cảnh báo.
- Mã nhân viên dài hơn 255 byte: bản Python báo lỗi và dừng, add-in bỏ qua mã đó và hiện cảnh báo.

## Khắc phục sự cố

- **Add-in không hiện trong SHARED FOLDER**: kiểm tra đường dẫn catalog là đường dẫn mạng (`\\...`), thư mục đã được chia sẻ, và đã khởi động lại Excel.
- **Khung bên phải trắng hoặc báo lỗi tải**: kiểm tra link `taskpane.html` trên GitHub Pages mở được trong trình duyệt, và `YOUR-USERNAME`/`YOUR-REPO` đã được thay hết trong `manifest.xml`.
- **Bấm tải zip mà không thấy file**: dùng **Lưu vào thư mục result** hoặc **Tải từng file**.
- **Sau khi cập nhật code mà vẫn thấy bản cũ**: đóng hẳn Excel, mở lại; nếu vẫn cũ thì tăng `<Version>` trong `manifest.xml`.
