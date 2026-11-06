# HƯỚNG DẪN ĐƯA ỨNG DỤNG LÊN VERCEL (DEPLOY TO VERCEL)

Ứng dụng **So Sánh & Kiểm Thử Dữ Liệu Excel** đã được đóng gói hoàn chỉnh thành phiên bản Web App Client-Side tại thư mục `VercelApp/` sẵn sàng triển khai lên **Vercel** chỉ trong 1 phút.

---

## 🌟 TẠI SAO PHIÊN BẢN NÀY HOÀN HẢO CHO VERCEL?

1. **Tương thích 100% với Vercel**: Vercel là nền tảng máy chủ Serverless/Frontend. Phiên bản này chạy trực tiếp trên trình duyệt bằng thư viện `SheetJS` và `ExcelJS`, không cần cấu hình Docker hay máy chủ backend phức tạp.
2. **Không giới hạn dung lượng upload (Vượt giới hạn 4.5MB của Vercel)**:
   - Các hàm Serverless trên Vercel có giới hạn request body tối đa **4.5 MB**.
   - Với ứng dụng này, dữ liệu Excel được đọc và tính toán **ngay trong bộ nhớ RAM trình duyệt của người dùng**, giúp xử lý các file Excel lớn (20MB - 50MB+) cực kỳ mượt mà mà không lo lỗi `413 Payload Too Large` hay Timeout.
3. **Bảo mật tuyệt đối**: Dữ liệu tài chính, kiểm toán, nhân sự không bao giờ bị tải lên bất kỳ máy chủ nào.
4. **Đầy đủ mọi tính năng & định dạng báo cáo**:
   - Tùy chọn làm tròn: **"Lấy hết (Không làm tròn - So khớp chính xác tuyệt đối)"** và các mức 0 đến 6 chữ số.
   - Nhận diện cột `MASO`, `Mã số`, `Code`, `ID`...
   - Chuẩn hóa số 0, số âm `(xxx)`, tiền tệ, phẩy/chấm kiểu Việt Nam và quốc tế.
   - Dashboard KPI thống kê trực quan (Tổng bản ghi, Trùng khớp, Sai lệch, Lệch dòng, Tỷ lệ %).
   - Bảng xem nhanh sai lệch (Live Preview) trực tiếp trên web.
   - Xuất file kết quả `.xlsx` chuẩn 3 sheet:
     - **Sheet 1 (`TongQuan`)**: Báo cáo kiểm thử có màu sắc, bảng tổng hợp số liệu và trạng thái PASS / FAIL.
     - **Sheet 2 (`KetQuaSoSanh`)**: Dữ liệu gốc tô màu ô sai (hồng), tô màu dòng lệch (vàng), cột Result, Freeze Panes, AutoFilter.
     - **Sheet 3 (`DanhSachSai`)**: Bảng chi tiết từng ô sai lệch, Freeze Panes, AutoFilter.

---

## 🚀 CÁCH 1: DEPLOY QUA GITHUB LÊN VERCEL (KHUYÊN DÙNG)

Đây là cách phổ biến nhất để có tên miền miễn phí dạng `https://ten-du-an.vercel.app` và tự động cập nhật khi bạn sửa code:

1. **Khởi tạo và đẩy code lên GitHub**:
   Mở terminal (PowerShell) tại thư mục `D:\LamViec_1\SoSanhKiemThuExcel`:
   ```bash
   git init
   git add .
   git commit -m "Khoi tao web app so sanh excel san sang cho Vercel"
   git branch -M main
   git remote add origin https://github.com/<tai-khoan-cua-ban>/<ten-repo>.git
   git push -u origin main
   ```

2. **Kết nối với Vercel**:
   - Truy cập: [https://vercel.com](https://vercel.com) và đăng nhập bằng tài khoản GitHub của bạn.
   - Bấm nút **"Add New..."** ở góc phải trên -> Chọn **"Project"**.
   - Tìm và bấm **"Import"** repository vừa tạo.
   - File cấu hình `vercel.json` ở thư mục gốc đã được thiết lập sẵn `outputDirectory: "VercelApp"`, bạn chỉ cần bấm **"Deploy"**!
   - Chờ khoảng 20-30 giây, Vercel sẽ cấp cho bạn một đường link web chính thức để sử dụng và chia sẻ cho người khác.

---

## ⚡ CÁCH 2: DEPLOY TRỰC TIẾP BẰNG VERCEL CLI (KHÔNG CẦN GITHUB)

Nếu bạn không muốn tạo GitHub repo, có thể deploy thẳng từ máy tính:

1. Mở PowerShell và di chuyển vào thư mục `VercelApp`:
   ```powershell
   cd D:\LamViec_1\SoSanhKiemThuExcel\VercelApp
   ```

2. Chạy lệnh deploy thông qua npx:
   ```powershell
   npx vercel
   ```

3. Làm theo hướng dẫn trên màn hình:
   - Đăng nhập vào tài khoản Vercel (Email hoặc GitHub).
   - Chọn `Set up and deploy?` -> Gõ `y` và nhấn Enter.
   - Giữ nguyên các giá trị mặc định bằng cách nhấn Enter cho đến khi hoàn tất.
   - Bạn sẽ nhận được đường dẫn trang web hoạt động ngay lập tức!

---

## 💻 CÁCH 3: CHẠY THỬ NGHIỆM TRÊN MÁY (LOCAL)

Bạn có thể chạy thử nghiệm ngay lập tức bằng một trong các cách sau:
- **Cách đơn giản nhất**: Mở trực tiếp file `D:\LamViec_1\SoSanhKiemThuExcel\VercelApp\index.html` bằng trình duyệt Google Chrome hoặc Microsoft Edge.
- **Dùng Node.js local server**:
  ```powershell
  npx serve D:\LamViec_1\SoSanhKiemThuExcel\VercelApp
  ```
