# Ứng Dụng Chuẩn Hóa Đề Thi & Văn Bản Hành Chính NĐ 30

Ứng dụng Web tĩnh (Client-Side) chất lượng cao giúp giáo viên và cán bộ văn phòng chuẩn hóa đề thi trắc nghiệm, bóc tách đáp án, dồn dòng phương án và chuẩn hóa văn bản hành chính theo Nghị định 30/2020/NĐ-CP trực tiếp trong trình duyệt mà không cần cài đặt phần mềm.

---

## 🚀 Hướng Dẫn Đưa Lên Netlify (Trong 30 Giây)

Ứng dụng được thiết kế **100% Client-Side** (HTML5, CSS3, JavaScript thuần + JSZip), hoàn toàn không cần server backend phức tạp:

### Cách 1: Kéo thả trực tiếp lên Netlify Drop (Nhanh nhất)
1. Truy cập vào trang: **[https://app.netlify.com/drop](https://app.netlify.com/drop)** (Đăng nhập tài khoản Netlify của bạn).
2. Kéo toàn bộ thư mục `App-chuan-hoa-dap-an-ND30` thả vào ô upload trên trang web.
3. Netlify sẽ cấp ngay một tên miền miễn phí (Ví dụ: `https://chuan-hoa-de-thi.netlify.app`) để bạn và đồng nghiệp sử dụng online mọi lúc, mọi nơi!

### Cách 2: Đẩy mã nguồn lên GitHub và kết nối Netlify
1. Tạo một repository mới trên GitHub và đẩy mã nguồn trong thư mục này lên.
2. Trên Netlify, chọn **Add new site** > **Import an existing project** > Chọn GitHub repo vừa tạo.
3. Để nguyên cấu hình mặc định (Publish directory: `.`) và bấm **Deploy**.

---

## 💻 Cách Chạy Trực Tiếp Trên Máy Tính (Offline)

- **Cách 1 (Nhanh nhất):** Nhấp đúp chuột vào file [`Mo_Ung_Dung.bat`](file:///d:/App-chuan-hoa-dap-an-ND30/Mo_Ung_Dung.bat) để tự động mở ứng dụng trên trình duyệt web mặc định.
- **Cách 2 (Tạo Icon Desktop):** Nhấp đúp vào file [`Tao_Icon_Desktop.bat`](file:///d:/App-chuan-hoa-dap-an-ND30/Tao_Icon_Desktop.bat) để tạo biểu tượng phím tắt ra màn hình chính Desktop giúp mở nhanh mọi lúc.
- **Cách 3 (Chia sẻ mạng nội bộ/Wi-Fi):** Nhấp đúp vào file [`Chay_May_Chu_LAN.bat`](file:///d:/App-chuan-hoa-dap-an-ND30/Chay_May_Chu_LAN.bat) để các thầy cô trong cùng mạng Wi-Fi trường học có thể cùng truy cập sử dụng.
- **Cách 4:** Nhấp đúp trực tiếp vào file [`index.html`](file:///d:/App-chuan-hoa-dap-an-ND30/index.html).

---

## 🎯 4 Bộ Công Cụ Tích Hợp

### 1. Chuẩn Hóa Đề Thi Chung
- **Tự động bóc tách và xóa bảng đáp án** ở cuối đề thi (Phần 1: Trắc nghiệm, Phần 2: Đúng/Sai, Phần 3: Trả lời ngắn).
- **Gắn dấu hoa thị `*`** vào phương án đúng (A.*, B.*, C.*, D.* hoặc a*), b*)...).
- **Chuẩn hóa quy cách đề mục**: `Câu 1.`, `Câu 2.`, `PHẦN 1.`, `PHẦN 2.`, `PHẦN 3.`.
- **Tách phương án nằm ngang** thành các dòng dọc riêng biệt.
- **Tự động điền dòng `Đáp án: [Kết quả]`** cho câu hỏi trả lời ngắn Phần 3 nếu chưa có.
- **Tùy chọn nhận diện in đậm / gạch chân / chữ đỏ** làm đáp án đúng khi file gốc không có bảng đáp án.

### 2. Chuẩn Hóa Đề Thi Tiếng Anh
- Xử lý chuyên sâu cấu trúc đề thi Tiếng Anh.
- Tự động gom nhóm các câu hỏi theo bài đọc (Reading passages) và bài điền từ (Cloze tests).
- Bẻ gãy các bảng True/False và bảng câu hỏi tiếng Anh thành đoạn văn bản chuẩn.
- Bóc tách bảng đáp án tiếng Anh và gắn dấu `*` tự động.

### 3. Chuẩn Hóa Văn Bản Hành Chính NĐ 30/2020/NĐ-CP
- **Căn lề khổ giấy A4 chuẩn**: Lề trên 2.0cm, Lề dưới 2.0cm, Lề trái 3.0cm, Lề phải 1.5cm.
- **Đánh số trang tự động** ở đầu trang (Số Ả Rập, phông Times New Roman 13pt, căn giữa, không đánh số trang 1).
- **Chuẩn hóa Quốc hiệu & Tiêu ngữ**: Căn giữa, in đậm, phông 13/14pt.
- **Chuẩn hóa Tiêu đề & Đề mục**: Chương, Mục, Điều, I., 1., a)... đúng cấp độ in đậm/nghiêng.
- **Chuẩn hóa thân đoạn văn**: Căn đều 2 bên (Justified), thụt đầu dòng 1.0cm, giãn dòng 1.2, giãn đoạn 6pt, phông Times New Roman 13/14pt.
- **Định dạng bảng biểu**: Phông Times New Roman 12-14pt, không thụt đầu dòng.

### 4. Bố Cục / Dồn Dòng Phương Án Trắc Nghiệm
- **Chế độ Auto 4-2-1 thông minh**: Tự động đo độ dài trực quan của chữ, công thức toán học MathType / OMML và hình ảnh để xếp 4, 2 hoặc 1 dòng tối ưu nhất mà không bao giờ bị tràn dòng hay lệch lề.
- **Chế độ 4 phương án / dòng**: Tab Stop chuẩn ~3.7cm, ~7.4cm, ~11.1cm.
- **Chế độ 2 phương án / dòng**: Tab Stop chuẩn ~7.94cm (Dòng 1: A-B, Dòng 2: C-D).
- **Chế độ 1 phương án / dòng**: Tách riêng từng phương án trên 1 dòng.
- **Bảo toàn 100%** công thức toán học MathType OLE, OMML Math, hình ảnh sơ đồ, màu sắc, định dạng và dấu hoa thị `*`.

### 5. Xử Lý Hàng Loạt (Batch Queue)
- Kéo thả nhiều file `.docx` cùng lúc.
- Tự động chuẩn hóa song song toàn bộ file với thanh tiến trình trực quan.
- Tải về toàn bộ kết quả trong 1 file `.zip` duy nhất.

---

## 🔒 Bảo Mật & An Toàn 100%

- Toàn bộ quá trình giải nén, đọc cấu trúc XML OpenXML, xử lý và đóng gói file `.docx` được thực hiện **trực tiếp trong trình duyệt máy tính của bạn**.
- Đề thi và tài liệu nội bộ **hoàn toàn không được gửi lên bất kỳ máy chủ nào**, đảm bảo tính bảo mật và riêng tư tuyệt đối cho kỳ thi của trường học.
