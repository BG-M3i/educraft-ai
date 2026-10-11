# EduCraft AI — bản có backend

## Chức năng
- Tạo kế hoạch bài dạy, dàn ý slide và đề kiểm tra bằng AI.
- Lưu tài liệu vào SQLite trên máy chủ.
- Xuất DOCX, PPTX và PDF.
- API key được giữ ở backend, không đưa vào trình duyệt.

## Chạy thử trên máy tính
Yêu cầu Node.js 20 trở lên.

```bash
npm install
cp .env.example .env
```

Mở `.env`, thêm `AI_API_KEY` của nhà cung cấp tương thích OpenAI. Sau đó:

```bash
npm start
```

Mở http://localhost:3000

Nếu chưa cấu hình API key, tính năng tạo bằng AI sẽ báo rõ và không giả vờ đã chạy.

## Triển khai công khai
Có thể triển khai project lên Render, Railway hoặc máy chủ Node.js:
1. Tải project lên một repository riêng.
2. Tạo Web Service từ repository đó; Build command: `npm install`; Start command: `npm start`.
3. Thêm biến môi trường `AI_API_KEY`, `AI_BASE_URL`, `AI_MODEL`.
4. Gắn persistent disk và đặt `DATA_DIR` vào thư mục trên disk để dữ liệu SQLite không mất khi redeploy.
5. Mở URL HTTPS do nền tảng cấp để kiểm tra.

## Giới hạn và việc cần hoàn thiện trước khi mở cho nhiều người
- Đây là bản nền tảng để triển khai, chưa có đăng nhập và phân quyền người dùng. Không mở công khai cho tài liệu nhạy cảm trước khi thêm xác thực.
- SQLite lưu trên máy chủ; hãy cấu hình persistent disk và sao lưu.
- PDF dùng Helvetica mặc định có thể không hiển thị đầy đủ dấu tiếng Việt. Để có dấu chuẩn, cung cấp đường dẫn tới font TTF Unicode qua `PDF_FONT_PATH`.
- Chi phí AI phụ thuộc nhà cung cấp/model và lượng sử dụng.
- Nội dung giáo dục do AI tạo phải được giáo viên kiểm tra theo chương trình, sách và quy định hiện hành.
