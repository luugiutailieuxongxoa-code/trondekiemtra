Young Mix Pro — MVP
Bộ mã nguồn khởi đầu có giao diện tiếng Việt responsive cho máy tính, máy tính bảng và điện thoại.
Chạy frontend
Yêu cầu Node.js 20+:
```bash
npm install
npm run dev
```
Chạy backend (tùy chọn, cần để nhập DOCX)
Yêu cầu Python 3.11+:
```bash
cd backend
python -m venv .venv
# Windows PowerShell: .venv\Scripts\Activate.ps1
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
uvicorn main:app --reload --port 8000
```
API docs: http://localhost:8000/docs
Tạo `.env` ở thư mục gốc frontend:
```env
VITE_API_BASE_URL=http://localhost:8000
```
Khởi động lại Vite sau khi sửa `.env`.
Đưa lên GitHub và Vercel
Giải nén ZIP, tạo repository GitHub và tải toàn bộ nội dung thư mục `young-mix-pro` lên.
Import repository vào Vercel.
Framework preset: Vite; Build command `npm run build`; Output directory `dist`.
Nếu có backend ở host riêng, thêm biến môi trường `VITE_API_BASE_URL` là URL API thật rồi redeploy.
Deploy backend Python riêng và cấu hình `CORS_ORIGINS` với domain frontend. Không dùng `*` trong production.
MVP hiện có
Dashboard responsive, menu mobile, danh sách câu hỏi và tìm kiếm/lọc.
Thêm/sửa/xóa câu hỏi; lưu tạm trong localStorage trình duyệt.
Trộn g0 (giữ nguyên), g1 (đảo câu), g2 (đảo phương án), g3 (đảo cả hai).
Đáp án phương án trắc nghiệm gắn theo ID phương án ổn định.
Xem trước mã đề, xuất ngân hàng/mã đề JSON.
Khung ma trận, kiểm định, xuất hồ sơ, thi trực tuyến và cài đặt.
Backend: health check, parser DOCX cơ bản, engine trộn và xuất Excel tổng quát.
Cảnh báo quan trọng
Đây là MVP để làm nền, không phải sản phẩm thi chính thức hoàn thiện.
Parser DOCX hiện chỉ trích xuất văn bản đoạn cơ bản. Không cam kết bảo toàn MathType/OMML, hình ảnh nổi, biểu đồ, bảng, header/footer hay bố cục Word.
Không có parser PDF/OCR trong phiên bản này.
Chưa có xuất DOCX/PDF chính thức. Giao diện không giả lập rằng các chức năng này đã hoàn thành.
Excel backend là bảng tổng quát, CHƯA ánh xạ mẫu đáp án 81 cột đã thống nhất.
localStorage không đồng bộ giữa thiết bị và không phải cơ sở dữ liệu dùng chung.
Chưa có đăng nhập, phân quyền, database, lưu bài online, chấm thi hoặc bảo vệ đáp án production.
Trước khi phát hành, cần kiểm thử bằng đề Word thực tế và rà soát mọi công thức, hình ảnh, đáp án.
Quy tắc trộn
g0: không đảo câu/phương án.
g1: đảo câu hỏi.
g2: đảo phương án trắc nghiệm.
g3: đảo câu và phương án.
Các tiểu ý đúng/sai giữ nguyên thứ tự trong MVP.
Sau khi trộn, luôn đối chiếu đề gốc và xác nhận đáp án bằng giáo viên.
Giai đoạn 1 — Nền tảng API
Backend hiện cung cấp:
`GET /api/health`: health check.
`GET /api/config`: thông tin tính năng không chứa bí mật.
`POST /api/exams/parse`: nhận DOCX (tối đa mặc định 20 MB), trả cảnh báo cần rà soát.
`POST /api/exams/mix`: endpoint trộn đề MVP.
`POST /api/exams/export-answer-key`: xuất bảng Excel tổng quát.
Chạy kiểm thử backend
Từ thư mục `backend`:
```bash
pip install -r requirements.txt
pytest -q
```
Biến môi trường backend
Sao chép `backend/.env.example` thành `.env` khi chạy local và đặt các giá trị thực tế trong Environment Variables của Render khi triển khai. Không commit `.env`.
Cấu hình frontend
Tạo `.env.local` ở thư mục gốc:
```env
VITE_API_BASE_URL=http://127.0.0.1:8000
```
Khởi động lại Vite sau khi thay đổi biến môi trường. Trong Vercel, đặt `VITE_API_BASE_URL` bằng URL backend production rồi redeploy.
Lưu ý: `GET /api/config` phản ánh đúng giới hạn hiện tại; database, xác thực, PDF/OCR và xuất DOCX/PDF chưa được triển khai ở giai đoạn 1.
