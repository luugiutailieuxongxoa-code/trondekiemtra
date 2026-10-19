# Young Mix Pro — MVP

Bộ mã nguồn khởi đầu có giao diện tiếng Việt responsive cho máy tính, máy tính bảng và điện thoại.

## Chạy frontend
Yêu cầu Node.js 20+:
```bash
npm install
npm run dev
```

## Chạy backend (tùy chọn, cần để nhập DOCX)
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

## Đưa lên GitHub và Vercel
1. Giải nén ZIP, tạo repository GitHub và tải toàn bộ nội dung thư mục `young-mix-pro` lên.
2. Import repository vào Vercel.
3. Framework preset: Vite; Build command `npm run build`; Output directory `dist`.
4. Nếu có backend ở host riêng, thêm biến môi trường `VITE_API_BASE_URL` là URL API thật rồi redeploy.
5. Deploy backend Python riêng và cấu hình `CORS_ORIGINS` với domain frontend. Không dùng `*` trong production.

## MVP hiện có
- Dashboard responsive, menu mobile, danh sách câu hỏi và tìm kiếm/lọc.
- Thêm/sửa/xóa câu hỏi; lưu tạm trong localStorage trình duyệt.
- Trộn g0 (giữ nguyên), g1 (đảo câu), g2 (đảo phương án), g3 (đảo cả hai).
- Đáp án phương án trắc nghiệm gắn theo ID phương án ổn định.
- Xem trước mã đề, xuất ngân hàng/mã đề JSON.
- Khung ma trận, kiểm định, xuất hồ sơ, thi trực tuyến và cài đặt.
- Backend: health check, parser DOCX cơ bản, engine trộn và xuất Excel tổng quát.

## Cảnh báo quan trọng
Đây là MVP để làm nền, không phải sản phẩm thi chính thức hoàn thiện.
- Parser DOCX hiện chỉ trích xuất văn bản đoạn cơ bản. Không cam kết bảo toàn MathType/OMML, hình ảnh nổi, biểu đồ, bảng, header/footer hay bố cục Word.
- Không có parser PDF/OCR trong phiên bản này.
- Chưa có xuất DOCX/PDF chính thức. Giao diện không giả lập rằng các chức năng này đã hoàn thành.
- Excel backend là bảng tổng quát, CHƯA ánh xạ mẫu đáp án 81 cột đã thống nhất.
- localStorage không đồng bộ giữa thiết bị và không phải cơ sở dữ liệu dùng chung.
- Chưa có đăng nhập, phân quyền, database, lưu bài online, chấm thi hoặc bảo vệ đáp án production.
- Trước khi phát hành, cần kiểm thử bằng đề Word thực tế và rà soát mọi công thức, hình ảnh, đáp án.

## Quy tắc trộn
- g0: không đảo câu/phương án.
- g1: đảo câu hỏi.
- g2: đảo phương án trắc nghiệm.
- g3: đảo câu và phương án.
- Các tiểu ý đúng/sai giữ nguyên thứ tự trong MVP.
- Sau khi trộn, luôn đối chiếu đề gốc và xác nhận đáp án bằng giáo viên.

## Giai đoạn 1 — Nền tảng API

Backend hiện cung cấp:
- `GET /api/health`: health check.
- `GET /api/config`: thông tin tính năng không chứa bí mật.
- `POST /api/exams/parse`: nhận DOCX (tối đa mặc định 20 MB), trả cảnh báo cần rà soát.
- `POST /api/exams/mix`: endpoint trộn đề MVP.
- `POST /api/exams/export-answer-key`: xuất bảng Excel tổng quát.

### Chạy kiểm thử backend
Từ thư mục `backend`:
```bash
pip install -r requirements.txt
pytest -q
```

### Biến môi trường backend
Sao chép `backend/.env.example` thành `.env` khi chạy local và đặt các giá trị thực tế trong Environment Variables của Render khi triển khai. Không commit `.env`.

### Cấu hình frontend
Tạo `.env.local` ở thư mục gốc:
```env
VITE_API_BASE_URL=http://127.0.0.1:8000
```
Khởi động lại Vite sau khi thay đổi biến môi trường. Trong Vercel, đặt `VITE_API_BASE_URL` bằng URL backend production rồi redeploy.

Lưu ý: `GET /api/config` phản ánh đúng giới hạn hiện tại; database, xác thực, PDF/OCR và xuất DOCX/PDF chưa được triển khai ở giai đoạn 1.


## Giai đoạn 2 — Xuất DOCX giữ cấu trúc nguồn (bản nền tảng)

API mới: `POST /api/exams/export-docx` (multipart upload trường `file`; tùy chọn query `question_order` là mảng JSON ID câu, ví dụ `?question_order=<JSON-array>`). API xuất lại DOCX bằng cách sắp xếp các khối XML của câu hỏi trong tài liệu gốc, không dựng lại toàn bộ câu từ văn bản. Cách này giúp giữ tốt hơn định dạng run, công thức OMML, hình ảnh nội tuyến, bảng, header/footer và quan hệ media của DOCX nguồn.

Ví dụ: `POST /api/exams/export-docx?question_order=%5B%22Q-DOCX-2%22%2C%22Q-DOCX-1%22%5D` cùng file DOCX gốc. ID theo quy ước parser `Q-DOCX-{số câu gốc}`. Danh sách phải chứa đủ các câu được nhận diện, mỗi ID đúng một lần.

**Giới hạn cần biết:** đây là nền tảng xuất DOCX bảo toàn XML nguồn và đánh lại số câu trong trường hợp tiêu đề nằm trong một text run. API chưa xóa dấu đáp án đúng, chưa thay đầu trang theo mẫu riêng từng Sở/Trường, chưa bảo đảm nhận diện mọi kiểu tiêu đề câu (ví dụ tiêu đề nằm trong bảng hoặc chữ `Câu` bị chia thành nhiều run), và chưa xuất PDF. Giáo viên cần đối chiếu bản DOCX xuất ra với bản gốc trước khi dùng chính thức.

Kiểm thử backend: `cd backend && pytest -q`.

## Giai đoạn 3 — kiểm thử DOCX thực tế có hình/công thức

Bản này nâng cấp nhận diện cấu trúc DOCX theo từng phần và tạo ID duy nhất cho câu hỏi, ví dụ `Q-DOCX-PI-1`, `Q-DOCX-PII-1`. Việc này tránh trùng ID khi Phần I và Phần II đều có “Câu 1”.

- Backend giữ nguyên các khối XML trong DOCX nguồn khi xuất; không tái tạo công thức/hình từ văn bản trích xuất.
- Các khối câu hỏi chỉ được đổi thứ tự trong cùng phần; tiêu đề phần được giữ cố định.
- Đầu ra vẫn giữ đánh dấu đáp án nguồn; chưa xóa gạch chân/in đậm/màu đáp án đúng.
- Câu có phương án đảo thứ tự (`g2`, `g3`) chưa được xuất theo cách bảo toàn nguồn; giao diện sẽ chặn để tránh lệch đáp án.
- Nhận diện là sơ bộ: câu nằm trong hộp văn bản/SmartArt, công thức OLE/MathType, tiêu đề bị chia qua nhiều run, hoặc cấu trúc không theo “Câu N.” có thể cần xử lý riêng.
- PDF/OCR chưa được triển khai.

### Chạy kiểm thử backend

```bash
cd backend
pip install -r requirements.txt
pytest -q
```

Trước khi sử dụng đề thi chính thức, hãy mở DOCX xuất bằng Microsoft Word/LibreOffice và so sánh lần lượt công thức, hình, bảng, phân phần và đáp án với bản gốc.

## Giai đoạn 4 — nhập đề gốc, nhận diện phần và trộn trực tiếp

- Khi tải DOCX, giao diện thay bộ dữ liệu demo bằng câu hỏi nhận diện từ chính tệp vừa tải, hiển thị số câu, số phần, hình/đồ họa và bảng.
- Nguồn DOCX được giữ trong phiên trình duyệt để gọi endpoint `POST /api/exams/export-docx`; file xuất được tạo từ DOCX nguồn thay vì dựng lại từ văn bản xem trước.
- Với DOCX nguồn, hiện bật `g0` (giữ thứ tự) và `g1` (đảo câu trong từng phần). `g2`/`g3` bị khóa để không giả vờ đảo phương án khi chưa đồng bộ đáp án và nhãn phương án trong XML gốc.
- Cần mở DOCX xuất ra bằng Microsoft Word/LibreOffice và kiểm tra thủ công công thức, hình, bảng, tiêu đề phần, đáp án. Parser hiện nhận diện chủ yếu các tiêu đề dạng `PHẦN I` và `Câu 1.` nằm trong các đoạn văn cấp thân tài liệu; câu hỏi nằm trong hộp văn bản hoặc bảng phức tạp có thể chưa được nhận diện.
- PDF/OCR, đảo phương án A-D bảo toàn đáp án, làm sạch dấu đáp án đúng và xuất PDF chưa triển khai trong giai đoạn này.


## Giai đoạn 5 — Mẫu tiêu đề DOCX và làm sạch đáp án

- Endpoint `POST /api/exams/export-docx` nhận thêm query `header` dạng JSON với các trường `department`, `school`, `exam_title`, `grade`, `school_year`, `subject`, `duration`, `exam_code`.
- DOCX xuất ra có khung tiêu đề hai cột, dòng mã đề, dòng họ tên/SBD; khổ A4, lề trên/dưới/phải 2 cm, lề trái 2,5 cm và font tiêu đề Times New Roman.
- Xóa các marker kỹ thuật `<g1>`, `<g2>`, `<g3>` và marker đáp án nội bộ dạng `ĐA:...`.
- Gỡ các dấu định dạng thường dùng để đánh dấu đáp án đúng (đậm, nghiêng, gạch chân, màu chữ, tô sáng) trên các đoạn phương án A–D. Giữ lại nội dung phương án.
- Nếu phát hiện tiêu đề phần đáp án riêng (`ĐÁP ÁN`/`ANSWER KEY`), phần từ tiêu đề đó đến cuối tài liệu được bỏ khỏi đề xuất ra.
- Chưa xuất PDF; cần mở DOCX đầu ra kiểm tra trực quan trước khi sử dụng chính thức.

## Giai đoạn 6 — làm sạch đáp án câu trả lời ngắn và chân trang chính thức

- Trong phần có tiêu đề `PHẦN ... Câu trắc nghiệm trả lời ngắn`, xóa các dòng đáp án gợi ý dạng `A. 10`, `A. 25,5` nằm riêng thành đoạn. Quy tắc chỉ áp dụng trong phần trả lời ngắn; phương án `A.` của phần trắc nghiệm nhiều lựa chọn được giữ nguyên.
- Bổ sung chân trang Word: `ĐỀ THI CHÍNH THỨC • Trang {PAGE}`; số trang dùng trường PAGE của Word để tự cập nhật khi mở/in tài liệu.
- Cấu hình API hiển thị phiên bản `0.7.0` và các cờ tính năng chân trang/số trang/làm sạch đáp án ngắn.
- Kiểm thử backend: `13 passed`.

Lưu ý: quy tắc loại `A. ...` dựa vào tiêu đề phần có cụm “trả lời ngắn” và dòng đáp án đứng riêng. Cần kiểm tra đề thực tế nếu mẫu dùng tiêu đề khác hoặc đáp án được gõ cùng đoạn với nội dung câu hỏi.


## Giai đoạn 8 — quy tắc g0/g1/g2/g3 theo từng phần

- Bộ phân tích nhận marker `<g0>`, `<g1>`, `<g2>`, `<g3>` trước tiêu đề phần hoặc trong đoạn marker riêng. Marker được dùng làm chế độ mặc định cho phần kế tiếp. Phần IV được phân loại là tự luận.
- Giao diện Trộn đề hiển thị lựa chọn riêng cho từng phần đã nhận diện: g0 giữ nguyên, g1 đảo câu, g2 đảo phương án, g3 đảo cả câu và phương án. Giáo viên có thể ghi đè chế độ mặc định trước khi tạo mã đề.
- Giao diện có trường chỉnh sửa thông tin tiêu đề DOCX (Sở, trường, tên kỳ thi, môn, khối/lớp, năm học, thời gian).
- Khi xuất DOCX, backend đảo câu trong phạm vi từng phần và đảo các đoạn phương án A./B./C./D. nhận diện được; nhãn phương án được đánh lại A-D. Nội dung hình ảnh, công thức, bảng và đối tượng nhúng vẫn lấy từ DOCX nguồn.
- **Giới hạn:** chế độ g2/g3 hiện chỉ đảo được phương án nằm trong các đoạn Word riêng biệt có nhãn A./B./C./D. Các phương án nằm chung một đoạn, trong hộp văn bản, hoặc được dàn bằng bảng/cột cần xử lý riêng. Việc trộn phương án hiện chưa xuất bảng đáp án đã ánh xạ lại; giáo viên bắt buộc rà soát đáp án trước khi phát hành.
- Kiểm thử backend: `pytest -q` (15 bài). Frontend build chưa được xác nhận vì `npm install` hết thời gian trong môi trường đóng gói.


## Phase 9 — DOCX output fixes (2026-10-19)

- Normalize duplicate A/B/C/D labels in each question block in reading order, including two-column/tab-separated option lines.
- Bold only the visible A./B./C./D. and a)/b)/c)/d) markers where the labels are separate Word text runs; preserve formula/image/object XML.
- Set Word's `updateFields` setting so PAGE and NUMPAGES fields are refreshed when the document is opened in a Word-compatible editor. The final page count still depends on the editor's pagination engine.
- Backend tests: 16 passed.
