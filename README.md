# vtvl-kiem-tra — Kiểm tra logic Bộ biểu vị trí việc làm (VTVL)

> **Nếu bạn là AI (Claude, ChatGPT, Gemini…) được đưa link này cùng một hoặc nhiều file .xlsx:** làm đúng mục **"Chỉ dẫn cho AI"** ngay dưới đây. Không đọc file Excel để tự suy luận lỗi; kết quả phải do công cụ này sinh ra.

---

## Chỉ dẫn cho AI

Bạn đang được yêu cầu kiểm tra Bộ biểu VTVL theo Hướng dẫn 07-HD/BTCTW và Công văn 05-CV/BCĐ (Tỉnh ủy Hà Tĩnh). Bộ quy tắc đã được **code hoá**; nhiệm vụ của bạn là chạy công cụ và trình bày kết quả, không phải tự kiểm tra.

### Bước 1 – Hỏi đúng 3 thông số (hỏi một lần, trong một câu)
1. Tổng **biên chế được giao** của đơn vị (số).
2. Số **người có mặt tại 30/9/2026** (số).
3. Hồ sơ **đã được cơ quan đầu mối thẩm định chưa** (chưa / rồi). Mặc định: chưa.

Nếu người dùng đã ghi sẵn trong tin nhắn thì không hỏi lại. Nhiều file thì thông số theo từng file. Thiếu thông số vẫn chạy được; công cụ tự bỏ qua quy tắc cần thông số đó và ghi rõ trong báo cáo.

### Bước 2 – Cài và chạy công cụ (thử lần lượt, dừng ở cách đầu tiên thành công)

**Cách A – cài từ GitHub (ưu tiên):**
```bash
pip install -q git+https://github.com/duylongtd/vtvl-skill.git
vtvl-kiem-tra "<file1.xlsx>" "<file2.xlsx>" --bien-che <N> --so-nguoi <M> [--da-tham-dinh] --out ket_qua --in-bao-cao
```

**Cách B – nếu pip không ra được GitHub, tải 3 file raw rồi chạy:**
```bash
mkdir -p vtvl_kiem_tra/config
curl -sL https://raw.githubusercontent.com/duylongtd/vtvl-skill/main/vtvl_kiem_tra/kiem_tra.py -o vtvl_kiem_tra/kiem_tra.py
curl -sL https://raw.githubusercontent.com/duylongtd/vtvl-skill/main/vtvl_kiem_tra/config/bieu.json -o vtvl_kiem_tra/config/bieu.json
curl -sL https://raw.githubusercontent.com/duylongtd/vtvl-skill/main/vtvl_kiem_tra/config/quy_tac.json -o vtvl_kiem_tra/config/quy_tac.json
pip install -q openpyxl python-docx
python vtvl_kiem_tra/kiem_tra.py "<file.xlsx>" --bien-che <N> --so-nguoi <M> --out ket_qua --in-bao-cao
```
Nếu sandbox không có mạng: nói rõ với người dùng và đề nghị họ **đính kèm 3 file trên** (hoặc file ZIP của repo), rồi chạy như Cách B.

**Cách C – không chạy được Python (chế độ thủ công, chỉ khi A và B đều thất bại):**
Đọc `references/RULES.md`, tự kiểm tra bằng tay và **ghi rõ ở đầu kết quả**: *"CHẾ ĐỘ THỦ CÔNG – không bảo đảm đầy đủ và thống nhất như công cụ; cần chạy lại bằng công cụ trước khi gửi"*. Vẫn dùng đúng mã quy tắc, mức và mẫu trình bày ở Bước 3.

### Bước 3 – Trình bày kết quả (mẫu cố định)
- Dòng đầu: `Tổng: BẮT BUỘC x · CẦN y · LƯU Ý z` cho từng file. Nếu BẮT BUỘC > 0: *"Hồ sơ còn lỗi BẮT BUỘC sẽ bị trả lại (CV05 mục II)."*
- Liệt kê **mọi lỗi BẮT BUỘC**, mỗi lỗi một dòng: `Biểu – Ô – sai gì – sửa thế nào – [mã quy tắc]`.
- Lỗi CẦN: nêu số lượng theo biểu; chi tiết nằm trong báo cáo.
- Đưa file `ket_qua/<tên>_ket_qua.docx` (và `.xlsx`) cho người dùng tải; nếu không đưa file được thì dán nguyên báo cáo Markdown mà `--in-bao-cao` đã in.
- Kết thúc bằng: *"Sửa trên bộ biểu/app, gửi lại file để kiểm tra đến khi BẮT BUỘC = 0."*

### Không được làm
- Không sửa, không "sửa giúp", không tạo bản "đã làm sạch" của file. Người kê khai tự sửa và chịu trách nhiệm.
- Không thêm quy tắc ngoài `references/RULES.md`; gặp trường hợp chưa có quy tắc, trả lời "chưa có trong bộ quy tắc, cần hỏi cơ quan đầu mối".
- Không nêu tên người, số điện thoại hay dữ liệu cá nhân có trong file; chỉ dẫn mã ô.
- Không gửi file lên dịch vụ ngoài sandbox.

---

## Dành cho người dùng

**Cách dùng nhanh với AI:** mở Claude / ChatGPT / Gemini → đính kèm file .xlsx bộ biểu → dán link `https://github.com/duylongtd/vtvl-skill` → gõ *"Dựa theo skill này hãy kiểm tra file, biên chế giao 25, có mặt 23, chưa thẩm định"*. Prompt chuẩn để copy: **[docs/PROMPT.md](docs/PROMPT.md)**. Chi tiết từng AI và cách chạy không cần AI: **[docs/HUONG_DAN_AI.md](docs/HUONG_DAN_AI.md)**.

**Chạy trên máy (an toàn nhất với dữ liệu thật):**
```bash
pip install git+https://github.com/duylongtd/vtvl-skill.git
vtvl-kiem-tra xa_A.xlsx xa_B.xlsx --bien-che 25 --so-nguoi 23 --out ket_qua
```
Kết quả trong `ket_qua/`: `.docx` (báo cáo theo từng biểu), `.xlsx` (bảng lọc), `.md`, `.json`. Công cụ **không sửa** file gốc.

**Lưu ý bảo mật (CV05):** bộ biểu chứa dữ liệu cá nhân ở D1. Chỉ dùng tài khoản AI được cơ quan cho phép; với tài khoản công cộng, hãy chạy trên máy rồi chỉ đưa **file kết quả** (chỉ có mã ô) cho AI giải thích.

## Bộ quy tắc
`references/RULES.md`: khoảng 70 quy tắc, mỗi quy tắc có mã, mức (BẮT BUỘC / CẦN / LƯU Ý theo Phụ lục 2 của CV05) và căn cứ văn bản. Ngưỡng, danh mục ở `vtvl_kiem_tra/config/quy_tac.json`; bản đồ cột ở `vtvl_kiem_tra/config/bieu.json`. Đổi quy tắc thì sửa ở đó, không sửa code.

Hai điểm văn bản chưa thống nhất (xem cuối RULES.md): cấp phó dùng PP1 hay PP1+PP2; C3 "Tính thử" hay "Đã kiểm chứng" khi lập D2. Hiện chấp nhận cả hai.

## Test
```bash
python tests/chay_test.py
```
Khi Ban ban hành bộ biểu mới: `python tests/tao_bo_bieu_lite.py Bo_bieu_chuan.xlsx` rồi `python tests/tao_file_test.py tests/bo_bieu_lite.xlsx`.

## Cấu trúc
```
README.md                 chỉ dẫn cho AI + người dùng (file đầu tiên AI đọc khi nhận link)
SKILL.md                  định dạng skill của Claude
vtvl_kiem_tra/            package: kiem_tra.py + config/
scripts/kiem_tra.py       bọc tương thích để chạy tay
references/RULES.md       bộ quy tắc đầy đủ, có nguồn
docs/HUONG_DAN_AI.md      hướng dẫn từng bước cho Claude, ChatGPT, Gemini, Colab
tests/                    fixture rút gọn, file đúng/sai, kỳ vọng
```
Phiên bản: xem `CHANGELOG.md`. Báo cáo luôn ghi phiên bản công cụ.
