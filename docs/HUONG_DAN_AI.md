# Hướng dẫn dùng skill `vtvl-kiem-tra` với Claude, ChatGPT, Gemini

Đầu vào: **một hoặc nhiều** file bộ biểu VTVL `.xlsx`. Đầu ra: báo cáo `.docx` theo từng biểu, chỉ rõ **ô nào sai – sai logic gì – sửa thế nào – căn cứ văn bản nào**. Công cụ **chỉ cảnh báo, không tự sửa file**; cán bộ tự sửa trên bộ biểu hoặc app rồi chạy lại cho đến khi hết lỗi BẮT BUỘC.

## Cách nhanh nhất (mọi AI có chạy Python)

Đính kèm file .xlsx → dán link `https://github.com/duylongtd/vtvl-skill` → gõ *"Dựa theo skill này hãy kiểm tra file, biên chế giao 25, có mặt 23, chưa thẩm định"*. README của repo là chỉ dẫn cho AI: nó sẽ `pip install` công cụ, chạy và trả báo cáo. Nếu AI tự đọc file mà không chạy công cụ, nhắc: *"Hãy làm đúng mục Chỉ dẫn cho AI trong README, chạy lệnh vtvl-kiem-tra, không tự suy luận."*

Các mục dưới đây dành cho cài đặt cố định (Project/Gem/Skill) để không phải dán link mỗi lần.

## Trước khi bắt đầu (áp dụng cho cả 3 AI)

1. Tải skill: vào https://github.com/duylongtd/vtvl-skill → **Code → Download ZIP** → giải nén. Hoặc `git clone https://github.com/duylongtd/vtvl-skill.git`.
2. Chuẩn bị 3 thông số của đơn vị, AI sẽ hỏi:
   - Tổng **biên chế được giao** (ví dụ 25)
   - Số **người có mặt tại 30/9/2026** (ví dụ 23)
   - Hồ sơ **đã được cơ quan đầu mối thẩm định chưa** (thường là *chưa*)
3. Bộ biểu phải là file `.xlsx` gốc, **mở bằng Excel và bấm Lưu** trước khi gửi (để có giá trị tính của công thức). Không chuyển sang Google Sheets, .csv hay .xls.
4. **Bảo mật** (Công văn 05-CV/BCĐ): bộ biểu chứa dữ liệu cá nhân (D1). Chỉ dùng tài khoản AI được cơ quan cho phép (gói Enterprise/Team có cam kết không dùng dữ liệu để huấn luyện), hoặc dùng cách **chạy trên máy** ở cuối tài liệu. Không dùng tài khoản AI miễn phí công cộng cho file có dữ liệu thật.

---

## A. Claude (claude.ai hoặc app Claude)

**Cài một lần**
1. Mở Claude → **Settings → Capabilities → Skills** → **Upload skill** → chọn file ZIP đã tải (hoặc nén lại thư mục `vtvl-skill` thành .zip).
2. Bật skill `vtvl-kiem-tra`.
   - Nếu dùng Claude Code: chép thư mục vào `.claude/skills/vtvl-kiem-tra/` trong project.

**Mỗi lần kiểm tra**
1. Mở chat mới, **kéo thả 1 hoặc nhiều file .xlsx** vào.
2. Gõ: `Kiểm tra bộ biểu VTVL theo skill vtvl-kiem-tra. Biên chế giao 25, có mặt 23, chưa thẩm định.`
   (Nhiều xã: ghi thông số cho từng file, ví dụ `xa_A.xlsx: 25/23; xa_B.xlsx: 18/18`.)
3. Claude chạy script trong môi trường của nó và trả về:
   - File `..._ket_qua.docx` cho từng bộ biểu (tải về gửi cán bộ xã).
   - Trong chat: số lỗi BẮT BUỘC / CẦN / LƯU Ý và danh sách lỗi BẮT BUỘC theo mẫu *Biểu – Ô – sai gì – sửa thế nào*.
4. Sửa trên bộ biểu, gửi lại file đã sửa, lặp lại đến khi BẮT BUỘC = 0.

---

## B. ChatGPT (chatgpt.com)

ChatGPT không có "skill", dùng **Project** (hoặc Custom GPT) để lưu sẵn hướng dẫn và script.

**Cài một lần**
1. Thanh bên trái → **Projects → New project**, đặt tên `VTVL kiem tra`.
2. Trong project → **Files** → tải lên các file sau từ thư mục skill:
   `SKILL.md`, `vtvl_kiem_tra/kiem_tra.py`, `vtvl_kiem_tra/config/bieu.json`, `vtvl_kiem_tra/config/quy_tac.json`, `references/RULES.md`.
3. **Instructions** của project, dán nguyên văn:
   ```
   Bạn là công cụ kiểm tra Bộ biểu VTVL. Làm đúng SKILL.md trong Files.
   Khi người dùng gửi file .xlsx: (1) hỏi đúng 3 thông số: biên chế được giao, số người có mặt 30/9/2026, đã thẩm định chưa;
   (2) dùng Python (Advanced Data Analysis): tạo thư mục vtvl_kiem_tra/config/, chép kiem_tra.py vào vtvl_kiem_tra/, bieu.json và quy_tac.json vào vtvl_kiem_tra/config/,
   pip install openpyxl python-docx nếu thiếu, rồi chạy:
   python vtvl_kiem_tra/kiem_tra.py <các file .xlsx> --bien-che N --so-nguoi M [--da-tham-dinh] --out ket_qua --in-bao-cao
   (3) đưa link tải các file .docx trong ket_qua/ và liệt kê lỗi BẮT BUỘC theo mẫu: Biểu – Ô – sai gì – sửa thế nào.
   Không sửa file gốc. Không tự thêm quy tắc ngoài RULES.md. Không nêu tên người trong file.
   ```
4. Cần bật **Advanced Data Analysis / Code Interpreter** (gói Plus, Team hoặc Enterprise).

**Mỗi lần kiểm tra**
1. Vào project → chat mới → kéo thả 1 hoặc nhiều file `.xlsx`.
2. Gõ: `Kiểm tra các bộ biểu này. Biên chế giao 25, có mặt 23, chưa thẩm định.`
3. ChatGPT chạy script và trả về link tải `.docx` cho từng file.
4. Nếu ChatGPT báo thiếu `python-docx`: gõ `Chạy lại với --json rồi tự trình bày kết quả theo từng biểu` (vẫn ra đủ nội dung, chỉ không có file .docx).

Lưu ý: file trong Project có thể cần tải lại khi script đổi phiên bản (xem `CHANGELOG.md` trên GitHub).

---

## C. Gemini (gemini.google.com)

Gemini chạy Python nhưng môi trường thường **không có `python-docx`** và không cho tải file kết quả về; `--in-bao-cao` in báo cáo ra màn hình để Gemini trình bày lại. Hai cách:

**Cách 1 – Gem (nhanh, kết quả hiện trong chat)**
1. Thanh bên trái → **Gems → New Gem**, đặt tên `VTVL kiem tra`.
2. Mục **Knowledge** (hoặc đính kèm file): tải `SKILL.md`, `vtvl_kiem_tra/kiem_tra.py`, `vtvl_kiem_tra/config/bieu.json`, `vtvl_kiem_tra/config/quy_tac.json`, `references/RULES.md`.
3. **Instructions**: dán nguyên đoạn hướng dẫn ở mục B (ChatGPT) nhưng thay dòng lệnh bằng:
   `python vtvl_kiem_tra/kiem_tra.py <file.xlsx> --bien-che N --so-nguoi M --in-bao-cao`
   và thêm: `Nếu không chạy được Python, không được tự "đoán" lỗi; nói rõ cần chạy bằng Colab (cách 2).`
4. Mỗi lần: mở Gem → đính kèm file `.xlsx` → gõ thông số. Gemini in kết quả JSON và trình bày theo từng biểu trong chat. Không có file .docx; copy nội dung vào Word nếu cần.

**Cách 2 – Google Colab (đầy đủ, có file .docx, chạy được nhiều file)**
1. Mở https://colab.research.google.com → **New notebook**.
2. Ô lệnh 1:
   ```
   !pip -q install git+https://github.com/duylongtd/vtvl-skill.git
   ```
3. Biểu tượng thư mục bên trái → **Upload** các file `.xlsx`.
4. Ô lệnh 2 (sửa tên file và thông số):
   ```
   !vtvl-kiem-tra "xa_A.xlsx" "xa_B.xlsx" --bien-che 25 --so-nguoi 23 --out ket_qua
   ```
5. Thư mục `ket_qua/` có `.docx`, `.xlsx`, `.json` cho từng file. Chuột phải → Download.
6. Muốn Gemini giải thích: mở Gemini trong Colab (nút ✨) và hỏi `Giải thích các lỗi BẮT BUỘC trong ket_qua/xa_A_ket_qua.json`.

---

## D. Chạy trên máy, không cần AI (an toàn nhất với dữ liệu thật)

```
pip install git+https://github.com/duylongtd/vtvl-skill.git
vtvl-kiem-tra "xa_A.xlsx" "xa_B.xlsx" --bien-che 25 --so-nguoi 23 --out ket_qua
```
Windows: cài Python từ python.org (tích *Add to PATH*), mở thư mục skill, gõ `cmd` vào thanh địa chỉ, dán lệnh trên. Kết quả trong `ket_qua\`. Sau đó có thể đưa **file kết quả** (không chứa dữ liệu cá nhân, chỉ có mã ô) cho bất kỳ AI nào để nhờ giải thích hoặc soạn công văn trả lời xã.

---

## Đọc kết quả

| Mức | Phải làm gì |
|---|---|
| **BẮT BUỘC** | Sửa hết trước khi gửi; còn lỗi là hồ sơ bị trả lại (CV05 mục II) |
| **CẦN** | Rà soát, giải trình vào cột Nguồn/giải trình của biểu |
| **LƯU Ý** | Xem lại, có thể hợp lệ |

Mỗi lỗi có **mã** (ví dụ `C3-09`). Tra mã trong `references/RULES.md` để thấy nguyên văn căn cứ. Phần cuối báo cáo liệt kê các nội dung **máy không kiểm tra được** (phiếu đo có thật không, sao chép hồ sơ mẫu…), người kiểm tra tự đối chiếu.

## Khi AI trả lời không giống nhau
Kết quả do script sinh ra là giống nhau trên mọi AI. Nếu hai AI đưa ra kết luận khác nhau, nghĩa là một bên **không chạy script** mà tự đọc file. Yêu cầu AI đó: `Chạy đúng lệnh python scripts/kiem_tra.py rồi chỉ trình bày kết quả, không tự suy luận thêm.`
