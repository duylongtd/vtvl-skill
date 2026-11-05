---
name: vtvl-kiem-tra
description: Kiểm tra logic và tính hợp lý của Bộ biểu vị trí việc làm (VTVL) .xlsx theo Hướng dẫn 07-HD/BTCTW và Công văn 05-CV/BCĐ Tỉnh ủy Hà Tĩnh. Dùng khi người dùng gửi bộ biểu VTVL (A1–D2, Dự báo) và muốn biết ô nào sai, sai logic gì, sửa thế nào. Chỉ đọc, không sửa file; kết quả xuất .docx theo từng biểu.
---

# vtvl-kiem-tra – Kiểm tra logic Bộ biểu VTVL

## Mục đích
App kê khai VTVL chỉ kiểm tra **định dạng** (mã có tồn tại, ô bắt buộc, danh mục). Skill này kiểm tra **logic và tính hợp lý** của số liệu theo văn bản. Toàn bộ quy tắc nằm trong `references/RULES.md` và được code hoá trong `scripts/kiem_tra.py`; mọi AI chạy skill này trên cùng một file sẽ ra cùng một kết quả.

## Quy trình bắt buộc (AI làm theo đúng thứ tự)

1. **Nhận file** bộ biểu `.xlsx` từ người dùng. Không nhận file đã chuyển sang .csv hay Google Sheets (mất công thức, mất trang ẩn).
2. **Hỏi đúng 3 thông số**, hỏi một lần, không hỏi thêm gì khác:
   - Tổng **biên chế được giao** của đơn vị (số).
   - Số **người có mặt tại 30/9/2026** (số).
   - Hồ sơ **đã được cơ quan đầu mối thẩm định chưa** (chưa / rồi). Mặc định: chưa.
   Loại đơn vị mặc định là `xa`; đổi sang `tinh` hoặc `su_nghiep` nếu người dùng nói rõ.
3. **Chạy script**, không tự đọc file bằng cách khác:
   ```
   python scripts/kiem_tra.py <file.xlsx> --bien-che <N> --so-nguoi <M> [--da-tham-dinh] [--loai-don-vi xa] --out <ten_ket_qua>
   ```
   Thiếu thông số thì vẫn chạy; script tự bỏ qua quy tắc cần thông số đó và ghi rõ trong báo cáo.
4. **Trả kết quả** cho người dùng:
   - Gửi file `<ten_ket_qua>.docx` (báo cáo theo từng biểu, đúng form Yêu cầu kiểm tra) và `.xlsx` (để lọc).
   - Trong chat: nêu tổng số BẮT BUỘC / CẦN / LƯU Ý, rồi liệt kê các lỗi **BẮT BUỘC** theo mẫu một dòng: `Biểu – Ô – sai gì – sửa thế nào`. Lỗi CẦN chỉ nêu số lượng theo biểu, chi tiết ở file.
   - Nói rõ: hồ sơ còn lỗi BẮT BUỘC sẽ bị trả lại (CV05 mục II).
5. **Không làm**:
   - Không sửa, không "sửa giúp" bất kỳ ô nào trong bộ biểu. Người kê khai tự sửa và chịu trách nhiệm.
   - Không tự thêm quy tắc ngoài `references/RULES.md`. Nếu người dùng hỏi về trường hợp chưa có quy tắc, trả lời "chưa có trong bộ quy tắc, cần hỏi cơ quan đầu mối".
   - Không suy diễn từ tên người, không đưa dữ liệu cá nhân trong file vào câu trả lời ngoài mã ô.
   - Không gửi file lên dịch vụ ngoài. Script chạy hoàn toàn trên máy.

## Đọc kết quả
| Mức | Ý nghĩa | Nguồn |
|---|---|---|
| BẮT BUỘC | Hồ sơ bị trả lại nếu còn | CV05 Phụ lục 2 mức "Bắt buộc" |
| CẦN | Phải rà soát, giải trình | CV05 Phụ lục 2 mức "Cần" |
| LƯU Ý | Dấu hiệu bất thường, có thể hợp lệ | Người kiểm tra tự xem |

Mỗi phát hiện có **mã quy tắc** (ví dụ `C3-09`). Tra mã trong `references/RULES.md` để thấy căn cứ văn bản nguyên văn.

## Giới hạn
- File do phần mềm khác tạo (không phải Excel lưu) có thể **không có giá trị cache công thức**; script tự tính dự phòng cho B1.G, C3.G, D1.D nhưng không tính được các cột cảnh báo của bộ biểu (B1.R, C3.M, D1.J, D2.J). Khi đó khuyên người dùng mở file bằng Excel, bấm Lưu, rồi gửi lại.
- Các nội dung không tự động hoá được (phiếu đo có thật không, sao chép từ hồ sơ mẫu, định mức chuyên ngành) nằm ở cuối báo cáo .docx để người kiểm tra tự đối chiếu.
- Hai điểm văn bản còn mâu thuẫn (phương pháp cho cấp phó; trạng thái C3 khi lập D2) được để trong `config/quy_tac.json`, xem mục "Điểm cần xác nhận" trong RULES.md.

## Cấu trúc
```
SKILL.md                 hướng dẫn này
scripts/kiem_tra.py      bộ kiểm tra (Python 3.9+, openpyxl, python-docx)
config/bieu.json         bản đồ sheet → cột, cột công thức (đối chiếu bộ biểu 23/10/2026)
config/quy_tac.json      ngưỡng, danh mục, từ khoá chức danh, điểm chưa chốt
references/RULES.md      bộ quy tắc đầy đủ, có nguồn
tests/                   fixture rút gọn + script tạo file test + kỳ vọng
```
Cài thư viện: `pip install openpyxl python-docx`
