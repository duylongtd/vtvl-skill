# Prompt chuẩn để kiểm tra Bộ biểu VTVL bằng AI

Dùng chung cho Claude, ChatGPT, Gemini. Đính kèm file `.xlsx` trước, rồi dán nguyên prompt, chỉ sửa phần trong `[ ]`.

## Prompt đầy đủ (lần kiểm tra đầu tiên)

```
Dựa theo skill tại https://github.com/duylongtd/vtvl-skill, hãy kiểm tra logic Bộ biểu VTVL đính kèm.

Thông số:
- Đơn vị: [Đảng ủy xã ABC]
- Biên chế được giao: [25]
- Số người có mặt tại 30/9/2026: [23]
- Hồ sơ đã được cơ quan đầu mối thẩm định: [chưa]
- Loại đơn vị: [xã]

Yêu cầu:
1. Làm đúng mục "Chỉ dẫn cho AI" trong README của skill: cài công cụ bằng pip từ GitHub (hoặc dùng 3 file công cụ tôi đính kèm nếu không có mạng) và chạy lệnh vtvl-kiem-tra với các thông số trên, thêm --in-bao-cao. Không tự mở file Excel để suy luận lỗi.
2. Trả kết quả theo mẫu: dòng đầu "Tổng: BẮT BUỘC x · CẦN y · LƯU Ý z"; sau đó liệt kê MỌI lỗi BẮT BUỘC, mỗi lỗi một dòng "Biểu – Ô – sai gì – sửa thế nào – [mã quy tắc]"; lỗi CẦN chỉ nêu số lượng theo biểu; cuối cùng đưa file ket_qua .docx để tải (nếu không đưa file được thì dán nguyên báo cáo Markdown).
3. Không sửa file, không tạo bản "đã làm sạch", không thêm quy tắc ngoài references/RULES.md, không nêu tên người trong file.
4. Nếu không chạy được công cụ bằng cả hai cách, nói rõ lý do và ghi "CHẾ ĐỘ THỦ CÔNG" ở đầu kết quả thay vì âm thầm tự kiểm tra.
```

## Prompt ngắn (các lần kiểm tra lại, cùng một chat)

```
Tôi đã sửa theo kết quả lần trước, file mới đính kèm. Chạy lại vtvl-kiem-tra với cùng thông số, trả kết quả theo mẫu cũ và cho biết lỗi nào đã hết, lỗi nào còn.
```

## Prompt nhiều file một lần

```
Dựa theo skill tại https://github.com/duylongtd/vtvl-skill, kiểm tra các bộ biểu đính kèm. Thông số từng file:
- [xa_A.xlsx]: biên chế [25], có mặt [23], [chưa] thẩm định
- [xa_B.xlsx]: biên chế [18], có mặt [18], [chưa] thẩm định
Chạy vtvl-kiem-tra một lần cho tất cả file với --out ket_qua --in-bao-cao, rồi trả bảng tổng hợp (mỗi file một dòng: BẮT BUỘC / CẦN / LƯU Ý) và chi tiết lỗi BẮT BUỘC theo từng file. Các yêu cầu 1–4 như prompt chuẩn.
```

## Prompt nhờ giải thích kết quả (không cần gửi file bộ biểu, an toàn với dữ liệu cá nhân)

```
Đây là file kết quả ket_qua.md do công cụ vtvl-kiem-tra sinh ra. Với mỗi lỗi BẮT BUỘC, giải thích ngắn gọn cho cán bộ xã hiểu vì sao sai theo căn cứ ghi trong cột "Căn cứ", và soạn hướng dẫn sửa theo từng biểu. Không bịa thêm căn cứ, không thay đổi mã quy tắc.
```

## Khi AI trả lời sai cách

| AI làm gì | Gõ thêm |
|---|---|
| Tự mở Excel, không chạy công cụ | `Chưa đúng. Chạy lệnh vtvl-kiem-tra theo README rồi chỉ trình bày kết quả của nó.` |
| Báo không có mạng | Đính kèm thêm 3 file: `vtvl_kiem_tra/kiem_tra.py`, `vtvl_kiem_tra/config/bieu.json`, `vtvl_kiem_tra/config/quy_tac.json` và gõ `Dùng 3 file vừa gửi theo Cách B trong README.` |
| Báo thiếu python-docx | `Chạy lại với --in-bao-cao và dán nguyên báo cáo.` |
| Đề nghị "sửa giúp" file | `Không sửa file. Chỉ báo lỗi.` |

Lưu ý: đổi `[xã]` thành `[tỉnh]` hoặc `[sự nghiệp]` cho đơn vị cấp tỉnh, đơn vị sự nghiệp (ảnh hưởng ngoại lệ PP1 cho cấp phó HĐND xã).
