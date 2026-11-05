# Prompt chuẩn để kiểm tra Bộ biểu VTVL bằng AI

Dùng chung cho Claude, ChatGPT, Gemini. Đính kèm file `.xlsx` trước, rồi dán nguyên prompt, chỉ sửa phần trong `[ ]`.

## Prompt chuẩn (lần kiểm tra đầu tiên)

```
Dựa theo skill tại https://github.com/duylongtd/vtvl-skill, hãy kiểm tra logic Bộ biểu VTVL đính kèm.
Biên chế được giao: [25]; số người có mặt tại 30/9/2026: [23]. (Không biết thì bỏ dòng này.)

1. Làm đúng mục "Chỉ dẫn cho AI" trong README của skill: cài công cụ bằng pip từ GitHub (hoặc dùng 3 file công cụ tôi đính kèm nếu không có mạng), chạy lệnh vtvl-kiem-tra với --in-bao-cao. Không tự mở file Excel để suy luận lỗi.
2. Trả kết quả theo mẫu: "Tổng: BẮT BUỘC x · CẦN y · LƯU Ý z"; liệt kê MỌI lỗi BẮT BUỘC, mỗi lỗi một dòng "Biểu – Ô – sai gì – sửa thế nào – [mã]"; lỗi CẦN chỉ nêu số lượng theo biểu; đưa file nhật ký _NhatKy.xlsx (không đưa được thì dán nguyên báo cáo gọn).
3. Không sửa file, không thêm quy tắc ngoài RULES.md, không nêu tên người trong file.
4. Không chạy được công cụ thì nói rõ và ghi "CHẾ ĐỘ THỦ CÔNG" ở đầu kết quả.
```

### Về thông số
Chỉ hai thông số đáng ghi, cả hai đều **tuỳ chọn**:

| Thông số | Dùng cho | Bỏ trống thì |
|---|---|---|
| Biên chế được giao | C3-03 (tổng người cần bố trí ≤ biên chế), D2-02 | Bỏ qua 2 quy tắc này, báo cáo ghi rõ |
| Số người có mặt 30/9/2026 | D1-01 (số dòng D1 = số người có mặt) | Bỏ qua D1-01 |

Phải nhập tay vì nếu lấy ngay trong file thì file tự chứng nhận cho nó (CV05 PL2-8 cấm chỉnh số liệu cho khớp biên chế, số người hiện có). Hai tuỳ chọn khác chỉ dùng khi khác mặc định: thêm `--da-tham-dinh` nếu hồ sơ **đã** được cơ quan đầu mối thẩm định; thêm `--loai-don-vi tinh` hoặc `su_nghiep` nếu không phải xã.

## Prompt ngắn (các lần kiểm tra lại, cùng một chat)

```
Tôi đã sửa theo kết quả lần trước, file mới đính kèm. Chạy lại vtvl-kiem-tra với cùng thông số và cùng tên tệp nhật ký để tăng version; trả kết quả theo mẫu cũ và cho biết lỗi nào đã hết, còn, mới (sheet Da_het, cột "So với version trước").
```

## Prompt nhiều file một lần

```
Dựa theo skill tại https://github.com/duylongtd/vtvl-skill, kiểm tra các bộ biểu đính kèm. Thông số từng file:
- [xa_A.xlsx]: biên chế [25], có mặt [23]
- [xa_B.xlsx]: biên chế [18], có mặt [18]
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
