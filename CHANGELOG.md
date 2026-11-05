# Nhật ký phiên bản

## 0.2.0 – 05/11/2026
- Đóng gói thành package: `pip install git+https://github.com/duylongtd/vtvl-skill.git` → lệnh `vtvl-kiem-tra`. Code và config chuyển vào `vtvl_kiem_tra/`; `scripts/kiem_tra.py` giữ làm bọc tương thích.
- README mở đầu bằng "Chỉ dẫn cho AI": thang A (pip từ GitHub) → B (tải raw) → C (thủ công, có gắn nhãn); mẫu trình bày kết quả cố định.
- Thêm báo cáo `.md` và cờ `--in-bao-cao` để in báo cáo ra màn hình khi AI không tải được tệp.

## 0.1.1 – 05/11/2026
- Chạy nhiều file một lần; `--out` là thư mục, có `tong_hop.json`.
- Thêm `docs/HUONG_DAN_AI.md`: các bước dùng với Claude, ChatGPT, Gemini, Colab, và chạy trên máy.

## 0.1.0 – 05/11/2026
- Bản đầu. 70 quy tắc trong `references/RULES.md`, code hoá trong `scripts/kiem_tra.py`.
- Nguồn: HD 07-HD/BTCTW (qua tài liệu tập huấn và HD nhập liệu v1.1), CV 05-CV/BCĐ ngày 03/11/2026 (PL1, PL2), Yêu cầu kiểm tra dữ liệu bản 2, HD_DIEN_BIEU và QUY_TAC_MA trong bộ biểu 23/10/2026.
- Bản đồ cột đối chiếu với tệp `Bo_bieu_VTVL_30000_dong_ban_sua_de_nghi_UEB_v2_2026-10-23.xlsx`.
- Chưa có: Bộ phụ lục, Cẩm nang, Hướng dẫn chuẩn nhập dữ liệu (tải lên lỗi), sẽ bổ sung quy tắc ở bản 0.2.
- Điểm chưa chốt: phương pháp cấp phó (PP1 hay PP1+PP2); trạng thái C3 khi lập D2.
