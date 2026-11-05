# Bộ quy tắc kiểm tra logic Bộ biểu VTVL (v0.1.0)

App kê khai VTVL chỉ kiểm tra **định dạng** (mã có tồn tại, ô bắt buộc có giá trị, danh mục hợp lệ). Bộ quy tắc này kiểm tra **logic và tính hợp lý** của số liệu theo các văn bản dưới đây. Mỗi quy tắc đều dẫn nguồn; không có quy tắc nào do người viết skill tự đặt ra.

## Nguồn

| Ký hiệu | Văn bản |
|---|---|
| **[HD07]** | Hướng dẫn số 07-HD/BTCTW ngày 19/10/2026 của Ban Tổ chức Trung ương |
| **[CV05]** | Công văn số 05-CV/BCĐ ngày 03/11/2026 của Ban Chỉ đạo VTVL, biên chế Tỉnh ủy Hà Tĩnh. PL1 = Phụ lục 1 (lỗi thường gặp), PL2 = Phụ lục 2 (bảng tự kiểm tra) |
| **[YC]** | Yêu cầu kiểm tra dữ liệu (bản 2) |
| **[HDNL]** | Hướng dẫn nhập liệu 09 biểu VTVL A1–D2 v1.1 (Desktop App) |
| **[TH]** | Slide Tập huấn Phase 1 – Cập nhật 2.0 |

## Mức kết quả

Dùng đúng thuật ngữ của Bảng tự kiểm tra [CV05-PL2]:

| Mức | Ý nghĩa |
|---|---|
| **BẮT BUỘC** | Vi phạm nội dung mức "Bắt buộc" của PL2 hoặc yêu cầu bắt buộc khác. "Hồ sơ còn lỗi thuộc mức Bắt buộc sẽ bị trả lại" [CV05, mục II] |
| **CẦN** | Nội dung mức "Cần" của PL2, hoặc cảnh báo phải rà soát, giải trình |
| **LƯU Ý** | Dấu hiệu bất thường máy phát hiện được nhưng có thể hợp lệ; người kiểm tra xem lại |

## Nguyên tắc bắt buộc của skill

1. **Chỉ đọc**, tuyệt đối không sửa file [YC].
2. Mỗi phát hiện ghi **Biểu – Ô (cột + dòng) – Giá trị hiện tại – Sai gì – Sửa thế nào – Căn cứ**. Không viết lan man.
3. Kết quả xuất **file .docx, nhóm theo từng biểu** [YC].
4. Không đưa dữ liệu cá nhân, tài liệu nội bộ lên công cụ AI công cộng [CV05, Lưu ý mục I]. Script chạy hoàn toàn trên máy, không gửi dữ liệu đi đâu.

## Thông số người kiểm tra nhập trước khi chạy

| Tham số | Dùng cho quy tắc | Bắt buộc? |
|---|---|---|
| `bien_che_giao`: tổng biên chế được giao | C3-04 | Có (thiếu thì bỏ qua C3-04 và ghi rõ trong báo cáo) |
| `so_nguoi_co_mat`: số người có mặt tại 30/9/2026 | D1-01 | Có (thiếu thì bỏ qua D1-01) |
| `loai_don_vi`: `xa` / `tinh` / `su_nghiep` | C3-05, C3-11 | Mặc định `xa` |
| `da_tham_dinh`: hồ sơ đã được cơ quan đầu mối thẩm định chưa | TT-01..03 | Mặc định `khong` |

---

## T – Toàn tệp, cấu trúc mẫu biểu

| Mã | Phát hiện | Mức | Sửa | Căn cứ |
|---|---|---|---|---|
| T-01 | Thiếu trang tính so với bộ biểu chuẩn, hoặc có **trang ẩn** ngoài 4 trang ẩn theo thiết kế (13. K_KiemTra, 14. T_TienDo, 15. S_NhatKy, DU_BAO_5_NAM) [BB] | BẮT BUỘC | Nhận bộ biểu sạch, chép dữ liệu bằng Dán giá trị; quét virus nếu có trang ẩn lạ | CV05 mục I "Về mẫu biểu"; PL1-7; PL2-1 |
| T-02 | **Tiêu đề cột bị đổi** so với mẫu | BẮT BUỘC | Trả lại tiêu đề đúng mẫu hoặc chép sang bộ biểu sạch | CV05 "Về mẫu biểu"; PL2-1 |
| T-03 | Công thức có **liên kết đến tệp khác** (ký hiệu `[1]`, đường dẫn, External Link) | BẮT BUỘC | Ngắt liên kết, chép bằng Dán giá trị | CV05 "Về mẫu biểu"; PL1-7; PL2-1; YC |
| T-04 | **Cột tự động bị gõ đè**: ô thuộc cột công thức chỉ còn giá trị tĩnh | BẮT BUỘC | Khôi phục công thức từ bộ biểu sạch | CV05 "Về mẫu biểu"; PL2-2; YC; HDNL 15.1 |
| T-05 | **Gõ công thức vào cột nhập liệu** (ô cột ✎/☰ chứa công thức) | BẮT BUỘC | Thay bằng giá trị (Dán giá trị) | CV05 "Về mẫu biểu"; PL2-2 |
| T-06 | Công thức trong một cột tự động **khác mẫu** so với các dòng còn lại | BẮT BUỘC | Chép lại công thức đúng từ bộ biểu sạch | CV05 PL2-2 (không gõ đè cột tự động) |
| T-07 | Trang Kiểm tra hoặc ô bất kỳ có lỗi **#REF!, #VALUE!** (và #N/A, #DIV/0!, #NAME?) | BẮT BUỘC | Tìm ô nguồn gây lỗi, thường do xoá/chèn dòng hoặc gõ đè | CV05 PL2-2 |
| T-08 | Dấu hiệu **chèn, xoá dòng** trong vùng mẫu (số dòng khung hoặc công thức tham chiếu lệch so với mẫu) | BẮT BUỘC | Chép sang bộ biểu sạch | CV05 "Về mẫu biểu"; PL1-7 |
| T-09 | Số hiệu văn bản **tăng dần đều theo dòng** (vd. 12, 13, 14…), hoặc năm tăng dần: dấu hiệu kéo chuột sao chép | BẮT BUỘC | Đối chiếu bản gốc, ghi lại đúng số hiệu, ngày | CV05 "Về căn cứ, mô tả, mã"; PL1-8; PL2-11 |
| T-10 | Cùng một văn bản (cùng số hiệu) nhưng **ngày ban hành khác nhau** giữa các ô trong tệp | BẮT BUỘC | Thống nhất theo bản gốc toàn tệp | CV05 PL2-11 |
| T-11 | Chuẩn hoá: Unicode tổ hợp, khoảng trắng thừa, ký tự ẩn, số lưu dạng chữ, giá trị danh mục không khớp nguyên văn | LƯU Ý | Gõ lại ô theo Unicode dựng sẵn; nhập số không kèm chữ; chọn đúng giá trị danh mục | HDNL 14.1 (gõ đúng nguyên văn, cột số không kèm chữ) |
| T-12 | Trang **S_NhatKy** trống trong khi tệp là bản gửi lại / có thay đổi số liệu | CẦN | Ghi nhật ký mọi lần sửa số liệu, có căn cứ | CV05 "Về kết quả tính toán"; PL1-10; PL2-8 |

## TT – Trạng thái (khi hồ sơ chưa thẩm định)

| Mã | Phát hiện | Mức | Sửa | Căn cứ |
|---|---|---|---|---|
| TT-01 | C3 cột L = "Đã kiểm chứng" khi hồ sơ **chưa thẩm định** | BẮT BUỘC | Để "Tính thử" cho đến khi được thẩm định | CV05 "Về trạng thái"; PL1-6; PL2-9 |
| TT-02 | Dự báo cột F = "Đã kiểm chứng" khi chưa thẩm định | BẮT BUỘC | Để "Chưa kiểm chứng" | như trên |
| TT-03 | D2 cột I = "Đã hoàn thành ở đơn vị" / "Đã thẩm định" / "Đã được quyết định" khi chưa thẩm định | BẮT BUỘC | Để "Đang xây dựng" | như trên |
| TT-04 | B1/B2/C1/D1: trạng thái "Đã kiểm chứng"/"Đã xác nhận" nhưng **cột nguồn trống hoặc chung chung** (không có số, ngày, người/cơ quan xác nhận) | BẮT BUỘC | Chỉ chọn khi đã có đối chiếu, xác nhận thật; ghi đủ nguồn | CV05 "Về trạng thái", "Về kiểm chứng" |

> Ghi chú: theo [HDNL], app chỉ chuyển J của C3 sang Dự báo/D2 khi C3.L = "Đã kiểm chứng". [CV05] yêu cầu để "Tính thử" đến khi thẩm định. Hai yêu cầu này mâu thuẫn ở bước lập Dự báo/D2; xem mục "Điểm cần xác nhận".

## A1 – Danh mục nhiệm vụ

| Mã | Phát hiện | Mức | Sửa | Căn cứ |
|---|---|---|---|---|
| A1-01 | Cột D (Căn cứ) không có **loại văn bản + số hiệu** | BẮT BUỘC | Ghi đủ số, ngày, cơ quan ban hành | YC; CV05 PL1-8, PL2-11 |
| A1-02 | Cột D không ghi **điều, khoản (điểm)** | BẮT BUỘC | Bổ sung điều, khoản, điểm | YC; CV05 4.2(1) "căn cứ đến điều, khoản"; PL1-8 |
| A1-03 | Cột G (Nguồn) chung chung, còn "…", "theo quy định", không dẫn văn bản cụ thể | BẮT BUỘC | Ghi nguồn cụ thể | YC; CV05 PL1-8 |
| A1-04 | Nhiệm vụ **không có công việc nào** ở A2 và không ghi lý do "không phát sinh" | CẦN | Bổ sung công việc ở A2 hoặc ghi "không phát sinh" kèm lý do | CV05 PL2-12; CV05 4.2(1) |
| A1-05 | Thiếu căn cứ, kỳ, trạng thái hoặc nguồn | CẦN | Nhập đủ | CV05 PL2-12 |

## A2 – Công việc và sản phẩm

| Mã | Phát hiện | Mức | Sửa | Căn cứ |
|---|---|---|---|---|
| A2-01 | Cột H (Căn cứ, nguồn) không có văn bản + số hiệu, hoặc không có **điều, khoản, điểm** | BẮT BUỘC | Ghi rõ điều, khoản, điểm của văn bản nào | YC; CV05 PL1-8 |
| A2-02 | **Tên sản phẩm (E) trùng** giữa các dòng | BẮT BUỘC | Đặt tên sản phẩm cụ thể, riêng cho từng công việc; nếu cùng sản phẩm thì mỗi vị trí chỉ tính phần việc mình trực tiếp làm | YC; CV05 PL1-8, "Về sản lượng" |
| A2-03 | Cùng một sản phẩm gán cho **nhiều vị trí** (C khác nhau) | BẮT BUỘC | Mỗi vị trí chỉ tính phần việc mình trực tiếp làm, không tính trùng | CV05 "Về sản lượng"; PL1-8 |
| A2-04 | Mã nhiệm vụ (B) không có ở A1, hoặc mã VTVL (C) không có ở C1 | BẮT BUỘC | Sửa mã theo A1, C1 | CV05 PL1-8, PL2-12; YC (liên kết theo mã) |
| A2-05 | **Trùng mã công việc** hoặc lặp nguyên dòng | CẦN | Xoá dòng lặp, sửa mã | CV05 PL2-13 |
| A2-06 | Tên công việc **cụt** (quá ngắn, chỉ 1–2 từ) | CẦN | Ghi tên công việc cụ thể, đo được | CV05 PL1-8, PL2-13 |
| A2-07 | Sản phẩm (E), đơn vị tính (F), chất lượng (G) **ghi một kiểu** cho nhiều dòng (cùng giá trị lặp lại) | CẦN | Ghi cụ thể cho từng công việc | CV05 PL1-8, PL2-13 |
| A2-08 | Công việc gắn vị trí lãnh đạo nhưng **cùng sản phẩm** lại xuất hiện ở vị trí chuyên môn | BẮT BUỘC | Phần việc lãnh đạo trực tiếp làm chỉ tính một lần ở vị trí lãnh đạo | CV05 "Về chức danh lãnh đạo"; PL2-10 |

## C1 – Danh mục vị trí việc làm / C2 – Khung năng lực

| Mã | Phát hiện | Mức | Sửa | Căn cứ |
|---|---|---|---|---|
| C1-01 | Cột D dẫn **quyết định chức năng, nhiệm vụ** thay vì quyết định phê duyệt VTVL | CẦN | Dẫn quyết định phê duyệt danh mục VTVL | CV05 PL1-9, PL2-14 |
| C1-02 | Chế độ "Đặc thù" nhưng thiếu căn cứ quỹ giờ (F) | BẮT BUỘC | Bổ sung văn bản quy định chế độ, giờ riêng | HDNL 4.4; CV05 "Về định mức" |
| C2-01 | Mã VTVL ở C2 **không có ở C1** | BẮT BUỘC | Dán đúng mã theo C1 | CV05 "Về căn cứ, mô tả, mã"; PL1-9 |
| C2-02 | VTVL ở C1 **không có dòng C2** (thiếu), hoặc C2 thừa vị trí | CẦN | Đủ vị trí ở cả hai biểu | CV05 PL1-9, PL2-14 |
| C2-03 | Thứ tự mã C2 lệch so với C1 (lệch dòng) | CẦN | Sắp cùng thứ tự với C1 | CV05 PL1-9, PL2-14 |
| C2-04 | Cột Chuyên môn (D) ghi **bằng cấp** (Đại học, Cử nhân, Thạc sĩ, trình độ…) thay vì năng lực | CẦN | Ghi năng lực và mức yêu cầu | CV05 PL1-9, PL2-14 |

## B1 – Khối lượng và thời gian

| Mã | Phát hiện | Mức | Sửa | Căn cứ |
|---|---|---|---|---|
| B1-01 | Giờ tổng (G) **≥ 1.000 giờ** | CẦN | Rà lại sản lượng, thời gian; giải trình | YC; CV05 "Về sản lượng"; PL1-3; PL2-7 (mức đề nghị của BCĐ) |
| B1-02 | Phút/SP (D) là **bội số của 480** (= số ngày × 8 giờ) | BẮT BUỘC | Ghi thời gian người thực sự làm từng phần việc, không nhân ngày × 8 giờ, không tính thời gian chờ | CV05 "Về thời gian"; PL1-1; PL2-5 |
| B1-03 | Cột I (Nguồn thời gian) **không có số, ngày, người/cơ quan xác nhận** (chỉ ghi "phiếu đo", "ước lượng"…) | BẮT BUỘC | Ghi số phiếu đo/phiếu khảo sát/văn bản định mức, ngày, người xác nhận | CV05 "Về thời gian"; PL1-1; PL2-5 |
| B1-04 | Cột L (Mã SP danh mục) có giá trị nhưng **không thuộc danh mục dùng chung** đã được cơ quan chuyên ngành xác nhận (danh mục tự lập) | BẮT BUỘC | Xoá danh mục tự lập và cột L; để trống L | CV05 "Về định mức"; PL1-2; PL2-6 |
| B1-05 | Cột S (Giải trình khác chuẩn) không bắt đầu bằng "Đã xác nhận:" khi đã dùng phút khác chuẩn | CẦN | Ghi đúng mẫu "Đã xác nhận: …" | HDNL 6.3 |
| B1-06 | Số SP (C) **= 0 hoặc trống mà R = ĐẠT**, không có giờ bổ sung có căn cứ | BẮT BUỘC | Thống kê sản lượng từ sổ, phần mềm; chỉ ghi 0 khi đã xác minh không phát sinh | CV05 "Về sản lượng"; PL1-3; PL2-7 |
| B1-07 | Nguồn sản lượng (H) chung chung, không chỉ ra sổ/phần mềm, kỳ | CẦN | Ghi nguồn cụ thể (tên sổ, phần mềm, kỳ 01/10/2025–30/9/2026) | CV05 "Về sản lượng"; PL1-3 |
| B1-08 | Giờ bổ sung (E) **gán cố định**: cùng một giá trị lặp ở nhiều dòng | CẦN | Chỉ nhập khi có căn cứ, đã xác nhận | CV05 "Về thời gian"; PL1-3 |
| B1-09 | Công việc tham mưu, nghiên cứu, thẩm định, kiểm tra, giám sát **được điền phút** (nên áp dụng PP3) | CẦN | Việc khó chuẩn hoá thời gian xác định theo PP3, không bắt buộc quy ra phút | CV05 "Về thời gian"; PL1-12; PL2-19 |
| B1-10 | Phút/SP có dấu hiệu **chỉnh để tránh số tròn** (vd. 58, 118, 238 = số tròn trừ vài phút), lặp ở nhiều dòng | LƯU Ý | Giữ đúng số đo, không chỉnh phút | CV05 "Về thời gian"; PL1-12 |
| B1-11 | Tổng giờ B1 của một VTVL **> số người hiện có (D1) × 1.904** | CẦN | Rà lại; có thể tính trùng giữa các vị trí | CV05 "Về sản lượng"; PL2-15 |
| B1-12 | Tổng giờ B1 của một VTVL **trùng đúng số người × 1.904** (hoặc × quỹ giờ C1) | BẮT BUỘC | Không chỉnh phút, sản lượng để tổng giờ khớp số người; giữ kết quả lẻ | CV05 "Về kết quả tính toán"; PL1-12; PL2-19 |
| B1-13 | Tổng giờ B1 ÷ 1.904 chênh lớn so với số người D1 của VTVL | CẦN | Rà lại, giải trình | CV05 PL2-15 |

## B2 – Tác động thay đổi

| Mã | Phát hiện | Mức | Sửa | Căn cứ |
|---|---|---|---|---|
| B2-01 | B2 chi tiết **bỏ trống** và không có danh sách nội dung chưa đủ căn cứ | CẦN | Kê B2 nếu có số liệu; chưa có thì ghi vào danh sách nội dung chưa đủ căn cứ | CV05 PL1-10; PL2-16 |
| B2-02 | SP áp dụng (E) **> Số SP** của công việc ở B1 | BẮT BUỘC | Sửa E không vượt B1 | HDNL 7.4 |
| B2-03 | Trạng thái (M) = "Ước lượng tham khảo" mà vẫn được tính (N hoặc O ≠ 0) | BẮT BUỘC | Ước lượng tham khảo không được tính | HDNL 7.4 |

## C3 – Nhu cầu nhân lực

| Mã | Phát hiện | Mức | Sửa | Căn cứ |
|---|---|---|---|---|
| C3-01 | S = Σ G (dòng B = PP2) + Σ H (dòng B = PP1 hoặc PP1+PP2) **< Σ J** | CẦN | J vượt số tính được theo phương pháp; rà lại | YC |
| C3-02 | S **> 1,2 × Σ J** (chênh quá **20%**) | CẦN | Rà lại kết luận J, giải trình | YC (bản 2: ngưỡng 20%) |
| C3-03 | Σ J **> biên chế được giao** | BẮT BUỘC | Rà lại; nhu cầu tăng ghi ở phần giải trình, chỉ đề nghị trên số liệu đã kiểm chứng | YC; CV05 "Về kiểm chứng" |
| C3-04 | VTVL **đứng đầu** (Giám đốc, Chánh Văn phòng, Trưởng phòng, Trưởng ban, Chi cục trưởng, Hiệu trưởng…) không dùng **PP1** | BẮT BUỘC | Chọn PP1 | YC; CV05 "Về chức danh lãnh đạo"; PL1-5; PL2-10 |
| C3-05 | VTVL **cấp phó** dùng phương pháp không đúng (xem "Điểm cần xác nhận" số 1) | BẮT BUỘC | Theo cấu hình đã thống nhất | YC; CV05 PL1-5 |
| C3-06 | VTVL **chuyên môn** (không giữ chức danh lãnh đạo, quản lý) không dùng **PP2** | CẦN | Chọn PP2 (trừ việc tham mưu khó chuẩn hoá đã chuyển PP3 có hồ sơ) | YC; CV05 PL1-12 |
| C3-07 | Dùng **PP3** hoặc tổ hợp có PP3 | CẦN | Cảnh báo; phải có hồ sơ PP3 (danh mục sản phẩm, phân tích phần việc, phương án phân công, đối chiếu nhiệm vụ tương tự, trình thẩm định) | YC; CV05 "Về thời gian" |
| C3-08 | Vị trí chuyên môn PP2 có **J > G làm tròn lên** (làm tròn tăng từng dòng) | BẮT BUỘC | Cột J theo kết quả quy đổi, không làm tròn tăng từng dòng | CV05 PL1-5 |
| C3-09 | Cột K (Nguồn/giải trình) **giống hệt nhau** ở nhiều vị trí | BẮT BUỘC | Viết giải trình riêng cho từng vị trí bằng số liệu của chính vị trí đó | CV05 "Về chức danh lãnh đạo"; PL1-5; PL2-18 |
| C3-10 | Vị trí lãnh đạo: cột K không dẫn **văn bản quy định số lượng** (số, ngày, điều, khoản) | BẮT BUỘC | Ghi văn bản quy định số lượng, quyết định của cấp có thẩm quyền, phạm vi phụ trách từng cấp phó | CV05 "Về chức danh lãnh đạo"; PL1-5 |
| C3-11 | Cấp phó: H **= số người đang bố trí** (D1) mà K không nêu phạm vi phụ trách | CẦN | Số cấp phó là số thực cần theo nhiệm vụ, không lấy số đang bố trí, không mặc nhiên bằng mức tối đa | CV05 "Về chức danh lãnh đạo"; PL2-10 |
| C3-12 | J **bằng đúng số người hiện có** (D1) ở nhiều vị trí trong khi G lệch rõ | CẦN | Không điều chỉnh để khớp số người hiện có; giữ kết quả quy đổi, giải trình | CV05 "Về kết quả tính toán"; PL2-8 |
| C3-13 | Còn cảnh báo ở cột M (≠ ĐẠT) mà K không giải trình | BẮT BUỘC | Xử lý hết cảnh báo hoặc giải trình cụ thể | CV05 mục II; PL2-18 |

## D1 – Nhân lực hiện có

| Mã | Phát hiện | Mức | Sửa | Căn cứ |
|---|---|---|---|---|
| D1-01 | Số dòng (người) **≠ số người có mặt** tại 30/9/2026 | BẮT BUỘC | Mỗi người có mặt là 01 dòng | YC; CV05 "Về nhân lực"; PL2-3 |
| D1-02 | Có dòng **"chưa bố trí người"**, mã nguồn lực là mã vị trí, tên trống/"dự kiến" (kê vị trí thay cho người) | BẮT BUỘC | Không kê vị trí thay người, không thêm dòng cho vị trí chưa có người | CV05 "Về nhân lực"; PL1-4; PL2-3 |
| D1-03 | Một người xuất hiện ở **nhiều dòng** (trùng mã nguồn lực) | BẮT BUỘC | Mỗi người 01 dòng; kiêm nhiệm ghi ở cột F | CV05 "Về nhân lực"; HDNL 11.5 |
| D1-04 | Giờ cho VTVL (E) **lấy bằng công thức**, chia đều, hoặc ghi đồng loạt cùng một số ở nhiều dòng | BẮT BUỘC | Ghi theo văn bản phân công thực tế | CV05 "Về nhân lực"; PL1-4; PL2-4 |
| D1-05 | E + F **> quỹ giờ (D)** ("Vượt quỹ giờ") mà chưa xử lý/giải trình | BẮT BUỘC | Giảm E/F cho đúng thực tế | HDNL 11.4; CV05 PL2-18 |
| D1-06 | Có giờ việc khác (F > 0) nhưng không nêu rõ kiêm nhiệm gì | CẦN | Ghi rõ nội dung kiêm nhiệm | CV05 "Về nhân lực"; PL2-4 |
| D1-07 | Thiếu nơi tính biên chế (G) hoặc nguồn (H) | BẮT BUỘC | Ghi đúng hồ sơ | CV05 PL2-4 |
| D1-08 | (Đối chiếu nhiều tệp) cùng một người có ở D1 của **khối Đảng và khối chính quyền** cùng cấp | BẮT BUỘC | Đối chiếu để mỗi người chỉ tính một lần | CV05 "Về nhân lực"; PL2-4 |

## DB – Dự báo 2027–2031

| Mã | Phát hiện | Mức | Sửa | Căn cứ |
|---|---|---|---|---|
| DB-01 | VTVL ở C3 **thiếu năm** (không đủ 5 dòng 2027–2031) | CẦN | Mỗi cặp VTVL – năm một dòng | HDNL 12.5 |
| DB-02 | Biến động (C) trống | BẮT BUỘC | Nhập 0 nếu không đổi, kèm giải trình | HDNL 12.4 |

## D2 – Phương án biên chế

| Mã | Phát hiện | Mức | Sửa | Căn cứ |
|---|---|---|---|---|
| D2-01 | Cột C (tự lấy từ Dự báo) **nhập tay đè công thức** | BẮT BUỘC | Lập D2 sau khi có C3, dự báo; khôi phục công thức | CV05 PL1-6 |
| D2-02 | Biên chế được giao (E) **chia vào từng vị trí** / lặp tổng biên chế vào từng dòng | BẮT BUỘC | Biên chế giao tổng ghi ở cấp giao, không lặp vào từng dòng | CV05 PL1-6; HDNL 13.6 |
| D2-03 | Đề nghị điều chỉnh tăng (G > 0) khi C3 còn **"Tính thử"** / số liệu chưa kiểm chứng | BẮT BUỘC | Chỉ đề nghị tăng trên số liệu đã kiểm chứng; ưu tiên bố trí lại, kiêm nhiệm, điều chuyển | CV05 "Về kiểm chứng"; PL1-13; PL2-20 |
| D2-04 | G **= C − D** ở mọi dòng (đề nghị mặc nhiên theo chênh lệch) | CẦN | Chỉ đề nghị sau soát xét, giải trình, thẩm định | HDNL 13.6 |
| D2-05 | **Hợp đồng lao động** (D1 cột C) được tính vào số bố trí/biên chế giao | BẮT BUỘC | Hợp đồng lao động ghi riêng, không tính vào biên chế giao | CV05 PL1-13 |

## LB – Liên biểu (chuỗi nhiệm vụ → … → phương án)

Theo [YC]: "Kiểm tra liên kết đầy đủ dữ liệu nhiệm vụ → công việc → sản phẩm → khối lượng → VTVL → nhu cầu → nguồn lực → phương án; dữ liệu mỗi trường chỉ có một nơi nhập chính, biểu sau tự liên kết theo mã."

| Mã | Phát hiện | Mức |
|---|---|---|
| LB-01 | Công việc A2 **không có dòng B1** | CẦN |
| LB-02 | VTVL C1 **không có công việc** ở A2 | CẦN |
| LB-03 | VTVL C1 **không có dòng C3** | CẦN |
| LB-04 | VTVL C3 có người cần bố trí nhưng **không có người ở D1** và không giải trình | CẦN |
| LB-05 | VTVL C3 **không có dòng Dự báo / D2** | CẦN |
| LB-06 | Cột tự lấy (B1.B, B2.B, C3.F, D1.D, Dự báo.G, D2.C/K/L) **không khớp** giá trị ở biểu nguồn (bị nhập tay) | BẮT BUỘC |

---

## Không tự động hoá được (đưa vào checklist cho người kiểm tra)

- Phiếu đo, phiếu khảo sát ghi ở B1 cột I **có tồn tại thật** và đúng nội dung (BCĐ kiểm tra chọn mẫu) [CV05 4.1].
- Số liệu, giải trình **sao chép từ hồ sơ mẫu** của đơn vị khác [CV05 "Về sử dụng hồ sơ mẫu"; PL2-20]. (Skill chỉ phát hiện được khi có tệp mẫu để so.)
- Đơn vị sự nghiệp có định mức chuyên ngành: phép tính PP1 đúng theo văn bản ngành [CV05; PL1-11; PL2-17].
- Người đứng đầu đã ký xác nhận từng biểu và Bảng tự kiểm tra [CV05 mục II].

## Điểm cần xác nhận trước khi chốt v1.0

1. **Phương pháp cho cấp phó:** [YC] ghi "PP1+PP2 chỉ áp dụng cho các VTVL cấp phó", nhưng [CV05] (mục I "Về chức danh lãnh đạo, quản lý" và PL1-5) ghi lãnh đạo, quản lý **chọn PP1**, liệt kê "cấp phó chọn PP2 hoặc PP1+PP2" là **lỗi**. Skill để tham số `cap_pho_phuong_phap` (mặc định chưa chốt).
2. **Trạng thái C3:** [CV05] yêu cầu để "Tính thử" đến khi thẩm định, nhưng app ([HDNL] 10.5, 12.2) chỉ chuyển J sang Dự báo/D2 khi C3.L = "Đã kiểm chứng". Cần thống nhất đơn vị được chọn "Đã kiểm chứng" ở bước nào.
3. **Ngưỡng C3-02:** đã dùng 20% theo [YC] bản 2 (bản trước là 30%).
4. **D1 một người một dòng:** [CV05] mục I và PL2-3 ghi "mỗi người có mặt là 01 dòng"; nhưng HD_DIEN_BIEU cột B của D1 trong bộ biểu [BB] ghi "người kiêm nhiệm có thể có nhiều dòng phần việc nhưng số người/biên chế chỉ tính một lần". Quy tắc D1-03 theo [CV05] (cấp tỉnh, ban hành sau). Nếu Ban chấp nhận nhiều dòng cho kiêm nhiệm thì hạ D1-03 xuống LƯU Ý.
5. Thiếu 3 tài liệu (Bộ phụ lục, Cẩm nang, Hướng dẫn chuẩn nhập dữ liệu) do tải lên lỗi; sẽ bổ sung ở bản 0.2.
6. Bộ biểu [BB] là "bản sửa đề nghị UEB v2", chưa phải bản Ban ban hành. Khi có bản chính thức: chạy `tests/tao_bo_bieu_lite.py` để cập nhật fixture và đối chiếu lại `config/bieu.json`.
