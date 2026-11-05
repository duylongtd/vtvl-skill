#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tạo 2 tệp test từ bộ biểu chuẩn (chỉ ghi vào bản SAO, không đụng tệp gốc):
  tests/mau_dung.xlsx  – dữ liệu giả định hợp lệ (kỳ vọng: không có lỗi BẮT BUỘC)
  tests/mau_sai.xlsx   – cài lỗi cố ý, mỗi lỗi ghi vào tests/mau_sai_ky_vong.json

Cách chạy:  python tests/tao_file_test.py <bo_bieu_chuan.xlsx>
Dữ liệu là GIẢ ĐỊNH, chỉ để thử công cụ; không nộp.
"""
import json
import os
import shutil
import sys

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
R0 = 3  # dòng dữ liệu đầu


def ghi(ws, r, data):
    for col, v in data.items():
        ws[f"{col}{r}"] = v


def dung(wb):
    a1, a2, b1, c1, c2, c3, d1, db, d2 = (wb[n] for n in ("1. A1_NhiemVu", "2. A2_CongViec", "3. B1_KhoiLuong", "6. C1_VTVL", "7. C2_NangLuc", "8. C3_NhuCau", "9. D1_NhanLuc", "DU_BAO_5_NAM", "10. D2_PhuongAn"))
    ky = "01/10/2025 – 30/09/2026"
    cq = "Đảng ủy xã Thử Nghiệm"
    ghi(a1, R0, {"A": "VD-NV-01", "B": "Tham mưu công tác tổ chức xây dựng Đảng", "C": cq, "D": "Quy định số 01-QĐ/ĐU ngày 05/01/2026, Điều 4, khoản 2", "E": ky, "F": "Tiếp tục thực hiện", "G": "Hồ sơ CV 2342-CV/BTCTW, trang NGUON_NHIEM_VU dòng 5"})
    ghi(a1, R0 + 1, {"A": "VD-NV-02", "B": "Tham mưu công tác kiểm tra, giám sát", "C": cq, "D": "Quy định số 01-QĐ/ĐU ngày 05/01/2026, Điều 6, khoản 1", "E": ky, "F": "Tiếp tục thực hiện", "G": "Kế thừa danh mục nhiệm vụ 2025, tệp DM_NV_2025.xlsx"})
    ghi(c1, R0, {"A": "VD-VTVL-01", "B": "Phó Bí thư Đảng ủy xã", "C": "Cán bộ", "D": "Quyết định số 12-QĐ/ĐU ngày 15/3/2026 phê duyệt danh mục VTVL", "F": "Mức chung 1.904 giờ – Bộ luật Lao động", "G": "Tiếp tục sử dụng", "J": "Phổ biến (40 giờ/tuần)"})
    ghi(c1, R0 + 1, {"A": "VD-VTVL-02", "B": "Chuyên viên tổ chức đảng", "C": "Công chức", "D": "Quyết định số 12-QĐ/ĐU ngày 15/3/2026 phê duyệt danh mục VTVL", "F": "Mức chung 1.904 giờ – Bộ luật Lao động", "G": "Tiếp tục sử dụng", "J": "Phổ biến (40 giờ/tuần)"})
    ghi(c1, R0 + 2, {"A": "VD-VTVL-03", "B": "Chuyên viên kiểm tra đảng", "C": "Công chức", "D": "Quyết định số 12-QĐ/ĐU ngày 15/3/2026 phê duyệt danh mục VTVL", "F": "Mức chung 1.904 giờ – Bộ luật Lao động", "G": "Tiếp tục sử dụng", "J": "Phổ biến (40 giờ/tuần)"})
    cv = [
        ("VD-CV-NV01-01", "VD-NV-01", "VD-VTVL-02", "Thẩm tra hồ sơ đề nghị kết nạp đảng viên", "Báo cáo thẩm tra hồ sơ kết nạp", "Hồ sơ", "Đúng quy trình, đúng hạn", "Quy định số 01-QĐ/ĐU ngày 05/01/2026, Điều 4, khoản 2, điểm a"),
        ("VD-CV-NV01-02", "VD-NV-01", "VD-VTVL-02", "Tổng hợp số liệu đảng viên định kỳ", "Báo cáo thống kê đảng viên", "Báo cáo", "Đúng mẫu, đúng hạn", "Hướng dẫn số 02-HD/ĐU ngày 10/01/2026, Điều 3"),
        ("VD-CV-NV02-01", "VD-NV-02", "VD-VTVL-03", "Thẩm tra đơn thư, dấu hiệu vi phạm của đảng viên", "Kết luận thẩm tra đơn thư", "Vụ việc", "Đúng quy trình, đúng thẩm quyền", "Quy định số 01-QĐ/ĐU ngày 05/01/2026, Điều 6, khoản 1"),
        ("VD-CV-NV02-02", "VD-NV-02", "VD-VTVL-01", "Chủ trì giao ban công tác xây dựng Đảng hằng tháng", "Thông báo kết luận giao ban", "Phiên họp", "Đúng lịch, có kết luận", "Quy chế làm việc số 03-QC/ĐU ngày 12/01/2026, Điều 8, khoản 3"),
    ]
    for i, row in enumerate(cv):
        ghi(a2, R0 + i, dict(zip("ABCDEFGH", row)))
    b1rows = [
        ("VD-CV-NV01-01", 120, 45, 0, "Sổ theo dõi hồ sơ kết nạp 2025–2026, Văn phòng Đảng ủy", "Phiếu đo thời gian số 01/PĐ ngày 10/9/2026, xác nhận: Phó Bí thư", "Kiểm đếm, ước lượng thực tế"),
        ("VD-CV-NV01-02", 12, 180, 24, "Phần mềm Sổ đảng viên điện tử, kỳ 01/10/2025–30/9/2026", "Phiếu đo thời gian số 02/PĐ ngày 10/9/2026, xác nhận: Phó Bí thư", "Kiểm đếm, ước lượng thực tế"),
        ("VD-CV-NV02-01", 9, 420, 0, "Sổ tiếp nhận đơn thư 2025–2026, UBKT Đảng ủy", "Phiếu khảo sát số 03/PKS ngày 12/9/2026, xác nhận: Chủ nhiệm UBKT", "Kiểm đếm, ước lượng thực tế"),
        ("VD-CV-NV02-02", 12, 150, 10, "Lịch giao ban 2025–2026, Văn phòng Đảng ủy", "Phiếu đo thời gian số 04/PĐ ngày 12/9/2026, xác nhận: Bí thư", "Kiểm đếm, ước lượng thực tế"),
    ]
    for i, (a, c, d, e, h, ii, j) in enumerate(b1rows):
        ghi(b1, R0 + i, {"A": a, "C": c, "D": d, "E": e, "H": h, "I": ii, "J": j, "K": ky})
    for i, k in enumerate(("VD-VTVL-01", "VD-VTVL-02", "VD-VTVL-03")):
        ghi(c2, R0 + i, {"A": k, "B": "Chịu trách nhiệm chất lượng sản phẩm của vị trí", "C": "Giao tiếp, phối hợp, sử dụng phần mềm", "D": "Nghiệp vụ công tác đảng, mức thành thạo", "E": "Không áp dụng" if i else "Lãnh đạo, điều hành cấp ủy", "F": "Quyết định số 12-QĐ/ĐU ngày 15/3/2026", "G": "Hồ sơ VTVL số HS-0%d" % (i + 1)})
    ghi(c3, R0, {"A": "VD-VTVL-01", "B": "PP1", "H": 1, "J": 1, "K": "Quy định số 329-QĐ/TW ngày 20/5/2025, Điều 5, khoản 2: 01 Phó Bí thư; phạm vi phụ trách: công tác xây dựng Đảng", "L": "Tính thử"})
    ghi(c3, R0 + 1, {"A": "VD-VTVL-02", "B": "PP2", "J": 1, "K": "Theo giờ B1 của vị trí: 126 giờ, quy đổi 0,07; bố trí 1 người kiêm nhiệm", "L": "Tính thử"})
    ghi(c3, R0 + 2, {"A": "VD-VTVL-03", "B": "PP2", "J": 1, "K": "Theo giờ B1 của vị trí: 63 giờ, quy đổi 0,03; kiêm nhiệm", "L": "Tính thử"})
    for i, (a, b, e) in enumerate((("VD-NL-01", "VD-VTVL-01", 1904), ("VD-NL-02", "VD-VTVL-02", 1500), ("VD-NL-03", "VD-VTVL-03", 1700))):
        ghi(d1, R0 + i, {"A": a, "B": b, "C": "Công chức" if i else "Cán bộ", "E": e, "F": 1904 - e, "G": cq, "H": "Hồ sơ cán bộ chốt 30/9/2026; Quyết định phân công số 05-QĐ/ĐU ngày 01/02/2026", "I": "Chưa kiểm chứng"})
    r = R0
    for k in ("VD-VTVL-01", "VD-VTVL-02", "VD-VTVL-03"):
        for y in range(2027, 2032):
            ghi(db, r, {"A": k, "B": y, "C": 0, "D": "Xác nhận không thay đổi so với năm cơ sở", "E": "Kế hoạch công tác giai đoạn 2027–2031, Văn phòng Đảng ủy", "F": "Chưa kiểm chứng"})
            ghi(d2, r, {"A": k, "B": y, "D": 1, "F": "0", "G": 0, "H": "Nhu cầu bằng năm cơ sở, đủ người; chưa đề nghị điều chỉnh", "I": "Đang xây dựng"})
            r += 1


def sai(wb):
    """Cài lỗi cố ý lên tệp đã có dữ liệu đúng. Trả về danh sách mã quy tắc kỳ vọng bắt được."""
    kv = []
    a1, a2, b1, c1, c2, c3, d1, db, d2 = (wb[n] for n in ("1. A1_NhiemVu", "2. A2_CongViec", "3. B1_KhoiLuong", "6. C1_VTVL", "7. C2_NangLuc", "8. C3_NhuCau", "9. D1_NhanLuc", "DU_BAO_5_NAM", "10. D2_PhuongAn"))
    a1["D3"] = "Theo quy định hiện hành"; kv += ["A1-01", "A1-03"]
    a1["D4"] = "Quy định số 01-QĐ/ĐU ngày 05/01/2026"; kv += ["A1-02"]
    a2["H3"] = "Quy trình nội bộ"; kv += ["A2-01"]
    a2["E4"] = a2["E3"].value; kv += ["A2-02"]
    a2["B5"] = "VD-NV-99"; kv += ["A2-04"]
    a2["D6"] = "Giao ban"; kv += ["A2-06"]
    b1["D3"] = 960; kv += ["B1-02"]
    b1["I4"] = "phiếu đo"; kv += ["B1-03"]
    b1["C5"] = 2000; b1["D5"] = 60; kv += ["B1-01"]
    b1["F6"] = 25; kv += ["T-04"]  # gõ đè cột công thức
    b1["C7"] = "=C3*2"; b1["A7"] = "VD-CV-NV01-01"; b1["H7"] = "Sổ theo dõi hồ sơ kết nạp 2025–2026"; b1["I7"] = "Phiếu đo thời gian số 01/PĐ ngày 10/9/2026, xác nhận: Phó Bí thư"; b1["J7"] = "Kiểm đếm, ước lượng thực tế"; b1["K7"] = "01/10/2025 – 30/09/2026"; b1["D7"] = 45; kv += ["T-05"]
    c1["D4"] = "Quyết định số 07-QĐ/ĐU ngày 02/02/2026 về chức năng, nhiệm vụ"; kv += ["C1-01"]
    c2["D3"] = "Đại học, Cao cấp lý luận chính trị"; kv += ["C2-04"]
    c2["A5"] = "VD-VTVL-09"; kv += ["C2-01", "C2-02"]
    c3["B3"] = "PP2"; c3["K3"] = "Theo giờ B1"; kv += ["C3-05", "C3-10"]  # cấp phó dùng PP2
    c3["B4"] = "PP3"; c3["I4"] = 2; c3["J4"] = 2; kv += ["C3-07"]
    c3["L5"] = "Đã kiểm chứng"; kv += ["TT-01"]
    c3["K4"] = c3["K5"].value; kv += ["C3-09"]
    d1["A4"] = "VD-NL-01"; kv += ["D1-03"]
    d1["A5"] = "Vị trí chưa bố trí người"; kv += ["D1-02"]
    d1["E5"] = 2500; kv += ["D1-05"]
    db["C3"] = None; kv += ["DB-02"]
    db["F4"] = "Đã kiểm chứng"; kv += ["TT-02"]
    d2["I3"] = "Đã hoàn thành ở đơn vị"; kv += ["TT-03"]
    d2["G4"] = 1; kv += ["D2-03"]
    d2["C5"] = 7; kv += ["T-04"]
    return sorted(set(kv))


def main():
    if len(sys.argv) < 2:
        sys.exit("Dùng: python tests/tao_file_test.py <bo_bieu_chuan.xlsx>")
    src = sys.argv[1]
    for name, fn in (("mau_dung.xlsx", None), ("mau_sai.xlsx", sai)):  # tệp nguồn nên là tests/bo_bieu_lite.xlsx để chạy nhanh
        dst = os.path.join(HERE, name)
        shutil.copy(src, dst)
        wb = openpyxl.load_workbook(dst)
        dung(wb)
        kv = fn(wb) if fn else []
        wb.save(dst)
        if fn:
            with open(os.path.join(HERE, "mau_sai_ky_vong.json"), "w", encoding="utf-8") as f:
                json.dump(kv, f, ensure_ascii=False, indent=1)
        print("Đã tạo", dst, f"({len(kv)} quy tắc kỳ vọng)" if fn else "")


if __name__ == "__main__":
    main()
