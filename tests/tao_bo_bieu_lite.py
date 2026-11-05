#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tạo tests/bo_bieu_lite.xlsx: bản RÚT GỌN của bộ biểu chuẩn để làm fixture test (nhỏ, không chứa dữ liệu).
Giữ: tên và thứ tự sheet, trạng thái ẩn, 2 dòng tiêu đề, công thức của N dòng đầu, các sheet tham số/hướng dẫn.
Dùng:  python tests/tao_bo_bieu_lite.py <bo_bieu_chuan.xlsx> [so_dong=60]
"""
import os
import sys

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
BIEU_LON = ("1. A1_NhiemVu", "2. A2_CongViec", "3. B1_KhoiLuong", "4. B2_CT_TacDong", "5. B2_ThayDoi", "6. C1_VTVL", "7. C2_NangLuc",
            "8. C3_NhuCau", "9. D1_NhanLuc", "10. D2_PhuongAn", "11. DM_SAN_PHAM", "12. CHIA_VUNG_CAP_XA", "15. S_NhatKy", "DU_BAO_5_NAM")


def main():
    src = sys.argv[1]
    n = int(sys.argv[2]) if len(sys.argv) > 2 else 60
    wb_in = openpyxl.load_workbook(src, read_only=True)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for ws_in in wb_in.worksheets:
        ws = wb.create_sheet(ws_in.title)
        ws.sheet_state = ws_in.sheet_state
        lim = (2 + n) if ws_in.title in BIEU_LON else 300
        for row in ws_in.iter_rows(min_row=1, max_row=lim):
            for c in row:
                if c.value is not None:
                    ws.cell(row=c.row, column=c.column, value=c.value)
    # DM_THAM_SO có tên vùng; khai lại tên vùng QUY_GIO_PP2 cho công thức tham chiếu
    try:
        from openpyxl.workbook.defined_name import DefinedName
        for name, ref in (("QUY_GIO_PP2", "DM_THAM_SO!$B$3"),):
            wb.defined_names[name] = DefinedName(name, attr_text=ref)
    except Exception:
        pass
    out = os.path.join(HERE, "bo_bieu_lite.xlsx")
    wb.save(out)
    print("Đã tạo", out, os.path.getsize(out) // 1024, "KB")


if __name__ == "__main__":
    main()
