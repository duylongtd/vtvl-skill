# -*- coding: utf-8 -*-
"""
Xuất NHẬT KÝ RÀ SOÁT dạng Excel, gom lỗi cùng loại thành một dòng, theo dõi qua từng version.

Tệp: <tên bộ biểu>_NhatKy.xlsx (hoặc --out). Nếu tệp đã tồn tại từ lần chạy trước,
đọc lại sheet Chi_tiet và Lich_su để: tăng version, đánh dấu lỗi MỚI / CÒN, liệt kê lỗi ĐÃ HẾT.

Sheet:
  Tong_quan          tệp, version, tham số, điểm, tổng theo mức, so với version trước
  Kiem_tra_noi_dung  lỗi nội dung số liệu (A1..D2), gom theo biểu + mã
  Kiem_tra_cau_truc  lỗi cấu trúc tệp, công thức, trạng thái (T-*, TT-*)
  Can_don_vi_xu_ly   việc đơn vị phải làm thêm (liên biểu LB-*, biểu trống, nhật ký)
  Kiem_soat_dong     số dòng dữ liệu từng biểu
  Chi_tiet           mọi phát hiện, một dòng một ô (để lọc)
  Da_het             lỗi của version trước không còn ở version này
  Lich_su            một dòng mỗi version
"""
import os
import re
from collections import Counter, OrderedDict, defaultdict
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

MUC_BAT_BUOC, MUC_CAN, MUC_LUU_Y = "BẮT BUỘC", "CẦN", "LƯU Ý"
MAU = {MUC_BAT_BUOC: "F8CBAD", MUC_CAN: "FFE699", MUC_LUU_Y: "DDEBF7"}
XEP_LOAI = ((90, "Tốt"), (70, "Khá"), (50, "Trung bình"), (0, "Chưa đạt"))
HEAD_CT = ["Mức", "Mã QT", "Biểu", "Ô", "Giá trị hiện tại", "Sai gì", "Sửa thế nào", "Căn cứ", "So với version trước"]
HEAD_GOM = ["Cấp độ", "Mã QT", "Biểu", "Nội dung kiểm tra", "Số ô", "Vị trí", "Ví dụ giá trị", "Việc cần làm", "Căn cứ", "So với version trước"]
HEAD_LS = ["Version", "Thời điểm", "Tệp", "Công cụ", "Điểm", "Xếp loại", "BẮT BUỘC", "CẦN", "LƯU Ý", "Số loại lỗi", "Đã hết", "Mới", "Biên chế giao", "Có mặt 30/9", "Ghi chú"]


def _nen_vi_tri(cells, toi_da=12):
    """['D10','D11','D12','H5'] → 'D10–D12 (3 ô); H5'. Ô không có dạng cột+dòng giữ nguyên."""
    ref = []
    khac = []
    for c in cells:
        m = re.fullmatch(r"([A-Z]{1,3})(\d+)", str(c).strip())
        (ref if m else khac).append((m.group(1), int(m.group(2))) if m else str(c))
    out = []
    by_col = defaultdict(list)
    for col, r in ref:
        by_col[col].append(r)
    for col in sorted(by_col, key=lambda x: (len(x), x)):
        rs = sorted(set(by_col[col]))
        i = 0
        while i < len(rs):
            j = i
            while j + 1 < len(rs) and rs[j + 1] == rs[j] + 1:
                j += 1
            out.append(f"{col}{rs[i]}" if i == j else f"{col}{rs[i]}–{col}{rs[j]} ({j - i + 1} ô)")
            i = j + 1
    out += khac
    if len(out) > toi_da:
        out = out[:toi_da] + [f"… +{len(out) - toi_da} vị trí khác (xem Chi_tiet)"]
    return "; ".join(out)


def _nhom(ma):
    if ma.startswith("LB-") or ma in ("B2-01", "T-12", "A1-04", "C2-02", "DB-01"):
        return "don_vi"
    if ma.startswith("T-") or ma.startswith("TT-"):
        return "cau_truc"
    return "noi_dung"


def _diem(ph):
    loai = defaultdict(set)
    for p in ph:
        loai[p.muc].add(p.ma)
    d = 100 - 6 * len(loai[MUC_BAT_BUOC]) - 2 * len(loai[MUC_CAN]) - 1 * len(loai[MUC_LUU_Y])
    d = max(0, d)
    xl = next(t for nguong, t in XEP_LOAI if d >= nguong)
    return d, xl


def _doc_cu(path):
    """Đọc nhật ký cũ: (version_cu, set khoá phát hiện cũ, rows Lich_su, rows Chi_tiet cũ)."""
    if not path or not os.path.exists(path):
        return 0, set(), [], []
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return 0, set(), [], []
    ls, ct, keys, ver = [], [], set(), 0
    if "Lich_su" in wb.sheetnames:
        for i, row in enumerate(wb["Lich_su"].iter_rows(values_only=True)):
            if i == 0 or row[0] is None:
                continue
            ls.append(list(row))
            try:
                ver = max(ver, int(row[0]))
            except (TypeError, ValueError):
                pass
    if "Chi_tiet" in wb.sheetnames:
        for i, row in enumerate(wb["Chi_tiet"].iter_rows(values_only=True)):
            if i == 0 or row[1] is None:
                continue
            ct.append(list(row))
            keys.add((row[1], row[2], row[3], row[5]))
    return ver, keys, ls, ct


def _version_tu_ten(path):
    m = re.search(r"[Vv]ersion[_ -]?(\d+)", os.path.basename(path))
    return int(m.group(1)) if m else None


def _ghi_bang(ws, head, rows, widths=None, mau_cot=None):
    ws.append(head)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="D9D9D9")
        c.alignment = Alignment(wrap_text=True, vertical="top")
    for r in rows:
        ws.append(r)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
        if mau_cot is not None and row[mau_cot].value in MAU:
            row[mau_cot].fill = PatternFill("solid", fgColor=MAU[row[mau_cot].value])
    for i, w in enumerate(widths or [], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    if rows:
        ws.auto_filter.ref = ws.dimensions


def xuat_nhat_ky(kt, meta, out_path, nhat_ky_truoc=None, version=None):
    """kt: KiemTra đã chạy. Trả về dict tóm tắt (version, điểm, đã hết, mới)."""
    truoc = nhat_ky_truoc or out_path
    ver_cu, keys_cu, ls_cu, ct_cu = _doc_cu(truoc)
    if version is None:
        version = _version_tu_ten(kt.path) or (ver_cu + 1 if ver_cu else 1)
    keys_moi = {(p.ma, p.bieu, p.o, p.noi_dung) for p in kt.ph}
    trang_thai = {}
    for p in kt.ph:
        k = (p.ma, p.bieu, p.o, p.noi_dung)
        trang_thai[k] = ("CÒN" if k in keys_cu else "MỚI") if keys_cu else ""
    da_het = [r for r in ct_cu if (r[1], r[2], r[3], r[5]) not in keys_moi] if keys_cu else []
    n_moi = sum(1 for v in trang_thai.values() if v == "MỚI")
    diem, xl = _diem(kt.ph)
    cnt = Counter(p.muc for p in kt.ph)

    wb = openpyxl.Workbook()
    # --- Tong_quan
    ws = wb.active
    ws.title = "Tong_quan"
    so_loai = len({p.ma for p in kt.ph})
    rows = [
        ["Tệp kiểm tra", os.path.basename(kt.path)],
        ["Version", version],
        ["Thời điểm", meta["thoi_diem"]],
        ["Công cụ", f"vtvl-kiem-tra v{meta['version']}"],
        ["Căn cứ", "HD 07-HD/BTCTW; CV 05-CV/BCĐ (PL1, PL2); Yêu cầu kiểm tra dữ liệu; HD nhập liệu v1.1"],
        ["Tham số", f"biên chế giao = {meta['bien_che']}; có mặt 30/9/2026 = {meta['so_nguoi']}; loại đơn vị = {meta['loai_don_vi']}; đã thẩm định = {'có' if meta['da_tham_dinh'] else 'chưa'}"],
        ["Quy tắc bỏ qua", "; ".join(kt.bo_qua) or "không"],
        [],
        ["ĐIỂM", f"{diem}/100 – {xl}"],
        ["Cách tính điểm", "100 − 6 × số LOẠI lỗi BẮT BUỘC − 2 × số loại CẦN − 1 × số loại LƯU Ý. Điểm chỉ để theo dõi tiến độ giữa các version, không phải đánh giá chính thức."],
        ["BẮT BUỘC", cnt.get(MUC_BAT_BUOC, 0)],
        ["CẦN", cnt.get(MUC_CAN, 0)],
        ["LƯU Ý", cnt.get(MUC_LUU_Y, 0)],
        ["Số loại lỗi", so_loai],
        [],
        ["So với version trước", f"đã hết {len(da_het)} · còn {sum(1 for v in trang_thai.values() if v == 'CÒN')} · mới {n_moi}" if keys_cu else "không có nhật ký trước (version đầu)"],
        ["Kết luận", "Hồ sơ còn lỗi BẮT BUỘC sẽ bị trả lại, không đưa vào tổng hợp (CV05 mục II)." if cnt.get(MUC_BAT_BUOC) else "Không còn lỗi BẮT BUỘC. Rà các mục CẦN và nội dung không kiểm tra tự động được."],
        [],
        ["Nguyên tắc", "Công cụ chỉ cảnh báo, không sửa dữ liệu. Người kê khai tự sửa trên bộ biểu/app và chịu trách nhiệm."],
        ["Không kiểm tra tự động được", "Phiếu đo có thật (CV05 4.1); không sao chép hồ sơ mẫu (PL2-20); định mức chuyên ngành (PL1-11); người đứng đầu ký xác nhận (mục II); đối chiếu D1 khối Đảng – chính quyền (PL2-4)."],
    ]
    for r in rows:
        ws.append(r)
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 110
    for row in ws.iter_rows():
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
        if row[0].value:
            row[0].font = Font(bold=True)
    ws["B9"].font = Font(bold=True, size=13)

    # --- gom theo (nhóm, biểu, mã)
    gom = OrderedDict()
    for p in kt.ph:
        k = (_nhom(p.ma), p.bieu, p.ma, p.muc)
        g = gom.setdefault(k, {"cells": [], "vals": [], "noi_dung": p.noi_dung, "sua": p.sua, "can_cu": p.can_cu, "tt": set()})
        g["cells"].append(p.o)
        if p.gia_tri and p.gia_tri not in g["vals"] and len(g["vals"]) < 3:
            g["vals"].append(p.gia_tri)
        g["tt"].add(trang_thai[(p.ma, p.bieu, p.o, p.noi_dung)])

    def rows_nhom(nhom):
        out = []
        thu_tu = {MUC_BAT_BUOC: 0, MUC_CAN: 1, MUC_LUU_Y: 2}
        for (nh, bieu, ma, muc), g in sorted(gom.items(), key=lambda kv: (kv[0][1], thu_tu.get(kv[0][3], 9), kv[0][2])):
            if nh != nhom:
                continue
            nd = g["noi_dung"]
            if len(g["cells"]) > 1:
                nd = re.sub(r"\s*\(.*?\)\s*$", "", nd) if nd.count("(") == 1 and nd.endswith(")") else nd
            tt = ", ".join(sorted(x for x in g["tt"] if x))
            out.append([muc, ma, bieu, nd, len(g["cells"]), _nen_vi_tri(g["cells"]), " | ".join(g["vals"]), g["sua"], g["can_cu"], tt])
        return out

    w_gom = [11, 8, 30, 48, 6, 36, 34, 52, 30, 12]
    _ghi_bang(wb.create_sheet("Kiem_tra_noi_dung"), HEAD_GOM, rows_nhom("noi_dung"), w_gom, 0)
    _ghi_bang(wb.create_sheet("Kiem_tra_cau_truc"), HEAD_GOM, rows_nhom("cau_truc"), w_gom, 0)
    _ghi_bang(wb.create_sheet("Can_don_vi_xu_ly"), HEAD_GOM, rows_nhom("don_vi"), w_gom, 0)

    # --- Kiem_soat_dong
    ks = []
    for key, b in kt.B.items():
        if b:
            ks.append([f"{b.ten} [{b.sheet_name}]", len(b.rows), b.rownums[0] if b.rownums else "", b.rownums[-1] if b.rownums else ""])
        else:
            ks.append([b.ten, "KHÔNG TÌM THẤY SHEET", "", ""])
    _ghi_bang(wb.create_sheet("Kiem_soat_dong"), ["Biểu", "Số dòng có dữ liệu", "Dòng đầu", "Dòng cuối"], ks, [44, 20, 10, 10])

    # --- Chi_tiet
    ct = [[p.muc, p.ma, p.bieu, p.o, p.gia_tri, p.noi_dung, p.sua, p.can_cu, trang_thai[(p.ma, p.bieu, p.o, p.noi_dung)]] for p in kt.ph]
    _ghi_bang(wb.create_sheet("Chi_tiet"), HEAD_CT, ct, [11, 8, 30, 14, 30, 50, 52, 30, 12], 0)

    # --- Da_het
    _ghi_bang(wb.create_sheet("Da_het"), HEAD_CT[:8], [r[:8] for r in da_het], [11, 8, 30, 14, 30, 50, 52, 30], 0)

    # --- Lich_su
    ls_moi = ls_cu + [[version, meta["thoi_diem"], os.path.basename(kt.path), f"v{meta['version']}", diem, xl,
                       cnt.get(MUC_BAT_BUOC, 0), cnt.get(MUC_CAN, 0), cnt.get(MUC_LUU_Y, 0), so_loai,
                       len(da_het) if keys_cu else "", n_moi if keys_cu else "", meta["bien_che"], meta["so_nguoi"], ""]]
    _ghi_bang(wb.create_sheet("Lich_su"), HEAD_LS, ls_moi, [9, 17, 36, 9, 7, 11, 10, 8, 8, 11, 8, 8, 12, 12, 30])

    wb.save(out_path)
    return {"version": version, "diem": diem, "xep_loai": xl, "da_het": len(da_het), "moi": n_moi, "con": sum(1 for v in trang_thai.values() if v == "CÒN"), "so_loai": so_loai}
