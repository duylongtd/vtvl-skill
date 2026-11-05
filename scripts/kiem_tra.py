#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
kiem_tra.py – Kiểm tra logic Bộ biểu VTVL (chỉ đọc, không sửa file).

Cách dùng:
  python kiem_tra.py bo_bieu.xlsx --bien-che 25 --so-nguoi 23 --loai-don-vi xa
  python kiem_tra.py bo_bieu.xlsx --da-tham-dinh --out ket_qua

Đầu ra (cùng tên --out):
  <out>.docx  báo cáo theo từng biểu (bắt buộc theo Yêu cầu kiểm tra)
  <out>.xlsx  bảng phát hiện để lọc, sắp xếp
  <out>.json  dữ liệu máy đọc

Mọi quy tắc đều ghi trong references/RULES.md. Script KHÔNG tự suy luận thêm
ngoài những gì cấu hình trong config/*.json.
"""
import argparse
import json
import os
import re
import sys
import unicodedata
from collections import Counter, defaultdict
from datetime import datetime

try:
    import openpyxl
    from openpyxl.utils import get_column_letter, column_index_from_string
except ImportError:
    sys.exit("Thiếu thư viện openpyxl. Cài: pip install openpyxl python-docx")

VERSION = "0.1.1"
HERE = os.path.dirname(os.path.abspath(__file__))
CONFIG_DIR = os.path.join(os.path.dirname(HERE), "config")

MUC_BAT_BUOC, MUC_CAN, MUC_LUU_Y = "BẮT BUỘC", "CẦN", "LƯU Ý"
MUC_THU_TU = {MUC_BAT_BUOC: 0, MUC_CAN: 1, MUC_LUU_Y: 2}


# ----------------------------------------------------------------- tiện ích
def load_json(name):
    with open(os.path.join(CONFIG_DIR, name), encoding="utf-8") as f:
        return json.load(f)


def nfc(s):
    return unicodedata.normalize("NFC", s)


def chuan(v):
    """Chuẩn hoá chuỗi để so sánh (R00): NFC, bỏ khoảng trắng thừa, chữ thường."""
    if v is None:
        return ""
    s = nfc(str(v)).replace(" ", " ").replace("\t", " ").replace("\r", " ").replace("\n", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s.lower()


def chuan_ma(v):
    return chuan(v).upper().replace(" ", "")


def la_so(v):
    if v is None or v == "":
        return False
    if isinstance(v, (int, float)):
        return True
    try:
        float(str(v).replace(".", "").replace(",", "."))
        return True
    except ValueError:
        return False


def so(v, default=None):
    """Đổi về float; chấp nhận '1.904', '1,5756', '30 phút' → None (không phải số thuần)."""
    if v is None or v == "":
        return default
    if isinstance(v, bool):
        return default
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if re.fullmatch(r"-?\d{1,3}(\.\d{3})+(,\d+)?", s):  # 1.904 hoặc 1.904,5
        s = s.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"-?\d+,\d+", s):
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return default


def trong(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


_REF = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")


def mau_cong_thuc(f, r):
    """Thay số dòng TƯƠNG ĐỐI bằng {r} để so mẫu công thức giữa các dòng.
    Giữ nguyên $A$3 (tuyệt đối) và không đụng vào chuỗi trong dấu ngoặc kép ("PP3")."""
    def rep(m):
        if m.group(3) == "" and int(m.group(4)) == r:
            return f"{m.group(1)}{m.group(2)}{{r}}"
        return m.group(0)
    out, i, n = [], 0, len(f)
    while i < n:
        ch = f[i]
        if ch in "\"'":  # chuỗi "..." hoặc tên sheet '...': giữ nguyên
            j = f.find(ch, i + 1)
            j = n if j < 0 else j
            out.append(f[i:j + 1])
            i = j + 1
        else:
            j = i
            while j < n and f[j] not in "\"'":
                j += 1
            out.append(_REF.sub(rep, f[i:j]))
            i = j
    return "".join(out)


class PhatHien:
    __slots__ = ("ma", "bieu", "o", "gia_tri", "noi_dung", "sua", "muc", "can_cu")

    def __init__(self, ma, bieu, o, gia_tri, noi_dung, sua, muc, can_cu):
        self.ma, self.bieu, self.o = ma, bieu, o
        self.gia_tri = "" if gia_tri is None else str(gia_tri)[:120]
        self.noi_dung, self.sua, self.muc, self.can_cu = noi_dung, sua, muc, can_cu

    def row(self):
        return [self.muc, self.ma, self.bieu, self.o, self.gia_tri, self.noi_dung, self.sua, self.can_cu]


# ----------------------------------------------------------------- đọc workbook
class Bieu:
    """Một sheet đã đọc: rows = list[dict col_letter -> value], frows = công thức tương ứng."""

    def __init__(self, key, cfg, ws_val, ws_f, header_vals=None):
        self.key, self.cfg = key, cfg
        self.ten = cfg["ten"]
        self.sheet_name = ws_val.title if ws_val else None
        self.rows, self.frows, self.rownums = [], [], []
        self.headers = header_vals or {}
        if ws_val is None:
            return
        hr = cfg["header_row"]
        cols = list(cfg["cols"].keys())
        idx = {c: column_index_from_string(c) - 1 for c in cols}
        maxc = max(idx.values()) + 1
        blank_run = 0
        it_v = ws_val.iter_rows(min_row=hr, max_col=maxc, values_only=True)
        it_f = ws_f.iter_rows(min_row=hr, max_col=maxc, values_only=True)
        for r, (rv, rf) in enumerate(zip(it_v, it_f), start=hr):
            if r == hr:
                self.headers = {c: (rv[idx[c]] if idx[c] < len(rv) else None) for c in cols}
                continue
            vals = {c: (rv[idx[c]] if idx[c] < len(rv) else None) for c in cols}
            fs = {c: (rf[idx[c]] if idx[c] < len(rf) else None) for c in cols}
            if all(trong(v) for v in vals.values()):
                blank_run += 1
                if blank_run >= 20:
                    break
                continue
            blank_run = 0
            self.rows.append(vals)
            self.frows.append(fs)
            self.rownums.append(r)
        self.max_data_row = self.rownums[-1] if self.rownums else hr

    def __bool__(self):
        return self.sheet_name is not None

    def o(self, i, col):
        return f"{col}{self.rownums[i]}"

    def by_key(self, col):
        d = defaultdict(list)
        for i, row in enumerate(self.rows):
            k = chuan_ma(row.get(col))
            if k:
                d[k].append(i)
        return d


def tim_sheet(wb, candidates):
    names = {chuan(n): n for n in wb.sheetnames}
    for c in candidates:
        if chuan(c) in names:
            return names[chuan(c)]
    # khớp mờ: tên sheet chứa ứng viên (bỏ số thứ tự "1. ")
    for c in candidates:
        cc = re.sub(r"^\d+\.\s*", "", chuan(c))
        for n0, n in names.items():
            if cc and cc in re.sub(r"^\d+\.\s*", "", n0):
                return n
    return None


# ----------------------------------------------------------------- bộ kiểm tra
class KiemTra:
    def __init__(self, path, args):
        self.path = path
        self.args = args
        self.cfg = load_json("bieu.json")
        self.qt = load_json("quy_tac.json")
        self.ph = []
        self.bo_qua = []
        self.wb_f = openpyxl.load_workbook(path, data_only=False, read_only=True)
        self.wb_v = openpyxl.load_workbook(path, data_only=True, read_only=True)
        self.B = {}
        for key, scfg in self.cfg["sheets"].items():
            name = tim_sheet(self.wb_v, scfg["sheet_candidates"])
            self.B[key] = Bieu(key, scfg, self.wb_v[name] if name else None, self.wb_f[name] if name else None)

    def _gioi_han_dong(self, title):
        """Số dòng cần quét của một sheet: tới dòng dữ liệu cuối + 20 cho biểu đã đọc; 300 cho sheet khác."""
        for b in self.B.values():
            if b and b.sheet_name == title:
                return b.max_data_row + 20
        return 300

    # ---- ghi phát hiện
    def add(self, ma, bieu, o, gia_tri, noi_dung, sua, muc, can_cu):
        self.ph.append(PhatHien(ma, bieu, o, gia_tri, noi_dung, sua, muc, can_cu))

    def ten(self, key):
        b = self.B[key]
        return f"{b.ten} [{b.sheet_name}]" if b else b.ten

    # ---- giá trị dự phòng khi tệp chưa được Excel tính lại (không có cache công thức)
    def _quy_gio(self, ma_vtvl):
        c1 = self.B["C1"]
        qg = self.cfg["quy_gio_pho_bien"]
        if c1:
            for row in c1.rows:
                if chuan_ma(row.get("A")) == ma_vtvl:
                    e = so(row.get("E"))
                    if e:
                        return e
                    h = so(row.get("H"))
                    if "đặc thù" in chuan(row.get("J")) and h and chuan(row.get("I")) == "đã kiểm chứng":
                        return h
                    return qg
        return qg

    def _b1_gio_tong(self, row):
        """B1.G; nếu trống thì tính C × (N hoặc D) ÷ 60 + E."""
        g = so(row.get("G"))
        if g is not None:
            return g
        c, d, n, e = so(row.get("C")), so(row.get("D")), so(row.get("N")), so(row.get("E"))
        p = n if n else d
        if c is not None and p is not None:
            return c * p / 60 + (e or 0)
        if e:
            return e
        return None

    def _b1_gio_theo_vtvl(self):
        b1, a2 = self.B["B1"], self.B["A2"]
        out = defaultdict(float)
        if not b1:
            return out
        vt = {chuan_ma(r.get("A")): chuan_ma(r.get("C")) for r in a2.rows} if a2 else {}
        for row in b1.rows:
            k = chuan_ma(row.get("B")) or vt.get(chuan_ma(row.get("A")), "")
            g = self._b1_gio_tong(row)
            if k and g:
                out[k] += g
        return out

    def _c3_G(self, row, gio_vtvl):
        g = so(row.get("G"))
        if g is not None:
            return g
        pp = chuan(row.get("B")).upper()
        if "PP2" not in pp:
            return None
        k = chuan_ma(row.get("A"))
        e = so(row.get("E"))
        if e is None:
            c = so(row.get("C"))
            if c is None:
                c = gio_vtvl.get(k)
            if c is None:
                return None
            e = c + (so(row.get("D")) or 0)
        f = so(row.get("F")) or self._quy_gio(k)
        return e / f if f else None

    # ================================================================ T – toàn tệp
    def kt_T(self):
        qt = self.qt
        # T-01 thiếu sheet / sheet ẩn
        for key, b in self.B.items():
            if not b:
                self.add("T-01", b.ten, "-", "", "Không tìm thấy trang tính của biểu này (các tên thử: %s)" % ", ".join(b.cfg["sheet_candidates"]),
                         "Dùng đúng bộ biểu do cơ quan đầu mối giao, không xoá/đổi tên trang tính; hoặc sửa config/bieu.json nếu tên sheet khác",
                         MUC_BAT_BUOC, "CV05 mục I 'Về mẫu biểu'; PL2-1")
        an_ok = {chuan(x) for x in self.cfg.get("sheet_an_theo_thiet_ke", [])}
        for name in self.cfg.get("sheet_bat_buoc", []):
            if not tim_sheet(self.wb_v, [name]):
                self.add("T-01", "Toàn tệp", name, "", "Thiếu trang tính so với bộ biểu chuẩn",
                         "Dùng bộ biểu sạch do cơ quan đầu mối giao; không xoá, đổi tên trang", MUC_BAT_BUOC, "CV05 'Về mẫu biểu'; PL2-1")
        for ws in self.wb_f.worksheets:
            if ws.sheet_state != "visible" and chuan(ws.title) not in an_ok:
                self.add("T-01", "Toàn tệp", ws.title, ws.sheet_state, f"Trang tính bị ẩn ({ws.sheet_state})",
                         "Không ẩn trang; nếu là trang ẩn lạ thì quét virus và chép sang bộ biểu sạch", MUC_BAT_BUOC, "CV05 'Về mẫu biểu'; PL1-7; PL2-1")
        # T-03 liên kết ngoài
        ext = getattr(self.wb_f, "_external_links", None) or []
        if ext:
            self.add("T-03", "Toàn tệp", "-", f"{len(ext)} liên kết", "Tệp có liên kết đến tệp khác (External Links)",
                     "Ngắt liên kết (Data → Edit Links → Break), chép dữ liệu bằng Dán giá trị", MUC_BAT_BUOC, "CV05 'Về mẫu biểu'; PL1-7; PL2-1; YC")
        for ws in self.wb_f.worksheets:
            lim = self._gioi_han_dong(ws.title)
            for row in ws.iter_rows(max_row=lim):
                for c in row:
                    v = c.value
                    if isinstance(v, str) and v.startswith("="):
                        if re.search(r"\[\d+\]|\[[^\]]+\.xls[xm]?\]|'[A-Za-z]:\\|\\\\|https?://", v):
                            self.add("T-03", "Toàn tệp", f"{ws.title}!{c.coordinate}", v, "Công thức tham chiếu tệp ngoài (có ký hiệu [1] hoặc đường dẫn)",
                                     "Ngắt liên kết, thay bằng giá trị", MUC_BAT_BUOC, "CV05 PL2-1; YC")
                    if isinstance(v, str) and v.startswith("#") and v.rstrip("!?/0").upper() in ("#REF", "#VALUE", "#N/A", "#DIV", "#NAME", "#NUM", "#NULL"):
                        pass
        # T-07 lỗi #REF!... (đọc ở bản giá trị)
        for ws in self.wb_v.worksheets:
            lim = self._gioi_han_dong(ws.title)
            for row in ws.iter_rows(max_row=lim):
                for c in row:
                    v = c.value
                    if isinstance(v, str) and re.match(r"#(REF|VALUE|N/A|DIV/0|NAME|NUM|NULL)", v):
                        self.add("T-07", "Toàn tệp", f"{ws.title}!{c.coordinate}", v, "Ô báo lỗi công thức",
                                 "Tìm ô nguồn gây lỗi (thường do xoá/chèn dòng hoặc gõ đè cột tự động)", MUC_BAT_BUOC, "CV05 PL2-2")
        # T-02 tiêu đề cột
        for key, b in self.B.items():
            if not b:
                continue
            for col, ten in b.cfg["cols"].items():
                h = b.headers.get(col)
                if trong(h):
                    self.add("T-02", self.ten(key), f"{col}{b.cfg['header_row']}", "", f"Tiêu đề cột trống (mẫu: '{ten}')",
                             "Giữ nguyên tiêu đề cột của mẫu; nếu bộ biểu thật khác thì cập nhật config/bieu.json", MUC_BAT_BUOC, "CV05 'Về mẫu biểu'; PL2-1")
        # T-04/T-05/T-06: công thức
        for key, b in self.B.items():
            if not b or not b.rows:
                continue
            fcols = set(b.cfg["formula_cols"])
            for col in b.cfg["cols"]:
                fs = [b.frows[i].get(col) for i in range(len(b.rows))]
                is_f = [isinstance(x, str) and x.startswith("=") for x in fs]
                n_f = sum(is_f)
                if col in fcols:
                    if n_f == 0:
                        # cả cột không có công thức: có thể là bản "Dán giá trị" của app xuất → LƯU Ý một lần
                        if any(not trong(b.rows[i].get(col)) for i in range(len(b.rows))):
                            self.add("T-04", self.ten(key), f"cột {col}", "", f"Cột tự động '{b.cfg['cols'][col]}' không còn công thức ở dòng nào (toàn cột là giá trị tĩnh)",
                                     "Nếu tệp này là bản gửi để nạp app thì chấp nhận; nếu là bộ biểu Excel thì khôi phục công thức từ bộ biểu sạch",
                                     MUC_LUU_Y, "CV05 'Về mẫu biểu'; PL2-2; YC")
                        continue
                    # chuẩn mẫu công thức: thay số dòng bằng {r}
                    pats = Counter()
                    for i, f in enumerate(fs):
                        if is_f[i]:
                            pats[mau_cong_thuc(f, b.rownums[i])] += 1
                    mau = pats.most_common(1)[0][0] if pats else None
                    for i in range(len(b.rows)):
                        if not is_f[i]:
                            if not trong(b.rows[i].get(col)):
                                self.add("T-04", self.ten(key), b.o(i, col), b.rows[i].get(col), f"Ô thuộc cột tự động '{b.cfg['cols'][col]}' bị gõ đè giá trị, mất công thức",
                                         "Khôi phục công thức (chép từ dòng còn công thức hoặc bộ biểu sạch)", MUC_BAT_BUOC, "CV05 'Về mẫu biểu'; PL2-2; YC")
                        else:
                            p = mau_cong_thuc(fs[i], b.rownums[i])
                            if mau and p != mau and pats[p] < pats[mau]:
                                self.add("T-06", self.ten(key), b.o(i, col), fs[i], "Công thức khác mẫu so với các dòng khác trong cột",
                                         f"Chép lại công thức đúng (mẫu: {mau.replace('{r}', str(b.rownums[i]))})", MUC_BAT_BUOC, "CV05 PL2-2")
                else:
                    for i in range(len(b.rows)):
                        if is_f[i]:
                            self.add("T-05", self.ten(key), b.o(i, col), fs[i], f"Cột nhập liệu '{b.cfg['cols'][col]}' chứa công thức",
                                     "Thay bằng giá trị (Dán giá trị)", MUC_BAT_BUOC, "CV05 'Về mẫu biểu'; PL2-2")
        # T-09 số hiệu tăng dần (kéo chuột)
        n_min = qt["nguong"]["T09_so_dong_lien_tiep_tang_dan"]
        for key, cols in (("A1", ["D"]), ("A2", ["H"]), ("C1", ["D"]), ("C2", ["F"])):
            b = self.B[key]
            if not b:
                continue
            for col in cols:
                prev, run, start = None, 1, 0
                for i, row in enumerate(b.rows):
                    m = re.search(r"\b(\d{1,4})\s*[-/]\s*[A-ZĐ]{1,6}", str(row.get(col) or ""))
                    cur = int(m.group(1)) if m else None
                    if prev is not None and cur is not None and cur == prev + 1:
                        run += 1
                    else:
                        if run >= n_min:
                            self.add("T-09", self.ten(key), f"{col}{b.rownums[start]}:{col}{b.rownums[i-1]}", "", f"Số hiệu văn bản tăng dần đều qua {run} dòng liên tiếp: dấu hiệu kéo chuột sao chép",
                                     "Đối chiếu bản gốc, ghi lại đúng số hiệu, ngày từng văn bản", MUC_BAT_BUOC, "CV05 'Về căn cứ, mô tả, mã'; PL1-8; PL2-11")
                        run, start = 1, i
                    prev = cur
                if run >= n_min:
                    self.add("T-09", self.ten(key), f"{col}{b.rownums[start]}:{col}{b.rownums[-1]}", "", f"Số hiệu văn bản tăng dần đều qua {run} dòng liên tiếp: dấu hiệu kéo chuột sao chép",
                             "Đối chiếu bản gốc, ghi lại đúng số hiệu, ngày từng văn bản", MUC_BAT_BUOC, "CV05 PL1-8; PL2-11")
        # T-10 cùng số hiệu khác ngày
        sohieu = defaultdict(set)
        for key, cols in (("A1", ["D"]), ("A2", ["H"]), ("C1", ["D"]), ("C2", ["F"]), ("C3", ["K"])):
            b = self.B[key]
            if not b:
                continue
            for col in cols:
                for i, row in enumerate(b.rows):
                    s = str(row.get(col) or "")
                    m = re.search(r"(\d{1,4}/[A-ZĐ0-9\-]{2,}[A-ZĐ0-9\-/]*)", s.upper())
                    d = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", s)
                    if m and d:
                        sohieu[m.group(1)].add((d.group(1), f"{self.ten(key)} {b.o(i, col)}"))
        for sh, items in sohieu.items():
            dates = {d for d, _ in items}
            if len(dates) > 1:
                self.add("T-10", "Toàn tệp", "; ".join(sorted(o for _, o in items))[:200], sh, f"Cùng số hiệu '{sh}' nhưng ghi {len(dates)} ngày khác nhau: {', '.join(sorted(dates))}",
                         "Thống nhất số, ngày theo bản gốc trong toàn tệp", MUC_BAT_BUOC, "CV05 PL2-11")
        # T-11 chuẩn hoá
        for key, b in self.B.items():
            if not b:
                continue
            for i, row in enumerate(b.rows):
                for col, v in row.items():
                    if isinstance(v, str):
                        if nfc(v) != v:
                            self.add("T-11", self.ten(key), b.o(i, col), v, "Chữ gõ kiểu Unicode tổ hợp (nhìn giống nhưng máy so sánh khác)",
                                     "Gõ lại bằng Unicode dựng sẵn (Unikey: bảng mã Unicode)", MUC_LUU_Y, "HDNL 14.1")
                        elif v != v.strip() or " " in v or "  " in v:
                            self.add("T-11", self.ten(key), b.o(i, col), v, "Khoảng trắng thừa / ký tự ẩn đầu-cuối ô",
                                     "Xoá khoảng trắng thừa", MUC_LUU_Y, "HDNL 14.1")
        # T-12 nhật ký
        nk = tim_sheet(self.wb_v, self.cfg["sheets_khac"]["NHAT_KY"])
        if nk:
            ws = self.wb_v[nk]
            n = sum(1 for r in ws.iter_rows(min_row=3, max_row=200, values_only=True) if any(not trong(v) for v in r))
            if n == 0:
                self.add("T-12", f"Nhật ký [{nk}]", "-", "", "Trang nhật ký (S_NhatKy) trống",
                         "Ghi nhật ký mọi lần sửa số liệu, có căn cứ (bắt buộc với bản gửi lại)", MUC_CAN, "CV05 'Về kết quả tính toán'; PL1-10; PL2-8")

    # ================================================================ TT – trạng thái
    def kt_TT(self):
        if self.args.da_tham_dinh:
            return
        ok = self.qt["trang_thai_chua_tham_dinh"]
        for ma, key, col, allowed, sua in (
            ("TT-01", "C3", "L", ok["C3_L"], "Để 'Tính thử' cho đến khi được thẩm định"),
            ("TT-02", "DB", "F", ok["DB_F"], "Để 'Chưa kiểm chứng' cho đến khi được thẩm định"),
            ("TT-03", "D2", "I", ok["D2_I"], "Để 'Đang xây dựng' cho đến khi được thẩm định"),
        ):
            b = self.B[key]
            if not b:
                continue
            al = {chuan(x) for x in allowed}
            for i, row in enumerate(b.rows):
                v = row.get(col)
                if not trong(v) and chuan(v) not in al:
                    self.add(ma, self.ten(key), b.o(i, col), v, "Tự chọn trạng thái đã kiểm chứng/hoàn thành khi hồ sơ chưa được thẩm định",
                             sua + " (chạy lại với --da-tham-dinh nếu hồ sơ đã thẩm định)", MUC_BAT_BUOC, "CV05 'Về trạng thái'; PL1-6; PL2-9")
        # TT-04: trạng thái kiểm chứng nhưng nguồn trống
        for key, col_tt, col_nguon, kc in (("B1", "J", "I", None), ("B2", "M", "J", "đã kiểm chứng"), ("C1", "I", "F", "đã kiểm chứng"), ("D1", "I", "H", "đã kiểm chứng")):
            b = self.B[key]
            if not b:
                continue
            for i, row in enumerate(b.rows):
                tt = chuan(row.get(col_tt))
                if kc and tt == kc and trong(row.get(col_nguon)):
                    self.add("TT-04", self.ten(key), b.o(i, col_nguon), "", f"Trạng thái '{row.get(col_tt)}' nhưng cột nguồn/căn cứ trống",
                             "Chỉ chọn 'Đã kiểm chứng' khi có đối chiếu, xác nhận thật; ghi đủ nguồn (số, ngày, người xác nhận)", MUC_BAT_BUOC, "CV05 'Về trạng thái', 'Về kiểm chứng'")

    # ================================================================ căn cứ
    def _kt_can_cu(self, ma_prefix, key, col, bat_buoc_dieu_khoan=True):
        b = self.B[key]
        if not b:
            return
        cc = self.qt["can_cu"]
        for i, row in enumerate(b.rows):
            v = row.get(col)
            s = chuan(v)
            if trong(v):
                self.add(f"{ma_prefix}-01", self.ten(key), b.o(i, col), "", "Căn cứ/nguồn trống", "Ghi loại văn bản, số hiệu, ngày, điều, khoản, điểm", MUC_BAT_BUOC, "YC; CV05 PL1-8")
                continue
            co_vb = any(k in s for k in cc["loai_van_ban"]) and re.search(r"\d", s)
            if not co_vb:
                self.add(f"{ma_prefix}-01", self.ten(key), b.o(i, col), v, "Căn cứ không có loại văn bản + số hiệu",
                         "Ghi đủ số, ngày, cơ quan ban hành (ví dụ: Quyết định số 12-QĐ/ĐU ngày 15/3/2026)", MUC_BAT_BUOC, "YC; CV05 PL1-8; PL2-11")
            elif bat_buoc_dieu_khoan and not any(k in s for k in cc["dieu_khoan"]):
                self.add(f"{ma_prefix}-02", self.ten(key), b.o(i, col), v, "Căn cứ không ghi điều, khoản, điểm",
                         "Bổ sung điều, khoản, điểm của văn bản", MUC_BAT_BUOC, "YC; CV05 4.2(1); PL1-8")
            if any(k in s for k in cc["chung_chung"]):
                self.add(f"{ma_prefix}-03", self.ten(key), b.o(i, col), v, "Căn cứ/nguồn ghi chung chung hoặc còn ký tự giữ chỗ",
                         "Dẫn văn bản cụ thể", MUC_BAT_BUOC, "YC; CV05 PL1-8")

    # ================================================================ A1
    def kt_A1(self):
        b = self.B["A1"]
        if not b:
            return
        self._kt_can_cu("A1", "A1", "D")
        cc = self.qt["can_cu"]
        for i, row in enumerate(b.rows):
            g = chuan(row.get("G"))
            if trong(row.get("G")) or any(k in g for k in cc["chung_chung"]):
                self.add("A1-03", self.ten("A1"), b.o(i, "G"), row.get("G"), "Nguồn trống hoặc chung chung", "Ghi nguồn chứng minh/kế thừa cụ thể (văn bản, danh mục, năm)", MUC_BAT_BUOC, "YC; CV05 PL1-8")
            for col in ("E", "F"):
                if trong(row.get(col)):
                    self.add("A1-05", self.ten("A1"), b.o(i, col), "", f"Thiếu {b.cfg['cols'][col].lower()}", "Nhập đủ", MUC_CAN, "CV05 PL2-12")
        # A1-04 nhiệm vụ không có công việc
        a2 = self.B["A2"]
        if a2:
            used = a2.by_key("B")
            for i, row in enumerate(b.rows):
                k = chuan_ma(row.get("A"))
                if k and k not in used and "không phát sinh" not in chuan(row.get("G")) + chuan(row.get("B")):
                    self.add("A1-04", self.ten("A1"), b.o(i, "A"), row.get("A"), "Nhiệm vụ không có công việc nào ở A2 và không ghi 'không phát sinh'",
                             "Bổ sung công việc ở A2, hoặc ghi 'không phát sinh' kèm lý do ở cột Nguồn", MUC_CAN, "CV05 PL2-12; 4.2(1)")
        # trùng mã
        for k, idx in b.by_key("A").items():
            if len(idx) > 1:
                self.add("A1-05", self.ten("A1"), ", ".join(b.o(i, "A") for i in idx), k, "Trùng mã nhiệm vụ", "Sửa mã hoặc xoá dòng thừa", MUC_BAT_BUOC, "HDNL 3.4")

    # ================================================================ A2
    def kt_A2(self):
        b = self.B["A2"]
        if not b:
            return
        qt = self.qt
        self._kt_can_cu("A2", "A2", "H")
        a1 = self.B["A1"].by_key("A") if self.B["A1"] else None
        c1 = self.B["C1"].by_key("A") if self.B["C1"] else None
        sp = defaultdict(list)
        for i, row in enumerate(b.rows):
            if a1 is not None and chuan_ma(row.get("B")) and chuan_ma(row.get("B")) not in a1:
                self.add("A2-04", self.ten("A2"), b.o(i, "B"), row.get("B"), "Mã nhiệm vụ không có ở A1", "Chọn đúng mã ở A1", MUC_BAT_BUOC, "CV05 PL1-8; PL2-12; YC")
            if c1 is not None and chuan_ma(row.get("C")) and chuan_ma(row.get("C")) not in c1:
                self.add("A2-04", self.ten("A2"), b.o(i, "C"), row.get("C"), "Mã VTVL không có ở C1", "Chọn đúng mã ở C1", MUC_BAT_BUOC, "CV05 PL1-8; YC")
            if not trong(row.get("E")):
                sp[chuan(row.get("E"))].append(i)
            ten = chuan(row.get("D"))
            if ten and len(ten.split()) < qt["nguong"]["A2_ten_cong_viec_toi_thieu_tu"]:
                self.add("A2-06", self.ten("A2"), b.o(i, "D"), row.get("D"), "Tên công việc cụt", "Ghi tên công việc cụ thể, đo được", MUC_CAN, "CV05 PL1-8; PL2-13")
        for k, idx in b.by_key("A").items():
            if len(idx) > 1:
                self.add("A2-05", self.ten("A2"), ", ".join(b.o(i, "A") for i in idx), k, "Trùng mã công việc", "Sửa mã hoặc xoá dòng lặp", MUC_CAN, "CV05 PL2-13")
        for k, idx in sp.items():
            if len(idx) > 1:
                vtvl = {chuan_ma(b.rows[i].get("C")) for i in idx}
                o = ", ".join(b.o(i, "E") for i in idx)
                if len(vtvl) > 1:
                    self.add("A2-03", self.ten("A2"), o, b.rows[idx[0]].get("E"), f"Cùng sản phẩm gán cho {len(vtvl)} vị trí: {', '.join(sorted(vtvl))}",
                             "Mỗi vị trí chỉ tính phần việc mình trực tiếp làm; đặt tên sản phẩm riêng cho phần việc đó", MUC_BAT_BUOC, "CV05 'Về sản lượng'; PL1-8")
                else:
                    self.add("A2-02", self.ten("A2"), o, b.rows[idx[0]].get("E"), f"Tên sản phẩm trùng ở {len(idx)} dòng",
                             "Đặt tên sản phẩm cụ thể, không trùng giữa các công việc", MUC_BAT_BUOC, "YC; CV05 PL1-8")
        # A2-07 ghi một kiểu
        for col in ("E", "F", "G"):
            cnt = Counter(chuan(r.get(col)) for r in b.rows if not trong(r.get(col)))
            n = len([1 for r in b.rows if not trong(r.get(col))])
            for v, c in cnt.items():
                if col != "F" and n >= 5 and c >= max(5, n // 2):
                    self.add("A2-07", self.ten("A2"), f"cột {col}", v, f"'{b.cfg['cols'][col]}' ghi cùng một nội dung ở {c}/{n} dòng",
                             "Ghi cụ thể cho từng công việc", MUC_CAN, "CV05 PL1-8; PL2-13")
        # A2-08 sản phẩm lãnh đạo lặp ở chuyên môn
        if self.B["C1"]:
            loai = self._loai_chuc_danh()
            for k, idx in sp.items():
                if len(idx) > 1:
                    kinds = {loai.get(chuan_ma(b.rows[i].get("C")), "cm") for i in idx}
                    if "ld" in kinds and "cm" in kinds:
                        self.add("A2-08", self.ten("A2"), ", ".join(b.o(i, "E") for i in idx), b.rows[idx[0]].get("E"),
                                 "Sản phẩm vừa gắn vị trí lãnh đạo vừa gắn vị trí chuyên môn", "Phần việc lãnh đạo trực tiếp làm chỉ tính một lần ở vị trí lãnh đạo", MUC_BAT_BUOC, "CV05 'Về chức danh lãnh đạo'; PL2-10")

    # ================================================================ C1 / C2
    def _loai_chuc_danh(self):
        """Trả về {mã VTVL: 'truong'|'pho'|'cm'} và map 'ld' cho lãnh đạo."""
        out = {}
        c1 = self.B["C1"]
        if not c1:
            return out
        cd = self.qt["chuc_danh"]
        for row in c1.rows:
            ten = chuan(row.get("B"))
            k = chuan_ma(row.get("A"))
            if not k:
                continue
            if any(ten.startswith(x) or f" {x}" in f" {ten}" for x in cd["cap_pho"]):
                out[k] = "pho"
            elif any(ten.startswith(x) or f" {x} " in f" {ten} " for x in cd["dung_dau"]):
                out[k] = "truong"
            else:
                out[k] = "cm"
        return out

    def _la_hdnd(self, ma_vtvl):
        c1 = self.B["C1"]
        for row in c1.rows:
            if chuan_ma(row.get("A")) == ma_vtvl:
                return any(k in chuan(row.get("B")) for k in self.qt["chuc_danh"]["hdnd_tu_khoa"])
        return False

    def kt_C1C2(self):
        c1, c2 = self.B["C1"], self.B["C2"]
        if c1:
            for i, row in enumerate(c1.rows):
                d = chuan(row.get("D"))
                if d and ("chức năng" in d or "nhiệm vụ" in d) and "vị trí" not in d:
                    self.add("C1-01", self.ten("C1"), c1.o(i, "D"), row.get("D"), "Dẫn quyết định chức năng, nhiệm vụ thay vì quyết định phê duyệt VTVL",
                             "Dẫn quyết định phê duyệt danh mục vị trí việc làm", MUC_CAN, "CV05 PL1-9; PL2-14")
                if "đặc thù" in chuan(row.get("J")) and trong(row.get("F")):
                    self.add("C1-02", self.ten("C1"), c1.o(i, "F"), "", "Chế độ đặc thù nhưng thiếu căn cứ quỹ giờ", "Ghi văn bản quy định chế độ/giờ riêng", MUC_BAT_BUOC, "HDNL 4.4; CV05 'Về định mức'")
            for k, idx in c1.by_key("A").items():
                if len(idx) > 1:
                    self.add("C1-02", self.ten("C1"), ", ".join(c1.o(i, "A") for i in idx), k, "Trùng mã VTVL", "Sửa mã", MUC_BAT_BUOC, "HDNL 4.4")
        if c1 and c2:
            k1 = c1.by_key("A")
            k2 = c2.by_key("A")
            for i, row in enumerate(c2.rows):
                k = chuan_ma(row.get("A"))
                if k and k not in k1:
                    self.add("C2-01", self.ten("C2"), c2.o(i, "A"), row.get("A"), "Mã VTVL ở C2 không có ở C1", "Dán đúng mã theo C1", MUC_BAT_BUOC, "CV05 'Về căn cứ, mô tả, mã'; PL1-9")
                d = chuan(row.get("D"))
                if any(x in d for x in self.qt["bang_cap_tu_khoa"]):
                    self.add("C2-04", self.ten("C2"), c2.o(i, "D"), row.get("D"), "Cột Chuyên môn ghi bằng cấp/trình độ thay vì năng lực",
                             "Ghi năng lực chuyên môn và mức yêu cầu", MUC_CAN, "CV05 PL1-9; PL2-14")
                for col in c2.cfg["required"]:
                    if trong(row.get(col)) and col != "A":
                        self.add("C2-02", self.ten("C2"), c2.o(i, col), "", f"Thiếu '{c2.cfg['cols'][col]}'", "Nhập đủ (cột Lãnh đạo ghi 'Không áp dụng' nếu không phải lãnh đạo)", MUC_CAN, "HDNL 9.4")
            for k in k1:
                if k not in k2:
                    self.add("C2-02", self.ten("C2"), "-", k, f"VTVL '{k}' ở C1 không có dòng C2", "Dùng 'Tạo dòng cho mọi VTVL chưa có' rồi điền", MUC_CAN, "CV05 PL1-9; PL2-14")
            # C2-03 lệch thứ tự
            order1 = [chuan_ma(r.get("A")) for r in c1.rows if chuan_ma(r.get("A"))]
            order2 = [chuan_ma(r.get("A")) for r in c2.rows if chuan_ma(r.get("A")) in k1]
            common = [k for k in order1 if k in k2]
            if order2 and common and order2[: len(common)] != common and sorted(order2) == sorted(common):
                self.add("C2-03", self.ten("C2"), "cột A", "", "Thứ tự mã VTVL ở C2 lệch so với C1", "Sắp cùng thứ tự với C1 để dễ đối chiếu", MUC_CAN, "CV05 PL1-9; PL2-14")

    # ================================================================ B1
    def kt_B1(self):
        b = self.B["B1"]
        if not b:
            return
        qt, ng = self.qt, self.qt["nguong"]
        nt = qt["nguon_thoi_gian"]
        a2 = self.B["A2"]
        cv_ten = {chuan_ma(r.get("A")): chuan(r.get("D")) for r in a2.rows} if a2 else {}
        vt_a2 = {chuan_ma(r.get("A")): chuan_ma(r.get("C")) for r in a2.rows} if a2 else {}
        quy_gio = self.cfg["quy_gio_pho_bien"]
        gio_vtvl = defaultdict(float)
        e_vals = Counter()
        for i, row in enumerate(b.rows):
            g = self._b1_gio_tong(row)
            d = so(row.get("D"))
            c = so(row.get("C"))
            if g is not None and g >= ng["B1_gio_tong_canh_bao"]:
                self.add("B1-01", self.ten("B1"), b.o(i, "G"), row.get("G"), f"Giờ tổng {g:,.0f} ≥ {ng['B1_gio_tong_canh_bao']:,} giờ".replace(",", "."),
                         "Rà lại sản lượng, thời gian; giải trình hoặc tách công việc", MUC_CAN, "YC; CV05 'Về sản lượng'; PL1-3; PL2-7")
            if d is not None and d >= 480 and d % 480 == 0:
                self.add("B1-02", self.ten("B1"), b.o(i, "D"), row.get("D"), f"Phút/SP = {d:.0f} là bội số của 480 (= số ngày × 8 giờ)",
                         "Ghi thời gian người thực sự làm từng phần việc; không nhân ngày × 8 giờ, không tính thời gian chờ", MUC_BAT_BUOC, "CV05 'Về thời gian'; PL1-1; PL2-5")
            if d is not None and d > 0 and d % 10 != 0 and (d + 2) % 10 == 0 or (d is not None and (d + 1) % 10 == 0 and d > 10):
                self.add("B1-10", self.ten("B1"), b.o(i, "D"), row.get("D"), "Phút/SP là số tròn trừ 1–2 phút: dấu hiệu chỉnh để tránh số tròn",
                         "Giữ đúng số đo thực tế", MUC_LUU_Y, "CV05 'Về thời gian'; PL1-12")
            # B1-03 nguồn thời gian
            if not trong(row.get("D")) or not trong(row.get("E")):
                s = str(row.get("I") or "")
                if trong(row.get("I")):
                    self.add("B1-03", self.ten("B1"), b.o(i, "I"), "", "Nguồn thời gian trống", "Ghi số phiếu đo/phiếu khảo sát/văn bản định mức, ngày, người xác nhận", MUC_BAT_BUOC, "CV05 PL1-1; PL2-5")
                elif not (re.search(nt["mau_ngay"], s, re.I) and re.search(r"\d", s)) or not any(k in chuan(s) for k in nt["tu_khoa_phieu"]):
                    self.add("B1-03", self.ten("B1"), b.o(i, "I"), s, "Nguồn thời gian không có số phiếu/văn bản + ngày (+ người xác nhận)",
                             "Ví dụ: 'Phiếu đo thời gian số 03/PĐ ngày 10/9/2026, xác nhận: Trưởng phòng …'", MUC_BAT_BUOC, "CV05 'Về thời gian'; PL1-1; PL2-5")
            # B1-06 sản lượng 0 mà ĐẠT
            if (c is None or c == 0) and chuan(row.get("R")) == "đạt" and (so(row.get("E")) or 0) == 0:
                self.add("B1-06", self.ten("B1"), b.o(i, "C"), row.get("C"), "Sản lượng 0/trống mà trạng thái ĐẠT, không có giờ bổ sung",
                         "Thống kê từ sổ, phần mềm; chỉ ghi 0 khi đã xác minh không phát sinh", MUC_BAT_BUOC, "CV05 'Về sản lượng'; PL1-3; PL2-7")
            # B1-07 nguồn SL
            h = chuan(row.get("H"))
            if c and (trong(row.get("H")) or any(k in h for k in qt["can_cu"]["chung_chung"]) or len(h) < 8):
                self.add("B1-07", self.ten("B1"), b.o(i, "H"), row.get("H"), "Nguồn sản lượng trống hoặc chung chung",
                         "Ghi tên sổ/phần mềm và kỳ 01/10/2025–30/9/2026", MUC_CAN, "CV05 'Về sản lượng'; PL1-3")
            # B1-05 giải trình
            if not trong(row.get("S")) and not chuan(row.get("S")).startswith("đã xác nhận"):
                self.add("B1-05", self.ten("B1"), b.o(i, "S"), row.get("S"), "Giải trình khác chuẩn không bắt đầu bằng 'Đã xác nhận:'", "Ghi đúng mẫu 'Đã xác nhận: …'", MUC_CAN, "HDNL 6.3")
            # B1-09 việc tham mưu điền phút
            ten_cv = cv_ten.get(chuan_ma(row.get("A")), "")
            if d and any(k in ten_cv for k in qt["pp3_tu_khoa_cong_viec"]):
                self.add("B1-09", self.ten("B1"), b.o(i, "D"), row.get("D"), f"Công việc '{ten_cv[:60]}' thuộc nhóm tham mưu/thẩm định/kiểm tra nhưng được điền phút/SP",
                         "Việc khó chuẩn hoá thời gian xác định theo PP3 (có hồ sơ), không bắt buộc quy ra phút", MUC_CAN, "CV05 'Về thời gian'; PL1-12; PL2-19")
            if so(row.get("E")):
                e_vals[so(row.get("E"))] += 1
            kb = chuan_ma(row.get("B")) or vt_a2.get(chuan_ma(row.get("A")), "")
            if g and kb:
                gio_vtvl[kb] += g
        # B1-08 giờ bổ sung gán cố định
        for v, n in e_vals.items():
            if n >= ng["B1_so_dong_lap_gio_bo_sung"]:
                self.add("B1-08", self.ten("B1"), "cột E", v, f"Giờ bổ sung = {v:g} lặp ở {n} dòng: dấu hiệu gán cố định", "Chỉ nhập giờ bổ sung khi có căn cứ, đã xác nhận cho từng công việc", MUC_CAN, "CV05 PL1-3")
        # B1-11/12/13 so với D1
        d1 = self.B["D1"]
        if d1:
            nguoi = Counter(chuan_ma(r.get("B")) for r in d1.rows if chuan_ma(r.get("A")))
            for k, tong in gio_vtvl.items():
                n = nguoi.get(k, 0)
                quy_gio = self._quy_gio(k)
                if n and abs(tong - n * quy_gio) < 0.5:
                    self.add("B1-12", self.ten("B1"), f"VTVL {k}", f"{tong:g}", f"Tổng giờ B1 của vị trí = {n} người × {quy_gio} (khớp tuyệt đối)",
                             "Không chỉnh phút, sản lượng để tổng giờ khớp số người; giữ kết quả lẻ", MUC_BAT_BUOC, "CV05 'Về kết quả tính toán'; PL1-12; PL2-19")
                elif n and tong > n * quy_gio:
                    self.add("B1-11", self.ten("B1"), f"VTVL {k}", f"{tong:g}", f"Tổng giờ B1 {tong:,.0f} > {n} người hiện có × {quy_gio} = {n*quy_gio:,}".replace(",", "."),
                             "Rà lại; có thể tính trùng giữa các vị trí", MUC_CAN, "CV05 'Về sản lượng'; PL2-15")
                elif n and tong / quy_gio < n * 0.5:
                    self.add("B1-13", self.ten("B1"), f"VTVL {k}", f"{tong:g}", f"Tổng giờ B1 ÷ {quy_gio} = {tong/quy_gio:.2f} người, chênh lớn so với {n} người ở D1",
                             "Rà lại, giải trình chênh lệch", MUC_CAN, "CV05 PL2-15")

    # ================================================================ B2
    def kt_B2(self):
        b, b1 = self.B["B2"], self.B["B1"]
        if not b:
            return
        if not b.rows:
            self.add("B2-01", self.ten("B2"), "-", "", "B2 chi tiết bỏ trống", "Kê B2 nếu có số liệu; chưa có thì ghi vào danh sách nội dung chưa đủ căn cứ", MUC_CAN, "CV05 PL1-10; PL2-16")
            return
        sp_b1 = {}
        if b1:
            for r in b1.rows:
                sp_b1[chuan_ma(r.get("A"))] = so(r.get("C"), 0)
        for i, row in enumerate(b.rows):
            k = chuan_ma(row.get("A"))
            e = so(row.get("E"))
            if e is not None and k in sp_b1 and e > sp_b1[k]:
                self.add("B2-02", self.ten("B2"), b.o(i, "E"), row.get("E"), f"SP áp dụng {e:g} > Số SP ở B1 ({sp_b1[k]:g})", "Sửa SP áp dụng không vượt B1", MUC_BAT_BUOC, "HDNL 7.4")
            if chuan(row.get("M")) == "ước lượng tham khảo" and ((so(row.get("N")) or 0) != 0 or (so(row.get("O")) or 0) != 0):
                self.add("B2-03", self.ten("B2"), b.o(i, "M"), row.get("M"), "Ước lượng tham khảo nhưng vẫn có giờ hợp lệ (N/O ≠ 0)", "Ước lượng tham khảo không được tính", MUC_BAT_BUOC, "HDNL 7.4")

    # ================================================================ C3
    def kt_C3(self):
        b = self.B["C3"]
        if not b:
            return
        qt, ng = self.qt, self.qt["nguong"]
        loai = self._loai_chuc_danh()
        S = T = 0.0
        giai_trinh = defaultdict(list)
        d1 = self.B["D1"]
        nguoi = Counter(chuan_ma(r.get("B")) for r in d1.rows if chuan_ma(r.get("A"))) if d1 else Counter()
        gio_vtvl = self._b1_gio_theo_vtvl()
        for i, row in enumerate(b.rows):
            k = chuan_ma(row.get("A"))
            pp = chuan(row.get("B")).upper().replace(" ", "")
            G, H, J = self._c3_G(row, gio_vtvl), so(row.get("H")), so(row.get("J"))
            if pp == "PP2" and G:
                S += G
            if pp in ("PP1", "PP1+PP2") and H:
                S += H
            if J:
                T += J
            kind = loai.get(k, "cm")
            if kind == "truong" and pp != "PP1":
                self.add("C3-04", self.ten("C3"), b.o(i, "B"), row.get("B"), "Vị trí đứng đầu đơn vị/bộ phận không dùng PP1", "Chọn PP1", MUC_BAT_BUOC, "YC; CV05 'Về chức danh lãnh đạo'; PL1-5; PL2-10")
            elif kind == "pho":
                allowed = qt["hdnd_xa_cap_pho_chap_nhan"] if (self.args.loai_don_vi == "xa" and self._la_hdnd(k)) else qt["cap_pho_phuong_phap"]["chap_nhan"]
                if pp not in [a.upper() for a in allowed]:
                    self.add("C3-05", self.ten("C3"), b.o(i, "B"), row.get("B"), f"Vị trí cấp phó dùng {pp or '(trống)'}; chấp nhận: {', '.join(allowed)}",
                             "Chọn phương pháp đúng cho cấp phó (xem config/quy_tac.json: cap_pho_phuong_phap)", MUC_BAT_BUOC, "YC; CV05 PL1-5")
            elif kind == "cm" and pp and pp != "PP2" and "PP3" not in pp:
                self.add("C3-06", self.ten("C3"), b.o(i, "B"), row.get("B"), "Vị trí chuyên môn không dùng PP2", "Chọn PP2 (trừ việc tham mưu khó chuẩn hoá đã chuyển PP3 có hồ sơ)", MUC_CAN, "YC; CV05 PL1-12")
            if "PP3" in pp:
                self.add("C3-07", self.ten("C3"), b.o(i, "B"), row.get("B"), "Sử dụng PP3", "Phải có hồ sơ PP3: danh mục sản phẩm, phân tích phần việc, phương án phân công, đối chiếu nhiệm vụ tương tự, trình thẩm định", MUC_CAN, "YC; CV05 'Về thời gian'")
            if kind == "cm" and pp == "PP2" and G is not None and J is not None and J > G and J - G >= 1:
                self.add("C3-08", self.ten("C3"), b.o(i, "J"), row.get("J"), f"J = {J:g} lớn hơn người quy đổi G = {G:.4f} từ 1 người trở lên (làm tròn tăng)",
                         "Cột J theo kết quả quy đổi, không làm tròn tăng từng dòng; giải trình nếu khác", MUC_BAT_BUOC, "CV05 PL1-5")
            kk = chuan(row.get("K"))
            if kk:
                giai_trinh[kk].append(i)
            if kind in ("truong", "pho"):
                cc = qt["can_cu"]
                if not (any(x in kk for x in cc["loai_van_ban"]) and re.search(r"\d", kk)):
                    self.add("C3-10", self.ten("C3"), b.o(i, "K"), row.get("K"), "Vị trí lãnh đạo: giải trình không dẫn văn bản quy định số lượng (số, ngày, điều, khoản)",
                             "Ghi văn bản quy định số lượng, quyết định của cấp có thẩm quyền, phạm vi phụ trách từng cấp phó", MUC_BAT_BUOC, "CV05 'Về chức danh lãnh đạo'; PL1-5")
                if kind == "pho" and H is not None and nguoi.get(k) == int(H) and "phụ trách" not in kk:
                    self.add("C3-11", self.ten("C3"), b.o(i, "H"), row.get("H"), "Số cấp phó = số người đang bố trí (D1), giải trình không nêu phạm vi phụ trách",
                             "Số cấp phó là số thực cần theo nhiệm vụ, phạm vi phụ trách; không lấy số đang bố trí, không mặc nhiên bằng mức tối đa", MUC_CAN, "CV05 'Về chức danh lãnh đạo'; PL2-10")
            if kind == "cm" and J is not None and G is not None and nguoi.get(k) and J == nguoi.get(k) and abs(G - J) >= 0.5:
                self.add("C3-12", self.ten("C3"), b.o(i, "J"), row.get("J"), f"J = số người hiện có ({nguoi.get(k)}) trong khi G = {G:.4f} lệch rõ",
                         "Không điều chỉnh để khớp số người hiện có; giữ kết quả quy đổi, giải trình", MUC_CAN, "CV05 'Về kết quả tính toán'; PL2-8")
            m = chuan(row.get("M"))
            if m and m != "đạt" and not kk:
                self.add("C3-13", self.ten("C3"), b.o(i, "M"), row.get("M"), "Còn cảnh báo ở cột M mà không có giải trình ở K", "Xử lý hết cảnh báo hoặc giải trình cụ thể", MUC_BAT_BUOC, "CV05 mục II; PL2-18")
        for kk, idx in giai_trinh.items():
            if len(idx) >= ng["C3_so_dong_lap_giai_trinh"]:
                self.add("C3-09", self.ten("C3"), ", ".join(b.o(i, "K") for i in idx), kk[:60], f"Giải trình giống hệt nhau ở {len(idx)} vị trí",
                         "Viết giải trình riêng cho từng vị trí bằng số liệu của chính vị trí đó", MUC_BAT_BUOC, "CV05 'Về chức danh lãnh đạo'; PL1-5; PL2-18")
        # tổng
        if T:
            if S < T:
                self.add("C3-01", self.ten("C3"), "tổng cột G/H/J", f"S={S:.2f}; ΣJ={T:g}", "Tổng G (PP2) + H (PP1, PP1+PP2) nhỏ hơn tổng J", "Rà lại kết luận J: vượt số tính được theo phương pháp", MUC_CAN, "YC")
            elif S > T * (1 + ng["C3_chenh_lech_phan_tram"] / 100):
                self.add("C3-02", self.ten("C3"), "tổng cột G/H/J", f"S={S:.2f}; ΣJ={T:g}", f"Tổng G+H vượt tổng J quá {ng['C3_chenh_lech_phan_tram']}% ({(S/T-1)*100:.1f}%)", "Rà lại, giải trình chênh lệch", MUC_CAN, "YC")
            if self.args.bien_che is not None:
                if T > self.args.bien_che:
                    self.add("C3-03", self.ten("C3"), "tổng cột J", f"ΣJ={T:g}; biên chế giao={self.args.bien_che:g}", "Tổng người cần bố trí vượt biên chế được giao",
                             "Rà lại; nhu cầu tăng ghi ở phần giải trình, chỉ đề nghị trên số liệu đã kiểm chứng", MUC_BAT_BUOC, "YC; CV05 'Về kiểm chứng'")
            else:
                self.bo_qua.append("C3-03 (chưa nhập --bien-che)")

    # ================================================================ D1
    def kt_D1(self):
        b = self.B["D1"]
        if not b:
            return
        qt, ng = self.qt, self.qt["nguong"]
        n = len([r for r in b.rows if chuan_ma(r.get("A"))])
        if self.args.so_nguoi is not None:
            if n != self.args.so_nguoi:
                self.add("D1-01", self.ten("D1"), "số dòng", f"{n} dòng; có mặt {self.args.so_nguoi}", "Số dòng D1 khác số người có mặt tại 30/9/2026",
                         "Mỗi người có mặt là 01 dòng; không kê vị trí thay người", MUC_BAT_BUOC, "YC; CV05 'Về nhân lực'; PL2-3")
        else:
            self.bo_qua.append("D1-01 (chưa nhập --so-nguoi)")
        e_vals = Counter()
        c1 = self.B["C1"].by_key("A") if self.B["C1"] else None
        for i, row in enumerate(b.rows):
            a = chuan(row.get("A"))
            if any(x in a for x in qt["D1_dau_hieu_khong_phai_nguoi"]) or (c1 is not None and chuan_ma(row.get("A")) in c1):
                self.add("D1-02", self.ten("D1"), b.o(i, "A"), row.get("A"), "Dòng kê vị trí thay cho người / vị trí chưa có người", "Xoá dòng; chỉ kê người thực có mặt", MUC_BAT_BUOC, "CV05 'Về nhân lực'; PL1-4; PL2-3")
            E, F, D = so(row.get("E")), so(row.get("F")), so(row.get("D"))
            if D is None:
                D = so(row.get("K")) or self._quy_gio(chuan_ma(row.get("B")))
            if E is not None:
                e_vals[E] += 1
            if D and E is not None and (E + (F or 0)) > D + 0.01:
                self.add("D1-05", self.ten("D1"), b.o(i, "E"), f"E={E:g}, F={F or 0:g}, D={D:g}", "Vượt quỹ giờ (E + F > D)", "Giảm E/F cho đúng thực tế phân công", MUC_BAT_BUOC, "HDNL 11.4; CV05 PL2-18")
            if F and F > 0 and trong(row.get("H")):
                self.add("D1-06", self.ten("D1"), b.o(i, "F"), row.get("F"), "Có giờ việc khác nhưng không nêu kiêm nhiệm gì (cột Nguồn trống)", "Ghi rõ nội dung kiêm nhiệm, văn bản phân công", MUC_CAN, "CV05 'Về nhân lực'; PL2-4")
            for col in ("G", "H"):
                if trong(row.get(col)):
                    self.add("D1-07", self.ten("D1"), b.o(i, col), "", f"Thiếu '{b.cfg['cols'][col]}'", "Ghi đúng hồ sơ cán bộ", MUC_BAT_BUOC, "CV05 PL2-4")
        for k, idx in b.by_key("A").items():
            if len(idx) > 1:
                self.add("D1-03", self.ten("D1"), ", ".join(b.o(i, "A") for i in idx), k, "Một người ở nhiều dòng", "Mỗi người 01 dòng; kiêm nhiệm ghi ở cột F", MUC_BAT_BUOC, "CV05 'Về nhân lực'; HDNL 11.5")
        for v, c in e_vals.items():
            if c >= ng["D1_so_dong_lap_gio"] and n >= 4 and c >= n * 0.6 and v != self.cfg["quy_gio_pho_bien"]:
                self.add("D1-04", self.ten("D1"), "cột E", v, f"Giờ cho VTVL = {v:g} ghi đồng loạt ở {c}/{n} người", "Ghi theo văn bản phân công thực tế từng người", MUC_BAT_BUOC, "CV05 'Về nhân lực'; PL1-4; PL2-4")

    # ================================================================ Dự báo, D2, liên biểu
    def kt_DB_D2(self):
        db, d2, c3 = self.B["DB"], self.B["D2"], self.B["C3"]
        nam = self.cfg["nam_du_bao"]
        if db:
            seen = defaultdict(set)
            for i, row in enumerate(db.rows):
                k = chuan_ma(row.get("A"))
                y = so(row.get("B"))
                if k and y:
                    seen[k].add(int(y))
                if k and trong(row.get("C")):
                    self.add("DB-02", self.ten("DB"), db.o(i, "C"), "", "Biến động trống", "Nhập 0 nếu không đổi, kèm giải trình", MUC_BAT_BUOC, "HDNL 12.4")
            if c3:
                for k in c3.by_key("A"):
                    thieu = [y for y in nam if y not in seen.get(k, set())]
                    if thieu and len(thieu) < len(nam):
                        self.add("DB-01", self.ten("DB"), f"VTVL {k}", "", f"Thiếu năm {', '.join(map(str, thieu))}", "Mỗi cặp VTVL – năm một dòng, đủ 5 năm", MUC_CAN, "HDNL 12.5")
                    elif thieu:
                        self.add("LB-05", self.ten("DB"), f"VTVL {k}", "", "VTVL ở C3 không có dòng Dự báo", "Lập dự báo 2027–2031", MUC_CAN, "YC; HDNL 12")
        if d2:
            c3_tt = {chuan_ma(r.get("A")): chuan(r.get("L")) for r in c3.rows} if c3 else {}
            e_vals = Counter()
            g_eq = 0
            n_rows = 0
            for i, row in enumerate(d2.rows):
                k = chuan_ma(row.get("A"))
                if not k:
                    continue
                n_rows += 1
                C, D, E, G = so(row.get("C")), so(row.get("D")), so(row.get("E")), so(row.get("G"))
                if E:
                    e_vals[E] += 1
                if G and G > 0 and c3_tt.get(k, "") != "đã kiểm chứng":
                    self.add("D2-03", self.ten("D2"), d2.o(i, "G"), row.get("G"), f"Đề nghị tăng khi C3 của VTVL còn '{c3_tt.get(k, '(không có)')}'",
                             "Chỉ đề nghị tăng trên số liệu đã kiểm chứng; ưu tiên bố trí lại, kiêm nhiệm, điều chuyển", MUC_BAT_BUOC, "CV05 'Về kiểm chứng'; PL1-13; PL2-20")
                if C is not None and D is not None and G is not None and G != 0 and abs(G - (C - D)) < 1e-9:
                    g_eq += 1
            if n_rows and g_eq >= max(3, n_rows * 0.6):
                self.add("D2-04", self.ten("D2"), "cột G", "", f"G = C − D ở {g_eq}/{n_rows} dòng (đề nghị mặc nhiên theo chênh lệch)", "Chỉ đề nghị sau soát xét, giải trình, thẩm định", MUC_CAN, "HDNL 13.6")
            if self.args.bien_che and e_vals:
                for v, c in e_vals.items():
                    if abs(v - self.args.bien_che) < 1e-9 and c >= 2:
                        self.add("D2-02", self.ten("D2"), "cột E", v, f"Tổng biên chế được giao ({v:g}) lặp vào {c} dòng vị trí", "Biên chế giao tổng ghi ở cấp giao, không lặp vào từng vị trí", MUC_BAT_BUOC, "CV05 PL1-6; HDNL 13.6")
                tong_E_nam = defaultdict(float)
                for row in d2.rows:
                    if so(row.get("B")) and so(row.get("E")):
                        tong_E_nam[int(so(row.get("B")))] += so(row.get("E"))
                for y, t in tong_E_nam.items():
                    if t > self.args.bien_che + 1e-9:
                        self.add("D2-02", self.ten("D2"), f"năm {y}", f"{t:g}", f"Tổng biên chế giao các vị trí năm {y} = {t:g} > biên chế giao {self.args.bien_che:g}", "Chia đúng, không vượt tổng", MUC_BAT_BUOC, "CV05 PL1-6")
            d1 = self.B["D1"]
            if d1:
                hd = {chuan_ma(r.get("B")) for r in d1.rows if "hợp đồng" in chuan(r.get("C"))}
                for i, row in enumerate(d2.rows):
                    if chuan_ma(row.get("A")) in hd and so(row.get("E")):
                        self.add("D2-05", self.ten("D2"), d2.o(i, "E"), row.get("E"), "Vị trí có hợp đồng lao động (D1) nhưng ghi biên chế giao", "Hợp đồng lao động ghi riêng, không tính vào biên chế giao", MUC_BAT_BUOC, "CV05 PL1-13")

    def kt_LB(self):
        a2, b1, c1, c3, d1, db, d2 = (self.B[k] for k in ("A2", "B1", "C1", "C3", "D1", "DB", "D2"))
        if a2 and b1:
            used = b1.by_key("A")
            for i, row in enumerate(a2.rows):
                k = chuan_ma(row.get("A"))
                if k and k not in used:
                    self.add("LB-01", self.ten("A2"), a2.o(i, "A"), row.get("A"), "Công việc không có dòng khối lượng ở B1", "Kê khối lượng ở B1 hoặc xoá công việc không phát sinh", MUC_CAN, "YC")
            # LB-06 cột tự lấy
            vt = {chuan_ma(r.get("A")): chuan_ma(r.get("C")) for r in a2.rows}
            for key in ("B1", "B2"):
                b = self.B[key]
                if not b:
                    continue
                for i, row in enumerate(b.rows):
                    k = chuan_ma(row.get("A"))
                    if k in vt and chuan_ma(row.get("B")) and chuan_ma(row.get("B")) != vt[k]:
                        self.add("LB-06", self.ten(key), b.o(i, "B"), row.get("B"), f"Mã VTVL không khớp A2 ({vt[k]}): cột tự lấy bị nhập tay", "Khôi phục công thức/giá trị tự lấy từ A2", MUC_BAT_BUOC, "YC; HDNL 6.3")
        if c1:
            k1 = c1.by_key("A")
            if a2:
                used = a2.by_key("C")
                for k in k1:
                    if k not in used:
                        self.add("LB-02", self.ten("C1"), c1.o(k1[k][0], "A"), k, "VTVL không có công việc nào ở A2", "Gán công việc cho vị trí hoặc xem lại danh mục VTVL", MUC_CAN, "YC")
            if c3:
                k3 = c3.by_key("A")
                for k in k1:
                    if k not in k3:
                        self.add("LB-03", self.ten("C1"), c1.o(k1[k][0], "A"), k, "VTVL không có dòng nhu cầu ở C3", "Lập dòng C3 cho vị trí", MUC_CAN, "YC")
                qg = {chuan_ma(r.get("A")): so(r.get("E")) for r in c1.rows}
                for i, row in enumerate(c3.rows):
                    k = chuan_ma(row.get("A"))
                    F = so(row.get("F"))
                    if k in qg and qg[k] and F and abs(F - qg[k]) > 0.5:
                        self.add("LB-06", self.ten("C3"), c3.o(i, "F"), row.get("F"), f"Giờ PP2 không khớp C1 cột E ({qg[k]:g})", "Khôi phục công thức tự lấy từ C1", MUC_BAT_BUOC, "YC; HDNL 10.3")
        if c3 and d1:
            nguoi = Counter(chuan_ma(r.get("B")) for r in d1.rows if chuan_ma(r.get("A")))
            for i, row in enumerate(c3.rows):
                k = chuan_ma(row.get("A"))
                if k and so(row.get("J")) and not nguoi.get(k) and not chuan(row.get("K")):
                    self.add("LB-04", self.ten("C3"), c3.o(i, "A"), k, "Vị trí có người cần bố trí nhưng không có người ở D1 và không giải trình", "Giải trình ở cột K", MUC_CAN, "YC")
        if c3 and d2:
            kd2 = d2.by_key("A")
            for k, idx in c3.by_key("A").items():
                if k not in kd2:
                    self.add("LB-05", self.ten("D2"), f"VTVL {k}", "", "VTVL ở C3 không có dòng D2", "Lập phương án D2 cho từng năm", MUC_CAN, "YC; HDNL 13")
        if db and d2:
            g = {(chuan_ma(r.get("A")), int(so(r.get("B")) or 0)): so(r.get("H")) for r in db.rows}
            for i, row in enumerate(d2.rows):
                key = (chuan_ma(row.get("A")), int(so(row.get("B")) or 0))
                C = so(row.get("C"))
                if key in g and g[key] is not None and C is not None and abs(C - g[key]) > 1e-9:
                    self.add("LB-06", self.ten("D2"), d2.o(i, "C"), row.get("C"), f"Số người cần bố trí không khớp Dự báo cột H ({g[key]:g})", "Khôi phục công thức tự lấy từ Dự báo", MUC_BAT_BUOC, "YC; CV05 PL1-6")

    # ================================================================ chạy
    def chay(self):
        for fn in (self.kt_T, self.kt_TT, self.kt_A1, self.kt_A2, self.kt_C1C2, self.kt_B1, self.kt_B2, self.kt_C3, self.kt_D1, self.kt_DB_D2, self.kt_LB):
            try:
                fn()
            except Exception as e:  # không để một quy tắc hỏng làm mất cả báo cáo
                self.add("LỖI-SCRIPT", fn.__name__, "-", "", f"Quy tắc {fn.__name__} gặp lỗi: {e!r}", "Báo cho người bảo trì skill", MUC_LUU_Y, "-")
        # khử trùng lặp
        seen, out = set(), []
        for p in self.ph:
            k = (p.ma, p.bieu, p.o, p.noi_dung)
            if k not in seen:
                seen.add(k)
                out.append(p)
        self.ph = sorted(out, key=lambda p: (p.bieu, MUC_THU_TU[p.muc] if p.muc in MUC_THU_TU else 9, p.ma, p.o))
        return self.ph


# ----------------------------------------------------------------- xuất
HEAD = ["Mức", "Mã QT", "Biểu", "Ô", "Giá trị hiện tại", "Sai gì", "Sửa thế nào", "Căn cứ"]


def xuat_xlsx(kt, path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Phat_hien"
    ws.append(HEAD)
    for p in kt.ph:
        ws.append(p.row())
    for col, w in zip("ABCDEFGH", (11, 9, 34, 16, 28, 60, 60, 40)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)


def xuat_docx(kt, path, meta):
    try:
        from docx import Document
        from docx.shared import Pt
    except ImportError:
        print("Thiếu python-docx, bỏ qua .docx (pip install python-docx)")
        return False
    d = Document()
    st = d.styles["Normal"]
    st.font.name = "Times New Roman"
    st.font.size = Pt(12)
    d.add_heading("BÁO CÁO KIỂM TRA LOGIC BỘ BIỂU VỊ TRÍ VIỆC LÀM", 1)
    d.add_paragraph(f"Tệp kiểm tra: {os.path.basename(kt.path)}")
    d.add_paragraph(f"Thời điểm: {meta['thoi_diem']}    Công cụ: vtvl-kiem-tra v{VERSION}")
    d.add_paragraph("Căn cứ: Hướng dẫn số 07-HD/BTCTW ngày 19/10/2026; Công văn số 05-CV/BCĐ ngày 03/11/2026 (Phụ lục 1, 2); Yêu cầu kiểm tra dữ liệu; Hướng dẫn nhập liệu 09 biểu v1.1.")
    tham_so = f"Tham số: biên chế được giao = {meta['bien_che']}; số người có mặt 30/9/2026 = {meta['so_nguoi']}; loại đơn vị = {meta['loai_don_vi']}; đã thẩm định = {'có' if meta['da_tham_dinh'] else 'chưa'}."
    d.add_paragraph(tham_so)
    cnt = Counter(p.muc for p in kt.ph)
    d.add_paragraph(f"Tổng số phát hiện: {len(kt.ph)}  (BẮT BUỘC: {cnt.get(MUC_BAT_BUOC,0)} · CẦN: {cnt.get(MUC_CAN,0)} · LƯU Ý: {cnt.get(MUC_LUU_Y,0)})")
    if cnt.get(MUC_BAT_BUOC, 0):
        d.add_paragraph("Hồ sơ còn lỗi thuộc mức BẮT BUỘC sẽ bị trả lại, không đưa vào tổng hợp (CV05, mục II).")
    if kt.bo_qua:
        d.add_paragraph("Quy tắc bỏ qua do thiếu tham số: " + "; ".join(kt.bo_qua))
    d.add_paragraph("Nguyên tắc: báo cáo này chỉ cảnh báo, không sửa dữ liệu. Người kê khai tự sửa trên bộ biểu/app và chịu trách nhiệm về số liệu.")
    by = defaultdict(list)
    for p in kt.ph:
        by[p.bieu].append(p)
    for bieu in by:
        d.add_heading(bieu, 2)
        items = by[bieu]
        c = Counter(p.muc for p in items)
        d.add_paragraph(f"{len(items)} phát hiện (BẮT BUỘC {c.get(MUC_BAT_BUOC,0)}, CẦN {c.get(MUC_CAN,0)}, LƯU Ý {c.get(MUC_LUU_Y,0)})")
        t = d.add_table(rows=1, cols=7)
        t.style = "Table Grid"
        for j, h in enumerate(["Mức", "Mã", "Ô", "Giá trị", "Sai gì", "Sửa thế nào", "Căn cứ"]):
            t.rows[0].cells[j].text = h
        for p in items:
            r = t.add_row().cells
            for j, v in enumerate([p.muc, p.ma, p.o, p.gia_tri, p.noi_dung, p.sua, p.can_cu]):
                r[j].text = str(v)
        for row in t.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(9)
    d.add_heading("Nội dung không kiểm tra tự động được (người kiểm tra tự đối chiếu)", 2)
    for s in ("Phiếu đo, phiếu khảo sát ghi ở B1 cột I có tồn tại thật và đúng nội dung (CV05 4.1: kiểm tra chọn mẫu).",
              "Số liệu, giải trình không sao chép từ hồ sơ mẫu đơn vị khác (CV05 PL2-20).",
              "Đơn vị sự nghiệp có định mức chuyên ngành: phép tính PP1 theo văn bản ngành (CV05 PL1-11; PL2-17).",
              "Người đứng đầu đã ký xác nhận từng biểu và Bảng tự kiểm tra (CV05 mục II).",
              "Đối chiếu D1 giữa khối Đảng và khối chính quyền cùng cấp để mỗi người chỉ tính một lần (CV05 PL2-4): cần cả hai tệp."):
        d.add_paragraph(s, style="List Bullet")
    d.save(path)
    return True


def main():
    ap = argparse.ArgumentParser(description="Kiểm tra logic Bộ biểu VTVL (chỉ đọc).")
    ap.add_argument("file", nargs="+", help="Một hoặc nhiều bộ biểu .xlsx")
    ap.add_argument("--bien-che", type=float, default=None, help="Tổng biên chế được giao (cho C3-03, D2-02)")
    ap.add_argument("--so-nguoi", type=int, default=None, help="Số người có mặt tại 30/9/2026 (cho D1-01)")
    ap.add_argument("--loai-don-vi", choices=["xa", "tinh", "su_nghiep"], default="xa")
    ap.add_argument("--da-tham-dinh", action="store_true", help="Hồ sơ đã được cơ quan đầu mối thẩm định (bỏ qua TT-01..03)")
    ap.add_argument("--out", default=None, help="Tên tệp kết quả, không đuôi (mặc định: <tên tệp>_ket_qua). Nhiều tệp đầu vào: --out là THƯ MỤC chứa kết quả")
    ap.add_argument("--json", action="store_true", help="In JSON ra stdout thay vì bảng tóm tắt")
    args = ap.parse_args()
    files = args.file
    nhieu = len(files) > 1
    if nhieu and args.out:
        os.makedirs(args.out, exist_ok=True)
    tong_hop = []
    for f in files:
        if not os.path.exists(f):
            print(f"Không thấy tệp: {f}")
            continue
        if nhieu:
            out = os.path.join(args.out or os.path.dirname(f) or ".", os.path.splitext(os.path.basename(f))[0] + "_ket_qua")
        else:
            out = args.out or os.path.splitext(f)[0] + "_ket_qua"
        data = chay_mot(f, args, out)
        tong_hop.append(data)
    if nhieu:
        print("\n=== TỔNG HỢP %d tệp ===" % len(tong_hop))
        for d in tong_hop:
            t = d["tong"]
            print(f"  {os.path.basename(d['file'])}: BẮT BUỘC {t.get(MUC_BAT_BUOC,0)} / CẦN {t.get(MUC_CAN,0)} / LƯU Ý {t.get(MUC_LUU_Y,0)}")
        with open(os.path.join(args.out or ".", "tong_hop.json"), "w", encoding="utf-8") as fh:
            json.dump([{"file": d["file"], "tong": d["tong"], "bo_qua": d["bo_qua"]} for d in tong_hop], fh, ensure_ascii=False, indent=1)


def chay_mot(f, args, out):
    kt = KiemTra(f, args)
    kt.chay()
    meta = {"thoi_diem": datetime.now().strftime("%d/%m/%Y %H:%M"), "bien_che": args.bien_che, "so_nguoi": args.so_nguoi,
            "loai_don_vi": args.loai_don_vi, "da_tham_dinh": args.da_tham_dinh, "version": VERSION}
    xuat_xlsx(kt, out + ".xlsx")
    co_docx = xuat_docx(kt, out + ".docx", meta)
    data = {"meta": meta, "file": f, "bo_qua": kt.bo_qua,
            "tong": dict(Counter(p.muc for p in kt.ph)),
            "phat_hien": [dict(zip(["muc", "ma", "bieu", "o", "gia_tri", "noi_dung", "sua", "can_cu"], p.row())) for p in kt.ph]}
    with open(out + ".json", "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=1)
    if args.json:
        print(json.dumps(data, ensure_ascii=False, indent=1))
        return data
    cnt = data["tong"]
    print(f"vtvl-kiem-tra v{VERSION} – {os.path.basename(f)}")
    print(f"Phát hiện: {len(kt.ph)}  |  BẮT BUỘC {cnt.get(MUC_BAT_BUOC,0)}  CẦN {cnt.get(MUC_CAN,0)}  LƯU Ý {cnt.get(MUC_LUU_Y,0)}")
    if kt.bo_qua:
        print("Bỏ qua:", "; ".join(kt.bo_qua))
    by = defaultdict(Counter)
    for p in kt.ph:
        by[p.bieu][p.muc] += 1
    for b, c in by.items():
        print(f"  {b}: BB {c.get(MUC_BAT_BUOC,0)} / CẦN {c.get(MUC_CAN,0)} / LƯU Ý {c.get(MUC_LUU_Y,0)}")
    print("Kết quả:", out + ".docx" if co_docx else "(không có docx)", "|", out + ".xlsx", "|", out + ".json")
    return data


if __name__ == "__main__":
    main()
