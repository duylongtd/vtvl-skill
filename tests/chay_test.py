#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Chạy test: mau_dung.xlsx không có BẮT BUỘC; mau_sai.xlsx bắt đủ mọi mã trong mau_sai_ky_vong.json."""
import json
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPT = os.path.join(os.path.dirname(HERE), "scripts", "kiem_tra.py")


def chay(f, extra=()):
    out = os.path.join(tempfile.mkdtemp(), "kq")
    subprocess.run([sys.executable, SCRIPT, f, "--bien-che", "3", "--so-nguoi", "3", "--out", out, *extra], check=True, stdout=subprocess.DEVNULL)
    with open(out + ".json", encoding="utf-8") as fh:
        return json.load(fh)


def main():
    loi = 0
    d = chay(os.path.join(HERE, "mau_dung.xlsx"))
    bb = [p for p in d["phat_hien"] if p["muc"] == "BẮT BUỘC"]
    if bb:
        loi += 1
        print("FAIL mau_dung: có BẮT BUỘC:", [(p["ma"], p["o"]) for p in bb])
    else:
        print("OK   mau_dung: 0 BẮT BUỘC,", d["tong"])
    s = chay(os.path.join(HERE, "mau_sai.xlsx"))
    with open(os.path.join(HERE, "mau_sai_ky_vong.json"), encoding="utf-8") as fh:
        kv = set(json.load(fh))
    got = {p["ma"] for p in s["phat_hien"]}
    sot = sorted(kv - got)
    if sot:
        loi += 1
        print("FAIL mau_sai: bỏ sót", sot)
    else:
        print(f"OK   mau_sai: bắt đủ {len(kv)}/{len(kv)} quy tắc kỳ vọng; tổng", s["tong"])
    sys.exit(1 if loi else 0)


if __name__ == "__main__":
    main()
