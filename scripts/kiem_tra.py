#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Bọc tương thích: python scripts/kiem_tra.py ... (code thật ở vtvl_kiem_tra/kiem_tra.py)."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from vtvl_kiem_tra.kiem_tra import main  # noqa: E402

if __name__ == "__main__":
    main()
