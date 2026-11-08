# -*- coding: utf-8 -*-
"""
品种 / 品系判定工具

与母猪综合查询、猪只转移使用完全相同的判定规则 (cleaning.py):
    品种 = classify_breed_label()    品系 = classify_line_label()
另外给出解析出的品系编码和血统构成, 方便核对。
"""
from __future__ import annotations

import os
import re

import pandas as pd

from .cleaning import (classify_breed_label, classify_line_label, decode_line_lineage,
                       normalize_breed_scalar, resolve_line_code)
from .rules import BreedRules

# 自定义规则的保存位置(与界面设置放在一起)
RULES_PATH = os.path.join(os.path.expanduser("~"), ".fpf_breed_rules.json")


def load_saved_rules(path: str = RULES_PATH) -> BreedRules | None:
    """读取已保存的自定义规则; 没有保存过返回 None"""
    if not os.path.exists(path):
        return None
    return BreedRules.load(path)

COLUMNS = ["输入", "品种", "品系", "品系编码", "血统构成"]


def split_inputs(text: str) -> list[str]:
    """一行一个, 也可以用逗号 / 顿号 / 分号 / 制表符分隔; 去掉空项"""
    parts = re.split(r"[\r\n,，、;；\t]+", text or "")
    return [p.strip() for p in parts if p.strip()]


def lineage_text(x: str | None) -> str:
    """'S332' -> '托佩克 75% + 美系 25%'; 无法解析返回空串"""
    comp = decode_line_lineage(resolve_line_code(normalize_breed_scalar(x)))
    if not comp:
        return ""
    items = sorted(comp.items(), key=lambda kv: -kv[1])
    return " + ".join(f"{k} {v * 100:g}%" for k, v in items)


def judge(values) -> pd.DataFrame:
    s = pd.Series(list(values), dtype="object")
    return pd.DataFrame({
        "输入": s,
        "品种": classify_breed_label(s),
        "品系": classify_line_label(s),
        "品系编码": s.map(lambda v: resolve_line_code(normalize_breed_scalar(v)) or ""),
        "血统构成": s.map(lineage_text),
    }, columns=COLUMNS)
