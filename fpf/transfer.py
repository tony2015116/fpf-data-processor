# -*- coding: utf-8 -*-
"""
猪只转移数据处理

对应的 R 代码:
    mutate(all_info = get_path_info(`_file`, 5)) |>
    extract(col = all_info,
            into = c("单据编号","转移类型","转移方式","转出场","转入场","猪只类型","转移数量","转移时间"),
            regex = "^([^_]+)_([^_]+)_([^_]+)_([^-]+)-([^_]+)_([^_]+)_(\\d+)头_(\\d{8})\\.xlsx$",
            remove = TRUE) |>
    mutate(转移数量 = as.integer(转移数量),
           转移时间 = as.Date(转移时间, format = "%Y%m%d")) |>
    mutate(品种 = classify_breed_label(品种品系),
           品系 = classify_line_label(品种品系))

与 R 版的一处差别: R 的 get_path_info(`_file`, 5) 取路径的第 5 段, 只有文件恰好在第 5 层
目录时才是文件名; 这里直接取文件名, 文件放在哪一层都能解析。
"""
from __future__ import annotations

import datetime as _dt
import os
import re

import numpy as np
import pandas as pd

from .cleaning import classify_breed_label, classify_line_label
from .report import as_reporter, esc, fmt_n
from .rules import active_rules

TRANSFER_FIELDS = ["单据编号", "转移类型", "转移方式", "转出场", "转入场",
                   "猪只类型", "转移数量", "转移时间"]

# 与 R 正则相同; \d 写成 [0-9], 避免 Python 把全角数字也当数字
FILENAME_RE = re.compile(
    r"([^_]+)_([^_]+)_([^_]+)_([^-]+)-([^_]+)_([^_]+)_([0-9]+)头_([0-9]{8})\.xlsx")
FILENAME_PATTERN_HINT = "单据编号_转移类型_转移方式_转出场-转入场_猪只类型_数量头_YYYYMMDD.xlsx"

REQUIRED = ["品种品系"]


def parse_transfer_filename(name: str | None) -> list | None:
    """'ZY001_调拨_汽运_金沙一场-郭祥一场_后备母猪_120头_20260105.xlsx' -> 8 个字段; 不匹配返回 None"""
    if not name:
        return None
    m = FILENAME_RE.fullmatch(name)
    return list(m.groups()) if m else None


def _to_int(x: str | None):
    if x is None:
        return pd.NA
    v = int(x)
    return v if v <= 2147483647 else pd.NA      # R as.integer 超出范围 -> NA


def _to_date(x: str | None):
    if x is None:
        return pd.NaT
    try:
        return pd.Timestamp(_dt.datetime.strptime(x, "%Y%m%d"))
    except ValueError:                           # 如 20261340 -> NA, 与 R as.Date 一致
        return pd.NaT


def _set_col(df: pd.DataFrame, name: str, values) -> None:
    """同 dplyr::mutate: 已有同名列则原位替换, 否则加在最后"""
    df[name] = values


def add_breed_line(df: pd.DataFrame, rep, src: str = "品种品系") -> None:
    """按当前生效的规则(内置或自定义)判定 品种 / 品系 两列, 并输出汇总; 原位修改 df
    猪只转移、母猪档案等共用"""
    _set_col(df, "品种", classify_breed_label(df[src]).to_numpy())
    _set_col(df, "品系", classify_line_label(df[src]).to_numpy())
    for col in ("品种", "品系"):
        vc = df[col].value_counts()
        rep.info(f"{{.field {col}}}: " + ", ".join(f"{esc(k)} {fmt_n(v)}" for k, v in vc.items()))
    fb = active_rules().breed_fallback
    n_other = int((df["品种"] == fb).sum())
    if n_other:
        ex = df.loc[df["品种"] == fb, src].dropna().astype(str).unique()[:5]
        rep.warn(f"{fmt_n(n_other)} 条记录的品种无法识别, 记为 {{.val {esc(fb)}}}",
                 ["如 " + "、".join(f"{{.val {esc(x)}}}" for x in ex)] if len(ex) else [])


def process_transfer(raw: pd.DataFrame, file_col: str = "_file", verbose: bool = False,
                     log=print, **_) -> pd.DataFrame:
    rep = as_reporter(log if verbose else None)

    missing = [c for c in REQUIRED + [file_col] if c not in raw.columns]
    if missing:
        raise ValueError("猪只转移数据缺少必需列, 请检查文件是否为猪只转移导出、标题是否在第 1 行:\n"
                         + "\n".join(f"  - {m}" for m in missing))
    df = raw.copy()

    # ---------------- 1. 从文件名拆出 8 个字段 ----------------
    names = df[file_col].map(lambda p: None if p is None or (isinstance(p, float) and np.isnan(p))
                             else os.path.basename(str(p)))
    parsed = {n: parse_transfer_filename(n) for n in names.dropna().unique()}
    cols = {f: [] for f in TRANSFER_FIELDS}
    for n in names:
        g = parsed.get(n) if n is not None else None
        for i, f in enumerate(TRANSFER_FIELDS):
            cols[f].append(g[i] if g else None)

    bad = [n for n, g in parsed.items() if g is None]
    if bad:
        n_rows = int(names.isin(bad).sum())
        rep.warn(f"{len(bad)} 个文件名不符合命名规则, 这 {fmt_n(n_rows)} 条记录的单据信息为 NA:",
                 [f"{{.file {esc(b)}}}" for b in bad[:5]]
                 + ([f"… 还有 {len(bad) - 5} 个"] if len(bad) > 5 else [])
                 + [f"应为 {{.dim {esc(FILENAME_PATTERN_HINT)}}}"])

    clash = [f for f in TRANSFER_FIELDS if f in df.columns]
    if clash:
        rep.warn("原始数据中已有同名列, 已用文件名中的信息覆盖:",
                 [f"{{.field {esc(c)}}}" for c in clash])

    for f in TRANSFER_FIELDS:
        if f == "转移数量":
            vals = pd.array([_to_int(x) for x in cols[f]], dtype="Int32")
        elif f == "转移时间":
            vals = pd.to_datetime(pd.Series([_to_date(x) for x in cols[f]], dtype="object"))
            vals = vals.astype("datetime64[ns]").to_numpy()
        else:
            vals = pd.Series(cols[f], dtype="object").to_numpy()
        _set_col(df, f, vals)

    n_ok = sum(1 for g in parsed.values() if g)
    n_date_bad = sum(1 for n, g in parsed.items() if g and _to_date(g[7]) is pd.NaT)
    if n_date_bad:
        rep.warn(f"{n_date_bad} 个文件名中的日期无效, {{.field 转移时间}} 为 NA")
    if n_ok:
        n_doc = len({g[0] for g in parsed.values() if g})
        rep.info(f"从 {n_ok} 个文件名解析出单据信息 (单据 {n_doc} 个)")

    # ---------------- 2. 品种 / 品系 ----------------
    add_breed_line(df, rep)
    return df
