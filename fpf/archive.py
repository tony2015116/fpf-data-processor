# -*- coding: utf-8 -*-
"""
母猪档案数据处理

文件名: 母猪档案_猪场_生产阶段_N条_YYYYMMDD_HHMM.xlsx
    例: 母猪档案_大庆一场_后备生产阶段_1990条_20261101_0805.xlsx
只读取文件名符合规则的文件(在 pipeline.TASKS 中筛选), 一行标题。

处理:
  1. 从文件名拆出 档案猪场 / 档案阶段 / 档案条数 / 导出日期 / 导出时间
     (加"档案"前缀, 不覆盖表格里原有的 猪场 等列)
  2. 核对: 文件名中的条数与实际读到的行数是否一致
  3. 按当前规则(内置或自定义)判定 品种 / 品系, 来源为 品种品系 列
"""
from __future__ import annotations

import datetime as _dt
import os
import re

import numpy as np
import pandas as pd

from .report import as_reporter, esc, fmt_n
from .transfer import _set_col, add_breed_line

ARCHIVE_FIELDS = ["档案猪场", "档案阶段", "档案条数", "导出日期", "导出时间"]
FILENAME_RE = re.compile(r"母猪档案_([^_]+)_([^_]+)_([0-9]+)条_([0-9]{8})_([0-9]{4})\.xlsx")
FILENAME_PATTERN_HINT = "母猪档案_猪场_生产阶段_N条_YYYYMMDD_HHMM.xlsx"
REQUIRED = ["品种品系"]


def parse_archive_filename(name: str | None) -> list | None:
    """'母猪档案_大庆一场_后备生产阶段_1990条_20261101_0805.xlsx'
       -> ['大庆一场', '后备生产阶段', 1990, Timestamp('2026-11-01'), '08:05'];  不匹配返回 None"""
    if not name:
        return None
    m = FILENAME_RE.fullmatch(name)
    if not m:
        return None
    farm, stage, n, ymd, hm = m.groups()
    try:
        day = pd.Timestamp(_dt.datetime.strptime(ymd, "%Y%m%d"))
    except ValueError:
        day = pd.NaT
    hh, mm = int(hm[:2]), int(hm[2:])
    t = f"{hh:02d}:{mm:02d}" if hh < 24 and mm < 60 else None
    cnt = int(n) if int(n) <= 2147483647 else None
    return [farm, stage, cnt, day, t]


def process_archive(raw: pd.DataFrame, file_col: str = "_file", verbose: bool = False,
                    log=print, **_) -> pd.DataFrame:
    rep = as_reporter(log if verbose else None)
    missing = [c for c in REQUIRED + [file_col] if c not in raw.columns]
    if missing:
        raise ValueError("母猪档案数据缺少必需列, 请检查文件是否为母猪档案导出、标题是否在第 1 行:\n"
                         + "\n".join(f"  - {m}" for m in missing))
    df = raw.copy()

    # ---------------- 1. 文件名信息 ----------------
    names = df[file_col].map(lambda p: None if p is None or (isinstance(p, float) and np.isnan(p))
                             else os.path.basename(str(p)))
    parsed = {n: parse_archive_filename(n) for n in names.dropna().unique()}
    rows = [parsed.get(n) if n is not None else None for n in names]
    for i, f in enumerate(ARCHIVE_FIELDS):
        vals = [r[i] if r else None for r in rows]
        if f == "档案条数":
            vals = pd.array([pd.NA if v is None else v for v in vals], dtype="Int32")
        elif f == "导出日期":
            vals = pd.to_datetime(pd.Series([pd.NaT if v is None else v for v in vals], dtype="object"))
            vals = vals.astype("datetime64[ns]").to_numpy()
        else:
            vals = pd.Series(vals, dtype="object").to_numpy()
        _set_col(df, f, vals)

    good = {n: g for n, g in parsed.items() if g}
    bad = [n for n, g in parsed.items() if g is None]
    if bad:
        rep.warn(f"{len(bad)} 个文件名不符合命名规则, 档案信息为 NA:",
                 [f"{{.file {esc(b)}}}" for b in bad[:5]]
                 + [f"应为 {{.dim {esc(FILENAME_PATTERN_HINT)}}}"])
    bad_dt = [n for n, g in good.items() if g[3] is pd.NaT or g[4] is None]
    if bad_dt:
        rep.warn(f"{len(bad_dt)} 个文件名中的日期或时间无效, 记为 NA:",
                 [f"{{.file {esc(b)}}}" for b in bad_dt[:5]])
    if good:
        farms = sorted({g[0] for g in good.values()})
        rep.info(f"从 {len(good)} 个文件名解析出档案信息: 猪场 " + "、".join(esc(f) for f in farms))

    # ---------------- 2. 条数核对 ----------------
    counts = names.value_counts()
    mismatch = [(n, g[2], int(counts.get(n, 0))) for n, g in good.items()
                if g[2] is not None and int(counts.get(n, 0)) != g[2]]
    if mismatch:
        rep.warn(f"{len(mismatch)} 个文件实际读到的行数与文件名中的条数不一致, 请检查是否导出完整:",
                 [f"{{.file {esc(n)}}}: 文件名 {fmt_n(c)} 条, 实际 {fmt_n(a)} 行"
                  for n, c, a in mismatch[:8]]
                 + ([f"… 还有 {len(mismatch) - 8} 个"] if len(mismatch) > 8 else []))
    elif good:
        rep.info(f"条数核对: {len(good)} 个文件的行数与文件名一致")

    # ---------------- 3. 品种 / 品系 ----------------
    add_breed_line(df, rep)
    return df
