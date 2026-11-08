# -*- coding: utf-8 -*-
"""
import_xlsx —— mintyr 0.2.1 `import_xlsx()` 的 Python 移植版

与 R 版保持一致的行为:
  * skip        : 标题行上方要跳过的行数(如一行大标题 -> skip=1)
  * header_rows : 标题占几行; 1 = 普通表, 2 = 两行合并表头 ...
  * header_sep  : 多行标题拼接符, 默认 "_"  (例: "配种信息" + "配种日期" -> "配种信息_配种日期")
  * header_fill : "auto"  读取工作簿里真实的合并单元格范围(横向/纵向填充), 失败时退回 "right"
                  "right" 上层标题向右填充(启发式)
                  "none"  不填充
  * 列名去重同 R make.unique(sep=header_sep): a, a_1, a_2
  * 空标题列 -> col_<j>
  * 多行标题时, 完全空白的数据行会被丢弃
  * 类型推断看全部行(同 mintyr 设定 guess_max 为最大值): 逻辑 < 数值 < 日期 < 文本
  * 多表合并时: 某表整列为空 -> 采用其它表的类型; 类型真冲突 -> 统一转文本并给出警告
  * 追加来源列 excel_name / sheet_name (可选 path_col)

读取引擎用 python-calamine(Rust 实现), 速度远快于 openpyxl。
"""
from __future__ import annotations

import datetime as _dt
import math
import os
import re
from concurrent.futures import ProcessPoolExecutor
from typing import Sequence

import numpy as np
import pandas as pd
from python_calamine import CalamineWorkbook

from .report import Reporter, as_reporter, esc, fmt_n



# ----------------------------------------------------------------------------
# 单元格 -> 类型判断
# ----------------------------------------------------------------------------
_EXCEL_EPOCH = _dt.datetime(1899, 12, 30)


_TYPE_KIND = {float: "number", int: "number", bool: "logical",
              _dt.datetime: "date", _dt.date: "date"}


def _cell_kind(v) -> str:
    """返回 blank / logical / number / date / text"""
    t = type(v)
    if t is str:
        return "text" if v and not v.isspace() else "blank"
    k = _TYPE_KIND.get(t)
    if k is not None:
        if t is float and v != v:
            return "blank"
        return k
    return _cell_kind_slow(v)


def _nonblank(v) -> bool:
    if type(v) is str:
        return bool(v) and not v.isspace()
    return _cell_kind(v) != "blank"


def _cell_kind_slow(v) -> str:
    if v is None:
        return "blank"
    if isinstance(v, str):
        return "blank" if v.strip() == "" else "text"
    if isinstance(v, bool):
        return "logical"
    if isinstance(v, (int, float)):
        if isinstance(v, float) and math.isnan(v):
            return "blank"
        return "number"
    if isinstance(v, (_dt.datetime, _dt.date)):
        return "date"
    if isinstance(v, (_dt.time, _dt.timedelta)):
        return "number"  # readxl 把纯时间读作数值(日的小数)
    return "text"


def _num_to_text(x: float) -> str:
    if isinstance(x, float) and x.is_integer() and abs(x) < 1e15:
        return str(int(x))
    return repr(float(x)) if isinstance(x, float) else str(x)


def _to_datetime(v):
    if isinstance(v, _dt.datetime):
        return v
    if isinstance(v, _dt.date):
        return _dt.datetime(v.year, v.month, v.day)
    return None


def _time_to_num(v) -> float:
    if isinstance(v, _dt.time):
        return (v.hour * 3600 + v.minute * 60 + v.second + v.microsecond / 1e6) / 86400
    if isinstance(v, _dt.timedelta):
        return v.total_seconds() / 86400
    return float(v)


def _cell_text(v) -> str | None:
    k = _cell_kind(v)
    if k == "blank":
        return None
    if k == "text":
        return str(v).strip()           # readxl trim_ws = TRUE
    if k == "logical":
        return "TRUE" if v else "FALSE"
    if k == "number":
        return _num_to_text(_time_to_num(v) if not isinstance(v, (int, float)) else v)
    if k == "date":
        d = _to_datetime(v)
        # 注: readxl 在文本列里会把日期写成序列号("45296"), 这里改写成
        # "YYYY-MM-DD"(非零点时带时间), 后续清洗更容易识别。
        s = d.strftime("%Y-%m-%d %H:%M:%S")
        return s[:-9] if s.endswith(" 00:00:00") else s
    return str(v)


_RANK = {"blank": 0, "logical": 1, "number": 2, "date": 3, "text": 4}


def _build_column(values: Sequence) -> pd.Series:
    """按 readxl 规则(全部行猜类型)把一列单元格转换为 pandas Series"""
    kinds = [_cell_kind(v) for v in values]
    top = max((_RANK[k] for k in kinds), default=0)
    n = len(values)

    if top == 0:  # 整列为空 -> 逻辑型 NA (与 readxl 一致)
        return pd.Series(pd.array([pd.NA] * n, dtype="boolean"))

    if top == 1:
        return pd.Series(pd.array([v if k == "logical" else pd.NA for v, k in zip(values, kinds)],
                                  dtype="boolean"))

    if top == 2:
        out = np.full(n, np.nan)
        for i, (v, k) in enumerate(zip(values, kinds)):
            if k == "number":
                out[i] = _time_to_num(v) if not isinstance(v, (int, float)) else float(v)
            elif k == "logical":
                out[i] = 1.0 if v else 0.0
        return pd.Series(out, dtype="float64")

    if top == 3:
        out = []
        for v, k in zip(values, kinds):
            if k == "date":
                out.append(_to_datetime(v))
            elif k == "number":  # 日期列里的数值按 Excel 序列号换算
                try:
                    out.append(_EXCEL_EPOCH + _dt.timedelta(days=_time_to_num(v)))
                except (OverflowError, ValueError):
                    out.append(None)
            else:
                out.append(None)
        s = pd.to_datetime(pd.Series(out, dtype="object"), errors="coerce")
        return s.dt.tz_localize("UTC")  # readxl 返回 POSIXct(UTC)

    # 文本
    return pd.Series([_cell_text(v) for v in values], dtype="object")


# ----------------------------------------------------------------------------
# 列名处理
# ----------------------------------------------------------------------------
def _make_unique(names: list[str], sep: str = "_") -> list[str]:
    """R make.unique(names, sep): 第二次出现起追加 sep+1, sep+2 ..."""
    seen: set[str] = set(names)
    cnt: dict[str, int] = {}
    out, used = [], set()
    for nm in names:
        if nm not in used:
            out.append(nm)
            used.add(nm)
            continue
        k = cnt.get(nm, 0)
        while True:
            k += 1
            cand = f"{nm}{sep}{k}"
            if cand not in used and cand not in seen:
                break
        cnt[nm] = k
        out.append(cand)
        used.add(cand)
    return out


def _repair_unique_readxl(names: list[str | None]) -> list[str]:
    """tibble/vctrs 的 'unique' 命名修复(readxl 单行标题时使用):
    空名 -> ...j;  重复名 -> name...j (j 为列位置)"""
    base = [re.sub(r"\.\.\.\d+$", "", n) if n else "" for n in names]
    counts: dict[str, int] = {}
    for b in base:
        counts[b] = counts.get(b, 0) + 1
    out = []
    for j, b in enumerate(base, start=1):
        if b == "" or counts[b] > 1:
            out.append(f"{b}...{j}")
        else:
            out.append(b)
    return out


def _combine_headers(hdr: list[list], sep: str, fill: str, merges, first_row: int) -> list[str]:
    """mintyr:::.combine_headers 的移植
    hdr: 标题块(行 x 列, 第 0 列 = Excel A 列), first_row: 标题块第一行(0 起)"""
    n_lev = len(hdr)
    n_col = max((len(r) for r in hdr), default=0)
    m = [[None] * n_col for _ in range(n_lev)]
    for i, row in enumerate(hdr):
        for j, v in enumerate(row):
            t = _cell_text(v)
            if t is not None:
                t = re.sub(r"[\r\n]+", " ", t).strip()
                if t == "":
                    t = None
            m[i][j] = t

    if merges is not None:
        last_row = first_row + n_lev - 1
        inside = [mg for mg in merges if first_row <= mg[0] <= last_row]
        if inside:
            width = max(n_col, max(mg[3] for mg in inside) + 1)
            if width > n_col:
                for row in m:
                    row.extend([None] * (width - n_col))
                n_col = width
            for r1, c1, r2, c2 in inside:
                rr = range(r1 - first_row, min(r2, last_row) - first_row + 1)
                val = m[r1 - first_row][c1]
                for r in rr:
                    for c in range(c1, c2 + 1):
                        m[r][c] = val
    elif fill == "right" and n_lev > 1:
        orig = [[v is not None for v in row] for row in m]
        for k in range(n_lev - 1):
            starts = [False] * n_col if k == 0 else [
                any(orig[i][j] for i in range(k)) for j in range(n_col)]
            cur = None
            for j in range(n_col):
                if starts[j]:
                    cur = None
                if m[k][j] is None:
                    m[k][j] = cur
                else:
                    cur = m[k][j]

    names = []
    for j in range(n_col):
        parts = [m[i][j] for i in range(n_lev) if m[i][j] is not None]
        dedup = [p for idx, p in enumerate(parts) if idx == 0 or p != parts[idx - 1]]
        names.append(sep.join(dedup))
    names = [nm if nm else f"col_{j + 1}" for j, nm in enumerate(names)]
    return _make_unique(names, sep)


# ----------------------------------------------------------------------------
# 读取单个工作表
# ----------------------------------------------------------------------------
def _trailing_trim(rows: list[list]) -> list[list]:
    while rows and not any(_nonblank(v) for v in rows[-1]):
        rows.pop()
    return rows


def detect_header_from_rows(rows: list[list], merges, required: Sequence[str],
                            header_sep: str = "_", max_skip: int = 3, max_rows: int = 3):
    """尝试 skip=0..max_skip, header_rows=1..max_rows, 返回匹配必需列最多的 (skip, header_rows, 命中数)"""
    req = set(required)
    best = (0, 1, -1)
    for hr in range(1, max_rows + 1):
        for sk in range(0, max_skip + 1):
            hdr = rows[sk:sk + hr]
            if len(hdr) < hr:
                continue
            if hr == 1:
                names = _repair_unique_readxl([_cell_text(v) for v in hdr[0]])
            else:
                names = _combine_headers(hdr, header_sep, "right", merges, first_row=sk)
            hit = len(req & set(names))
            if hit > best[2]:
                best = (sk, hr, hit)
    return best


def _read_sheet(path: str, sheet: str, skip: int, header_rows: int,
                header_sep: str, header_fill: str, required: Sequence[str] | None = None,
                info: dict | None = None) -> pd.DataFrame:
    auto = skip is None or header_rows is None
    wb = CalamineWorkbook.from_path(path)
    try:
        sh = wb.get_sheet_by_name(sheet)
        rows = sh.to_python(skip_empty_area=False)
        merges = None
        if (auto or header_rows > 1) and header_fill == "auto" and \
                re.search(r"\.xls[xm]$", path, re.I):
            try:
                mr = sh.merged_cell_ranges
                if mr is not None:
                    merges = [(a[0], a[1], b[0], b[1]) for a, b in mr]
            except Exception:
                merges = None
    finally:
        try:
            wb.close()
        except Exception:
            pass

    # calamine 的 to_python 从工作表 "start" 开始; 统一补齐到 A1 起
    start_r, start_c = (0, 0)
    try:
        start_r, start_c = sh.start  # type: ignore[attr-defined]
    except Exception:
        pass
    if start_c:
        rows = [[""] * start_c + list(r) for r in rows]
    if start_r:
        rows = [[] for _ in range(start_r)] + rows
    width = max((len(r) for r in rows), default=0)
    rows = [list(r) + [""] * (width - len(r)) for r in rows]

    if auto:
        sk, hr, hit = detect_header_from_rows(rows, merges, required or [], header_sep)
        skip = sk if skip is None else skip
        header_rows = hr if header_rows is None else header_rows
        if info is not None:
            info.update(skip=skip, header_rows=header_rows, hit=hit)

    if header_rows == 1:
        body = rows[skip:]
        # readxl: 自动跳过开头的空行; 只读取非空单元格所在的矩形
        while body and not any(_nonblank(v) for v in body[0]):
            body.pop(0)
        body = _trailing_trim(body)
        if not body:
            return pd.DataFrame()
        nonblank_cols = [j for j in range(width)
                         if any(_nonblank(r[j]) for r in body)]
        c0, c1 = nonblank_cols[0], nonblank_cols[-1]
        header = [_cell_text(v) for v in body[0][c0:c1 + 1]]
        names = _repair_unique_readxl(header)
        data = [r[c0:c1 + 1] for r in body[1:]]
        cols = {nm: _build_column([r[j] for r in data]) for j, nm in enumerate(names)}
        return pd.DataFrame(cols)

    # 多行标题
    hdr = rows[skip:skip + header_rows]
    if merges is None and header_fill == "auto":
        fill = "right"
    else:
        fill = "none" if header_fill == "none" else "right"
    names = _combine_headers(hdr, header_sep, fill, merges, first_row=skip)

    data = _trailing_trim(rows[skip + header_rows:])
    n = 0
    if data:
        n = max((j + 1 for r in data for j, v in enumerate(r) if _nonblank(v)),
                default=0)
    if n == 0:
        return pd.DataFrame({nm: pd.Series([], dtype="boolean") for nm in names})
    if len(names) < n:
        names = _make_unique(names + [f"col_{j + 1}" for j in range(len(names), n)], header_sep)
    names = names[:n]
    data = [r[:n] for r in data if any(_nonblank(v) for v in r[:n])]
    cols = {nm: _build_column([r[j] for r in data]) for j, nm in enumerate(names)}
    return pd.DataFrame(cols)


def _read_task(task: dict):
    try:
        info: dict = {}
        df = _read_sheet(task["path"], task["sheet"], task["skip"], task["header_rows"],
                         task["header_sep"], task["header_fill"], task.get("required"), info)
        if info:
            df.attrs["auto_header"] = info
        return df
    except Exception as e:  # 与 R 版一样: 错误先收集, 最后统一报告
        return _ReadError(describe_read_error(task["path"], e))


class _ReadError:
    def __init__(self, message: str):
        self.message = message


# ----------------------------------------------------------------------------
# 跨表类型对齐 (mintyr:::.align_col_types)
# ----------------------------------------------------------------------------
def _kind_of(s: pd.Series) -> str:
    if isinstance(s.dtype, pd.DatetimeTZDtype) or pd.api.types.is_datetime64_any_dtype(s):
        return "date-time"
    if pd.api.types.is_bool_dtype(s):
        return "logical"
    if pd.api.types.is_numeric_dtype(s):
        return "number"
    return "text"


def _as_text(s: pd.Series) -> pd.Series:
    k = _kind_of(s)
    if k == "date-time":
        out = s.dt.strftime("%Y-%m-%d %H:%M:%S").str.replace(r" 00:00:00$", "", regex=True)
        return out.astype("object").where(s.notna(), None)
    if k == "number":
        return pd.Series([None if pd.isna(v) else _num_to_text(float(v)) for v in s], dtype="object")
    if k == "logical":
        return pd.Series([None if pd.isna(v) else ("TRUE" if v else "FALSE") for v in s],
                         dtype="object")
    return s


_KIND_CN = {"date-time": "日期", "text": "文本", "number": "数值", "logical": "逻辑"}


def _align_col_types(tables: list[pd.DataFrame], labels: list[str], rep: Reporter) -> None:
    if len(tables) < 2:
        return
    cols: list[str] = []
    for t in tables:
        for c in t.columns:
            if c not in cols:
                cols.append(c)
    report = []
    for col in cols:
        has = [i for i, t in enumerate(tables) if col in t.columns]
        if len(has) < 2:
            continue
        kinds = [_kind_of(tables[i][col]) for i in has]
        if len(set(kinds)) == 1:
            continue
        empty = [kinds[k] == "logical" and tables[i][col].isna().all() for k, i in enumerate(has)]
        filled = sorted({kinds[k] for k in range(len(has)) if not empty[k]})
        if not filled or all(f in ("logical", "number") for f in filled):
            # 逻辑/数值混合: 与 rbindlist 一样按数值合并
            if "number" in kinds:
                for k, i in enumerate(has):
                    if kinds[k] == "logical":
                        tables[i][col] = tables[i][col].astype("Float64").astype("float64")
            continue
        if len(filled) == 1:
            ref = next(tables[i][col] for k, i in enumerate(has) if not empty[k])
            for k, i in enumerate(has):
                if empty[k]:
                    tables[i][col] = pd.Series([pd.NA] * len(tables[i]), index=tables[i].index,
                                               dtype=ref.dtype) if len(tables[i]) else ref.iloc[:0]
            continue
        where: dict[str, list[str]] = {}
        for k, i in enumerate(has):
            if not empty[k]:
                where.setdefault(kinds[k], []).append(labels[i])
            if kinds[k] != "text":
                tables[i][col] = _as_text(tables[i][col]).values
        parts = [f"{_KIND_CN.get(t, t)} {len(l)} 个表" for t, l in where.items()]
        few = min(where.values(), key=len)       # 列出占少数的来源, 方便定位
        if len(few) <= 3:
            parts[-1 if few is list(where.values())[-1] else 0] += (
                " {.dim (" + esc(", ".join(few)) + ")}")
        report.append(f"{{.field {esc(col)}}}: " + " / ".join(parts))
    if report:
        rep.warn(f"{len(report)} 列在不同来源中类型不同, 已统一按文本合并:", report)


# ----------------------------------------------------------------------------
# 对外接口
# ----------------------------------------------------------------------------
def _strip_ext(name: str) -> str:
    return re.sub(r"(?<=[^/\\.])\.[^./\\]+$", "", name)


def describe_read_error(path: str, e: Exception) -> str:
    """把读取异常翻译成能直接处理的中文说明"""
    if isinstance(e, PermissionError) or getattr(e, "errno", None) == 13 or \
            "being used by another process" in str(e) or "Permission denied" in str(e):
        return "文件被占用或没有读取权限 (若正在 Excel 中打开, 请先关闭)"
    if isinstance(e, FileNotFoundError):
        return "文件不存在"
    try:
        size = os.path.getsize(path)
        with open(path, "rb") as fh:
            head = fh.read(512)
    except PermissionError:
        return "文件被占用或没有读取权限 (若正在 Excel 中打开, 请先关闭)"
    except OSError:
        return _short_err(e)
    if size == 0:
        return "文件为空 (0 字节)"
    is_x = path.lower().endswith((".xlsx", ".xlsm"))
    if head.startswith(b"\xd0\xcf\x11\xe0") and is_x:
        return "文件可能设置了打开密码, 或是把 .xls 直接改成了 .xlsx; 请用 Excel 打开后另存为 .xlsx"
    txt = head.lstrip(b"\xef\xbb\xbf \t\r\n").lower()
    if txt.startswith((b"<", b"<!doctype", b"<html", b"<?xml")):
        return "实际是网页 (HTML) 格式, 不是真正的 Excel 文件; 请用 Excel 打开后另存为 .xlsx"
    if head.startswith(b"PK") and "zip" in str(e).lower():
        return "文件不完整或已损坏 (可能下载 / 复制未完成); 请重新导出"
    if "zip" in str(e).lower() or "cannot detect" in str(e).lower():
        return "不是有效的 Excel 文件 (可能是改了扩展名的其他文件)"
    return _short_err(e)


def _fail_line(path: str, sheet, reason: str) -> str:
    where = f" › {esc(sheet)}" if sheet else ""
    folder = os.path.dirname(os.path.abspath(path))
    return f"{{.file {esc(os.path.basename(path))}}}{where}: {esc(reason)} {{.dim ({esc(folder)})}}"


def _fail_plain(path: str, sheet, reason: str) -> str:
    where = f" [{sheet}]" if sheet else ""
    return f"  - {os.path.basename(path)}{where}: {reason}  ({os.path.dirname(os.path.abspath(path))})"


def _short_err(e: Exception) -> str:
    msg = str(e).strip().splitlines()
    return msg[0] if msg else type(e).__name__


def _sheet_label(t: dict, k: int, n: int) -> str:
    return (f"{{.dim [{k}/{n}]}} {{.file {esc(os.path.basename(t['path']))}}} "
            f"› {esc(t['sheet'])}")


def _dims(d: pd.DataFrame) -> str:
    s = f"{fmt_n(len(d))} 行 × {d.shape[1]} 列"
    if len(d) == 0:
        s += " {.dim (空表)}"
    ah = d.attrs.get("auto_header")
    if ah:
        s += f" {{.dim (自动识别: 跳过 {ah['skip']} 行, 标题 {ah['header_rows']} 行)}}"
    return s


def list_sheets(path: str) -> list[str]:
    wb = CalamineWorkbook.from_path(path)
    try:
        return list(wb.sheet_names)
    finally:
        try:
            wb.close()
        except Exception:
            pass


def import_xlsx(file: str | Sequence[str],
                combine: bool = True,
                sheet: Sequence[int] | Sequence[str] | int | str | None = None,
                skip: int | None = 0,
                header_rows: int | None = 1,
                header_sep: str = "_",
                header_fill: str = "auto",
                file_col: str | None = "excel_name",
                sheet_col: str | None = "sheet_name",
                path_col: str | None = None,
                workers: int = 1,
                verbose: bool = False,
                on_error: str = "stop",
                required: Sequence[str] | None = None,
                log=print,
                summary: bool = True,
                failures: list | None = None):
    """读取一个或多个 xlsx, 参数含义与 mintyr::import_xlsx 相同。

    combine=True 返回一个 DataFrame; combine=False 返回 {"<文件名>_<表名>": DataFrame}

    扩展(非 mintyr): skip / header_rows 传 None 表示按工作表自动识别, 依据是
    `required` 列名能匹配上多少(每个工作表单独判断, 不同表头格式的文件可以混读)。
    log 可以是 report.Reporter 或普通函数(如 print); summary=False 时不输出最后的汇总行。
    failures: 传入列表时, 读取失败的 (路径, 工作表或 None, 原因) 追加到其中、由调用方统一报告;
              不传时在这里输出警告。
    """
    rep = as_reporter(log)
    files = [file] if isinstance(file, str) else list(file)
    if not files:
        raise ValueError("`file` 不能为空")
    missing = [f for f in files if not os.path.exists(f)]
    if missing and on_error == "stop":
        raise FileNotFoundError(f"找不到 {len(missing)} 个文件:\n"
                                + "\n".join(_fail_plain(f, None, "文件不存在") for f in missing))
    if header_fill not in ("auto", "right", "none"):
        raise ValueError("header_fill 只能是 'auto' / 'right' / 'none'")
    if on_error not in ("stop", "warn"):
        raise ValueError("on_error 只能是 'stop' / 'warn'")
    if (skip is not None and skip < 0) or (header_rows is not None and header_rows < 1):
        raise ValueError("skip >= 0 且 header_rows >= 1")
    track = [c for c in (file_col, sheet_col, path_col) if c is not None]
    if len(set(track)) != len(track):
        raise ValueError("file_col / sheet_col / path_col 不能重名")
    if isinstance(sheet, (int, str)):
        sheet = [sheet]

    bases = [_strip_ext(os.path.basename(f)) for f in files]
    if len(set(bases)) < len(bases) and (file_col is not None or not combine):
        rep.warn("有多个文件同名(在不同文件夹), {.field excel_name} 无法区分, 已自动加后缀",
                 ["如需区分请设置 {.field path_col}"])
        bases = _make_unique(bases, "_")

    # 1. 工作表目录 (打不开的文件记下原因)
    file_errors, sheets_cache = [], {}          # file_errors: [(路径, None, 原因)]
    for i, f in enumerate(files):
        if f in missing:
            file_errors.append((f, None, "文件不存在"))
            continue
        try:
            sheets_cache[i] = list_sheets(f)
        except Exception as e:
            file_errors.append((f, None, describe_read_error(f, e)))
    if file_errors and (on_error == "stop" or len(file_errors) == len(files)):
        raise RuntimeError(f"{len(file_errors)} 个文件无法读取"
                           + (" (全部文件都读不了)" if len(file_errors) == len(files) else "") + ":\n"
                           + "\n".join(_fail_plain(*x) for x in file_errors))

    # 2. (文件, 表) 任务
    tasks = []
    for i, f in enumerate(files):
        if i not in sheets_cache:
            continue
        avail = sheets_cache[i]
        if sheet is None:
            sel = avail
        elif all(isinstance(s, str) for s in sheet):
            miss = [s for s in sheet if s not in avail]
            if miss:
                raise ValueError(f"{os.path.basename(f)} 中找不到工作表: {', '.join(miss)}")
            sel = list(sheet)
        else:
            idx = [int(s) for s in sheet]
            if max(idx) > len(avail):
                raise ValueError(f"工作表序号 {max(idx)} 超过 {os.path.basename(f)} 的工作表数 {len(avail)}")
            sel = [avail[k - 1] for k in idx]
        full = os.path.abspath(f).replace("\\", "/")
        for s in sel:
            tasks.append(dict(path=f, full=full, base=bases[i], sheet=s, skip=skip,
                              header_rows=header_rows, header_sep=header_sep,
                              header_fill=header_fill, required=list(required or [])))

    # 3. 读取(串行 / 多进程)
    use_workers = min(max(int(workers), 1), len(tasks)) if tasks else 1
    if use_workers > 1:
        with ProcessPoolExecutor(max_workers=use_workers) as ex:
            results = list(ex.map(_read_task, tasks))
    else:
        results = []
        n = len(tasks)
        for k, t in enumerate(tasks, 1):
            if not verbose:
                results.append(_read_task(t))
                continue
            label = _sheet_label(t, k, n)
            with rep.step(f"读取 {label}") as st:
                r = _read_task(t)
                if isinstance(r, _ReadError):
                    st.warn = True
                    st.done = f"{label}  读取失败: {esc(r.message)}"
                else:
                    st.done = f"{label}  {_dims(r)}"
            results.append(r)

    failed = [i for i, r in enumerate(results) if isinstance(r, _ReadError)]
    if failed or file_errors:
        info = file_errors + [(tasks[i]["path"], tasks[i]["sheet"], results[i].message)
                              for i in failed]
        if on_error == "stop" or (failed and len(failed) == len(results)):
            raise RuntimeError(f"{len(info)} 个文件 / 工作表无法读取:\n"
                               + "\n".join(_fail_plain(*x) for x in info))
        if failures is not None:
            failures.extend(info)
        else:
            rep.warn(f"{len(info)} 个文件 / 工作表无法读取, 已跳过:", [_fail_line(*x) for x in info])
        keep = [i for i in range(len(results)) if i not in failed]
        results = [results[i] for i in keep]
        tasks = [tasks[i] for i in keep]

    if verbose and use_workers > 1:   # 并行时读完再逐表报告
        for k, (t, d) in enumerate(zip(tasks, results), 1):
            rep.success(f"{_sheet_label(t, k, len(tasks))}  {_dims(d)}")
    if verbose and summary:
        total = sum(len(d) for d in results)
        rep.success(f"共导入 {len(tasks)} 个工作表 (来自 {len(files)} 个文件), 合计 {fmt_n(total)} 行")

    if not combine:
        out, names = {}, _make_unique([f"{t['base']}_{t['sheet']}" for t in tasks], "_")
        for nm, d in zip(names, results):
            out[nm] = d
        return out

    for t, d in zip(tasks, results):
        if file_col:
            d[file_col] = t["base"]
        if sheet_col:
            d[sheet_col] = t["sheet"]
        if path_col:
            d[path_col] = t["full"]
    _align_col_types(results, [f"{os.path.basename(t['path'])} › {t['sheet']}" for t in tasks], rep)
    results = [d for d in results if d.shape[1] > 0]
    if not results:
        return pd.DataFrame()
    out = pd.concat(results, ignore_index=True, sort=False) if len(results) > 1 else results[0]
    front = [c for c in (file_col, sheet_col, path_col) if c and c in out.columns]
    out = out[front + [c for c in out.columns if c not in front]]
    return out.reset_index(drop=True)
