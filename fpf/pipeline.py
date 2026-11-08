# -*- coding: utf-8 -*-
"""读取 -> 处理 -> 导出 一体化流程"""
from __future__ import annotations

import glob
import os
import time
from dataclasses import dataclass, field
from typing import Callable

import pandas as pd

from . import archive, cleaning, transfer
from .importer import _fail_line, import_xlsx
from .rds_writer import write_rds
from .report import as_reporter, esc, fmt_n, fmt_size
from .rules import BreedRules, use_rules

@dataclass
class Task:
    """一个处理任务(对应界面上的一个页签)"""
    key: str                       # 命令行 --task 用
    name: str                      # 页签 / 日志标题
    func: Callable                 # func(raw, verbose=, log=, **options) -> DataFrame
    required: list = field(default_factory=list)   # 必需列(用于自动识别标题、报错提示)
    skip: int | None = None        # 默认跳过行数; None = 自动识别
    header_rows: int | None = None # 默认标题行数; None = 自动识别
    path_col: str | None = "_file" # 读取时追加的完整路径列(同 mintyr import_xlsx 的 path_col)
    default_out: str = "结果"       # 默认输出文件名(不含扩展名)
    file_filter: Callable | None = None   # 文件名筛选: filter(文件名) -> bool; 不符合的不读取
    filter_hint: str = ""                 # 筛选规则说明(用于提示)


TASKS = {
    "sow": Task("sow", "母猪综合查询", cleaning.clean_production_data,
                required=[src for _, src in cleaning.SELECT_MAP],
                default_out="母猪综合查询_清洗"),
    "transfer": Task("transfer", "猪只转移", transfer.process_transfer,
                     required=transfer.REQUIRED, skip=0, header_rows=1, path_col="_file",
                     default_out="猪只转移_清洗",
                     file_filter=lambda name: transfer.parse_transfer_filename(name) is not None,
                     filter_hint=transfer.FILENAME_PATTERN_HINT),
    "archive": Task("archive", "母猪档案", archive.process_archive,
                    required=archive.REQUIRED, skip=0, header_rows=1, path_col="_file",
                    default_out="母猪档案_清洗",
                    file_filter=lambda name: archive.parse_archive_filename(name) is not None,
                    filter_hint=archive.FILENAME_PATTERN_HINT),
}

# 导出格式: 名称 -> 扩展名
EXPORT_FORMATS = {
    "rds": ".rds",           # R 原生, gzip 压缩(与 saveRDS 默认相同)
    "csv": ".csv",           # 通用(UTF-8 BOM, Excel 可直接打开)
}


def scan_inputs(paths: list[str]) -> tuple[list[str], list[str]]:
    """文件 + 文件夹(递归) -> (xlsx 列表, 被忽略的 Excel 临时文件 ~$xxx.xlsx)"""
    out = []
    for p in paths:
        if os.path.isdir(p):
            found = []
            for ext in ("xlsx", "xlsm", "xls", "XLSX", "XLSM", "XLS"):
                found += glob.glob(os.path.join(p, "**", f"*.{ext}"), recursive=True)
            out += sorted(set(found))   # 文件夹内按路径排序, 结果顺序固定
        else:
            out.append(p)
    out = list(dict.fromkeys(os.path.normpath(f) for f in out))
    temp = [f for f in out if os.path.basename(f).startswith("~$")]
    return [f for f in out if f not in temp], temp


def collect_xlsx(paths: list[str]) -> list[str]:
    """文件 + 文件夹(递归) -> xlsx 列表; 忽略 Excel 临时文件 ~$xxx.xlsx"""
    return scan_inputs(paths)[0]


def _name_list(paths: list[str], limit: int = 8) -> list[str]:
    items = [f"{{.file {esc(os.path.basename(f))}}} {{.dim ({esc(os.path.dirname(os.path.abspath(f)))})}}"
             for f in paths[:limit]]
    if len(paths) > limit:
        items.append(f"… 还有 {len(paths) - limit} 个")
    return items


def export_data(df: pd.DataFrame, path: str, fmt: str) -> str:
    fmt = fmt.lower()
    if fmt not in EXPORT_FORMATS:
        raise ValueError(f"不支持的导出格式: {fmt}, 可选: {', '.join(EXPORT_FORMATS)}")
    root, ext = os.path.splitext(path)
    want = EXPORT_FORMATS[fmt]
    if ext.lower() != want:
        path = root + want if ext.lower() in (".rds", ".csv") \
            else path + want
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)

    if fmt == "rds":
        write_rds(df, path, compress=True, level=1)
    else:
        out = df.copy()
        for c in out.columns:
            if pd.api.types.is_datetime64_any_dtype(out[c]) and \
                    not isinstance(out[c].dtype, pd.DatetimeTZDtype):
                nn = out[c].dropna()
                if (nn.dt.normalize() == nn).all():
                    out[c] = out[c].dt.strftime("%Y-%m-%d")
        out.to_csv(path, index=False, encoding="utf-8-sig")
    return path


def run_pipeline(files: list[str], task: str = "sow", process: bool = True,
                 skip: int | str | None = "default", header_rows: int | str | None = "default",
                 header_fill: str = "auto", sheet=None, on_error: str = "warn",
                 workers: int = 1, options: dict | None = None,
                 out_path: str | None = None, fmt: str = "rds",
                 rules: BreedRules | None = None,
                 log=print, **legacy) -> tuple[pd.DataFrame, str | None]:
    """读取 -> 处理 -> 导出

    task:    TASKS 中的键 ("sow" 母猪综合查询 / "transfer" 猪只转移)
    process: False 时只读取合并, 不做处理
    skip / header_rows: "default" 用任务默认值, "auto" 自动识别, 或具体数字
    options: 传给处理函数的参数, 如母猪综合查询的 data_type / drop_missing_date
    rules:   品种品系判定规则; None = 内置默认
    log:     report.Reporter 或普通函数(如 print)
    """
    rep = as_reporter(log)
    t0 = time.perf_counter()
    tk = TASKS[task]
    options = dict(options or {})
    # 兼容旧参数: processor="仅读取(不处理)", data_type=..., drop_missing_date=...
    if legacy.pop("processor", None) == "仅读取(不处理)":
        process = False
    options.update({k: legacy.pop(k) for k in ("data_type", "drop_missing_date") if k in legacy})
    if legacy:
        raise TypeError(f"未知参数: {', '.join(legacy)}")

    files, temp = scan_inputs(files)
    if not files:
        raise ValueError("没有找到 xlsx 文件")

    def _opt(v, default):
        if v is None or str(v).strip().lower() == "default":
            return default
        return None if str(v).strip().lower() in ("auto", "") else int(v)
    skip, header_rows = _opt(skip, tk.skip), _opt(header_rows, tk.header_rows)

    # ---------------- 读取 ----------------
    rep.h1(f"读取: {tk.name}")
    hdr = ("自动识别" if skip is None and header_rows is None else
           f"跳过 {'自动' if skip is None else skip} 行, "
           f"标题 {'自动' if header_rows is None else header_rows} 行")
    if temp:
        rep.info(f"忽略 {len(temp)} 个 Excel 临时文件 ({{.file ~$}} 开头, 是 Excel 打开文件时自动生成的)")
    if tk.file_filter is not None:
        keep = [f for f in files if tk.file_filter(os.path.basename(f))]
        skipped = [f for f in files if f not in keep]
        if skipped:
            rep.info(f"跳过 {len(skipped)} 个文件名不符合规则的文件 "
                     f"(应为 {{.dim {esc(tk.filter_hint)}}}):", _name_list(skipped))
        if not keep:
            raise ValueError(f"没有文件名符合规则的文件, 文件名应为:\n  - {tk.filter_hint}")
        files = keep
    rep.info(f"共 {len(files)} 个文件, 标题: {hdr}"
             + (f", 并行 {workers} 进程" if workers > 1 else ""))

    failures: list = []
    raw = import_xlsx(files, combine=True, sheet=sheet, skip=skip,
                      header_rows=header_rows, header_fill=header_fill,
                      file_col=None, sheet_col=None, path_col=tk.path_col,
                      workers=workers, verbose=True, on_error=on_error,
                      required=tk.required, log=rep, summary=False, failures=failures)
    if failures:
        rep.warn(f"{len(failures)} 个文件 / 工作表无法读取, 已跳过:",
                 [_fail_line(*x) for x in failures])
    t1 = time.perf_counter()
    rep.success(f"读取完成: {fmt_n(raw.shape[0])} 行 × {raw.shape[1]} 列", elapsed=t1 - t0)

    # ---------------- 处理 ----------------
    if process:
        rep.h1(f"处理: {tk.name}")
        if rules is None:
            rep.info("品种品系规则: 内置默认")
        else:
            when = f", 保存于 {esc(rules.saved_at)}" if rules.saved_at else ""
            rep.info(f"品种品系规则: {{.strong 自定义}} ({esc(rules.summary())}{when})")
        with use_rules(rules):
            df = tk.func(raw, verbose=True, log=rep, **options)
        t2 = time.perf_counter()
        rep.success(f"处理完成: {fmt_n(df.shape[0])} 行 × {df.shape[1]} 列", elapsed=t2 - t1)
    else:
        df = raw

    # ---------------- 导出 ----------------
    written = None
    if out_path:
        rep.h1("导出")
        with rep.step(f"写出 {fmt}") as st:
            written = export_data(df, out_path, fmt)
            st.done = (f"已导出 {{.file {esc(written)}}} "
                       f"{{.dim ({fmt_size(os.path.getsize(written))})}}")
    rep.text("")
    if failures:
        rep.warn(f"有 {len(failures)} 个文件 / 工作表没有读取, 结果中不包含它们 (原因见「读取」部分)")
    rep.success(f"{{.strong 全部完成}}: {fmt_n(df.shape[0])} 行 × {df.shape[1]} 列",
                elapsed=time.perf_counter() - t0)
    df.attrs["unread"] = [(p, sh, why) for p, sh, why in failures]
    return df, written
