# -*- coding: utf-8 -*-
"""命令行与图形界面共用: 日志文件、错误显示"""
from __future__ import annotations

import os
import time

from .report import esc

APP_NAME = "FPF 数据处理工具"
APP_VERSION = "1.1.0"


def parse_sheet(text):
    """'1' -> [1];  'Sheet1,数据' -> ['Sheet1', '数据'];  空 -> None(全部)"""
    if not text or not str(text).strip():
        return None
    parts = [t.strip() for t in str(text).replace("，", ",").split(",") if t.strip()]
    if all(t.isdigit() for t in parts):
        return [int(t) for t in parts]
    return parts


def log_path_for(out: str | None) -> str | None:
    return os.path.splitext(out)[0] + ".log" if out else None


def write_log_file(path: str | None, lines: list[str], tb: str | None = None):
    """纯文本日志: 与输出文件同名的 .log"""
    if not path:
        return
    try:
        os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            f.write(f"{APP_NAME} v{APP_VERSION}  {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("\n".join(lines).lstrip("\n") + "\n")
            if tb:
                f.write("\n── 详细错误信息 ──\n" + tb)
    except Exception:
        pass


def report_error(rep, e: Exception, log_file: str | None):
    """把异常显示为 cli 风格: ✖ 第一行, 其余 '  - xxx' 行作为条目"""
    lines = [ln for ln in str(e).splitlines() if ln.strip()] or [type(e).__name__]
    bullets = [esc(ln.strip()[2:] if ln.strip().startswith("- ") else ln.strip()) for ln in lines[1:]]
    rep.text("")
    rep.danger(esc(lines[0]), bullets)
    if log_file:
        rep.info(f"详细错误信息已写入 {{.file {esc(log_file)}}}")
