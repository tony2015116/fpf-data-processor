# -*- coding: utf-8 -*-
"""
仿 R cli 包风格的信息输出

    ── 读取 ─────────────────────────────────────────
    ℹ 共 3 个文件
    ✔ A场_2025.xlsx › Sheet1  1,484 行 × 39 列 [45ms]
    ! 5 列在不同来源中类型不同, 已统一按文本合并:
      • 出生日期: 日期 (A场_2025.xlsx [Sheet1]); 文本 (B场_2026.xlsx [数据])
    ✖ 读取失败

各模块只产生结构化消息 Msg, 由不同的"输出端"负责呈现:
    PlainSink    纯文本(日志文件 / 普通 print)
    ConsoleSink  终端, 带 ANSI 颜色; 进行中的步骤原地刷新
    图形界面     main.py 中用 Text 控件的颜色标签呈现

行内标记(与 cli 相同的写法):
    {.file 路径}  {.val 值}  {.field 列名}  {.strong 文本}  {.dim 文本}
需要原样输出的变量先用 esc() 转义花括号。
"""
from __future__ import annotations

import os
import re
import sys
import time
import unicodedata
from contextlib import contextmanager
from dataclasses import dataclass, field
from typing import Callable, Iterable

# ----------------------------------------------------------------------------
# 消息
# ----------------------------------------------------------------------------
KINDS = ("h1", "info", "success", "warn", "danger", "text", "step")


@dataclass
class Msg:
    kind: str                       # 见 KINDS
    text: str = ""
    bullets: list[str] = field(default_factory=list)
    elapsed: float | None = None    # 秒; 有值时在行尾显示 [耗时]
    replace: bool = False           # True: 替换上一条 step 行


SYMBOLS = {"success": "✔", "info": "ℹ", "warn": "!", "danger": "✖",
           "step": "ℹ", "bullet": "•", "rule": "─", "sep": "›"}
# Windows 默认字体里 ✔ ℹ ✖ 可能缺字, 用 cli 包在 Windows 下的替代符号
SYMBOLS_WIN = {**SYMBOLS, "success": "√", "info": "i", "danger": "×", "step": "i"}

RULE_WIDTH = 64


def esc(x) -> str:
    """把任意值转成可安全放进消息的文本(转义花括号)"""
    return str(x).replace("{", "{{").replace("}", "}}")


def fmt_n(n) -> str:
    """千分位: 3762 -> 3,762"""
    try:
        return f"{int(n):,}"
    except (TypeError, ValueError):
        return str(n)


def fmt_time(sec: float) -> str:
    """cli 风格耗时: 45ms / 1.2s / 1m 5s"""
    if sec < 1:
        return f"{sec * 1000:.0f}ms"
    if sec < 60:
        return f"{sec:.1f}s"
    m, s = divmod(int(round(sec)), 60)
    return f"{m}m {s}s"


def fmt_size(nbytes: int) -> str:
    for unit in ("B", "kB", "MB", "GB"):
        if nbytes < 1000 or unit == "GB":
            return f"{nbytes:.0f} {unit}" if unit == "B" else f"{nbytes:.1f} {unit}"
        nbytes /= 1000
    return str(nbytes)


def text_width(s: str) -> int:
    """显示宽度(中文等全角字符算 2)"""
    return sum(2 if unicodedata.east_asian_width(c) in "WF" else 1 for c in s)


# ----------------------------------------------------------------------------
# 行内标记解析
# ----------------------------------------------------------------------------
_MARK = re.compile(r"\{\.(file|val|field|strong|dim) ((?:[^{}]|\{\{|\}\})*)\}")


def _unesc(s: str) -> str:
    return s.replace("{{", "{").replace("}}", "}")


def parse(text: str) -> list[tuple[str | None, str]]:
    """'读取 {.file a.xlsx} 完成' -> [(None,'读取 '), ('file','a.xlsx'), (None,' 完成')]"""
    out, pos = [], 0
    for m in _MARK.finditer(text):
        if m.start() > pos:
            out.append((None, _unesc(text[pos:m.start()])))
        out.append((m.group(1), _unesc(m.group(2))))
        pos = m.end()
    if pos < len(text):
        out.append((None, _unesc(text[pos:])))
    return out


_PLAIN_WRAP = {"field": "`{}`", "val": '"{}"'}


def plain(text: str) -> str:
    """无颜色时: 列名加反引号、值加引号(与 cli 包无颜色输出一致), 其余原样"""
    return "".join(_PLAIN_WRAP.get(st, "{}").format(t) if st else t for st, t in parse(text))


def rule_line(title: str, sym=SYMBOLS) -> str:
    head = f"{sym['rule'] * 2} {title} "
    return head + sym["rule"] * max(4, RULE_WIDTH - text_width(head))


def render_plain(m: Msg, sym=SYMBOLS) -> str:
    """一条消息 -> 纯文本(可多行)"""
    if m.kind == "h1":
        return "\n" + rule_line(plain(m.text), sym)
    line = plain(m.text)
    if m.kind in sym and m.kind != "text":
        line = f"{sym[m.kind]} {line}"
    if m.kind == "step":
        line += " ..."
    if m.elapsed is not None:
        line += f" [{fmt_time(m.elapsed)}]"
    lines = [line] + [f"  {sym['bullet']} {plain(b)}" for b in m.bullets]
    return "\n".join(lines)


# ----------------------------------------------------------------------------
# Reporter: 各模块调用的接口
# ----------------------------------------------------------------------------
class Reporter:
    def __init__(self, sink: Callable[[Msg], None] | None = None):
        self.sink = sink or (lambda m: None)

    # 兼容旧接口: log("文本")
    def __call__(self, text: str):
        self.info(text)

    def emit(self, m: Msg):
        try:
            self.sink(m)
        except Exception:
            pass  # 输出失败不能影响数据处理

    def h1(self, text: str):
        self.emit(Msg("h1", text))

    def info(self, text: str, bullets: Iterable[str] = (), elapsed=None):
        self.emit(Msg("info", text, list(bullets), elapsed))

    def success(self, text: str, bullets: Iterable[str] = (), elapsed=None):
        self.emit(Msg("success", text, list(bullets), elapsed))

    def warn(self, text: str, bullets: Iterable[str] = (), elapsed=None):
        self.emit(Msg("warn", text, list(bullets), elapsed))

    def danger(self, text: str, bullets: Iterable[str] = (), elapsed=None):
        self.emit(Msg("danger", text, list(bullets), elapsed))

    def text(self, text: str, bullets: Iterable[str] = ()):
        self.emit(Msg("text", text, list(bullets)))

    @contextmanager
    def step(self, text: str):
        """进行中显示 'ℹ text ...', 完成后原地替换为 '✔ done [耗时]'

            with rep.step("读取 a.xlsx") as st:
                ...
                st.done = "读取 a.xlsx  100 行"     # 可选: 修改完成时的文字
        """
        st = _Step(text)
        self.emit(Msg("step", text))
        t0 = time.perf_counter()
        try:
            yield st
        except BaseException:
            self.emit(Msg("danger", st.fail or text, elapsed=time.perf_counter() - t0,
                          replace=True))
            raise
        kind = "warn" if st.warn else "success"
        self.emit(Msg(kind, st.done or text, st.bullets, time.perf_counter() - t0, replace=True))


class _Step:
    def __init__(self, text):
        self.done: str | None = None
        self.fail: str | None = None
        self.warn = False
        self.bullets: list[str] = []
        self.text = text


def as_reporter(log) -> Reporter:
    """接受 Reporter / 普通函数(如 print) / None, 统一成 Reporter"""
    if isinstance(log, Reporter):
        return log
    if log is None:
        return Reporter()
    return Reporter(PlainSink(log))


# ----------------------------------------------------------------------------
# 输出端
# ----------------------------------------------------------------------------
class PlainSink:
    """纯文本: 每条完成的消息调用 write(str); 进行中的 step 不输出"""

    def __init__(self, write: Callable[[str], None], sym=SYMBOLS):
        self.write, self.sym = write, sym
        self.lines: list[str] = []

    def __call__(self, m: Msg):
        if m.kind == "step":
            return
        s = render_plain(m, self.sym)
        self.lines.append(s)
        self.write(s)


class Tee:
    def __init__(self, *sinks):
        self.sinks = sinks

    def __call__(self, m: Msg):
        for s in self.sinks:
            try:
                s(m)
            except Exception:
                pass


_ANSI = {"success": "32", "info": "36", "step": "36", "warn": "33", "danger": "31",
         "file": "34", "val": "34", "field": "32", "strong": "1", "dim": "90", "h1": "1"}


def _a(code: str, s: str) -> str:
    return f"\x1b[{code}m{s}\x1b[0m"


class ConsoleSink:
    """终端输出; 支持颜色时上色, 是交互终端时进行中的步骤原地刷新"""

    def __init__(self, stream=None, color: bool | None = None):
        self.stream = stream or sys.stdout
        isatty = bool(getattr(self.stream, "isatty", lambda: False)())
        self.color = isatty if color is None else color
        self.live = isatty
        legacy_win = sys.platform.startswith("win") and not os.environ.get("WT_SESSION")
        ok = _can_encode(self.stream, "✔ℹ✖•─›") and not legacy_win
        self.sym = SYMBOLS if ok else SYMBOLS_WIN
        self._pending = False

    def _style(self, text: str) -> str:
        if not self.color:
            return plain(text)
        return "".join(_a(_ANSI[st], t) if st else t for st, t in parse(text))

    def _line(self, m: Msg) -> str:
        if m.kind == "h1":
            r = rule_line(plain(m.text), self.sym)
            return "\n" + (_a("1", r) if self.color else r)
        sym = self.sym.get(m.kind, "")
        head = (_a(_ANSI[m.kind], sym) if self.color else sym) + " " if sym else ""
        s = head + self._style(m.text)
        if m.kind == "step":
            s += " ..."
        if m.elapsed is not None:
            t = f" [{fmt_time(m.elapsed)}]"
            s += _a("90", t) if self.color else t
        for b in m.bullets:
            s += f"\n  {self.sym['bullet']} " + self._style(b)
        return s

    def __call__(self, m: Msg):
        w = self.stream.write
        if m.kind == "step":
            if self.live:
                w(self._line(m))
                self.stream.flush()
                self._pending = True
            return
        if self._pending:
            w("\r\x1b[2K" if m.replace else "\n")
            self._pending = False
        w(self._line(m) + "\n")
        self.stream.flush()


def _can_encode(stream, chars: str) -> bool:
    enc = getattr(stream, "encoding", None) or "ascii"
    try:
        chars.encode(enc)
        return True
    except (UnicodeEncodeError, LookupError):
        return False
