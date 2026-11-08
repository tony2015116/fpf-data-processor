# -*- coding: utf-8 -*-
"""
图形界面: 多页签
    母猪综合查询   读取 -> 清洗分类 -> 导出
    猪只转移       读取 -> 文件名拆分 + 品种品系判定 -> 导出
    母猪档案       读取 -> 文件名拆分 + 条数核对 + 品种品系判定 -> 导出
    品种品系判定   编辑判定规则(可应用到其他页签), 输入任意写法测试

新增"读取 -> 处理 -> 导出"类功能: 在 pipeline.TASKS 登记任务, 再在 TAB_SPECS 里
写一个参数区构建函数即可(见 _sow_params / _transfer_params)。
"""
from __future__ import annotations

import json
import os
import queue
import sys
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .app_io import log_path_for, parse_sheet, report_error, write_log_file
import time

from .breed import RULES_PATH, judge, load_saved_rules, split_inputs
from .pipeline import EXPORT_FORMATS, TASKS, collect_xlsx, run_pipeline
from .report import (SYMBOLS, SYMBOLS_WIN, Msg, PlainSink, Reporter, Tee, fmt_time,
                     parse, plain, rule_line)
from .rules import BREED_MODES, CODE_KINDS, COMBO_MODES, DEFAULT_RULES, BreedRules, use_rules
from .archive import FILENAME_PATTERN_HINT as ARCHIVE_PATTERN_HINT
from .transfer import FILENAME_PATTERN_HINT

PAD = dict(padx=6, pady=3)
MONO = ("Consolas", 10)


# ==============================================================================
# 日志框: 按 cli 风格上色
# ==============================================================================
class LogView:
    COLORS = {"success": "#1a7f37", "info": "#0969da", "step": "#0969da", "warn": "#bf8700",
              "danger": "#cf222e", "file": "#0550ae", "val": "#0550ae", "field": "#1a7f37",
              "dim": "#8c959f", "dangertext": "#cf222e"}

    def __init__(self, parent, height=10):
        frame = ttk.Frame(parent)
        self.frame = frame
        self.txt = tk.Text(frame, height=height, wrap="word", font=MONO, relief="flat",
                           background="#fbfbfb", padx=8, pady=6, spacing1=1)
        sb = ttk.Scrollbar(frame, orient="vertical", command=self.txt.yview)
        self.txt.configure(yscrollcommand=sb.set)
        self.txt.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        for tag, color in self.COLORS.items():
            self.txt.tag_configure(tag, foreground=color)
        self.txt.tag_configure("strong", font=MONO + ("bold",))
        self.txt.tag_configure("h1", foreground="#24292f", font=MONO + ("bold",))
        self.sym = SYMBOLS_WIN if sys.platform.startswith("win") else SYMBOLS
        self.step_line = None

    def clear(self):
        self.txt.delete("1.0", "end")
        self.step_line = None

    def _ins(self, text, base=None):
        for st, t in parse(text):
            self.txt.insert("end", t, tuple(x for x in (base, st) if x))

    def show(self, m: Msg):
        t, sym = self.txt, self.sym
        if self.step_line is not None:
            if m.replace:
                t.delete(f"{self.step_line}.0", "end-1c")
            self.step_line = None
        if m.kind == "h1":
            if t.index("end-1c") != "1.0":
                t.insert("end", "\n")
            t.insert("end", rule_line(plain(m.text), sym) + "\n", ("h1",))
        else:
            if m.kind == "step":
                self.step_line = int(t.index("end-1c").split(".")[0])
            if m.kind in sym and m.kind != "text":
                t.insert("end", sym[m.kind] + " ", (m.kind,))
            self._ins(m.text, "dangertext" if m.kind == "danger" else None)
            if m.kind == "step":
                t.insert("end", " ...", ("dim",))
            if m.elapsed is not None:
                t.insert("end", f" [{fmt_time(m.elapsed)}]", ("dim",))
            t.insert("end", "\n")
            for b in m.bullets:
                t.insert("end", f"  {sym['bullet']} ", ("dim",))
                self._ins(b)
                t.insert("end", "\n")
        t.see("end")


# ==============================================================================
# 读取 -> 处理 -> 导出 页签
# ==============================================================================
class TaskTab:
    def __init__(self, app, nb: ttk.Notebook, task_key: str, build_params):
        self.app, self.task = app, TASKS[task_key]
        self.key = task_key
        cfg = app.cfg.get(task_key, {})
        self.frame = ttk.Frame(nb, padding=8)
        nb.add(self.frame, text=f"  {self.task.name}  ")
        self.msgq: queue.Queue = queue.Queue()
        self.running = False
        self.files: list[str] = list(cfg.get("files", []))

        tk_default = lambda v: "auto" if v is None else str(v)   # noqa: E731
        self.v = {
            "skip": tk.StringVar(value=cfg.get("skip", tk_default(self.task.skip))),
            "header_rows": tk.StringVar(value=cfg.get("header_rows", tk_default(self.task.header_rows))),
            "header_fill": tk.StringVar(value=cfg.get("header_fill", "auto")),
            "sheet": tk.StringVar(value=cfg.get("sheet", "")),
            "workers": tk.StringVar(value=str(cfg.get("workers", 1))),
            "skip_bad": tk.BooleanVar(value=cfg.get("skip_bad", cfg.get("on_error_warn", True))),
            "read_only": tk.BooleanVar(value=cfg.get("read_only", False)),
            "format": tk.StringVar(value=cfg.get("format", "rds")),
            "out": tk.StringVar(value=cfg.get("out", "")),
            "rules": tk.StringVar(value=cfg.get("rules", "builtin")),   # builtin / custom
        }
        if self.v["format"].get() not in EXPORT_FORMATS:
            self.v["format"].set("rds")
        self.opt_vars: dict[str, tk.Variable] = {}   # 处理函数参数(各页签自定义)

        self._build_files()
        self._build_params(build_params, cfg)
        self._build_export()
        self._build_run()
        self._refresh()
        self._poll()

    # ---------------- 1. 文件 ----------------
    def _build_files(self):
        f = ttk.LabelFrame(self.frame, text="1. 数据文件 (xlsx)", padding=6)
        f.pack(fill="x")
        self.lb = tk.Listbox(f, height=4, selectmode="extended")
        self.lb.grid(row=0, column=0, rowspan=4, sticky="nsew")
        sb = ttk.Scrollbar(f, orient="vertical", command=self.lb.yview)
        sb.grid(row=0, column=1, rowspan=4, sticky="ns")
        self.lb.configure(yscrollcommand=sb.set)
        f.columnconfigure(0, weight=1)
        for i, (txt, cmd) in enumerate([("添加文件…", self._add_files), ("添加文件夹…", self._add_dir),
                                        ("移除选中", self._remove), ("清空", self._clear)]):
            ttk.Button(f, text=txt, command=cmd, width=12).grid(row=i, column=2, sticky="ew", padx=6, pady=1)
        self.lbl_count = ttk.Label(f, foreground="#555")
        self.lbl_count.grid(row=4, column=0, sticky="w")

    def _refresh(self):
        self.lb.delete(0, "end")
        for f in self.files:
            self.lb.insert("end", f)
        self.lbl_count.configure(text=f"共 {len(self.files)} 个文件")

    def _add_files(self):
        sel = filedialog.askopenfilenames(title="选择 xlsx 文件",
                                          filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("全部", "*.*")])
        for f in sel:
            f = os.path.normpath(f)
            if f not in self.files:
                self.files.append(f)
        self._refresh()

    def _add_dir(self):
        d = filedialog.askdirectory(title="选择文件夹(含子文件夹)")
        if d:
            for f in collect_xlsx([d]):
                if f not in self.files:
                    self.files.append(f)
            self._refresh()

    def _remove(self):
        for i in reversed(self.lb.curselection()):
            self.files.pop(i)
        self._refresh()

    def _clear(self):
        self.files.clear()
        self._refresh()

    # ---------------- 2. 参数 ----------------
    def _build_params(self, build_params, cfg):
        f = ttk.LabelFrame(self.frame, text="2. 读取与处理参数", padding=6)
        f.pack(fill="x", pady=(6, 0))
        self.pf = f
        build_params(self, f, cfg)

    def label(self, f, r, c, text):
        ttk.Label(f, text=text).grid(row=r, column=c, sticky="e", **PAD)

    def tip(self, f, r, c, text, span=1):
        ttk.Label(f, text=text, foreground="#888").grid(row=r, column=c, columnspan=span,
                                                        sticky="w", padx=(0, 6))

    def add_read_controls(self, f, row, header_choices=("auto", "1", "2", "3"), fill=True):
        """常用读取参数: 跳过行数 / 标题行数 / 合并单元格 / 工作表 / 并行 / 跳过坏文件 / 只读取"""
        self.label(f, row, 0, "跳过行数")
        ttk.Combobox(f, textvariable=self.v["skip"], values=["auto", "0", "1", "2", "3"],
                     width=7).grid(row=row, column=1, sticky="w", **PAD)
        self.label(f, row, 2, "标题行数")
        ttk.Combobox(f, textvariable=self.v["header_rows"], values=list(header_choices),
                     width=7).grid(row=row, column=3, sticky="w", **PAD)
        if fill:
            self.label(f, row, 4, "合并单元格")
            ttk.Combobox(f, textvariable=self.v["header_fill"], values=["auto", "right", "none"],
                         state="readonly", width=7).grid(row=row, column=5, sticky="w", **PAD)
        row += 1
        self.label(f, row, 0, "工作表")
        ttk.Entry(f, textvariable=self.v["sheet"], width=9).grid(row=row, column=1, sticky="w", **PAD)
        self.label(f, row, 2, "并行进程")
        ttk.Spinbox(f, textvariable=self.v["workers"], from_=1, to=16, width=6).grid(
            row=row, column=3, sticky="w", **PAD)
        self.tip(f, row, 4, "工作表留空=全部; 如 1 或 Sheet1,数据", span=3)
        row += 1
        ttk.Checkbutton(f, text="跳过读不了的文件", variable=self.v["skip_bad"]).grid(
            row=row, column=0, columnspan=2, sticky="w", **PAD)
        ttk.Checkbutton(f, text="只读取合并, 不处理", variable=self.v["read_only"]).grid(
            row=row, column=2, columnspan=2, sticky="w", **PAD)
        return row + 1

    def add_rules_control(self, f, row):
        """品种品系规则: 内置默认 / 自定义(在「品种品系判定」页签编辑并保存)"""
        self.label(f, row, 0, "品种品系规则")
        box = ttk.Frame(f)
        box.grid(row=row, column=1, columnspan=7, sticky="w", **PAD)
        ttk.Radiobutton(box, text="内置默认", value="builtin", variable=self.v["rules"]).pack(side="left")
        ttk.Radiobutton(box, text="自定义", value="custom", variable=self.v["rules"]).pack(
            side="left", padx=(10, 0))
        ttk.Label(box, text="  自定义规则在「品种品系判定」页签中编辑并保存", foreground="#888").pack(
            side="left")
        return row + 1

    def _resolve_rules(self):
        """返回 (是否继续, 规则); 规则 None = 内置默认"""
        if self.v["rules"].get() != "custom":
            return True, None
        try:
            rules = load_saved_rules()
        except Exception as e:
            messagebox.showerror(self.task.name, f"自定义规则文件有误:\n\n{e}")
            return False, None
        if rules is None:
            messagebox.showwarning(self.task.name, "还没有保存过自定义规则。\n\n"
                                   "请先在「品种品系判定」页签中编辑并点「保存规则」, "
                                   "或把本页的品种品系规则改为「内置默认」。")
            return False, None
        if self.app.breed.dirty and not messagebox.askyesno(
                self.task.name, "「品种品系判定」页签中有未保存的规则修改。\n\n"
                                "本次将使用上次保存的规则, 是否继续?"):
            return False, None
        return True, rules

    # ---------------- 3. 导出 ----------------
    def _build_export(self):
        f = ttk.LabelFrame(self.frame, text="3. 导出", padding=6)
        f.pack(fill="x", pady=(6, 0))
        self.label(f, 0, 0, "格式")
        ttk.Combobox(f, textvariable=self.v["format"], values=list(EXPORT_FORMATS),
                     state="readonly", width=7).grid(row=0, column=1, sticky="w", **PAD)
        self.fmt_tip = ttk.Label(f, foreground="#888")
        self.fmt_tip.grid(row=0, column=2, columnspan=2, sticky="w")
        self.label(f, 1, 0, "输出文件")
        ttk.Entry(f, textvariable=self.v["out"]).grid(row=1, column=1, columnspan=2, sticky="ew", **PAD)
        ttk.Button(f, text="浏览…", command=self._choose_out).grid(row=1, column=3, **PAD)
        f.columnconfigure(2, weight=1)
        self.v["format"].trace_add("write", self._on_fmt)
        self._on_fmt()

    def _on_fmt(self, *_):
        tips = {"rds": "R: readRDS(), 无需装包, 列类型原样保留", "csv": "通用, Excel 可直接打开"}
        self.fmt_tip.configure(text=tips.get(self.v["format"].get(), ""))
        out = self.v["out"].get().strip()
        if out:
            base, ext = os.path.splitext(out)
            if ext.lower() in (".rds", ".csv"):
                self.v["out"].set(base + EXPORT_FORMATS[self.v["format"].get()])

    def _choose_out(self):
        ext = EXPORT_FORMATS[self.v["format"].get()]
        name = self.task.default_out + ext
        init = self.v["out"].get() or (os.path.join(os.path.dirname(self.files[0]), name)
                                       if self.files else name)
        p = filedialog.asksaveasfilename(title="保存为", defaultextension=ext,
                                         initialdir=os.path.dirname(init) or None,
                                         initialfile=os.path.basename(init),
                                         filetypes=[(self.v["format"].get(), "*" + ext)])
        if p:
            self.v["out"].set(os.path.normpath(p))

    # ---------------- 运行 + 日志 ----------------
    def _build_run(self):
        bar = ttk.Frame(self.frame)
        bar.pack(fill="x", pady=(8, 0))
        self.btn_run = ttk.Button(bar, text="▶  开始处理", command=self.start)
        self.btn_run.pack(side="left")
        self.pb = ttk.Progressbar(bar, mode="indeterminate", length=180)
        self.pb.pack(side="left", padx=10)
        self.btn_open = ttk.Button(bar, text="打开输出文件夹", state="disabled", command=self._open_dir)
        self.btn_open.pack(side="right")
        self.log = LogView(self.frame)
        self.log.frame.pack(fill="both", expand=True, pady=(6, 0))

    def _poll(self):
        try:
            while True:
                kind, payload = self.msgq.get_nowait()
                if kind == "msg":
                    self.log.show(payload)
                elif kind == "out":
                    self.v["out"].set(payload)
                elif kind == "done":
                    self.running = False
                    self.pb.stop()
                    self.btn_run.configure(state="normal")
                    ok, info, partial = payload
                    if ok:
                        self.btn_open.configure(state="normal")
                        (messagebox.showwarning if partial else messagebox.showinfo)(
                            self.task.name, f"处理完成!\n\n{info}")
                    else:
                        messagebox.showerror(self.task.name, f"处理失败:\n\n{info}")
        except queue.Empty:
            pass
        self.frame.after(100, self._poll)

    def config(self) -> dict:
        d = {k: v.get() for k, v in self.v.items()}
        d.update({k: v.get() for k, v in self.opt_vars.items()})
        d["files"] = list(self.files)
        return d

    def start(self):
        if self.running:
            return
        if not self.files:
            messagebox.showwarning(self.task.name, "请先添加 xlsx 文件")
            return
        if not self.v["out"].get().strip():
            self._choose_out()
            if not self.v["out"].get().strip():
                return
        try:
            workers = max(1, int(self.v["workers"].get()))
        except ValueError:
            workers = 1
        ok, rules = self._resolve_rules()
        if not ok:
            return
        self.app.save_cfg()
        self.log.clear()
        self.btn_run.configure(state="disabled")
        self.btn_open.configure(state="disabled")
        self.pb.start(12)
        self.running = True
        opts = dict(task=self.key, process=not self.v["read_only"].get(),
                    skip=self.v["skip"].get().strip() or "auto",
                    header_rows=self.v["header_rows"].get().strip() or "auto",
                    header_fill=self.v["header_fill"].get(), sheet=parse_sheet(self.v["sheet"].get()),
                    on_error="warn" if self.v["skip_bad"].get() else "stop", workers=workers,
                    options={k: v.get() for k, v in self.opt_vars.items()},
                    out_path=self.v["out"].get().strip(), fmt=self.v["format"].get(),
                    rules=rules)
        files = list(self.files)
        threading.Thread(target=self._work, args=(files, opts), daemon=True).start()

    def _work(self, files, opts):
        file_sink = PlainSink(lambda s: None)
        rep = Reporter(Tee(lambda m: self.msgq.put(("msg", m)), file_sink))
        log_file = log_path_for(opts["out_path"])
        try:
            df, written = run_pipeline(files, log=rep, **opts)
            write_log_file(log_path_for(written), file_sink.lines)
            self.msgq.put(("out", written))
            unread = df.attrs.get("unread") or []
            info = f"{df.shape[0]:,} 行 × {df.shape[1]} 列\n{written}"
            if unread:
                names = "\n".join(os.path.basename(p) + (f" [{sh}]" if sh else "")
                                  for p, sh, _ in unread[:5])
                more = f"\n… 还有 {len(unread) - 5} 个" if len(unread) > 5 else ""
                info += (f"\n\n注意: 有 {len(unread)} 个文件 / 工作表无法读取, 结果中不包含:\n"
                         f"{names}{more}\n\n原因见下方日志")
            self.msgq.put(("done", (True, info, bool(unread))))
        except Exception as e:
            tb = traceback.format_exc()
            report_error(rep, e, log_file)
            write_log_file(log_file, file_sink.lines, tb)
            msg = str(e).splitlines()[0] if str(e) else type(e).__name__
            self.msgq.put(("done", (False, msg, False)))

    def _open_dir(self):
        p = self.v["out"].get().strip()
        d = os.path.dirname(os.path.abspath(p)) if p else None
        if d and os.path.isdir(d):
            if sys.platform.startswith("win"):
                os.startfile(d)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                os.system(f'open "{d}"')
            else:
                os.system(f'xdg-open "{d}"')


# ---------------- 各页签的参数区 ----------------
def _sow_params(tab: TaskTab, f, cfg):
    row = tab.add_read_controls(f, 0)
    v_dtype = tk.StringVar(value=cfg.get("data_type", "all"))
    v_drop = tk.BooleanVar(value=cfg.get("drop_missing_date", True))
    tab.opt_vars.update(data_type=v_dtype, drop_missing_date=v_drop)
    tab.label(f, 0, 6, "数据类型")
    ttk.Combobox(f, textvariable=v_dtype, values=["all", "配种", "分娩", "断奶"],
                 state="readonly", width=7).grid(row=0, column=7, sticky="w", **PAD)
    ttk.Checkbutton(f, text="删除关键日期缺失的记录", variable=v_drop).grid(
        row=row - 1, column=4, columnspan=4, sticky="w", **PAD)
    row = tab.add_rules_control(f, row)
    tab.tip(f, row, 0, "跳过行数 / 标题行数为 auto 时, 按每个工作表自动识别", span=8)


def _transfer_params(tab: TaskTab, f, cfg):
    row = tab.add_read_controls(f, 0, header_choices=("1", "2", "3"), fill=False)
    row = tab.add_rules_control(f, row)
    tab.tip(f, row, 0, "猪只转移文件为一行标题; 只读取文件名符合以下规则的文件, 并从文件名拆出单据信息:", span=8)
    ttk.Label(f, text="    " + FILENAME_PATTERN_HINT, foreground="#0550ae").grid(
        row=row + 1, column=0, columnspan=8, sticky="w", padx=6)


def _archive_params(tab: TaskTab, f, cfg):
    row = tab.add_read_controls(f, 0, header_choices=("1", "2", "3"), fill=False)
    row = tab.add_rules_control(f, row)
    tab.tip(f, row, 0, "母猪档案文件为一行标题; 只读取文件名符合以下规则的文件, 并从文件名拆出档案信息:", span=8)
    ttk.Label(f, text="    " + ARCHIVE_PATTERN_HINT, foreground="#0550ae").grid(
        row=row + 1, column=0, columnspan=8, sticky="w", padx=6)


TAB_SPECS = [("sow", _sow_params), ("transfer", _transfer_params), ("archive", _archive_params)]


# ==============================================================================
# 品种品系判定 页签: 规则编辑 + 测试
# ==============================================================================
class EditableTable:
    """可编辑表格: 双击单元格修改; 新增 / 删除 / 上移 / 下移
    数据保存在 Python 端(不经 Tcl 转换), 编码 08 这类前导 0 不会丢失"""

    def __init__(self, parent, columns, height=7, on_change=None):
        # columns: [(列名, 宽度, 可选值列表或 None)]
        self.columns, self.on_change = columns, on_change or (lambda: None)
        self.frame = ttk.Frame(parent)
        names = [c[0] for c in columns]
        self.tree = ttk.Treeview(self.frame, columns=names, show="headings", height=height,
                                 selectmode="browse")
        for name, width, _ in columns:
            self.tree.heading(name, text=name)
            self.tree.column(name, width=width, anchor="w", stretch=True)
        sb = ttk.Scrollbar(self.frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        btns = ttk.Frame(self.frame)
        btns.pack(side="left", fill="y", padx=(6, 0))
        for txt, cmd in [("新增", self.add), ("删除", self.delete), ("上移", lambda: self.move(-1)),
                         ("下移", lambda: self.move(1))]:
            ttk.Button(btns, text=txt, width=6, command=cmd).pack(pady=1)
        self.data: dict[str, list[str]] = {}
        self.editor = None
        self.tree.bind("<Double-1>", self._begin_edit)

    # ---------------- 数据 ----------------
    def set_rows(self, rows):
        self._end_edit(commit=False)
        self.tree.delete(*self.tree.get_children())
        self.data.clear()
        for r in rows:
            self._insert("end", [str(x) for x in r])

    def get_rows(self) -> list[tuple]:
        return [tuple(self.data[i]) for i in self.tree.get_children()]

    def _insert(self, index, values):
        iid = self.tree.insert("", index, values=values)
        self.data[iid] = list(values)
        return iid

    def _refresh(self, iid):
        self.tree.item(iid, values=self.data[iid])

    # ---------------- 操作 ----------------
    def add(self):
        self._end_edit()
        sel = self.tree.selection()
        index = self.tree.index(sel[0]) + 1 if sel else "end"
        vals = [c[2][0] if c[2] else "" for c in self.columns]
        iid = self._insert(index, vals)
        self.tree.selection_set(iid)
        self.tree.see(iid)
        self.on_change()
        first_free = next((k for k, c in enumerate(self.columns) if not c[2]), 0)
        self.frame.after(50, lambda: self._edit_cell(iid, first_free))

    def delete(self):
        self._end_edit(commit=False)
        for iid in self.tree.selection():
            self.tree.delete(iid)
            self.data.pop(iid, None)
            self.on_change()

    def move(self, step):
        self._end_edit()
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]
        idx = self.tree.index(iid) + step
        if 0 <= idx < len(self.tree.get_children()):
            self.tree.move(iid, "", idx)
            self.tree.see(iid)
            self.on_change()

    # ---------------- 单元格编辑 ----------------
    def _begin_edit(self, event):
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        iid = self.tree.identify_row(event.y)
        col = int(self.tree.identify_column(event.x)[1:]) - 1
        if iid:
            self._edit_cell(iid, col)

    def _edit_cell(self, iid, col):
        self._end_edit()
        bbox = self.tree.bbox(iid, col)
        if not bbox:
            return
        x, y, w, h = bbox
        choices = self.columns[col][2]
        var = tk.StringVar(value=self.data[iid][col])
        if choices:
            w_ = ttk.Combobox(self.tree, textvariable=var, values=list(choices), state="readonly")
            w_.bind("<<ComboboxSelected>>", lambda e: self._end_edit())
        else:
            w_ = ttk.Entry(self.tree, textvariable=var)
            w_.select_range(0, "end")
        w_.place(x=x, y=y, width=w, height=h)
        w_.focus_set()
        w_.bind("<Return>", lambda e: self._end_edit())
        w_.bind("<Escape>", lambda e: self._end_edit(commit=False))
        w_.bind("<FocusOut>", lambda e: self._end_edit())
        self.editor = (w_, var, iid, col)

    def _end_edit(self, commit=True):
        if not self.editor:
            return
        w_, var, iid, col = self.editor
        self.editor = None
        if commit and iid in self.data:
            new = var.get().strip()
            if new != self.data[iid][col]:
                self.data[iid][col] = new
                self._refresh(iid)
                self.on_change()
        w_.destroy()


class BreedTab:
    EXAMPLE = "S332\nSYX32\n托美\nX31\n长白\n二元\nP"
    RESULT_COLS = ["输入", "品种", "品系", "品系编码", "血统构成", "与内置默认对比"]

    def __init__(self, app, nb: ttk.Notebook):
        self.app = app
        cfg = app.cfg.get("breed", {})
        self.frame = ttk.Frame(nb, padding=8)
        nb.add(self.frame, text="  品种品系判定  ")
        self.dirty = False
        self.result = None
        self._build_editor()
        self._build_test(cfg)
        self._load_initial()

    # ======================== 规则编辑 ========================
    def _build_editor(self):
        box = ttk.LabelFrame(self.frame, text="判定规则 (双击单元格修改; 保存后可在其他页签选择「自定义」使用)",
                             padding=6)
        box.pack(fill="x")
        top = ttk.Frame(box)
        top.pack(fill="x")
        self.lbl_status = ttk.Label(top, foreground="#555")
        self.lbl_status.pack(side="left")
        self.lbl_dirty = ttk.Label(top, foreground="#cf222e")
        self.lbl_dirty.pack(side="left", padx=8)

        nb = ttk.Notebook(box)
        nb.pack(fill="x", pady=(4, 0))
        mark = lambda: self._set_dirty(True)   # noqa: E731

        # 品种规则
        f1 = ttk.Frame(nb, padding=4)
        nb.add(f1, text="品种规则")
        ttk.Label(f1, foreground="#666", text=(
            "按顺序匹配, 第一个命中的生效。匹配对象已去空格、转大写 (如 \" s 332\" → \"S332\")。"
            "包含 / 开头是 / 等于 可用 | 分隔多个关键词。")).pack(anchor="w")
        self.t_breed = EditableTable(f1, [("匹配方式", 80, BREED_MODES), ("匹配内容", 260, None),
                                          ("品种", 140, None)], on_change=mark)
        self.t_breed.frame.pack(fill="x", pady=2)
        fb = ttk.Frame(f1)
        fb.pack(anchor="w")
        ttk.Label(fb, text="都不匹配时:").pack(side="left")
        self.v_breed_fb = tk.StringVar()
        e = ttk.Entry(fb, textvariable=self.v_breed_fb, width=12)
        e.pack(side="left", padx=4)
        self.v_breed_fb.trace_add("write", lambda *_: mark())

        # 品系编码
        f2 = ttk.Frame(nb, padding=4)
        nb.add(f2, text="品系编码")
        ttk.Label(f2, foreground="#666", text=(
            "单一品系: 一位数字, 多位编码按级进杂交逐位拆解 (最左边是最近一次杂交); "
            "合成系: 多位数字整体代表一个品系。")).pack(anchor="w")
        self.t_codes = EditableTable(f2, [("编码", 80, None), ("名称", 200, None),
                                          ("类型", 100, CODE_KINDS)], on_change=mark)
        self.t_codes.frame.pack(fill="x", pady=2)

        # 中文名称
        f3 = ttk.Frame(nb, padding=4)
        nb.add(f3, text="中文名称")
        ttk.Label(f3, foreground="#666", text=(
            "写法里没有可识别的数字编码时, 按顺序看是否包含这些名称并换算成编码 (如 托美 → 32)。"
            "复合名称要排在单一名称前面。")).pack(anchor="w")
        self.t_text = EditableTable(f3, [("名称", 200, None), ("编码", 100, None)], on_change=mark)
        self.t_text.frame.pack(fill="x", pady=2)

        # 品系组合
        f4 = ttk.Frame(nb, padding=4)
        nb.add(f4, text="品系组合")
        ttk.Label(f4, foreground="#666", text=(
            "判定顺序: ① 编码等于 (如 80 → 含法)  ② 单一品系编码 (如 3、03 → 托佩克, 来自品系编码表)  "
            "③ 数字恰好是 / 包含数字 (只用于含两种及以上数字的编码, 按顺序第一个命中的生效)。"),
            wraplength=820).pack(anchor="w")
        self.t_combo = EditableTable(f4, [("条件", 100, COMBO_MODES), ("数字", 100, None),
                                          ("品系", 160, None)], on_change=mark)
        self.t_combo.frame.pack(fill="x", pady=2)
        fl = ttk.Frame(f4)
        fl.pack(anchor="w")
        ttk.Label(fl, text="都不匹配时:").pack(side="left")
        self.v_line_fb = tk.StringVar()
        ttk.Entry(fl, textvariable=self.v_line_fb, width=12).pack(side="left", padx=4)
        self.v_line_fb.trace_add("write", lambda *_: mark())

        bar = ttk.Frame(box)
        bar.pack(fill="x", pady=(6, 0))
        ttk.Button(bar, text="保存规则", command=self.save).pack(side="left")
        ttk.Button(bar, text="恢复为内置默认", command=self.reset_default).pack(side="left", padx=6)
        ttk.Button(bar, text="导入…", command=self.import_rules).pack(side="left")
        ttk.Button(bar, text="导出…", command=self.export_rules).pack(side="left", padx=6)

    def _set_dirty(self, v: bool):
        if getattr(self, "_loading", False):
            return
        self.dirty = v
        self.lbl_dirty.configure(text="● 有未保存的修改" if v else "")

    def _fill(self, r: BreedRules):
        self._loading = True
        try:
            self.t_breed.set_rows(r.breed_rules)
            self.v_breed_fb.set(r.breed_fallback)
            self.t_codes.set_rows(r.line_codes)
            self.t_text.set_rows(r.line_text)
            self.t_combo.set_rows(r.line_combos)
            self.v_line_fb.set(r.line_fallback)
        finally:
            self._loading = False

    def editor_rules(self) -> BreedRules:
        for t in (self.t_breed, self.t_codes, self.t_text, self.t_combo):
            t._end_edit()
        return BreedRules(breed_rules=self.t_breed.get_rows(), breed_fallback=self.v_breed_fb.get().strip(),
                          line_codes=self.t_codes.get_rows(), line_text=self.t_text.get_rows(),
                          line_combos=self.t_combo.get_rows(), line_fallback=self.v_line_fb.get().strip())

    def _status(self):
        try:
            saved = load_saved_rules()
        except Exception:
            saved = None
        if saved is None:
            self.lbl_status.configure(text="尚未保存自定义规则, 下面是内置默认规则, 修改后点「保存规则」")
        else:
            self.lbl_status.configure(text=f"已保存的自定义规则: {saved.summary()}"
                                      + (f"  (保存于 {saved.saved_at})" if saved.saved_at else ""))

    def _load_initial(self):
        try:
            saved = load_saved_rules()
        except Exception as e:
            saved = None
            messagebox.showwarning("品种品系判定", f"自定义规则文件有误, 已显示内置默认规则:\n\n{e}")
        self._fill(saved or DEFAULT_RULES)
        self._set_dirty(False)
        self._status()

    def _check(self, r: BreedRules) -> bool:
        errs = r.validate()
        if errs:
            more = f"\n… 还有 {len(errs) - 12} 处" if len(errs) > 12 else ""
            messagebox.showerror("品种品系判定", "规则有误, 请修改:\n\n" + "\n".join(errs[:12]) + more)
            return False
        return True

    def save(self):
        r = self.editor_rules()
        if not self._check(r):
            return
        r.saved_at = time.strftime("%Y-%m-%d %H:%M")
        try:
            r.save(RULES_PATH)
        except Exception as e:
            messagebox.showerror("品种品系判定", f"保存失败:\n\n{e}")
            return
        self._set_dirty(False)
        self._status()
        self.lbl.configure(text="规则已保存")

    def reset_default(self):
        if messagebox.askyesno("品种品系判定", "把编辑区恢复为内置默认规则?\n\n(点「保存规则」后才会生效)"):
            self._fill(DEFAULT_RULES)
            self._set_dirty(True)

    def import_rules(self):
        p = filedialog.askopenfilename(title="导入规则", filetypes=[("规则文件", "*.json")])
        if not p:
            return
        try:
            r = BreedRules.load(p)
        except Exception as e:
            messagebox.showerror("品种品系判定", f"导入失败:\n\n{e}")
            return
        self._fill(r)
        self._set_dirty(True)
        self.lbl.configure(text="已导入, 点「保存规则」后生效")

    def export_rules(self):
        r = self.editor_rules()
        if not self._check(r):
            return
        p = filedialog.asksaveasfilename(title="导出规则", defaultextension=".json",
                                         initialfile="品种品系规则.json", filetypes=[("规则文件", "*.json")])
        if p:
            r.saved_at = time.strftime("%Y-%m-%d %H:%M")
            r.save(p)
            self.lbl.configure(text=f"已导出 {p}")

    # ======================== 测试 ========================
    def _build_test(self, cfg):
        box = ttk.LabelFrame(self.frame, text="测试", padding=6)
        box.pack(fill="both", expand=True, pady=(8, 0))
        bar = ttk.Frame(box)
        bar.pack(fill="x")
        ttk.Label(bar, text="使用:").pack(side="left")
        self.v_use = tk.StringVar(value=cfg.get("use", "editor"))
        ttk.Radiobutton(bar, text="编辑中的规则", value="editor", variable=self.v_use).pack(side="left", padx=4)
        ttk.Radiobutton(bar, text="内置默认", value="builtin", variable=self.v_use).pack(side="left")
        ttk.Button(bar, text="▶  判定", command=self.run).pack(side="left", padx=(16, 0))
        ttk.Button(bar, text="复制结果", command=self.copy).pack(side="left", padx=6)
        ttk.Button(bar, text="导出 csv…", command=self.export).pack(side="left")
        self.lbl = ttk.Label(bar, foreground="#555")
        self.lbl.pack(side="right")

        body = ttk.Frame(box)
        body.pack(fill="both", expand=True, pady=(6, 0))
        left = ttk.Frame(body)
        left.pack(side="left", fill="y")
        ttk.Label(left, text="输入 (一行一个)", foreground="#555").pack(anchor="w")
        self.inp = tk.Text(left, height=8, width=18, font=MONO, wrap="none")
        self.inp.pack(fill="y", expand=True)
        self.inp.insert("1.0", cfg.get("text", self.EXAMPLE))
        self.inp.bind("<Control-Return>", lambda e: (self.run(), "break")[1])

        tf = ttk.Frame(body)
        tf.pack(side="left", fill="both", expand=True, padx=(8, 0))
        self.tree = ttk.Treeview(tf, columns=self.RESULT_COLS, show="headings", height=8)
        widths = {"输入": 100, "品种": 100, "品系": 70, "品系编码": 70, "血统构成": 190,
                  "与内置默认对比": 190}
        for c in self.RESULT_COLS:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=widths[c], anchor="w", stretch=c in ("血统构成", "与内置默认对比"))
        sb = ttk.Scrollbar(tf, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.tag_configure("other", foreground="#bf8700")
        self.tree.tag_configure("diff", foreground="#0550ae")

    def config(self) -> dict:
        return {"text": self.inp.get("1.0", "end-1c"), "use": self.v_use.get()}

    def run(self):
        values = split_inputs(self.inp.get("1.0", "end"))
        self.tree.delete(*self.tree.get_children())
        if not values:
            self.lbl.configure(text="请先输入")
            return
        use_editor = self.v_use.get() == "editor"
        rules = None
        if use_editor:
            rules = self.editor_rules()
            if not self._check(rules):
                return
        with use_rules(rules):
            res = judge(values)
        base = judge(values) if use_editor else res
        diff = []
        for i in range(len(res)):
            changed = [f"{c} {base.at[i, c]}" for c in ("品种", "品系")
                       if base.at[i, c] != res.at[i, c]]
            diff.append(("内置: " + ", ".join(changed)) if changed else "")
        res["与内置默认对比"] = diff
        self.result = res
        fb = (rules or DEFAULT_RULES)
        for _, r in res.iterrows():
            tags = ("diff",) if r["与内置默认对比"] else \
                (("other",) if r["品种"] == fb.breed_fallback or r["品系"] == fb.line_fallback else ())
            self.tree.insert("", "end", values=[r[c] for c in self.RESULT_COLS], tags=tags)
        n_diff = sum(1 for d in diff if d)
        n_other = int(((res["品种"] == fb.breed_fallback) | (res["品系"] == fb.line_fallback)).sum())
        msg = f"共 {len(values)} 个"
        if use_editor and n_diff:
            msg += f" · 与内置不同 {n_diff} (蓝)"
        if n_other:
            msg += f" · 未识别 {n_other} (黄)"
        self.lbl.configure(text=msg)
        self.app.save_cfg()

    def copy(self):
        if self.result is None:
            return
        self.frame.clipboard_clear()
        self.frame.clipboard_append(self.result.to_csv(sep="\t", index=False))
        self.lbl.configure(text="已复制, 可直接粘贴到 Excel")

    def export(self):
        if self.result is None:
            return
        p = filedialog.asksaveasfilename(title="导出判定结果", defaultextension=".csv",
                                         initialfile="品种品系判定.csv", filetypes=[("csv", "*.csv")])
        if p:
            self.result.to_csv(p, index=False, encoding="utf-8-sig")
            self.lbl.configure(text=f"已导出 {p}")


# ==============================================================================
# 主窗口
# ==============================================================================
class App:
    def __init__(self, app_name: str, version: str, config_path: str):
        self.config_path = config_path
        self.cfg = self._load_cfg()
        self.root = root = tk.Tk()
        root.title(f"{app_name}  v{version}")
        root.geometry("960x900")
        root.minsize(820, 720)
        if sys.platform.startswith("win"):
            try:  # 高分屏下不模糊
                import ctypes
                ctypes.windll.shcore.SetProcessDpiAwareness(1)
            except Exception:
                pass
        style = ttk.Style()
        for theme in ("vista", "clam"):
            if theme in style.theme_names():
                style.theme_use(theme)
                break
        style.configure("TNotebook.Tab", padding=(10, 4))

        self.nb = ttk.Notebook(root)
        self.nb.pack(fill="both", expand=True, padx=8, pady=8)
        self.tabs = {key: TaskTab(self, self.nb, key, builder) for key, builder in TAB_SPECS}
        self.breed = BreedTab(self, self.nb)
        frames = {k: t.frame for k, t in self.tabs.items()}
        frames["breed"] = self.breed.frame
        key = self.cfg.get("tab_key")
        if key is None and "tab" in self.cfg:            # 旧版按序号记录: 0 母猪综合 1 转移 2 判定
            key = {0: "sow", 1: "transfer", 2: "breed"}.get(self.cfg.get("tab"))
        if key in frames:
            self.nb.select(frames[key])
        self._frames = frames
        root.protocol("WM_DELETE_WINDOW", self._close)

    def _load_cfg(self) -> dict:
        try:
            with open(self.config_path, encoding="utf-8") as f:
                cfg = json.load(f)
        except Exception:
            return {}
        if "sow" not in cfg and "files" in cfg:      # 旧版单页设置 -> 母猪综合查询
            cfg = {"sow": cfg}
        return cfg

    def save_cfg(self):
        data = {k: t.config() for k, t in self.tabs.items()}
        data["breed"] = self.breed.config()
        try:
            cur = str(self.nb.select())
            data["tab_key"] = next((k for k, f in self._frames.items() if str(f) == cur), "sow")
        except Exception:
            pass
        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=1)
        except Exception:
            pass

    def _close(self):
        self.save_cfg()
        self.root.destroy()

    def run(self):
        auto = os.environ.get("FPF_AUTORUN", "")        # 自动化测试用: sow / transfer / archive / breed
        if auto:
            key = "sow" if auto == "1" else auto
            if key == "breed":
                self.nb.select(self.breed.frame)
                self.root.after(500, self.breed.run)
            elif key in self.tabs:
                self.nb.select(self.tabs[key].frame)
                self.root.after(500, self.tabs[key].start)
        self.root.mainloop()


def run_gui(app_name: str, version: str, config_path: str):
    App(app_name, version, config_path).run()
