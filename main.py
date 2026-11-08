# -*- coding: utf-8 -*-
"""
FPF 数据处理工具 —— 读取 / 处理 / 导出 一体化

双击运行: 打开图形界面(页签: 母猪综合查询 / 猪只转移 / 母猪档案 / 品种品系判定)
命令行  : FPF数据处理.exe --cli 文件或文件夹 ... -o 输出路径 [--task sow|transfer|archive] [选项]
          FPF数据处理.exe --judge S332 托美 SYX32      (品种品系判定)
          FPF数据处理.exe --help                        (查看全部选项)
"""
from __future__ import annotations

import argparse
import multiprocessing
import os
import sys
import traceback

from fpf.app_io import (APP_NAME, APP_VERSION, log_path_for, parse_sheet, report_error,
                        write_log_file)

CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".fpf_tool.json")


# ==============================================================================
# 命令行
# ==============================================================================
def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="FPF数据处理", description=f"{APP_NAME} v{APP_VERSION}")
    p.add_argument("--cli", nargs="+", metavar="PATH", help="xlsx 文件或文件夹(可多个)")
    p.add_argument("--task", default="sow", choices=["sow", "transfer", "archive"],
                   help="sow=母猪综合查询(默认), transfer=猪只转移, archive=母猪档案")
    p.add_argument("-o", "--out", required=False, help="输出文件路径")
    p.add_argument("-f", "--format", default="rds", choices=["rds", "csv"])
    p.add_argument("--read-only", action="store_true", help="只读取合并, 不处理")
    p.add_argument("--processor", default=None, help=argparse.SUPPRESS)   # 旧版参数, 兼容
    p.add_argument("--skip", default="default", help="跳过的标题行数: 数字或 auto; 默认按任务")
    p.add_argument("--header-rows", default="default", help="标题行数: 数字或 auto; 默认按任务")
    p.add_argument("--header-fill", default="auto", choices=["auto", "right", "none"])
    p.add_argument("--sheet", default=None, help="工作表名或序号, 逗号分隔; 默认全部")
    p.add_argument("--data-type", default="all", choices=["all", "配种", "分娩", "断奶"],
                   help="母猪综合查询: 按哪个日期筛选")
    p.add_argument("--keep-missing-date", action="store_true",
                   help="母猪综合查询: 不删除关键日期缺失的记录")
    p.add_argument("--workers", type=int, default=1, help="并行读取进程数")
    p.add_argument("--on-error", default="warn", choices=["warn", "stop"])
    p.add_argument("--judge", nargs="+", metavar="TEXT", help="品种品系判定, 如 --judge S332 托美")
    p.add_argument("--rules", metavar="JSON",
                   help="自定义品种品系规则文件(在界面「品种品系判定」页签中保存/导出); 默认用内置规则")
    return p


def run_cli(args) -> int:
    from fpf.pipeline import run_pipeline
    from fpf.report import ConsoleSink, PlainSink, Reporter, Tee
    from fpf.rules import BreedRules

    file_sink = PlainSink(lambda s: None)
    rep = Reporter(Tee(ConsoleSink(sys.stdout), file_sink))
    log_file = log_path_for(args.out)
    tb = None
    code = 0
    options = ({"data_type": args.data_type, "drop_missing_date": not args.keep_missing_date}
               if args.task == "sow" else {})
    process = not (args.read_only or args.processor == "仅读取(不处理)")
    try:
        rules = BreedRules.load(args.rules) if args.rules else None
        run_pipeline(args.cli, task=args.task, process=process, skip=args.skip,
                     header_rows=args.header_rows, header_fill=args.header_fill,
                     sheet=parse_sheet(args.sheet), on_error=args.on_error,
                     workers=args.workers, options=options,
                     out_path=args.out, fmt=args.format, rules=rules, log=rep)
    except Exception as e:
        tb = traceback.format_exc()
        report_error(rep, e, log_file)
        code = 1
    write_log_file(log_file, file_sink.lines, tb)
    return code


def run_judge(values, rules_path=None) -> int:
    from fpf.breed import judge, split_inputs
    from fpf.report import text_width
    from fpf.rules import BreedRules, use_rules
    items = [x for v in values for x in split_inputs(v)]
    with use_rules(BreedRules.load(rules_path) if rules_path else None):
        df = judge(items)
    widths = {c: max(text_width(str(c)), *(text_width(str(x)) for x in df[c])) for c in df.columns}
    line = lambda vals: "  ".join(str(v) + " " * (widths[c] - text_width(str(v)))   # noqa: E731
                                  for c, v in zip(df.columns, vals)).rstrip()
    print(line(df.columns))
    for _, r in df.iterrows():
        print(line(r.values))
    return 0


def attach_parent_console():
    """打包后的 exe 是窗口程序, 从 cmd / PowerShell 运行时默认看不到输出。
    这里接回启动它的命令行窗口, 并开启颜色(ANSI)和 UTF-8。"""
    if not (sys.platform.startswith("win") and getattr(sys, "frozen", False)):
        return
    try:
        import ctypes
        import msvcrt
        k32 = ctypes.windll.kernel32
        if not k32.AttachConsole(-1):          # ATTACH_PARENT_PROCESS
            return
        k32.SetConsoleOutputCP(65001)
        out = open("CONOUT$", "w", encoding="utf-8", errors="replace", buffering=1)
        h = msvcrt.get_osfhandle(out.fileno())
        mode = ctypes.c_uint32()
        if k32.GetConsoleMode(h, ctypes.byref(mode)):
            k32.SetConsoleMode(h, mode.value | 0x0004)   # ENABLE_VIRTUAL_TERMINAL_PROCESSING
        sys.stdout = sys.stderr = out
        out.write("\n")   # 命令提示符已经先返回, 换行后再输出
    except Exception:
        pass


def main():
    multiprocessing.freeze_support()  # 打包后多进程读取必需
    if len(sys.argv) > 1:
        attach_parent_console()
    # 无控制台 exe 中 stdout/stderr 为 None, 任何 print 都会报错
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w", encoding="utf-8")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w", encoding="utf-8")
    if len(sys.argv) > 1:
        args = build_parser().parse_args()
        if args.judge:
            sys.exit(run_judge(args.judge, args.rules))
        if args.cli:
            sys.exit(run_cli(args))
        build_parser().print_help()
        return
    from fpf.gui import run_gui
    run_gui(APP_NAME, APP_VERSION, CONFIG_PATH)


if __name__ == "__main__":
    main()
