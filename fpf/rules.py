# -*- coding: utf-8 -*-
"""
品种 / 品系判定规则

规则是一份数据(可保存为 JSON), 判定函数(cleaning.py)从"当前生效的规则"中读取:

    with use_rules(my_rules):          # 只在这个 with 块(及当前线程)内生效
        df = clean_production_data(raw)

不调用 use_rules 时使用内置默认规则 DEFAULT_RULES, 与 R 版
综合选择查询数据处理_通用.R 的判定逻辑完全一致。

规则的组成
----------
品种规则 breed_rules  : 按顺序匹配, 第一个命中的给出品种; 都不命中 -> breed_fallback
    每条 = (匹配方式, 匹配内容, 品种)
    匹配方式: 正则 / 包含 / 开头是 / 等于   (包含/开头是/等于 可用 | 分隔多个关键词)
    匹配对象是标准化后的写法: 去空格、字母转大写, 如 " s 332" -> "S332"

品系编码 line_codes   : (编码, 名称, 类型)
    类型 = 单一品系: 一位数字, 如 3 = 托佩克; 多位编码按级进杂交逐位拆解
    类型 = 合成系  : 多位数字整体代表一个品系, 如 30 = 合成系纯种/二元, 不逐位拆解

中文名称 line_text    : (名称, 编码) 按顺序匹配; 写法里没有可识别的数字编码时,
    看是否包含这些名称, 如 "托美" -> 32。复合名称要排在单一名称前面

品系组合 line_combos  : (条件, 内容, 品系) 判定顺序:
    1. 条件 = 编码等于   : 编码与内容完全相同, 如 80 -> 含法   (最先判断)
    2. 单一品系编码       : 只含一种数字, 如 3 / 03 -> 托佩克 (自动, 来自品系编码表)
    3. 条件 = 数字恰好是 : 编码中出现的数字集合与内容相同, 如 31 / 13 -> 托丹
       条件 = 包含数字   : 编码中出现了内容里的全部数字, 如 包含 8 -> 含法
       (3 只用于含两种及以上数字的编码, 按表中顺序, 第一个命中的生效)
    都不命中 -> line_fallback
"""
from __future__ import annotations

import contextvars
import json
import re
from contextlib import contextmanager
from dataclasses import dataclass, field

BREED_MODES = ("正则", "包含", "开头是", "等于")
CODE_KINDS = ("单一品系", "合成系")
COMBO_MODES = ("编码等于", "数字恰好是", "包含数字")


def _norm_kw(s: str) -> str:
    """关键词与输入用同样的方式标准化: 去空格、转大写"""
    return re.sub(r"\s+", "", str(s).strip().upper())


@dataclass
class BreedRules:
    breed_rules: list = field(default_factory=list)     # [(方式, 内容, 品种)]
    breed_fallback: str = "其它"
    line_codes: list = field(default_factory=list)      # [(编码, 名称, 类型)]
    line_text: list = field(default_factory=list)       # [(名称, 编码)]
    line_combos: list = field(default_factory=list)     # [(条件, 内容, 品系)]
    line_fallback: str = "其它"
    name: str = "自定义"
    saved_at: str = ""

    # ---------------- 派生(编译后)数据 ----------------
    def __post_init__(self):
        self.compile()

    def compile(self):
        self._breed = []
        for mode, pat, label in self.breed_rules:
            if mode == "正则":
                try:
                    rx = re.compile(pat)
                except re.error:          # 写错的正则不参与匹配, 由 validate() 报告
                    continue
                self._breed.append((lambda x, rx=rx: rx.search(x) is not None, label))
            else:
                kws = [_norm_kw(k) for k in str(pat).split("|") if _norm_kw(k)]
                if mode == "包含":
                    fn = lambda x, kws=kws: any(k in x for k in kws)          # noqa: E731
                elif mode == "开头是":
                    fn = lambda x, kws=kws: any(x.startswith(k) for k in kws)  # noqa: E731
                else:
                    fn = lambda x, kws=kws: x in kws                           # noqa: E731
                self._breed.append((fn, label))
        self.single = {str(c): n for c, n, k in self.line_codes if k == "单一品系"}
        self.synthetic = {str(c): n for c, n, k in self.line_codes if k == "合成系"}
        self.text = [(_norm_kw(n), str(c)) for n, c in self.line_text if _norm_kw(n)]
        self.exact = {str(v): lab for m, v, lab in self.line_combos if m == "编码等于"}
        self.combos = [(m, set(str(v)) - {"0"}, lab) for m, v, lab in self.line_combos
                       if m != "编码等于"]

    def line_name(self, digit: str, default: str) -> str:
        return self.single.get(digit, default)

    # ---------------- 校验 ----------------
    def validate(self) -> list[str]:
        err = []
        if not self.breed_fallback.strip():
            err.append("品种规则: 「都不匹配时」不能为空")
        if not self.line_fallback.strip():
            err.append("品系组合: 「都不匹配时」不能为空")
        for i, (mode, pat, label) in enumerate(self.breed_rules, 1):
            if mode not in BREED_MODES:
                err.append(f"品种规则第 {i} 行: 匹配方式应为 {'/'.join(BREED_MODES)}")
            if not str(pat).strip():
                err.append(f"品种规则第 {i} 行: 匹配内容为空")
            elif mode == "正则":
                try:
                    re.compile(pat)
                except re.error as e:
                    err.append(f"品种规则第 {i} 行: 正则写法有误 ({e})")
            if not str(label).strip():
                err.append(f"品种规则第 {i} 行: 品种为空")
        seen = set()
        for i, (code, name, kind) in enumerate(self.line_codes, 1):
            code = str(code).strip()
            if kind == "单一品系" and not re.fullmatch(r"[1-9]", code):
                err.append(f"品系编码第 {i} 行: 单一品系的编码应为一位数字 1-9")
            if kind == "合成系" and not re.fullmatch(r"[0-9]{2,}", code):
                err.append(f"品系编码第 {i} 行: 合成系的编码应为两位及以上数字")
            if kind not in CODE_KINDS:
                err.append(f"品系编码第 {i} 行: 类型应为 单一品系 / 合成系")
            if code in seen:
                err.append(f"品系编码第 {i} 行: 编码 {code} 重复")
            seen.add(code)
            if not str(name).strip():
                err.append(f"品系编码第 {i} 行: 名称为空")
        for i, (name, code) in enumerate(self.line_text, 1):
            if not str(name).strip():
                err.append(f"中文名称第 {i} 行: 名称为空")
            if not re.fullmatch(r"[0-9]+", str(code).strip()):
                err.append(f"中文名称第 {i} 行: 编码应为数字")
        for i, (mode, val, label) in enumerate(self.line_combos, 1):
            if mode not in COMBO_MODES:
                err.append(f"品系组合第 {i} 行: 条件应为 {'/'.join(COMBO_MODES)}")
            if not re.fullmatch(r"[0-9]+", str(val).strip()):
                err.append(f"品系组合第 {i} 行: 内容应为数字")
            if not str(label).strip():
                err.append(f"品系组合第 {i} 行: 品系为空")
        return err

    # ---------------- 保存 / 读取 ----------------
    def to_dict(self) -> dict:
        return {
            "name": self.name, "saved_at": self.saved_at,
            "breed_rules": [list(r) for r in self.breed_rules],
            "breed_fallback": self.breed_fallback,
            "line_codes": [list(r) for r in self.line_codes],
            "line_text": [list(r) for r in self.line_text],
            "line_combos": [list(r) for r in self.line_combos],
            "line_fallback": self.line_fallback,
        }

    @classmethod
    def from_dict(cls, d: dict) -> "BreedRules":
        t = lambda key, n: [tuple(str(x) for x in r[:n]) for r in d.get(key, [])]  # noqa: E731
        return cls(breed_rules=t("breed_rules", 3), breed_fallback=str(d.get("breed_fallback", "其它")),
                   line_codes=t("line_codes", 3), line_text=t("line_text", 2),
                   line_combos=t("line_combos", 3), line_fallback=str(d.get("line_fallback", "其它")),
                   name=str(d.get("name", "自定义")), saved_at=str(d.get("saved_at", "")))

    def save(self, path: str):
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, ensure_ascii=False, indent=1)

    @classmethod
    def load(cls, path: str) -> "BreedRules":
        with open(path, encoding="utf-8") as f:
            r = cls.from_dict(json.load(f))
        errs = r.validate()
        if errs:
            raise ValueError("规则文件有误:\n" + "\n".join(f"  - {e}" for e in errs))
        return r

    def summary(self) -> str:
        return (f"品种规则 {len(self.breed_rules)} 条, 品系编码 {len(self.line_codes)} 个, "
                f"中文名称 {len(self.line_text)} 个, 品系组合 {len(self.line_combos)} 条")


# ==============================================================================
# 内置默认规则 (与 R 版完全一致)
# ==============================================================================
DEFAULT_RULES = BreedRules(
    name="内置默认",
    breed_rules=[
        ("正则", "三元杂|商品", "商品猪"),
        ("正则", "^SYX", "杜洛克商品猪"),
        ("正则", "^PYX", "皮特兰商品猪"),
        ("正则", "^YX|二元", "二元"),
        ("正则", r"^S\d+$|杜洛克", "杜洛克"),
        ("正则", r"^X\d+$|长白", "长白"),
        ("正则", r"^Y\d+$|大白", "大白"),
        ("正则", r"^P\d+$|皮特兰", "皮特兰"),
    ],
    breed_fallback="其它",
    line_codes=[
        ("1", "丹系", "单一品系"), ("2", "美系", "单一品系"), ("3", "托佩克", "单一品系"),
        ("6", "华系", "单一品系"), ("8", "法系", "单一品系"), ("9", "EB5", "单一品系"),
        ("30", "合成系纯种/二元", "合成系"), ("23", "合成系杜洛克/三元", "合成系"),
    ],
    line_text=[
        ("托丹", "31"), ("丹托", "31"), ("托法", "38"), ("法托", "38"),
        ("托美", "32"), ("托华", "36"),
        ("合成系纯种", "30"), ("合成系二元", "30"), ("合成系杜洛克", "23"), ("合成系三元", "23"),
        ("丹系", "1"), ("美系", "2"), ("托佩克", "3"), ("华系", "6"), ("法系", "8"), ("EB5", "9"),
    ],
    line_combos=[
        ("编码等于", "80", "含法"),
        ("数字恰好是", "31", "托丹"),
        ("数字恰好是", "38", "法托"),
        ("包含数字", "8", "含法"),
        ("包含数字", "3", "含托"),
    ],
    line_fallback="其它",
)


# ==============================================================================
# 当前生效的规则 (按线程隔离: 两个页签同时运行、各用各的规则也不会互相影响)
# ==============================================================================
_ACTIVE: contextvars.ContextVar = contextvars.ContextVar("fpf_breed_rules", default=DEFAULT_RULES)


def active_rules() -> BreedRules:
    return _ACTIVE.get()


@contextmanager
def use_rules(rules: BreedRules | None):
    token = _ACTIVE.set(rules or DEFAULT_RULES)
    try:
        yield
    finally:
        _ACTIVE.reset(token)
