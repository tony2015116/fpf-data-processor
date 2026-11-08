# -*- coding: utf-8 -*-
"""
综合选择查询 —— 生产数据清洗
R 源: 综合选择查询数据处理_通用.R  (clean_production_data 及全部辅助函数)

逐条对应 R 版逻辑; 分类函数先在"去重后的取值"上计算再映射回整列, 大表也很快。
"""
from __future__ import annotations

import re

import numpy as np
import pandas as pd

from .report import as_reporter, esc, fmt_n
from .rules import DEFAULT_RULES, active_rules


# ==============================================================================
# 0. 品系代码字典
#    判定规则已移到 rules.py(可自定义); 这里保留内置规则的只读视图, 兼容旧代码
# ==============================================================================
LINE_CODE_MAP = dict(DEFAULT_RULES.single)
SYNTHETIC_LINE_MAP = dict(DEFAULT_RULES.synthetic)
LINE_TEXT_TO_CODE = list(DEFAULT_RULES.text)

FARM_ORDER = [
    "金沙一场", "龙马一场", "龙马二场", "香江场",
    "郭祥一场", "郭祥三场",
    "西江场", "郭祥二场", "大庆一场", "德邦场1线", "德邦场2线",
    "凤凰一场", "凤凰二场", "大庆二场", "雄桂场", "信宜场",
]

FARM_TYPE_LEVELS = ["核心育种场", "扩繁一胎场", "扩繁经产场", "商品一胎场", "商品经产场", "其他"]

_NA_TOKENS = {"", "--", "NA", "N/A", "NULL"}


# ==============================================================================
# 小工具
# ==============================================================================
def _is_na(v) -> bool:
    if v is None or v is pd.NA or v is pd.NaT:
        return True
    try:
        return bool(isinstance(v, float) and np.isnan(v))
    except Exception:
        return False


def _r_as_character(v) -> str | None:
    """近似 R as.character(): 整数值的浮点数不带 .0"""
    if _is_na(v):
        return None
    if isinstance(v, bool) or isinstance(v, np.bool_):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (float, np.floating)):
        f = float(v)
        if f.is_integer() and abs(f) < 1e15:
            return str(int(f))
        return repr(f)
    if isinstance(v, pd.Timestamp):
        s = v.strftime("%Y-%m-%d %H:%M:%S")
        return s[:-9] if s.endswith(" 00:00:00") else s
    return str(v)


def _uniq(s: pd.Series):
    """返回 (codes, uniques); NA 的 code 为 -1"""
    codes, uniques = pd.factorize(s, use_na_sentinel=True)
    return codes, list(uniques)


def _map_unique(s: pd.Series, fn) -> pd.Series:
    """对去重值计算 fn 再映射回去(NA 也作为一个键)"""
    codes, uniques = _uniq(s)
    vals = [fn(_r_as_character(u)) for u in uniques] + [fn(None)]
    arr = np.empty(len(vals), dtype=object)
    arr[:] = vals
    return pd.Series(arr[codes], index=s.index, dtype="object")


# ==============================================================================
# 1. 日期
# ==============================================================================
_DATE_RE = re.compile(
    r"^\s*(\d{2,4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{1,2})(?::(\d{1,2}(?:\.\d+)?))?)?\s*$")


def _parse_date_str(x: str | None):
    if x is None:
        return pd.NaT
    x = x.strip()
    if x in _NA_TOKENS:
        return pd.NaT
    x = re.sub(r"年|/|\.", "-", x)
    x = x.replace("月", "-").replace("日", "")
    m = _DATE_RE.match(x)
    if not m:
        return pd.NaT
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if len(m.group(1)) == 2:  # lubridate 'y': 两位年份, 00-68 -> 20xx
        y += 2000 if y <= 68 else 1900
    try:
        return pd.Timestamp(year=y, month=mo, day=d)
    except (ValueError, OverflowError):
        return pd.NaT


def as_date_safe(s: pd.Series) -> pd.Series:
    """统一转成日期(datetime64, 时间部分为 0)。支持 日期/日期时间/Excel 序列号/各种字符串写法"""
    if isinstance(s.dtype, pd.DatetimeTZDtype):
        return s.dt.tz_convert("UTC").dt.tz_localize(None).dt.normalize()
    if pd.api.types.is_datetime64_any_dtype(s):
        return s.dt.normalize()
    if pd.api.types.is_bool_dtype(s) and s.isna().all():
        return pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")
    if pd.api.types.is_numeric_dtype(s) and not pd.api.types.is_bool_dtype(s):
        out = pd.to_datetime("1899-12-30") + pd.to_timedelta(s.astype("float64").floor(), unit="D")
        return out
    codes, uniques = _uniq(s)
    vals = []
    for v in uniques:
        if isinstance(v, pd.Timestamp):
            vals.append(v.tz_localize(None).normalize() if v.tzinfo else v.normalize())
        elif hasattr(v, "year") and hasattr(v, "month") and not isinstance(v, str):
            vals.append(pd.Timestamp(v).normalize())
        elif isinstance(v, (int, float)) and not isinstance(v, bool) and not _is_na(v):
            vals.append(pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(v)))
        else:
            vals.append(_parse_date_str(None if _is_na(v) else str(v)))
    u = pd.to_datetime(pd.Series(vals + [pd.NaT], dtype="object")).to_numpy(dtype="datetime64[ns]")
    return pd.Series(u[codes], index=s.index)


# ==============================================================================
# 2. 品种品系字符串标准化
# ==============================================================================
def normalize_breed_scalar(x: str | None) -> str | None:
    if x is None:
        return None
    x = re.sub(r"\s+", "", x.strip().upper())
    return None if x in _NA_TOKENS else x


# ==============================================================================
# 2.1 品系编码解析
# ==============================================================================
def _digits_valid(d: str | None) -> bool:
    if d is None:
        return False
    r = active_rules()
    if d in r.synthetic:
        return True
    nz = [c for c in d if c != "0"]
    return len(nz) > 0 and all(c in r.single for c in nz)


def resolve_line_code(x: str | None) -> str | None:
    if x is None:
        return None
    m = re.search(r"[0-9]+$", x)
    digits = m.group(0) if m else None
    if _digits_valid(digits):
        return digits
    for nm, code in active_rules().text:
        if nm in x:
            return code
    return None


def decode_line_lineage(code: str | None) -> dict | None:
    """返回 {品系: 占比}; 无法识别返回 None"""
    if code is None or code == "":
        return None
    r = active_rules()
    if code in r.synthetic:
        return {r.synthetic[code]: 1.0}
    digits = [c for c in code if c != "0"]
    if not digits or not all(c in r.single for c in digits):
        return None
    comp = {r.single[digits[-1]]: 1.0}
    for c in reversed(digits[:-1]):
        new = r.single[c]
        comp = {k: v * 0.5 for k, v in comp.items()}
        comp[new] = comp.get(new, 0.0) + 0.5
    return comp


def _digit_set(code: str | None) -> list[str]:
    if code is None or code == "" or code in active_rules().synthetic:
        return []
    out = []
    for c in code:
        if c != "0" and c not in out:
            out.append(c)
    return out


def composite_line_label(code: str | None) -> str | None:
    """含两种及以上数字的编码, 按品系组合表判定 (内置: 31->托丹, 38->法托, 含8->含法, 含3->含托)"""
    ds = set(_digit_set(code))
    if len(ds) < 2:
        return None
    for mode, digits, label in active_rules().combos:
        if (mode == "数字恰好是" and ds == digits) or (mode == "包含数字" and digits <= ds):
            return label
    return None


# ==============================================================================
# 2.2 安全数值/整数转换
# ==============================================================================
_NUM_RE = re.compile(r"^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$")
_HEX_RE = re.compile(r"^[+-]?0[xX][0-9a-fA-F]+$")
_NUM_NA = {"", "--", "-", "NA", "N/A", "NULL", "无", "未知"}


def _to_number(t: str | None) -> float:
    if t is None:
        return np.nan
    t = t.strip().replace(",", "")
    t = re.sub(r"[，、]", "", t)
    t = re.sub(r"\s+", "", t)
    t = re.sub(r"[－—]", "-", t)
    if t in _NUM_NA:
        return np.nan
    if _NUM_RE.match(t):
        return float(t)
    if _HEX_RE.match(t):
        return float(int(t, 16))
    low = t.lower().lstrip("+-")
    if low in ("inf", "infinity"):
        return float("-inf") if t.startswith("-") else float("inf")
    return np.nan


def as_numeric_safe(s: pd.Series, verbose: bool = False, field_name: str = "",
                    log=print, collect: list | None = None) -> pd.Series:
    """collect 不为 None 时, 把无法识别的情况追加为 (列名, 个数, 示例列表), 不单独输出"""
    if pd.api.types.is_numeric_dtype(s) and not pd.api.types.is_bool_dtype(s):
        return s.astype("float64")
    codes, uniques = _uniq(s)
    u_txt = [_r_as_character(u) for u in uniques]
    u_num = np.array([_to_number(t) for t in u_txt] + [np.nan], dtype="float64")
    out = pd.Series(u_num[codes], index=s.index, dtype="float64")
    if verbose or collect is not None:
        counts = np.bincount(codes[codes >= 0], minlength=len(uniques)) if len(uniques) else []
        bad = [(t, int(c)) for t, c, r in zip(u_txt, counts, u_num[:-1])
               if t not in (None, "") and np.isnan(r)]
        if bad:
            n_bad = sum(c for _, c in bad)
            examples = [t for t, _ in sorted(bad, key=lambda x: -x[1])[:3]]
            if collect is not None:
                collect.append((field_name, n_bad, examples))
            else:
                as_reporter(log).warn(
                    f"{{.field {esc(field_name)}}}: {fmt_n(n_bad)} 个值无法识别为数值, 已转为 NA",
                    ["示例: " + "、".join(f"{{.val {esc(x)}}}" for x in examples)])
    return out


def as_integer_safe(s: pd.Series, verbose: bool = False, field_name: str = "",
                    log=print, collect: list | None = None) -> pd.Series:
    num = as_numeric_safe(s, verbose=verbose, field_name=field_name, log=log, collect=collect)
    tr = np.trunc(num.to_numpy())
    ok = np.isfinite(tr) & (np.abs(tr) <= 2147483647)
    arr = pd.array(np.where(ok, tr, 0).astype("int64"), dtype="Int32")
    arr[~ok] = pd.NA
    return pd.Series(arr, index=s.index)


# ==============================================================================
# 2.3 通用品种 / 品系判断 (标量版, 输入为 normalize 之后的字符串)
# ==============================================================================
def _breed_label(x: str | None, fallback: str | None = None) -> str:
    """品种规则按顺序匹配, 第一个命中的生效 (内置规则见 rules.DEFAULT_RULES)"""
    r = active_rules()
    fallback = r.breed_fallback if fallback is None else fallback
    if x is None:
        return fallback
    for match, label in r._breed:
        if match(x):
            return label
    return fallback


def _line_label(x: str | None, fallback: str | None = None) -> str:
    """判定顺序: 编码等于 -> 单一品系编码(如 3 / 03) -> 组合规则 -> 兜底"""
    r = active_rules()
    fallback = r.line_fallback if fallback is None else fallback
    if x is None:
        return fallback
    code = resolve_line_code(x)
    if code is None:
        return fallback
    if code in r.exact:
        return r.exact[code]
    if code in r.single:                                   # "3"
        return r.single[code]
    if len(code) == 2 and code[0] == "0" and code[1] in r.single:   # "03" 补零写法
        return r.single[code[1]]
    comp = composite_line_label(code)
    return comp if comp is not None else fallback


def classify_breed_label(s: pd.Series, fallback: str | None = None) -> pd.Series:
    """fallback=None 时使用规则中的「都不匹配时」(内置为 其它)"""
    return _map_unique(s, lambda v: _breed_label(normalize_breed_scalar(v), fallback))


def classify_line_label(s: pd.Series, fallback: str | None = None) -> pd.Series:
    return _map_unique(s, lambda v: _line_label(normalize_breed_scalar(v), fallback))


# ==============================================================================
# 3. 后裔
# ==============================================================================
def _offspring_type(x: str | None) -> str:
    if x is None:
        return "error"
    if re.search(r"^S(Y|X)|^P(Y|X)|三元杂|三元", x):
        return "商品"
    if re.search(r"^YX|二元", x):
        return "扩繁"
    if re.search(r"^S\d+$|^X\d+$|^Y\d+$|^P\d+$|杜洛克|长白|大白|长白猪|大白猪|皮特兰", x):
        return "纯繁"
    return "error"


def classify_offspring(s: pd.Series) -> pd.DataFrame:
    return pd.DataFrame({
        "后裔性质": _map_unique(s, lambda v: _offspring_type(normalize_breed_scalar(v))),
        "后裔品种": classify_breed_label(s, fallback="error"),
        "后裔品系": classify_line_label(s),
    }, index=s.index)


# ==============================================================================
# 4. 公猪
# ==============================================================================
def _boar_breed(x: str | None) -> str:
    if x is None:
        return "其它"
    if re.search(r"杜洛克|^S\d*$", x):
        return "杜洛克"
    if re.search(r"长白|^X\d*$", x):
        return "长白"
    if re.search(r"大白|^Y\d*$", x):
        return "大白"
    if re.search(r"皮特兰|^P\d*$", x):
        return "皮特兰"
    return "其它"


def classify_boar(s: pd.Series) -> pd.DataFrame:
    return pd.DataFrame({
        "公猪品种": _map_unique(s, lambda v: _boar_breed(normalize_breed_scalar(v))),
        "公猪品系": classify_line_label(s),
    }, index=s.index)


# ==============================================================================
# 5. 公猪血统特征
# ==============================================================================
_ANCESTRY_COLS = ["与配公猪品种品系_含托", "与配公猪品种品系_含法",
                  "与配公猪品种品系_托法", "与配公猪品种品系_托法皮"]


def _boar_ancestry(x: str | None) -> tuple:
    if x is None:
        return ("其它",) * 4
    code = resolve_line_code(x)
    comp = decode_line_lineage(code)
    breed = _boar_breed(x)

    # 托佩克 / 法系 按编码 3 / 8 在当前规则中的名称判断(内置即 托佩克 / 法系)
    r = active_rules()
    to_name, fa_name = r.line_name("3", "托佩克"), r.line_name("8", "法系")
    has_to = bool(comp and comp.get(to_name, 0) > 0) or bool(re.search(r"托佩克|托丹|托美|托华", x))
    has_fa = bool(comp and comp.get(fa_name, 0) > 0) or ("法系" in x)
    pure_to = bool(comp) and len(comp) == 1 and next(iter(comp)) == to_name
    pure_fa = bool(comp) and len(comp) == 1 and next(iter(comp)) == fa_name
    mix_to = has_to and not pure_to
    white = breed in ("长白", "大白")

    c1 = "含托佩克" if has_to else "其它"
    c2 = "法系" if has_fa else "其它"
    if breed == "皮特兰":
        c3 = "皮特兰"
    elif mix_to:
        c3 = "含托佩克"
    elif pure_to:
        c3 = "纯托佩克"
    elif pure_fa:
        c3 = "纯法系"
    else:
        c3 = "其它"
    if breed == "皮特兰":
        c4 = "法系皮特兰"
    elif mix_to:
        c4 = "含托佩克"
    elif pure_fa and breed == "杜洛克":
        c4 = "法系杜洛克"
    elif pure_to:
        c4 = "纯托佩克"
    elif pure_fa and white:
        c4 = "纯法系白猪"
    else:
        c4 = "其它"
    return (c1, c2, c3, c4)


def classify_boar_ancestry(s: pd.Series) -> pd.DataFrame:
    tup = _map_unique(s, lambda v: _boar_ancestry(normalize_breed_scalar(v)))
    arr = list(tup)
    return pd.DataFrame({c: [t[i] for t in arr] for i, c in enumerate(_ANCESTRY_COLS)},
                        index=s.index)


# ==============================================================================
# 5.5 母猪 / 6. 后裔品种总分类
# ==============================================================================
def classify_sow(breed: pd.Series, line: pd.Series) -> pd.DataFrame:
    return pd.DataFrame({"母猪品种": classify_breed_label(breed),
                         "母猪品系": classify_line_label(line)}, index=breed.index)


def _offspring_total(x: str) -> str:
    if x in ("大白", "长白"):
        return "白猪"
    if x == "二元":
        return "二元"
    if x == "杜洛克":
        return "杜洛克"
    if x is not None and "商品猪" in x:
        return "商品猪"
    return "其它"


# ==============================================================================
# 7. 日期衍生变量
# ==============================================================================
def _season(m):
    if pd.isna(m):
        return None
    m = int(m)
    if m in (3, 4, 5):
        return "春季"
    if m in (6, 7, 8):
        return "夏季"
    if m in (9, 10, 11):
        return "秋季"
    return "冬季"


def _ym(d: pd.Series) -> pd.Series:
    return d.dt.strftime("%Y-%m").astype("object").where(d.notna(), None)


def _int_part(x: pd.Series) -> pd.Series:
    return x.astype("Int32")


def create_date_variables(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for pre in ("配种", "分娩", "断奶"):
        col = f"{pre}日期"
        df[col] = as_date_safe(df[col])
        df[f"{pre}年月"] = _ym(df[col])
        df[f"{pre}年"] = _int_part(df[col].dt.year)
        df[f"{pre}月"] = _int_part(df[col].dt.month)
        if pre == "配种":
            df["配种季节"] = pd.Categorical(df["配种月"].map(_season, na_action="ignore"),
                                         categories=["春季", "夏季", "秋季", "冬季"])
    return df


# ==============================================================================
# 8. 主函数
# ==============================================================================
SELECT_MAP = [
    ("猪场", "猪场"), ("国标编码", "国标编码"), ("品种", "品种"), ("品系", "品系"),
    ("出生日期", "出生日期"),
    ("后备来源场", "后备信息_后备来源场"), ("后备批次", "后备信息_后备猪批次"),
    ("首配日期", "后备信息_首配日期"), ("首配日龄", "后备信息_首配日龄"),
    ("首配体重", "后备信息_首配体重"),
    ("配种日期", "配种信息_配种日期"), ("与配公猪国标编码", "配种信息_与配公猪国标编码"),
    ("与配公猪品种品系", "配种信息_与配公猪品种品系"), ("后裔品种品系", "配种信息_后裔品种品系"),
    ("配种背膘", "配种信息_配种背膘"), ("配种背膘测定结果", "配种信息_配种背膘测定结果"),
    ("首配背膘", "配种信息_首配背膘"), ("首配背膘测定结果", "配种信息_首配背膘测定结果"),
    ("配种结果日期", "配种结果_配种结果日期"), ("配种结果", "配种结果"),
    ("配种结果天数", "配种结果_结果天数"),
    ("分娩日期", "分娩信息_分娩日期"), ("分娩胎次", "分娩信息_分娩胎次"),
    ("分娩窝重", "分娩信息_分娩窝重（kg）"), ("总产仔数", "分娩信息_总产仔数"),
    ("活仔", "分娩信息_活仔"), ("健仔", "分娩信息_健仔"), ("死胎", "分娩信息_死胎"),
    ("木乃伊", "分娩信息_木乃伊"),
    ("断奶日期", "断奶信息_断奶日期"), ("断配间距", "断奶信息_断配间距"),
    ("断奶头数", "断奶信息_断奶头数"), ("断奶总重", "断奶信息_断奶总重"),
    ("离场日期", "离场信息_离场日期"), ("离场所在猪场", "离场信息_离场所在猪场"),
    ("离场类型", "离场信息_离场类型"), ("离场原因", "离场信息_离场原因"),
    ("转入场", "离场信息_转入场"),
]

INT_COLS = ["总产仔数", "活仔", "健仔", "死胎", "木乃伊", "断配间距", "断奶头数",
            "分娩胎次", "首配日龄", "配种结果天数"]

FARM_TYPE = {
    **{f: "核心育种场" for f in ("金沙一场", "龙马一场", "香江场")},
    "郭祥一场": "扩繁一胎场",
    "郭祥三场": "扩繁经产场",
    **{f: "商品一胎场" for f in ("郭祥二场", "大庆一场", "德邦场1线", "德邦场2线")},
    **{f: "商品经产场" for f in ("凤凰一场", "凤凰二场", "大庆二场", "雄桂场", "信宜场")},
}

DATA_TYPES = ("all", "配种", "分娩", "断奶")


def _is_text_col(s: pd.Series) -> bool:
    return s.dtype == object or pd.api.types.is_string_dtype(s)


def _date_to_str(s: pd.Series) -> pd.Series:
    """对应 R: format(as.Date(x), "%Y-%m-%d")"""
    d = as_date_safe(s)
    return d.dt.strftime("%Y-%m-%d").astype("object").where(d.notna(), None)


def clean_production_data(raw: pd.DataFrame, data_type: str = "all",
                          drop_missing_date: bool = True, verbose: bool = False,
                          log=print) -> pd.DataFrame:
    rep = as_reporter(log if verbose else None)
    if data_type not in DATA_TYPES:
        raise ValueError(f"data_type 只能是 {DATA_TYPES}")

    # ---------------- Step 1. 字段选择 + 基础清洗 ----------------
    missing = [src for _, src in SELECT_MAP if src not in raw.columns]
    if missing:
        raise ValueError(f"原始数据缺少 {len(missing)} 个必需列, 请检查文件是否为综合选择查询导出, "
                         f"或手动设置跳过行数 / 标题行数:\n"
                         + "\n".join(f"  - {m}" for m in missing[:6])
                         + (f"\n  - … 还有 {len(missing) - 6} 列" if len(missing) > 6 else ""))
    df = raw[[src for _, src in SELECT_MAP]].copy()
    df.columns = [dst for dst, _ in SELECT_MAP]

    n0 = len(df)
    df = df.drop_duplicates(keep="first").reset_index(drop=True)
    if n0 != len(df):
        rep.info(f"去重: 删除 {fmt_n(n0 - len(df))} 条完全重复的记录")

    n1 = len(df)
    df = df[df["猪场"].notna()].reset_index(drop=True)
    if n1 != len(df):
        rep.info(f"删除 {{.field 猪场}} 为空的记录 {fmt_n(n1 - len(df))} 条")

    for c in df.columns:
        if _is_text_col(df[c]):
            df[c] = df[c].where(df[c] != "--", None)

    bad_num: list = []
    for c in INT_COLS:
        df[c] = as_integer_safe(df[c], field_name=c, collect=bad_num)

    for c in ("分娩日期", "配种日期", "断奶日期", "出生日期", "首配日期"):
        df[c] = _date_to_str(df[c])

    for c in ("配种背膘", "首配背膘", "断奶总重"):
        df[c] = as_numeric_safe(df[c], field_name=c, collect=bad_num)
    if bad_num:
        rep.warn(f"{len(bad_num)} 列中有无法识别为数值的内容, 已转为 NA:", [
            f"{{.field {esc(c)}}}: {fmt_n(n)} 个, 如 "
            + "、".join(f"{{.val {esc(x)}}}" for x in ex)
            for c, n, ex in bad_num])

    farm_txt = df["猪场"].map(_r_as_character)
    df["猪场类型"] = pd.Categorical(farm_txt.map(lambda f: FARM_TYPE.get(f, "其他")),
                                categories=FARM_TYPE_LEVELS)
    unknown = sorted({f for f in farm_txt.dropna().unique() if f not in FARM_ORDER})
    if unknown:
        n_rows = int(farm_txt.isin(unknown).sum())
        rep.warn(f"{len(unknown)} 个猪场不在 {{.field farm_order}} 中, "
                 f"这 {fmt_n(n_rows)} 条记录的 {{.field 猪场}} 将为 NA (与 R 的 factor() 一致):",
                 [f"{{.val {esc(f)}}}: {fmt_n(int((farm_txt == f).sum()))} 条" for f in unknown])
    df["猪场"] = pd.Categorical(farm_txt, categories=FARM_ORDER)

    # ---------------- Step 2-6. 分类 ----------------
    df = pd.concat([
        df,
        classify_sow(df["品种"], df["品系"]),
        classify_offspring(df["后裔品种品系"]),
        classify_boar(df["与配公猪品种品系"]),
        classify_boar_ancestry(df["与配公猪品种品系"]),
    ], axis=1)
    df["后裔品种_总"] = df["后裔品种"].map(_offspring_total).astype("object")

    # ---------------- Step 7. 日期衍生变量 ----------------
    df = create_date_variables(df)

    # ---------------- Step 8. 按数据类型过滤 ----------------
    if drop_missing_date and data_type != "all":
        key = f"{data_type}日期"
        n2 = len(df)
        df = df[df[key].notna()].reset_index(drop=True)
        rep.info(f"按 {{.field {key}}} 筛选: 保留 {fmt_n(len(df))} 条, "
                 f"删除日期缺失 {fmt_n(n2 - len(df))} 条")

    return df
