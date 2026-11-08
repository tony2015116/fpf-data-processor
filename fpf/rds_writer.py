# -*- coding: utf-8 -*-
"""
极简 RDS 写出器 —— 把 pandas DataFrame 写成 R 原生 .rds(序列化格式 v3, XDR)

R 端直接 readRDS() 即可, 不需要安装任何包。支持的列类型:
  布尔/可空布尔 -> logical      Int/可空Int   -> integer
  float          -> double       文本          -> character (UTF-8)
  Categorical    -> factor        日期(零点)   -> Date
  带时区时间     -> POSIXct(保留时区, 无时区按 UTC)
"""
from __future__ import annotations

import gzip
import struct

import numpy as np
import pandas as pd

_NILVALUE = 254
_LGLSXP, _INTSXP, _REALSXP, _STRSXP, _VECSXP, _LISTSXP, _SYMSXP, _CHARSXP = 10, 13, 14, 16, 19, 2, 1, 9
_IS_OBJ, _HAS_ATTR, _HAS_TAG = 1 << 8, 1 << 9, 1 << 10
_UTF8 = 1 << 3
_NA_INT = -2147483648
_NA_REAL = struct.unpack(">d", bytes.fromhex("7FF00000000007A2"))[0]


_NA_CHARSXP = struct.pack(">ii", _CHARSXP, -1)
_PACK2 = struct.Struct(">ii").pack


def _charsxp_bytes(s) -> bytes:
    """一个字符串的完整 CHARSXP 字节(标记 + 长度 + UTF-8 内容)"""
    if s is None:
        return _NA_CHARSXP
    b = s.encode("utf-8")
    enc = 64 if b.isascii() else _UTF8          # 64 = ASCII 标记
    return _PACK2(_CHARSXP | (enc << 12), len(b)) + b


class _W:
    """把序列化结果按块写到文件对象, 不在内存里拼整个文件"""

    def __init__(self, fh):
        self.fh = fh

    def i(self, v: int):
        self.fh.write(struct.pack(">i", v))

    def raw(self, b: bytes):
        self.fh.write(b)

    def charsxp(self, s):
        self.fh.write(_charsxp_bytes(s))

    def strsxp(self, values, attrs=None):
        self.i(_STRSXP | (_HAS_ATTR if attrs else 0))
        self.i(len(values))
        self.fh.write(b"".join(_charsxp_bytes(v) for v in values))
        if attrs:
            self.attrs(attrs)

    def strsxp_series(self, s: pd.Series):
        """字符列: 每个不重复的值只编码一次, 再按位置拼接(重复值多的列快很多)"""
        codes, uniques = pd.factorize(s, use_na_sentinel=True)
        enc = np.empty(len(uniques) + 1, dtype=object)
        enc[:-1] = [_charsxp_bytes(str(u)) for u in uniques]
        enc[-1] = _NA_CHARSXP                       # code = -1 -> NA
        self.i(_STRSXP)
        self.i(len(codes))
        self.fh.write(b"".join(enc[codes].tolist()))

    def attrs(self, attrs: list[tuple[str, callable]]):
        for tag, writer in attrs:
            self.i(_LISTSXP | _HAS_TAG)
            self.i(_SYMSXP)
            self.charsxp(tag)
            writer()
        self.i(_NILVALUE)

    def vec(self, sxptype: int, arr: np.ndarray, attrs=None, is_obj=False):
        flags = sxptype | (_HAS_ATTR if attrs else 0) | (_IS_OBJ if is_obj else 0)
        self.i(flags)
        self.i(len(arr))
        self.raw(arr.tobytes())
        if attrs:
            self.attrs(attrs)


def _write_col(w: _W, s: pd.Series):
    dt = s.dtype
    if isinstance(dt, pd.CategoricalDtype):
        codes = s.cat.codes.to_numpy().astype(np.int64) + 1
        codes = np.where(codes == 0, _NA_INT, codes).astype(">i4")
        levels = [str(c) for c in s.cat.categories]
        w.vec(_INTSXP, codes, attrs=[
            ("levels", lambda: w.strsxp(levels)),
            ("class", lambda: w.strsxp(["factor"])),
        ], is_obj=True)
    elif pd.api.types.is_bool_dtype(dt):
        a = s.astype("boolean")
        arr = np.where(a.isna().to_numpy(), _NA_INT, a.fillna(False).to_numpy().astype(np.int64)).astype(">i4")
        w.vec(_LGLSXP, arr)
    elif pd.api.types.is_integer_dtype(dt):
        a = s.astype("Int64")
        arr = np.where(a.isna().to_numpy(), _NA_INT, a.fillna(0).to_numpy().astype(np.int64)).astype(">i4")
        w.vec(_INTSXP, arr)
    elif pd.api.types.is_float_dtype(dt):
        a = s.astype("float64").to_numpy(na_value=np.nan)
        arr = np.where(np.isnan(a), _NA_REAL, a).astype(">f8")
        w.vec(_REALSXP, arr)
    elif isinstance(dt, pd.DatetimeTZDtype) or pd.api.types.is_datetime64_any_dtype(dt):
        tz = None
        if isinstance(dt, pd.DatetimeTZDtype):
            tz = str(dt.tz)
            v = s.dt.tz_convert("UTC").dt.tz_localize(None)
        else:
            v = s
        ns = v.to_numpy(dtype="datetime64[ns]").astype(np.int64).astype(np.float64)
        na = v.isna().to_numpy()
        secs = ns / 1e9
        all_midnight = tz is None and bool(np.all((v.dropna().dt.normalize() == v.dropna())))
        if all_midnight:
            arr = np.where(na, _NA_REAL, np.floor(secs / 86400.0)).astype(">f8")
            w.vec(_REALSXP, arr, attrs=[("class", lambda: w.strsxp(["Date"]))], is_obj=True)
        else:
            arr = np.where(na, _NA_REAL, secs).astype(">f8")
            tzname = tz or "UTC"
            w.vec(_REALSXP, arr, attrs=[
                ("class", lambda: w.strsxp(["POSIXct", "POSIXt"])),
                ("tzone", lambda: w.strsxp([tzname])),
            ], is_obj=True)
    else:
        w.strsxp_series(s)


def write_rds(df: pd.DataFrame, path: str, compress: bool = True, level: int = 1) -> None:
    """compress=True: gzip 压缩(与 R saveRDS 默认相同, 文件小);
    compress=False: 不压缩(写入更快, 文件约大 4-8 倍)。两种 R 都直接 readRDS()。"""
    if compress:
        with gzip.open(path, "wb", compresslevel=level) as fh:
            _serialize(df, _W(fh))
    else:
        with open(path, "wb", buffering=1 << 20) as fh:
            _serialize(df, _W(fh))


def _serialize(df: pd.DataFrame, w: _W) -> None:
    w.raw(b"X\n")
    w.i(3)
    w.i(0x00040300)          # 写出端 R 版本 4.3.0
    w.i(0x00030500)          # 最低读取版本 3.5.0
    enc = b"UTF-8"
    w.i(len(enc))
    w.raw(enc)

    names = [str(c) for c in df.columns]
    w.i(_VECSXP | _HAS_ATTR | _IS_OBJ)
    w.i(df.shape[1])
    for c in df.columns:
        _write_col(w, df[c])
    n = len(df)
    w.attrs([
        ("names", lambda: w.strsxp(names)),
        ("class", lambda: w.strsxp(["data.frame"])),
        ("row.names", lambda: w.vec(_INTSXP, np.array([_NA_INT, -n], dtype=">i4"))),
    ])
