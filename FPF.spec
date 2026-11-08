# -*- mode: python ; coding: utf-8 -*-
# PyInstaller 打包配置:  pyinstaller FPF.spec --noconfirm
from PyInstaller.utils.hooks import collect_dynamic_libs

EXCLUDES = [
    # 用不到的大库
    "matplotlib", "IPython", "jupyter", "notebook", "scipy", "sklearn", "numba",
    "sqlalchemy", "PyQt5", "PySide2", "PySide6", "PyQt6", "sphinx", "pytest",
    "openpyxl", "xlrd", "xlsxwriter", "odf", "pyxlsb", "lxml", "bs4", "html5lib",
    "jinja2", "markupsafe", "tables", "h5py", "fsspec", "s3fs", "gcsfs", "botocore",
    "boto3", "psycopg2", "pymysql", "xarray", "numexpr", "bottleneck", "zstandard",
    "rdata", "pyreadr", "PIL",
    # 测试 / 文档 / 开发工具
    "pandas.tests", "numpy.tests", "numpy.f2py", "numpy.distutils", "numpy.testing",
    "tkinter.test", "lib2to3",
    "distutils", "setuptools", "pkg_resources", "email.test",
    "unittest", "pydoc", "doctest", "pdb",
    # 不使用 pyarrow; 即使打包环境里装了也不打进去
    "pyarrow",
]

a = Analysis(
    ["main.py"],
    pathex=["."],
    binaries=collect_dynamic_libs("python_calamine"),
    datas=[],
    hiddenimports=["python_calamine"],
    excludes=EXCLUDES,
    noarchive=False,
    optimize=1,         # 去掉 assert, 略小略快
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="FPF数据处理",
    debug=False,
    strip=False,
    upx=False,          # 不用 UPX: 杀毒软件误报更少, 启动更快
    console=False,      # 图形界面, 不弹黑框
    icon=None,
)
