@echo off
chcp 65001 >nul
REM ================================================================
REM  一键打包 FPF数据处理.exe
REM  前提: 已安装 64 位 Python 3.10 ~ 3.15 (python.org 官方安装包)
REM  结果: dist\FPF数据处理.exe  (单文件, 拷到任何 Windows 电脑都能直接运行)
REM  大小约 40 MB, 导出 rds / csv
REM ================================================================
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 goto :no_python

echo [检查] Python 环境:
python -c "import sys,struct,platform; print('       Python', sys.version.split()[0], '|', str(struct.calcsize('P')*8)+' 位', '|', platform.machine())"
if errorlevel 1 goto :no_python
python -c "import struct,sys; sys.exit(0 if struct.calcsize('P')==8 else 1)"
if errorlevel 1 goto :bit32
python -c "import sys; sys.exit(0 if (3,10) <= sys.version_info[:2] <= (3,15) else 1)"
if errorlevel 1 echo [提示] 建议使用 Python 3.10 ~ 3.15, 当前版本可能装不上依赖

if exist .venv\Scripts\python.exe goto :venv_ok
echo [1/3] 创建独立虚拟环境 .venv ...
python -m venv .venv
if errorlevel 1 goto :fail
:venv_ok
call .venv\Scripts\activate.bat

echo [2/3] 安装依赖 (依次尝试多个下载源) ...
call :try_install https://pypi.tuna.tsinghua.edu.cn/simple pypi.tuna.tsinghua.edu.cn && goto :installed
call :try_install https://mirrors.aliyun.com/pypi/simple mirrors.aliyun.com && goto :installed
call :try_install https://mirrors.cloud.tencent.com/pypi/simple mirrors.cloud.tencent.com && goto :installed
call :try_install https://pypi.org/simple pypi.org && goto :installed
echo.
echo [错误] 所有下载源都安装失败, 请检查网络或代理设置
goto :fail

:installed
echo [3/3] 打包中, 约 1-3 分钟 ...
pyinstaller FPF.spec --noconfirm --clean --log-level WARN
if errorlevel 1 goto :fail

echo.
echo 完成: %cd%\dist\FPF数据处理.exe
explorer dist
pause
exit /b 0

REM ---------------- 子程序: 用指定源安装依赖 ----------------
:try_install
echo       尝试 %~2 ...
python -m pip install -q --disable-pip-version-check --timeout 30 --upgrade pip -i %~1 --trusted-host %~2 >nul 2>nul
python -m pip install -q --disable-pip-version-check --timeout 30 -r requirements.txt pyinstaller -i %~1 --trusted-host %~2
set RC=%errorlevel%
if not "%RC%"=="0" echo       %~2 失败, 换下一个源
exit /b %RC%

:no_python
echo [错误] 找不到 python, 请从 python.org 安装 64 位 Python, 安装时勾选 "Add python.exe to PATH"
goto :fail

:bit32
echo [错误] 当前是 32 位 Python, pandas 新版本不再提供 32 位安装包
echo        请卸载后从 python.org 重新安装 64 位版本 (Windows installer 64-bit)
echo        装好后删除本文件夹下的 .venv 文件夹, 再双击本脚本
goto :fail

:fail
echo.
echo [错误] 打包失败, 请把上面的报错信息截图发出来
pause
exit /b 1
