@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo === 论文排版工具 - 一键运行 ===

where python >nul 2>&1
if errorlevel 1 (
    where py >nul 2>&1
    if errorlevel 1 (
        echo [错误] 未找到 Python。请先安装: https://www.python.org/downloads/
        echo 安装时勾选 "Add Python to PATH"
        pause
        exit /b 1
    )
    set PY=py -3
) else (
    set PY=python
)

if not exist ".venv\Scripts\python.exe" (
    echo [1/4] 正在创建虚拟环境 .venv ...
    %PY% -m venv .venv
    if errorlevel 1 (
        echo [错误] 创建虚拟环境失败
        pause
        exit /b 1
    )
)

echo [2/4] 正在安装依赖 ...
.venv\Scripts\python.exe -m pip install -r requirements.txt -q
if errorlevel 1 (
    echo [错误] 安装依赖失败
    pause
    exit /b 1
)

if not exist "examples\sample_input.docx" (
    echo [3/4] 正在生成测试文档 ...
    .venv\Scripts\python.exe scripts\create_sample.py
)

echo [4/4] 正在排版 ...
.venv\Scripts\python.exe main.py examples\sample_input.docx
if errorlevel 1 (
    echo [错误] 排版失败
    pause
    exit /b 1
)

echo.
echo === 完成 ===
echo 输出文件: %cd%\output\sample_input_formatted.docx
echo 用 Word 打开上述文件查看效果
echo.
pause
