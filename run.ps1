# 论文排版工具 - PowerShell 一键运行
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

Write-Host "=== 论文排版工具 - 一键运行 ===" -ForegroundColor Cyan

# 找 Python
$py = $null
if (Get-Command python -ErrorAction SilentlyContinue) { $py = "python" }
elseif (Get-Command py -ErrorAction SilentlyContinue) { $py = "py -3" }
else {
    Write-Host "[错误] 未找到 Python。请安装: https://www.python.org/downloads/" -ForegroundColor Red
    Write-Host "安装时勾选 Add Python to PATH"
    exit 1
}

# 创建 venv
if (-not (Test-Path ".venv\Scripts\python.exe")) {
    Write-Host "[1/4] 创建虚拟环境 .venv ..." -ForegroundColor Yellow
    Invoke-Expression "$py -m venv .venv"
}

$venvPython = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $venvPython)) {
    Write-Host "[错误] 虚拟环境创建失败，找不到 $venvPython" -ForegroundColor Red
    exit 1
}

Write-Host "[2/4] 安装依赖 ..." -ForegroundColor Yellow
& $venvPython -m pip install -r requirements.txt -q

if (-not (Test-Path "examples\sample_input.docx")) {
    Write-Host "[3/4] 生成测试文档 ..." -ForegroundColor Yellow
    & $venvPython scripts\create_sample.py
} else {
    Write-Host "[3/4] 测试文档已存在，跳过" -ForegroundColor Gray
}

Write-Host "[4/4] 运行排版 ..." -ForegroundColor Yellow
& $venvPython main.py examples\sample_input.docx

$out = Join-Path $PSScriptRoot "output\sample_input_formatted.docx"
Write-Host ""
Write-Host "=== 完成 ===" -ForegroundColor Green
Write-Host "输出文件: $out"
if (Test-Path $out) {
    Write-Host "文件大小: $((Get-Item $out).Length) 字节"
}
