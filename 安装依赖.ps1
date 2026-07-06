$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$OutputEncoding = [System.Text.UTF8Encoding]::new()

Write-Host "正在检查 Python..."
python --version

Write-Host "正在安装或更新 yt-dlp..."
python -m pip install -U yt-dlp
if ($LASTEXITCODE -ne 0) {
    Write-Host "默认 pip 源失败，改用官方 PyPI 重试..."
    python -m pip install -U yt-dlp --index-url https://pypi.org/simple
}

Write-Host "正在检查 FFmpeg..."
$ffmpeg = Get-Command ffmpeg -ErrorAction SilentlyContinue
if (-not $ffmpeg) {
    Write-Host "未检测到 FFmpeg，尝试通过 winget 安装..."
    winget install --id Gyan.FFmpeg -e --accept-source-agreements --accept-package-agreements
} else {
    Write-Host "已检测到 FFmpeg：$($ffmpeg.Source)"
}

if (-not (Test-Path -LiteralPath ".\urls.txt")) {
    Copy-Item -LiteralPath ".\urls.example.txt" -Destination ".\urls.txt"
    Write-Host "已创建 urls.txt。"
}

Write-Host ""
Write-Host "依赖检查完成。可以双击 启动B站批量下载器.cmd 使用工具。"

