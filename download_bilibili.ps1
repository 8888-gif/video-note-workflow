param(
    [string]$UrlFile = ".\urls.txt",
    [string]$OutputDir = ".\downloads",
    [switch]$UseCookies,
    [string]$CookiesFile = ".\cookies.txt",
    [switch]$HighQuality
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$OutputEncoding = [System.Text.UTF8Encoding]::new()

function Find-FFmpeg {
    $command = Get-Command ffmpeg -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    $wingetPath = Get-ChildItem -Path "$env:LOCALAPPDATA\Microsoft\WinGet\Packages" -Recurse -Filter ffmpeg.exe -ErrorAction SilentlyContinue |
        Select-Object -First 1 -ExpandProperty FullName
    if ($wingetPath) {
        return $wingetPath
    }

    return $null
}

if (-not (Test-Path -LiteralPath $UrlFile)) {
    Write-Host "未找到链接文件：$UrlFile"
    Write-Host "请把 B 站视频链接逐行写入 urls.txt 后再运行。"
    exit 1
}

$urls = Get-Content -LiteralPath $UrlFile |
    ForEach-Object { $_.Trim() } |
    Where-Object { $_ -and -not $_.StartsWith("#") }

if ($urls.Count -eq 0) {
    Write-Host "urls.txt 中还没有可下载链接。"
    Write-Host "请把 B 站视频链接一行一个写入 urls.txt 后再运行。"
    exit 0
}

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$ffmpegPath = Find-FFmpeg

Write-Host "正在检查 yt-dlp..."
python -m yt_dlp --version > $null 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "未检测到 yt-dlp，正在通过 pip 安装到当前 Python 环境..."
    python -m pip install -U yt-dlp
    if ($LASTEXITCODE -ne 0) {
        Write-Host "默认 pip 源安装失败，改用官方 PyPI 重试..."
        python -m pip install -U yt-dlp --index-url https://pypi.org/simple
    }
}

$downloadArgs = @(
    "-m", "yt_dlp",
    "--ignore-errors",
    "--no-overwrites",
    "--continue",
    "--retries", "10",
    "--fragment-retries", "10",
    "--sleep-interval", "1",
    "--max-sleep-interval", "3",
    "--download-archive", (Join-Path $OutputDir "downloaded.txt"),
    "-a", $UrlFile,
    "-o", (Join-Path $OutputDir "%(uploader)s/%(title)s [%(id)s].%(ext)s"),
    "--merge-output-format", "mp4"
)

if ($ffmpegPath) {
    Write-Host "已找到 FFmpeg，将自动合并音频和视频：$ffmpegPath"
    $downloadArgs += @("--ffmpeg-location", (Split-Path -Parent $ffmpegPath))
} else {
    Write-Host "未找到 FFmpeg。高清视频可能会下载成音频和视频分离文件。"
    Write-Host "建议安装 FFmpeg，或使用工具中的“合并已下载音视频”功能。"
}

if ($HighQuality) {
    $downloadArgs += @("-f", "bv*+ba/best")
} else {
    $downloadArgs += @("-f", "best[ext=mp4]/best")
}

if ($UseCookies -and -not (Test-Path -LiteralPath $CookiesFile)) {
    Write-Host "已启用 -UseCookies，但未找到 cookies 文件：$CookiesFile"
    Write-Host "请导出 cookies.txt 后放到本文件夹，或取消 -UseCookies。"
    exit 1
}

if ($UseCookies) {
    $downloadArgs += @("--cookies", $CookiesFile)
}

Write-Host "开始批量下载..."
python @downloadArgs

Write-Host ""
Write-Host "下载结束。文件目录：$OutputDir"
Write-Host "已下载记录：$(Join-Path $OutputDir 'downloaded.txt')"
