param(
    [string]$RootDir = ".\downloads",
    [switch]$KeepSourceFiles
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

$ffmpegPath = Find-FFmpeg
if (-not $ffmpegPath) {
    Write-Host "未找到 FFmpeg，无法合并音视频。"
    exit 1
}

if (-not (Test-Path -LiteralPath $RootDir)) {
    Write-Host "未找到下载目录：$RootDir"
    exit 1
}

$videoFiles = Get-ChildItem -LiteralPath $RootDir -Recurse -File -Filter "*.mp4" |
    Where-Object { $_.BaseName -match "\.f\d+$" }

$mergedCount = 0
foreach ($video in $videoFiles) {
    $prefix = $video.BaseName -replace "\.f\d+$", ""
    $audio = Get-ChildItem -LiteralPath $video.DirectoryName -File |
        Where-Object { $_.BaseName.StartsWith("$prefix.f") -and $_.Extension -in @(".m4a", ".aac", ".mp3", ".webm") } |
        Select-Object -First 1

    if (-not $audio) {
        continue
    }

    $output = Join-Path $video.DirectoryName "$prefix.mp4"
    if (Test-Path -LiteralPath $output) {
        Write-Host "已存在合并文件，跳过：$output"
        continue
    }

    Write-Host "正在合并：$($video.Name) + $($audio.Name)"
    & $ffmpegPath -hide_banner -loglevel error -y -i $video.FullName -i $audio.FullName -c copy $output

    if ($LASTEXITCODE -eq 0 -and (Test-Path -LiteralPath $output)) {
        $mergedCount += 1
        Write-Host "合并完成：$output"
        if (-not $KeepSourceFiles) {
            Remove-Item -LiteralPath $video.FullName -Force
            Remove-Item -LiteralPath $audio.FullName -Force
            Write-Host "已删除分离的源文件。"
        }
    } else {
        Write-Host "合并失败：$($video.FullName)"
    }
}

Write-Host ""
Write-Host "处理完成，共合并 $mergedCount 个视频。"
