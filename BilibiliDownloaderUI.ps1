Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = "Stop"
$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$urlFile = Join-Path $baseDir "urls.txt"
$scriptFile = Join-Path $baseDir "download_bilibili.ps1"
$mergeScriptFile = Join-Path $baseDir "合并已下载音视频.ps1"
$workflowDocFile = Join-Path $baseDir "视频转图文笔记_分享版流程.md"
$noteTemplateFile = Join-Path $baseDir "图文笔记模板.md"
$sharePptFile = Join-Path $baseDir "B站视频转图文笔记工作流_朋友圈分享版.pptx"
$defaultOutput = Join-Path $baseDir "downloads"
$cookiesFile = Join-Path $baseDir "cookies.txt"
$baiduNetdiskExe = "D:\baidu\BaiduNetdisk\BaiduNetdisk.exe"
$script:selectedUploadFiles = @()

if (-not (Test-Path -LiteralPath $urlFile)) {
    New-Item -ItemType File -Path $urlFile -Force | Out-Null
}

function Add-Log {
    param([string]$Message)
    $logBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')] $Message`r`n")
}

function Set-UploadFiles {
    param([string[]]$Files)

    $script:selectedUploadFiles = @($Files | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -Unique)
    if ($script:selectedUploadFiles.Count -eq 0) {
        $selectedFilesBox.Text = "尚未选择文件"
        return
    }

    $totalSize = ($script:selectedUploadFiles | ForEach-Object { (Get-Item -LiteralPath $_).Length } | Measure-Object -Sum).Sum
    $totalSizeGB = [math]::Round($totalSize / 1GB, 2)
    $selectedFilesBox.Text = "已选择 $($script:selectedUploadFiles.Count) 个文件，共 $totalSizeGB GB"
    [System.Windows.Forms.Clipboard]::SetText(($script:selectedUploadFiles -join "`r`n"))
    Add-Log "已选择 $($script:selectedUploadFiles.Count) 个文件，共 $totalSizeGB GB，文件路径已复制到剪贴板。"
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "B站批量下载器"
$form.Size = New-Object System.Drawing.Size(920, 860)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(820, 800)

$font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
$form.Font = $font

$labelUrls = New-Object System.Windows.Forms.Label
$labelUrls.Text = "B站视频链接（一行一个）"
$labelUrls.Location = New-Object System.Drawing.Point(16, 16)
$labelUrls.Size = New-Object System.Drawing.Size(280, 24)
$form.Controls.Add($labelUrls)

$urlBox = New-Object System.Windows.Forms.TextBox
$urlBox.Multiline = $true
$urlBox.ScrollBars = "Vertical"
$urlBox.AcceptsReturn = $true
$urlBox.AcceptsTab = $true
$urlBox.Location = New-Object System.Drawing.Point(16, 44)
$urlBox.Size = New-Object System.Drawing.Size(870, 260)
$urlBox.Anchor = "Top,Left,Right"
$urlBox.Text = (Get-Content -LiteralPath $urlFile -Raw -ErrorAction SilentlyContinue)
$form.Controls.Add($urlBox)

$labelOutput = New-Object System.Windows.Forms.Label
$labelOutput.Text = "下载保存位置"
$labelOutput.Location = New-Object System.Drawing.Point(16, 322)
$labelOutput.Size = New-Object System.Drawing.Size(120, 24)
$form.Controls.Add($labelOutput)

$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Location = New-Object System.Drawing.Point(136, 320)
$outputBox.Size = New-Object System.Drawing.Size(610, 26)
$outputBox.Anchor = "Top,Left,Right"
$outputBox.Text = $defaultOutput
$form.Controls.Add($outputBox)

$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "选择目录"
$browseButton.Location = New-Object System.Drawing.Point(758, 318)
$browseButton.Size = New-Object System.Drawing.Size(128, 30)
$browseButton.Anchor = "Top,Right"
$form.Controls.Add($browseButton)

$highQualityBox = New-Object System.Windows.Forms.CheckBox
$highQualityBox.Text = "尽量下载高清"
$highQualityBox.Location = New-Object System.Drawing.Point(16, 366)
$highQualityBox.Size = New-Object System.Drawing.Size(140, 28)
$form.Controls.Add($highQualityBox)

$cookiesBox = New-Object System.Windows.Forms.CheckBox
$cookiesBox.Text = "使用 cookies.txt"
$cookiesBox.Location = New-Object System.Drawing.Point(168, 366)
$cookiesBox.Size = New-Object System.Drawing.Size(160, 28)
$form.Controls.Add($cookiesBox)

$saveButton = New-Object System.Windows.Forms.Button
$saveButton.Text = "保存链接"
$saveButton.Location = New-Object System.Drawing.Point(16, 412)
$saveButton.Size = New-Object System.Drawing.Size(130, 36)
$form.Controls.Add($saveButton)

$openFolderButton = New-Object System.Windows.Forms.Button
$openFolderButton.Text = "打开目录"
$openFolderButton.Location = New-Object System.Drawing.Point(158, 412)
$openFolderButton.Size = New-Object System.Drawing.Size(130, 36)
$form.Controls.Add($openFolderButton)

$downloadButton = New-Object System.Windows.Forms.Button
$downloadButton.Text = "开始下载"
$downloadButton.Location = New-Object System.Drawing.Point(300, 412)
$downloadButton.Size = New-Object System.Drawing.Size(160, 36)
$form.Controls.Add($downloadButton)

$stopButton = New-Object System.Windows.Forms.Button
$stopButton.Text = "停止"
$stopButton.Location = New-Object System.Drawing.Point(472, 412)
$stopButton.Size = New-Object System.Drawing.Size(110, 36)
$stopButton.Enabled = $false
$form.Controls.Add($stopButton)

$mergeButton = New-Object System.Windows.Forms.Button
$mergeButton.Text = "合并已下载音视频"
$mergeButton.Location = New-Object System.Drawing.Point(594, 412)
$mergeButton.Size = New-Object System.Drawing.Size(170, 36)
$form.Controls.Add($mergeButton)

$workflowButton = New-Object System.Windows.Forms.Button
$workflowButton.Text = "打开分享流程"
$workflowButton.Location = New-Object System.Drawing.Point(16, 466)
$workflowButton.Size = New-Object System.Drawing.Size(150, 32)
$form.Controls.Add($workflowButton)

$templateButton = New-Object System.Windows.Forms.Button
$templateButton.Text = "打开笔记模板"
$templateButton.Location = New-Object System.Drawing.Point(178, 466)
$templateButton.Size = New-Object System.Drawing.Size(150, 32)
$form.Controls.Add($templateButton)

$pptButton = New-Object System.Windows.Forms.Button
$pptButton.Text = "打开分享PPT"
$pptButton.Location = New-Object System.Drawing.Point(340, 466)
$pptButton.Size = New-Object System.Drawing.Size(150, 32)
$form.Controls.Add($pptButton)

$uploadGroup = New-Object System.Windows.Forms.GroupBox
$uploadGroup.Text = "上传到百度网盘"
$uploadGroup.Location = New-Object System.Drawing.Point(16, 512)
$uploadGroup.Size = New-Object System.Drawing.Size(870, 154)
$uploadGroup.Anchor = "Top,Left,Right"
$form.Controls.Add($uploadGroup)

$selectedFilesBox = New-Object System.Windows.Forms.TextBox
$selectedFilesBox.Location = New-Object System.Drawing.Point(14, 28)
$selectedFilesBox.Size = New-Object System.Drawing.Size(530, 26)
$selectedFilesBox.ReadOnly = $true
$selectedFilesBox.Text = "尚未选择文件"
$selectedFilesBox.Anchor = "Top,Left,Right"
$uploadGroup.Controls.Add($selectedFilesBox)

$chooseFilesButton = New-Object System.Windows.Forms.Button
$chooseFilesButton.Text = "批量选择文件"
$chooseFilesButton.Location = New-Object System.Drawing.Point(558, 26)
$chooseFilesButton.Size = New-Object System.Drawing.Size(130, 30)
$chooseFilesButton.Anchor = "Top,Right"
$uploadGroup.Controls.Add($chooseFilesButton)

$selectAllVideosButton = New-Object System.Windows.Forms.Button
$selectAllVideosButton.Text = "全选下载视频"
$selectAllVideosButton.Location = New-Object System.Drawing.Point(704, 26)
$selectAllVideosButton.Size = New-Object System.Drawing.Size(140, 30)
$selectAllVideosButton.Anchor = "Top,Right"
$uploadGroup.Controls.Add($selectAllVideosButton)

$openBaiduButton = New-Object System.Windows.Forms.Button
$openBaiduButton.Text = "打开百度网盘"
$openBaiduButton.Location = New-Object System.Drawing.Point(704, 108)
$openBaiduButton.Size = New-Object System.Drawing.Size(140, 30)
$openBaiduButton.Anchor = "Top,Right"
$uploadGroup.Controls.Add($openBaiduButton)

$syncLabel = New-Object System.Windows.Forms.Label
$syncLabel.Text = "同步目录"
$syncLabel.Location = New-Object System.Drawing.Point(14, 72)
$syncLabel.Size = New-Object System.Drawing.Size(86, 24)
$uploadGroup.Controls.Add($syncLabel)

$syncFolderBox = New-Object System.Windows.Forms.TextBox
$syncFolderBox.Location = New-Object System.Drawing.Point(100, 70)
$syncFolderBox.Size = New-Object System.Drawing.Size(444, 26)
$syncFolderBox.Anchor = "Top,Left,Right"
$syncFolderBox.Text = Join-Path $baseDir "BaiduUploadQueue"
$uploadGroup.Controls.Add($syncFolderBox)

$browseSyncButton = New-Object System.Windows.Forms.Button
$browseSyncButton.Text = "选择目录"
$browseSyncButton.Location = New-Object System.Drawing.Point(558, 68)
$browseSyncButton.Size = New-Object System.Drawing.Size(130, 30)
$browseSyncButton.Anchor = "Top,Right"
$uploadGroup.Controls.Add($browseSyncButton)

$copyToSyncButton = New-Object System.Windows.Forms.Button
$copyToSyncButton.Text = "批量上传到同步目录"
$copyToSyncButton.Location = New-Object System.Drawing.Point(704, 68)
$copyToSyncButton.Size = New-Object System.Drawing.Size(140, 30)
$copyToSyncButton.Anchor = "Top,Right"
$uploadGroup.Controls.Add($copyToSyncButton)

$uploadTipLabel = New-Object System.Windows.Forms.Label
$uploadTipLabel.Text = "提示：批量选择文件时可按 Ctrl/Shift 多选；全选下载视频会自动选择下载目录下所有完整视频文件。"
$uploadTipLabel.Location = New-Object System.Drawing.Point(14, 112)
$uploadTipLabel.Size = New-Object System.Drawing.Size(670, 24)
$uploadTipLabel.Anchor = "Top,Left,Right"
$uploadGroup.Controls.Add($uploadTipLabel)

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.ReadOnly = $true
$logBox.Location = New-Object System.Drawing.Point(16, 682)
$logBox.Size = New-Object System.Drawing.Size(870, 124)
$logBox.Anchor = "Top,Left,Right,Bottom"
$form.Controls.Add($logBox)

$script:process = $null

$browseButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.SelectedPath = $outputBox.Text
    if ($dialog.ShowDialog() -eq "OK") {
        $outputBox.Text = $dialog.SelectedPath
    }
})

$saveButton.Add_Click({
    Set-Content -LiteralPath $urlFile -Value $urlBox.Text -Encoding UTF8
    Add-Log "链接已保存到：$urlFile"
})

$openFolderButton.Add_Click({
    $target = $outputBox.Text
    if (-not (Test-Path -LiteralPath $target)) {
        New-Item -ItemType Directory -Force -Path $target | Out-Null
    }
    Start-Process explorer.exe $target
})

$workflowButton.Add_Click({
    if (Test-Path -LiteralPath $workflowDocFile) {
        Start-Process -FilePath $workflowDocFile
    } else {
        [System.Windows.Forms.MessageBox]::Show("未找到分享流程文档：$workflowDocFile", "文件不存在") | Out-Null
    }
})

$templateButton.Add_Click({
    if (Test-Path -LiteralPath $noteTemplateFile) {
        Start-Process -FilePath $noteTemplateFile
    } else {
        [System.Windows.Forms.MessageBox]::Show("未找到笔记模板：$noteTemplateFile", "文件不存在") | Out-Null
    }
})

$pptButton.Add_Click({
    if (Test-Path -LiteralPath $sharePptFile) {
        Start-Process -FilePath $sharePptFile
    } else {
        [System.Windows.Forms.MessageBox]::Show("未找到分享PPT：$sharePptFile", "文件不存在") | Out-Null
    }
})

$mergeButton.Add_Click({
    if (-not (Test-Path -LiteralPath $mergeScriptFile)) {
        [System.Windows.Forms.MessageBox]::Show("未找到合并脚本：$mergeScriptFile", "缺少脚本") | Out-Null
        return
    }

    New-Item -ItemType Directory -Force -Path $outputBox.Text | Out-Null
    $args = @(
        "-NoExit",
        "-ExecutionPolicy", "Bypass",
        "-File", "`"$mergeScriptFile`"",
        "-RootDir", "`"$($outputBox.Text)`""
    )

    Add-Log "正在打开音视频合并窗口..."
    Start-Process powershell.exe -ArgumentList $args -WorkingDirectory $baseDir
})

$chooseFilesButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Multiselect = $true
    $dialog.Title = "选择要上传的已下载文件"
    $dialog.InitialDirectory = $outputBox.Text
    $dialog.Filter = "视频文件|*.mp4;*.mkv;*.flv;*.webm;*.mov;*.avi;*.m4a;*.mp3|所有文件|*.*"

    if ($dialog.ShowDialog() -eq "OK") {
        Set-UploadFiles -Files $dialog.FileNames
    }
})

$selectAllVideosButton.Add_Click({
    if (-not (Test-Path -LiteralPath $outputBox.Text)) {
        [System.Windows.Forms.MessageBox]::Show("下载目录不存在：$($outputBox.Text)", "目录不存在") | Out-Null
        return
    }

    $videoFiles = Get-ChildItem -LiteralPath $outputBox.Text -Recurse -File |
        Where-Object {
            $_.Extension -in @(".mp4", ".mkv", ".flv", ".webm", ".mov", ".avi") -and
            $_.BaseName -notmatch "\.f\d+$"
        } |
        Sort-Object FullName |
        Select-Object -ExpandProperty FullName

    if (-not $videoFiles -or $videoFiles.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("下载目录中没有找到可批量上传的完整视频文件。", "没有视频文件") | Out-Null
        return
    }

    Set-UploadFiles -Files $videoFiles
})

$browseSyncButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.SelectedPath = $syncFolderBox.Text
    if ($dialog.ShowDialog() -eq "OK") {
        $syncFolderBox.Text = $dialog.SelectedPath
    }
})

$openBaiduButton.Add_Click({
    if (Test-Path -LiteralPath $baiduNetdiskExe) {
        Start-Process -FilePath $baiduNetdiskExe
        Add-Log "已打开百度网盘客户端。"
    } else {
        Add-Log "未找到百度网盘客户端：$baiduNetdiskExe"
    }

    if ($script:selectedUploadFiles.Count -gt 0) {
        [System.Windows.Forms.Clipboard]::SetText(($script:selectedUploadFiles -join "`r`n"))
        $firstFolder = Split-Path -Parent $script:selectedUploadFiles[0]
        if (Test-Path -LiteralPath $firstFolder) {
            Start-Process explorer.exe $firstFolder
        }
        Add-Log "已复制所选文件路径。请在百度网盘中点击上传并选择这些文件。"
    }
})

$copyToSyncButton.Add_Click({
    if ($script:selectedUploadFiles.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("请先选择要上传的文件。", "尚未选择文件") | Out-Null
        return
    }

    $syncTarget = $syncFolderBox.Text
    if (-not (Test-Path -LiteralPath $syncTarget)) {
        New-Item -ItemType Directory -Force -Path $syncTarget | Out-Null
    }

    $copiedCount = 0
    foreach ($file in $script:selectedUploadFiles) {
        if (Test-Path -LiteralPath $file) {
            Copy-Item -LiteralPath $file -Destination $syncTarget -Force
            $copiedCount += 1
            Add-Log "已加入上传队列：$(Split-Path -Leaf $file)"
        }
    }

    Start-Process explorer.exe $syncTarget
    Add-Log "批量上传准备完成：已复制 $copiedCount 个文件到同步目录：$syncTarget"
    Add-Log "如果该目录已在百度网盘中设置为同步目录，客户端会自动开始上传。"
})

$downloadButton.Add_Click({
    Set-Content -LiteralPath $urlFile -Value $urlBox.Text -Encoding UTF8

    $urls = $urlBox.Lines | ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not $_.StartsWith("#") }
    if ($urls.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("请至少添加一个 B 站视频链接。", "没有链接") | Out-Null
        return
    }

    if ($cookiesBox.Checked -and -not (Test-Path -LiteralPath $cookiesFile)) {
        [System.Windows.Forms.MessageBox]::Show("未在工具目录中找到 cookies.txt：$baseDir", "缺少 cookies") | Out-Null
        return
    }

    New-Item -ItemType Directory -Force -Path $outputBox.Text | Out-Null

    $args = @(
        "-NoExit",
        "-ExecutionPolicy", "Bypass",
        "-File", "`"$scriptFile`"",
        "-UrlFile", "`"$urlFile`"",
        "-OutputDir", "`"$($outputBox.Text)`""
    )

    if ($highQualityBox.Checked) {
        $args += "-HighQuality"
    }
    if ($cookiesBox.Checked) {
        $args += @("-UseCookies", "-CookiesFile", "`"$cookiesFile`"")
    }

    Add-Log "正在打开 PowerShell 下载窗口..."
    $script:process = Start-Process powershell.exe -ArgumentList $args -WorkingDirectory $baseDir -PassThru
    $downloadButton.Enabled = $false
    $stopButton.Enabled = $true
})

$stopButton.Add_Click({
    if ($script:process -and -not $script:process.HasExited) {
        $script:process.Kill()
        Add-Log "下载进程已停止。"
    }
    $downloadButton.Enabled = $true
    $stopButton.Enabled = $false
})

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 1000
$timer.Add_Tick({
    if ($script:process -and $script:process.HasExited) {
        $downloadButton.Enabled = $true
        $stopButton.Enabled = $false
        $script:process = $null
        Add-Log "下载进程已结束。"
    }
})
$timer.Start()

$form.Add_Shown({
    Add-Log "准备就绪。工具目录：$baseDir"
    Add-Log "提示：高清下载可能需要安装 ffmpeg。"
})

[void]$form.ShowDialog()
