# my-workflow-template.ps1
# 個人工作流範本
# 請根據你的需求修改此腳本

param(
    [Parameter(Position=0)]
    [ValidateSet("convert", "batch", "watch", "report", "help")]
    [string]$Action = "help",

    [string]$InputPath = ".",
    [string]$OutputPath = ".\output"
)

# 顏色輸出函數
function Write-Success($msg) { Write-Host $msg -ForegroundColor Green }
function Write-Info($msg) { Write-Host $msg -ForegroundColor Cyan }
function Write-Warn($msg) { Write-Host $msg -ForegroundColor Yellow }

# 確保輸出資料夾存在
function Ensure-OutputFolder {
    if (!(Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        Write-Info "已建立輸出資料夾: $OutputPath"
    }
}

# 顯示說明
function Show-Help {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "    我的文件處理工作流" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "用法: .\my-workflow.ps1 -Action <動作> [選項]"
    Write-Host ""
    Write-Host "可用動作:"
    Write-Host "  convert  - 轉換單一檔案"
    Write-Host "  batch    - 批次轉換資料夾中的所有 .md"
    Write-Host "  watch    - 監控資料夾自動轉換"
    Write-Host "  report   - 生成處理報告"
    Write-Host "  help     - 顯示此說明"
    Write-Host ""
    Write-Host "範例:"
    Write-Host "  .\my-workflow.ps1 -Action batch -InputPath D:\文件"
    Write-Host ""
}

# 單一檔案轉換
function Convert-Single {
    Write-Info "請選擇要轉換的檔案..."

    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "Markdown (*.md)|*.md|所有檔案 (*.*)|*.*"

    if ($dialog.ShowDialog() -eq 'OK') {
        $input = $dialog.FileName
        $output = [System.IO.Path]::ChangeExtension($input, ".docx")

        Write-Info "轉換中: $input"
        pandoc $input -o $output

        if (Test-Path $output) {
            Write-Success "完成: $output"
        }
    }
}

# 批次轉換
function Convert-Batch {
    Ensure-OutputFolder

    $files = Get-ChildItem "$InputPath\*.md"

    if ($files.Count -eq 0) {
        Write-Warn "在 $InputPath 中找不到 .md 檔案"
        return
    }

    Write-Info "找到 $($files.Count) 個檔案"

    $success = 0
    foreach ($file in $files) {
        $output = Join-Path $OutputPath "$($file.BaseName).docx"
        try {
            pandoc $file.FullName -o $output
            Write-Success "  ✓ $($file.Name)"
            $success++
        } catch {
            Write-Warn "  ✗ $($file.Name): $_"
        }
    }

    Write-Host ""
    Write-Info "完成: $success / $($files.Count) 個檔案"
}

# 監控資料夾
function Watch-Folder {
    Ensure-OutputFolder

    Write-Info "監控中: $InputPath"
    Write-Info "輸出到: $OutputPath"
    Write-Host "按 Ctrl+C 停止"
    Write-Host ""

    $watcher = New-Object System.IO.FileSystemWatcher
    $watcher.Path = (Resolve-Path $InputPath).Path
    $watcher.Filter = "*.md"
    $watcher.EnableRaisingEvents = $true

    $action = {
        $file = $Event.SourceEventArgs.FullPath
        $name = $Event.SourceEventArgs.Name
        Start-Sleep -Seconds 1
        $output = Join-Path $using:OutputPath ([System.IO.Path]::ChangeExtension($name, ".docx"))
        pandoc $file -o $output
        Write-Host "$(Get-Date -Format 'HH:mm:ss') 已轉換: $name" -ForegroundColor Green
    }

    Register-ObjectEvent $watcher "Created" -Action $action | Out-Null

    while ($true) { Start-Sleep -Seconds 1 }
}

# 生成報告
function Generate-Report {
    $date = Get-Date -Format "yyyy-MM-dd"
    $reportFile = "report-$date.md"

    $mdFiles = Get-ChildItem "$InputPath\*.md" -ErrorAction SilentlyContinue
    $docxFiles = Get-ChildItem "$OutputPath\*.docx" -ErrorAction SilentlyContinue

    $report = @"
# 文件處理報告

**日期**: $date
**來源**: $InputPath
**輸出**: $OutputPath

## 統計

- Markdown 檔案: $($mdFiles.Count) 個
- Word 檔案: $($docxFiles.Count) 個

## 檔案列表

### 來源檔案
$($mdFiles | ForEach-Object { "- $($_.Name)" } | Out-String)

### 輸出檔案
$($docxFiles | ForEach-Object { "- $($_.Name)" } | Out-String)

---
*報告自動生成*
"@

    $report | Out-File $reportFile -Encoding UTF8
    Write-Success "報告已生成: $reportFile"
}

# 主程式
switch ($Action) {
    "convert" { Convert-Single }
    "batch"   { Convert-Batch }
    "watch"   { Watch-Folder }
    "report"  { Generate-Report }
    "help"    { Show-Help }
    default   { Show-Help }
}
