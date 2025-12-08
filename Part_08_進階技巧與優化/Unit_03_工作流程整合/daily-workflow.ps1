# daily-workflow.ps1
# 完整的每日文件處理工作流程
# 使用方式：.\daily-workflow.ps1

param(
    [string]$InputFolder = ".\input",
    [string]$OutputFolder = ".\output",
    [string]$ArchiveFolder = ".\archive",
    [switch]$Notify
)

$date = Get-Date -Format "yyyy-MM-dd"
$timestamp = Get-Date -Format "HH:mm:ss"

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "    每日文件處理工作流程" -ForegroundColor Cyan
Write-Host "    $date $timestamp" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 確保資料夾存在
foreach ($folder in @($InputFolder, $OutputFolder, $ArchiveFolder)) {
    if (!(Test-Path $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
}

# 步驟 1: 收集待處理檔案
Write-Host "[1/5] 收集待處理檔案..." -ForegroundColor Yellow
$files = Get-ChildItem "$InputFolder\*.md" -ErrorAction SilentlyContinue

if ($files.Count -eq 0) {
    Write-Host "      沒有待處理的檔案" -ForegroundColor Gray
    Write-Host ""
    Write-Host "工作流程完成 (無檔案需處理)" -ForegroundColor Green
    exit 0
}

Write-Host "      找到 $($files.Count) 個檔案" -ForegroundColor Green

# 步驟 2: 處理檔案
Write-Host ""
Write-Host "[2/5] 轉換檔案..." -ForegroundColor Yellow

$results = @{
    Success = @()
    Failed = @()
}

foreach ($file in $files) {
    Write-Host "      處理: $($file.Name)..." -NoNewline

    try {
        $outputFile = Join-Path $OutputFolder "$($file.BaseName).docx"
        pandoc $file.FullName -o $outputFile -ErrorAction Stop

        if (Test-Path $outputFile) {
            $results.Success += $file.Name
            Write-Host " OK" -ForegroundColor Green
        } else {
            throw "輸出檔案未產生"
        }
    } catch {
        $results.Failed += @{
            Name = $file.Name
            Error = $_.ToString()
        }
        Write-Host " 失敗" -ForegroundColor Red
    }
}

# 步驟 3: 歸檔已處理檔案
Write-Host ""
Write-Host "[3/5] 歸檔原始檔案..." -ForegroundColor Yellow

$archiveSubfolder = Join-Path $ArchiveFolder $date
if (!(Test-Path $archiveSubfolder)) {
    New-Item -ItemType Directory -Path $archiveSubfolder -Force | Out-Null
}

foreach ($fileName in $results.Success) {
    $sourceFile = Join-Path $InputFolder $fileName
    $destFile = Join-Path $archiveSubfolder $fileName
    Move-Item $sourceFile $destFile -Force
}
Write-Host "      已移動 $($results.Success.Count) 個檔案到 $archiveSubfolder" -ForegroundColor Green

# 步驟 4: 產生報告
Write-Host ""
Write-Host "[4/5] 產生處理報告..." -ForegroundColor Yellow

$reportContent = @"
# 文件處理報告

**日期**: $date
**時間**: $timestamp

## 處理結果

- 成功: $($results.Success.Count) 個
- 失敗: $($results.Failed.Count) 個

## 成功列表

$($results.Success | ForEach-Object { "- $_" } | Out-String)

## 失敗列表

$($results.Failed | ForEach-Object { "- $($_.Name): $($_.Error)" } | Out-String)

---
*報告自動產生於 $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")*
"@

$reportFile = Join-Path $OutputFolder "report-$date.md"
$reportContent | Out-File $reportFile -Encoding UTF8
Write-Host "      報告已儲存: $reportFile" -ForegroundColor Green

# 步驟 5: 發送通知
Write-Host ""
Write-Host "[5/5] 完成通知..." -ForegroundColor Yellow

if ($Notify) {
    try {
        $shell = New-Object -ComObject WScript.Shell
        $message = "處理完成！`n成功: $($results.Success.Count)`n失敗: $($results.Failed.Count)"
        $shell.Popup($message, 5, "每日文件處理", 64)
        Write-Host "      已發送桌面通知" -ForegroundColor Green
    } catch {
        Write-Host "      通知發送失敗: $_" -ForegroundColor Yellow
    }
} else {
    Write-Host "      (通知已停用，使用 -Notify 啟用)" -ForegroundColor Gray
}

# 完成摘要
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "    工作流程完成" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "成功處理: $($results.Success.Count) 個檔案" -ForegroundColor Green
if ($results.Failed.Count -gt 0) {
    Write-Host "處理失敗: $($results.Failed.Count) 個檔案" -ForegroundColor Red
}
Write-Host "輸出位置: $OutputFolder"
Write-Host ""
