# watch-folder.ps1
# 資料夾監控自動轉換腳本
# 使用方式：.\watch-folder.ps1 -WatchPath "D:\待處理" -OutputPath "D:\已完成"

param(
    [string]$WatchPath = ".\input",
    [string]$OutputPath = ".\output",
    [string]$OutputFormat = "docx"
)

# 確保路徑存在
if (!(Test-Path $WatchPath)) {
    New-Item -ItemType Directory -Path $WatchPath -Force | Out-Null
    Write-Host "已建立監控資料夾: $WatchPath" -ForegroundColor Yellow
}

if (!(Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    Write-Host "已建立輸出資料夾: $OutputPath" -ForegroundColor Yellow
}

# 取得完整路徑
$WatchPath = (Resolve-Path $WatchPath).Path
$OutputPath = (Resolve-Path $OutputPath).Path

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "    資料夾監控服務" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "監控資料夾: $WatchPath" -ForegroundColor Green
Write-Host "輸出資料夾: $OutputPath" -ForegroundColor Green
Write-Host "輸出格式: $OutputFormat"
Write-Host ""
Write-Host "服務已啟動，等待新檔案..."
Write-Host "按 Ctrl+C 停止服務"
Write-Host ""

# 建立檔案監控器
$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $WatchPath
$watcher.Filter = "*.md"
$watcher.IncludeSubdirectories = $false
$watcher.EnableRaisingEvents = $true

# 處理函數
$processFile = {
    param($source, $e)

    $inputFile = $e.FullPath
    $fileName = $e.Name
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
    $outputFile = Join-Path $OutputPath "$baseName.$using:OutputFormat"

    # 等待檔案寫入完成
    Start-Sleep -Seconds 1

    $timestamp = Get-Date -Format "HH:mm:ss"
    Write-Host "[$timestamp] 偵測到新檔案: $fileName" -ForegroundColor Yellow

    try {
        pandoc $inputFile -o $outputFile
        Write-Host "[$timestamp] 轉換完成: $baseName.$using:OutputFormat" -ForegroundColor Green
    } catch {
        Write-Host "[$timestamp] 轉換失敗: $_" -ForegroundColor Red
    }
}

# 註冊事件
Register-ObjectEvent $watcher "Created" -Action $processFile | Out-Null

# 保持腳本運行
try {
    while ($true) {
        Start-Sleep -Seconds 1
    }
} finally {
    # 清理
    $watcher.EnableRaisingEvents = $false
    $watcher.Dispose()
    Write-Host "`n服務已停止" -ForegroundColor Yellow
}
