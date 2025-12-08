# smart-convert.ps1 - 智慧轉換腳本
# 支援參數：輸入資料夾、輸出資料夾、輸出格式
# 使用範例：
#   .\smart-convert.ps1
#   .\smart-convert.ps1 -Format html
#   .\smart-convert.ps1 -InputFolder "D:\筆記" -OutputFolder "D:\輸出" -Format docx

param(
    [string]$InputFolder = ".",
    [string]$OutputFolder = "output",
    [string]$Format = "docx"
)

# 顯示設定
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "📁 輸入資料夾: $InputFolder" -ForegroundColor Cyan
Write-Host "📂 輸出資料夾: $OutputFolder" -ForegroundColor Cyan
Write-Host "📄 輸出格式: $Format" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# 建立輸出資料夾
if (!(Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
    Write-Host "✅ 已建立輸出資料夾" -ForegroundColor Green
}

# 取得檔案
$files = Get-ChildItem -Path "$InputFolder\*.md" -ErrorAction SilentlyContinue

if ($files.Count -eq 0) {
    Write-Host "❌ 在 $InputFolder 中找不到 .md 檔案" -ForegroundColor Red
    exit
}

Write-Host "🔍 找到 $($files.Count) 個檔案`n" -ForegroundColor Cyan

# 轉換
$success = 0
$failed = 0

foreach ($file in $files) {
    $outputPath = "$OutputFolder\$($file.BaseName).$Format"
    Write-Host "📄 $($file.Name)" -NoNewline

    try {
        if ($Format -eq "html") {
            pandoc $file.FullName -o $outputPath --standalone
        } else {
            pandoc $file.FullName -o $outputPath
        }
        Write-Host " → $($file.BaseName).$Format ✅" -ForegroundColor Green
        $success++
    } catch {
        Write-Host " ❌ 錯誤: $($_.Exception.Message)" -ForegroundColor Red
        $failed++
    }
}

# 統計
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "📊 轉換完成！" -ForegroundColor Cyan
Write-Host "   成功: $success" -ForegroundColor Green
Write-Host "   失敗: $failed" -ForegroundColor $(if ($failed -gt 0) { "Red" } else { "Green" })
Write-Host "   輸出位置: $OutputFolder" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
