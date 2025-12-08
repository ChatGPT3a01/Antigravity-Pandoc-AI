# convert.ps1 - 批次轉換腳本範例
# 用途：將當前資料夾所有 .md 檔案轉成 .docx
# 使用方式：在含有 .md 檔案的資料夾執行 .\convert.ps1

$ErrorActionPreference = "Continue"
$files = Get-ChildItem -Filter "*.md"

if ($files.Count -eq 0) {
    Write-Host "❌ 找不到任何 .md 檔案" -ForegroundColor Red
    exit
}

Write-Host "🔍 找到 $($files.Count) 個 Markdown 檔案" -ForegroundColor Cyan

$success = 0
$failed = 0

foreach ($file in $files) {
    Write-Host "📄 轉換: $($file.Name)" -NoNewline
    try {
        pandoc $file.Name -o ($file.BaseName + ".docx")
        Write-Host " ✅" -ForegroundColor Green
        $success++
    } catch {
        Write-Host " ❌" -ForegroundColor Red
        $failed++
    }
}

Write-Host "`n📊 完成！成功: $success, 失敗: $failed" -ForegroundColor Cyan
