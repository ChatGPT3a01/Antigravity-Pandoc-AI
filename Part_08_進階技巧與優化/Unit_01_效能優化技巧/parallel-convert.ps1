# parallel-convert.ps1
# 平行批次轉換腳本 (需要 PowerShell 7+)
# 使用方式：.\parallel-convert.ps1 -InputFolder "docs" -OutputFormat "docx"

param(
    [string]$InputFolder = ".",
    [string]$OutputFormat = "docx",
    [int]$ThrottleLimit = 4
)

# 檢查 PowerShell 版本
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "警告: 平行處理需要 PowerShell 7+" -ForegroundColor Yellow
    Write-Host "將使用逐一處理模式" -ForegroundColor Yellow
    $useParallel = $false
} else {
    $useParallel = $true
}

$files = Get-ChildItem "$InputFolder\*.md"

if ($files.Count -eq 0) {
    Write-Host "找不到 .md 檔案" -ForegroundColor Red
    exit 1
}

Write-Host "=== 批次轉換 ===" -ForegroundColor Cyan
Write-Host "找到 $($files.Count) 個檔案"
Write-Host "輸出格式: $OutputFormat"
Write-Host "平行處理: $(if($useParallel){'啟用 (x' + $ThrottleLimit + ')'}else{'停用'})"
Write-Host ""

$startTime = Get-Date

if ($useParallel) {
    # PowerShell 7+ 平行處理
    $files | ForEach-Object -Parallel {
        $output = [System.IO.Path]::ChangeExtension($_.FullName, $using:OutputFormat)
        pandoc $_.FullName -o $output
        Write-Host "完成: $($_.Name)" -ForegroundColor Green
    } -ThrottleLimit $ThrottleLimit
} else {
    # 逐一處理
    foreach ($file in $files) {
        $output = [System.IO.Path]::ChangeExtension($file.FullName, $OutputFormat)
        pandoc $file.FullName -o $output
        Write-Host "完成: $($file.Name)" -ForegroundColor Green
    }
}

$duration = (Get-Date) - $startTime

Write-Host ""
Write-Host "=== 完成 ===" -ForegroundColor Yellow
Write-Host "處理檔案: $($files.Count) 個"
Write-Host "總耗時: $($duration.TotalSeconds.ToString('F1')) 秒"
Write-Host "平均每檔: $(($duration.TotalSeconds / $files.Count).ToString('F2')) 秒"
