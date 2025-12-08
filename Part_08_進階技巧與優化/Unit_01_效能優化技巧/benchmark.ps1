# benchmark.ps1
# Pandoc 轉換效能測試腳本
# 使用方式：.\benchmark.ps1 -InputFile "test.md"

param(
    [Parameter(Mandatory=$true)]
    [string]$InputFile,
    [string]$OutputFormat = "docx",
    [int]$Iterations = 3
)

# 檢查檔案存在
if (!(Test-Path $InputFile)) {
    Write-Host "錯誤: 找不到檔案 $InputFile" -ForegroundColor Red
    exit 1
}

$outputFile = [System.IO.Path]::ChangeExtension($InputFile, $OutputFormat)
$inputSize = (Get-Item $InputFile).Length / 1KB

Write-Host "=== Pandoc 效能測試 ===" -ForegroundColor Cyan
Write-Host "輸入檔案: $InputFile ($($inputSize.ToString('F1')) KB)"
Write-Host "輸出格式: $OutputFormat"
Write-Host "測試次數: $Iterations"
Write-Host ""

$times = @()

for ($i = 1; $i -le $Iterations; $i++) {
    Write-Host "測試 $i/$Iterations..." -NoNewline

    $result = Measure-Command {
        pandoc $InputFile -o $outputFile
    }

    $times += $result.TotalSeconds
    Write-Host " $($result.TotalSeconds.ToString('F2')) 秒" -ForegroundColor Green
}

# 計算統計
$avg = ($times | Measure-Object -Average).Average
$min = ($times | Measure-Object -Minimum).Minimum
$max = ($times | Measure-Object -Maximum).Maximum
$outputSize = (Get-Item $outputFile).Length / 1KB

Write-Host ""
Write-Host "=== 測試結果 ===" -ForegroundColor Yellow
Write-Host "平均時間: $($avg.ToString('F2')) 秒"
Write-Host "最快: $($min.ToString('F2')) 秒"
Write-Host "最慢: $($max.ToString('F2')) 秒"
Write-Host "輸出大小: $($outputSize.ToString('F1')) KB"
Write-Host "處理速度: $((($inputSize / $avg)).ToString('F1')) KB/秒"
