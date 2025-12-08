# debug-pandoc.ps1
# Pandoc 除錯輔助工具
# 使用方式：.\debug-pandoc.ps1 -InputFile "problem.md"

param(
    [Parameter(Mandatory=$true)]
    [string]$InputFile
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "    Pandoc 除錯工具 v1.0" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 1. 檢查 Pandoc 安裝
Write-Host "[1] 檢查 Pandoc 安裝" -ForegroundColor Yellow
try {
    $version = pandoc --version | Select-Object -First 1
    Write-Host "    版本: $version" -ForegroundColor Green
} catch {
    Write-Host "    錯誤: 找不到 Pandoc，請確認已安裝並加入 PATH" -ForegroundColor Red
    exit 1
}

# 2. 檢查輸入檔案
Write-Host ""
Write-Host "[2] 檢查輸入檔案" -ForegroundColor Yellow
if (Test-Path $InputFile) {
    $file = Get-Item $InputFile
    Write-Host "    路徑: $($file.FullName)" -ForegroundColor Green
    Write-Host "    大小: $($file.Length) bytes"
    Write-Host "    修改: $($file.LastWriteTime)"
} else {
    Write-Host "    錯誤: 檔案不存在 - $InputFile" -ForegroundColor Red
    exit 1
}

# 3. 檢查檔案編碼
Write-Host ""
Write-Host "[3] 檢查檔案編碼" -ForegroundColor Yellow
$bytes = [System.IO.File]::ReadAllBytes($InputFile)
if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
    Write-Host "    編碼: UTF-8 with BOM" -ForegroundColor Green
} elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
    Write-Host "    編碼: UTF-16 LE" -ForegroundColor Yellow
    Write-Host "    建議: 轉換成 UTF-8 格式"
} else {
    Write-Host "    編碼: UTF-8 (無 BOM) 或其他" -ForegroundColor Green
}

# 4. 檢查檔案開頭
Write-Host ""
Write-Host "[4] 檔案開頭 (前 10 行)" -ForegroundColor Yellow
$lines = Get-Content $InputFile -TotalCount 10 -Encoding UTF8
$lineNum = 1
foreach ($line in $lines) {
    $display = if ($line.Length -gt 60) { $line.Substring(0, 60) + "..." } else { $line }
    Write-Host "    $($lineNum.ToString('D2')): $display"
    $lineNum++
}

# 5. 檢查 YAML 標頭
Write-Host ""
Write-Host "[5] 檢查 YAML 標頭" -ForegroundColor Yellow
$content = Get-Content $InputFile -Raw -Encoding UTF8
if ($content -match "^---") {
    if ($content -match "^---[\s\S]*?---") {
        Write-Host "    狀態: 找到 YAML 標頭" -ForegroundColor Green

        # 提取 YAML 內容
        $yaml = ($content | Select-String -Pattern "^---[\s\S]*?---" -AllMatches).Matches[0].Value
        Write-Host "    內容預覽:"
        $yaml -split "`n" | Select-Object -First 5 | ForEach-Object {
            Write-Host "        $_"
        }
    } else {
        Write-Host "    警告: 找到開頭 --- 但沒有結尾 ---" -ForegroundColor Red
        Write-Host "    這可能導致 YAML 解析錯誤"
    }
} else {
    Write-Host "    狀態: 沒有 YAML 標頭" -ForegroundColor Gray
}

# 6. 測試轉換
Write-Host ""
Write-Host "[6] 測試轉換 (HTML)" -ForegroundColor Yellow
$testOutput = "$env:TEMP\pandoc-debug-test.html"
try {
    $result = pandoc $InputFile -o $testOutput 2>&1
    if (Test-Path $testOutput) {
        $size = (Get-Item $testOutput).Length
        Write-Host "    結果: 轉換成功 ($size bytes)" -ForegroundColor Green
        Remove-Item $testOutput -Force
    } else {
        Write-Host "    結果: 轉換失敗 (無輸出檔)" -ForegroundColor Red
    }
} catch {
    Write-Host "    錯誤: $_" -ForegroundColor Red
}

# 7. 詳細模式測試
Write-Host ""
Write-Host "[7] 詳細模式輸出" -ForegroundColor Yellow
Write-Host "    執行: pandoc $InputFile --verbose" -ForegroundColor Gray
Write-Host ""
pandoc $InputFile -o "$env:TEMP\verbose-test.html" --verbose 2>&1 | ForEach-Object {
    Write-Host "    $_" -ForegroundColor Gray
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "    除錯完成" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
