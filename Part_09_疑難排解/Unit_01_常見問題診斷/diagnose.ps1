# diagnose.ps1
# Pandoc 問題診斷腳本
# 使用方式：.\diagnose.ps1 -File "your-file.md"

param(
    [string]$File
)

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "    Pandoc 問題診斷工具" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$issues = @()

# 1. 檢查 Pandoc 安裝
Write-Host "[1] 檢查 Pandoc 安裝" -ForegroundColor Yellow
try {
    $version = pandoc --version | Select-Object -First 1
    Write-Host "    狀態: 已安裝" -ForegroundColor Green
    Write-Host "    版本: $version"
} catch {
    Write-Host "    狀態: 未安裝或不在 PATH 中" -ForegroundColor Red
    $issues += "Pandoc 未安裝或未加入 PATH"
}

# 2. 檢查檔案（如有提供）
if ($File) {
    Write-Host ""
    Write-Host "[2] 檢查輸入檔案" -ForegroundColor Yellow

    if (Test-Path $File) {
        $fileInfo = Get-Item $File
        Write-Host "    狀態: 檔案存在" -ForegroundColor Green
        Write-Host "    大小: $($fileInfo.Length) bytes"
        Write-Host "    修改: $($fileInfo.LastWriteTime)"

        # 檢查編碼
        Write-Host ""
        Write-Host "[3] 檢查檔案編碼" -ForegroundColor Yellow
        $bytes = [System.IO.File]::ReadAllBytes($File)
        if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
            Write-Host "    編碼: UTF-8 (with BOM)" -ForegroundColor Green
        } elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
            Write-Host "    編碼: UTF-16 LE" -ForegroundColor Yellow
            Write-Host "    建議: 轉換為 UTF-8"
            $issues += "檔案編碼不是 UTF-8"
        } else {
            Write-Host "    編碼: UTF-8 (無 BOM) 或其他" -ForegroundColor Green
        }

        # 檢查 YAML 標頭
        Write-Host ""
        Write-Host "[4] 檢查 YAML 標頭" -ForegroundColor Yellow
        $content = Get-Content $File -Raw -Encoding UTF8
        if ($content -match "^---") {
            if ($content -match "^---[\s\S]*?---") {
                Write-Host "    狀態: YAML 標頭格式正確" -ForegroundColor Green
            } else {
                Write-Host "    狀態: YAML 標頭不完整 (缺少結尾 ---)" -ForegroundColor Red
                $issues += "YAML 標頭缺少結尾 ---"
            }
        } else {
            Write-Host "    狀態: 無 YAML 標頭" -ForegroundColor Gray
        }

        # 測試轉換
        Write-Host ""
        Write-Host "[5] 測試轉換" -ForegroundColor Yellow
        $testOutput = "$env:TEMP\pandoc-test-output.html"
        try {
            $result = pandoc $File -o $testOutput 2>&1
            if (Test-Path $testOutput) {
                Write-Host "    狀態: 轉換成功" -ForegroundColor Green
                Remove-Item $testOutput -Force
            } else {
                Write-Host "    狀態: 轉換失敗 (無輸出)" -ForegroundColor Red
                $issues += "轉換測試失敗"
            }
        } catch {
            Write-Host "    狀態: 轉換失敗" -ForegroundColor Red
            Write-Host "    錯誤: $_" -ForegroundColor Red
            $issues += "轉換測試失敗: $_"
        }

    } else {
        Write-Host "    狀態: 檔案不存在" -ForegroundColor Red
        $issues += "指定的檔案不存在"
    }
} else {
    Write-Host ""
    Write-Host "[提示] 未指定檔案，跳過檔案相關檢查" -ForegroundColor Gray
    Write-Host "       使用 -File 參數指定檔案進行完整檢查"
}

# 結果摘要
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "    診斷結果" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

if ($issues.Count -eq 0) {
    Write-Host "沒有發現問題！" -ForegroundColor Green
} else {
    Write-Host "發現 $($issues.Count) 個問題：" -ForegroundColor Red
    $issues | ForEach-Object {
        Write-Host "  - $_" -ForegroundColor Yellow
    }
}
Write-Host ""
