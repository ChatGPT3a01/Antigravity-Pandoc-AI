# check-env.ps1
# 環境設定檢查腳本
# 使用方式：.\check-env.ps1

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "    環境設定檢查" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 1. 系統資訊
Write-Host "[系統資訊]" -ForegroundColor Yellow
Write-Host "  作業系統: $([System.Environment]::OSVersion.VersionString)"
Write-Host "  PowerShell: $($PSVersionTable.PSVersion)"
Write-Host "  使用者: $env:USERNAME"
Write-Host ""

# 2. Pandoc
Write-Host "[Pandoc]" -ForegroundColor Yellow
$pandocCmd = Get-Command pandoc -ErrorAction SilentlyContinue
if ($pandocCmd) {
    $version = pandoc --version | Select-Object -First 1
    Write-Host "  狀態: 已安裝" -ForegroundColor Green
    Write-Host "  版本: $version"
    Write-Host "  位置: $($pandocCmd.Source)"
} else {
    Write-Host "  狀態: 未安裝或未在 PATH 中" -ForegroundColor Red
    Write-Host "  建議: 從 pandoc.org 下載安裝"
}
Write-Host ""

# 3. LaTeX
Write-Host "[LaTeX]" -ForegroundColor Yellow
$latexCmd = Get-Command pdflatex -ErrorAction SilentlyContinue
if ($latexCmd) {
    Write-Host "  狀態: 已安裝" -ForegroundColor Green
    Write-Host "  位置: $($latexCmd.Source)"
} else {
    Write-Host "  狀態: 未安裝" -ForegroundColor Yellow
    Write-Host "  影響: 無法直接轉換 PDF"
    Write-Host "  建議: 安裝 MiKTeX 或使用其他 PDF 方法"
}
Write-Host ""

# 4. PowerShell 執行原則
Write-Host "[執行原則]" -ForegroundColor Yellow
$policy = Get-ExecutionPolicy
Write-Host "  目前原則: $policy"
if ($policy -eq "Restricted") {
    Write-Host "  狀態: 受限" -ForegroundColor Red
    Write-Host "  影響: 無法執行 .ps1 腳本"
    Write-Host "  建議: 執行 Set-ExecutionPolicy RemoteSigned -Scope CurrentUser"
} else {
    Write-Host "  狀態: 正常" -ForegroundColor Green
}
Write-Host ""

# 5. PATH 檢查
Write-Host "[PATH 環境變數]" -ForegroundColor Yellow
$pathEntries = $env:Path -split ";"
$pandocPaths = $pathEntries | Where-Object { $_ -like "*pandoc*" -or $_ -like "*Pandoc*" }
if ($pandocPaths) {
    Write-Host "  Pandoc 路徑: 已設定" -ForegroundColor Green
    $pandocPaths | ForEach-Object {
        Write-Host "    - $_"
    }
} else {
    Write-Host "  Pandoc 路徑: 未在 PATH 中發現" -ForegroundColor Yellow
}
Write-Host ""

# 6. 常用資料夾
Write-Host "[常用路徑]" -ForegroundColor Yellow
Write-Host "  使用者目錄: $env:USERPROFILE"
Write-Host "  暫存目錄: $env:TEMP"
Write-Host "  目前目錄: $(Get-Location)"
Write-Host ""

# 7. 快速測試
Write-Host "[快速測試]" -ForegroundColor Yellow
if ($pandocCmd) {
    $testFile = "$env:TEMP\pandoc-test.md"
    $testOutput = "$env:TEMP\pandoc-test.html"
    "# Test`nHello World" | Out-File $testFile -Encoding UTF8

    try {
        pandoc $testFile -o $testOutput 2>&1 | Out-Null
        if (Test-Path $testOutput) {
            Write-Host "  基本轉換: 成功" -ForegroundColor Green
            Remove-Item $testFile, $testOutput -Force
        } else {
            Write-Host "  基本轉換: 失敗" -ForegroundColor Red
        }
    } catch {
        Write-Host "  基本轉換: 錯誤 - $_" -ForegroundColor Red
    }
} else {
    Write-Host "  跳過測試 (Pandoc 未安裝)"
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "    檢查完成" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
