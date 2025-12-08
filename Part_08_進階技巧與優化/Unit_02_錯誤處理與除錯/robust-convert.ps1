# robust-convert.ps1
# 具備完整錯誤處理的轉換腳本
# 使用方式：.\robust-convert.ps1 -InputFile "doc.md" -OutputFormat "docx"

param(
    [Parameter(Mandatory=$true)]
    [string]$InputFile,
    [string]$OutputFormat = "docx",
    [string]$LogFile = "convert.log"
)

# 日誌函數
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"

    # 輸出到螢幕
    switch ($Level) {
        "ERROR" { Write-Host $logEntry -ForegroundColor Red }
        "WARN"  { Write-Host $logEntry -ForegroundColor Yellow }
        "OK"    { Write-Host $logEntry -ForegroundColor Green }
        default { Write-Host $logEntry }
    }

    # 寫入日誌檔
    $logEntry | Out-File $LogFile -Append -Encoding UTF8
}

Write-Log "=== 開始轉換作業 ==="

# 步驟 1: 檢查輸入檔案
Write-Log "檢查輸入檔案: $InputFile"
if (!(Test-Path $InputFile)) {
    Write-Log "檔案不存在: $InputFile" "ERROR"
    exit 1
}

$fileInfo = Get-Item $InputFile
Write-Log "檔案大小: $($fileInfo.Length) bytes"
Write-Log "修改時間: $($fileInfo.LastWriteTime)"

# 步驟 2: 檢查檔案編碼
Write-Log "檢查檔案編碼..."
try {
    $content = Get-Content $InputFile -Raw -Encoding UTF8 -ErrorAction Stop
    Write-Log "檔案編碼正常" "OK"
} catch {
    Write-Log "讀取檔案時發生錯誤: $_" "ERROR"
    exit 1
}

# 步驟 3: 準備輸出路徑
$outputFile = [System.IO.Path]::ChangeExtension($InputFile, $OutputFormat)
$outputDir = Split-Path $outputFile -Parent

if ($outputDir -and !(Test-Path $outputDir)) {
    Write-Log "建立輸出目錄: $outputDir"
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

# 步驟 4: 執行轉換
Write-Log "開始轉換: $InputFile -> $outputFile"

try {
    $result = pandoc $InputFile -o $outputFile 2>&1

    # 檢查是否有錯誤輸出
    if ($LASTEXITCODE -ne 0) {
        throw "Pandoc 返回錯誤碼: $LASTEXITCODE`n$result"
    }

    # 驗證輸出檔案
    if (!(Test-Path $outputFile)) {
        throw "轉換完成但找不到輸出檔案"
    }

    $outputSize = (Get-Item $outputFile).Length
    if ($outputSize -eq 0) {
        Write-Log "警告: 輸出檔案大小為 0" "WARN"
    } else {
        Write-Log "轉換成功! 輸出大小: $outputSize bytes" "OK"
    }

} catch {
    Write-Log "轉換失敗: $_" "ERROR"

    # 清理可能的不完整檔案
    if (Test-Path $outputFile) {
        Remove-Item $outputFile -Force
        Write-Log "已刪除不完整的輸出檔案"
    }

    exit 1
}

Write-Log "=== 轉換作業完成 ==="
