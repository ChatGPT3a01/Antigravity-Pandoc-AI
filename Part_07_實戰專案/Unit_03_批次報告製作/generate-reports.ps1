# generate-reports.ps1
# 批次報告生成腳本
# 使用方式：.\generate-reports.ps1 -DataFile data.csv

param(
    [string]$DataFile = "sample-data.csv",
    [string]$Template = "report-template.md",
    [string]$OutputFolder = "output"
)

# 建立輸出資料夾
if (!(Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}

# 讀取資料和模板
$data = Import-Csv $DataFile -Encoding UTF8
$templateContent = Get-Content $Template -Raw -Encoding UTF8
$date = Get-Date -Format "yyyy/MM/dd"
$generated = Get-Date -Format "yyyy-MM-dd HH:mm"

Write-Host "📊 開始生成報告..." -ForegroundColor Cyan
Write-Host "   資料檔案: $DataFile" -ForegroundColor Gray
Write-Host "   模板: $Template" -ForegroundColor Gray
Write-Host "   輸出目錄: $OutputFolder" -ForegroundColor Gray
Write-Host ""

$count = 0

foreach ($row in $data) {
    # 替換模板變數
    $report = $templateContent
    $report = $report -replace '\{\{name\}\}', $row.name
    $report = $report -replace '\{\{department\}\}', $row.department
    $report = $report -replace '\{\{completed\}\}', $row.completed
    $report = $report -replace '\{\{pending\}\}', $row.pending
    $report = $report -replace '\{\{highlight\}\}', $row.highlight
    $report = $report -replace '\{\{date\}\}', $date
    $report = $report -replace '\{\{generated\}\}', $generated

    # 輸出 Markdown
    $mdFile = "$OutputFolder\report-$($row.name).md"
    $report | Out-File $mdFile -Encoding UTF8

    # 轉換成 Word
    $docxFile = "$OutputFolder\report-$($row.name).docx"
    pandoc $mdFile -o $docxFile

    Write-Host "✅ $($row.name)" -ForegroundColor Green
    $count++
}

Write-Host ""
Write-Host "📊 完成！共生成 $count 份報告" -ForegroundColor Cyan
Write-Host "📁 輸出位置: $OutputFolder\" -ForegroundColor Cyan
