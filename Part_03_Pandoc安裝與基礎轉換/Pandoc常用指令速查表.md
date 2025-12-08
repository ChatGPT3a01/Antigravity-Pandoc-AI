# Pandoc 常用指令速查表

## 基礎轉換指令

| 用途 | 指令 |
|------|------|
| MD → DOCX | `pandoc input.md -o output.docx` |
| MD → HTML | `pandoc input.md -o output.html` |
| MD → PDF | `pandoc input.md -o output.pdf` |
| DOCX → MD | `pandoc input.docx -o output.md` |
| HTML → MD | `pandoc input.html -o output.md` |

## 常用參數

| 參數 | 說明 | 範例 |
|------|------|------|
| `-o` | 指定輸出檔案 | `-o output.docx` |
| `-f` | 指定輸入格式 | `-f markdown` |
| `-t` | 指定輸出格式 | `-t docx` |
| `--standalone` | 產生完整文件 | `--standalone` |
| `--toc` | 加入目錄 | `--toc` |

## 批次轉換（PowerShell）

```powershell
# 轉換資料夾內所有 .md 檔案為 .docx
Get-ChildItem *.md | ForEach-Object {
    pandoc $_.Name -o ($_.BaseName + ".docx")
}
```

## 常見錯誤排除

| 錯誤訊息 | 原因 | 解決方法 |
|----------|------|----------|
| 'pandoc' 不是命令 | 未安裝或 PATH 未設定 | 重新安裝或重開終端機 |
| Could not find file | 檔案路徑錯誤 | 檢查檔名和路徑 |
| Unknown output format | 副檔名錯誤 | 確認輸出格式正確 |

## 在 Antigravity 中使用

1. 開啟終端機：`` Ctrl+` ``
2. 切換到檔案目錄：`cd 資料夾路徑`
3. 執行轉換指令

---

> 💡 **提示**：善用 Tab 鍵自動完成檔名，避免打字錯誤！
