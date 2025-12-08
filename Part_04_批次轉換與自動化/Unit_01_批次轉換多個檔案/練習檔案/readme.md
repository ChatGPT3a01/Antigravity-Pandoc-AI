# 批次轉換練習資料夾

請在這個資料夾中練習批次轉換指令。

## 練習步驟

1. 在此資料夾建立多個 .md 測試檔案
2. 使用批次轉換指令
3. 確認 .docx 檔案已產生

## 快速建立測試檔案

```powershell
1..5 | % { "# 測試 $_" | Out-File "test$_.md" -Encoding UTF8 }
```

## 批次轉換指令

```powershell
gci *.md | % { pandoc $_.Name -o ($_.BaseName + ".docx") }
```
