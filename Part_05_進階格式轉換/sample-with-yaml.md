---
title: Pandoc 進階教學
author: 學習者
date: 2025-01-01
abstract: 這是一份關於 Pandoc 進階功能的教學文件，包含 HTML 樣式、PDF 輸出和自訂模板。
keywords: [Pandoc, Markdown, HTML, PDF, 模板]
lang: zh-TW
toc: true
---

# 簡介

這份文件示範如何使用 YAML 元資料和自訂模板。

## 為什麼使用 YAML？

YAML 讓你可以在文件中定義：

- 標題和作者
- 日期和關鍵字
- 摘要內容
- 各種設定選項

## 程式碼範例

以下是一段 PowerShell 程式碼：

```powershell
# 批次轉換腳本
Get-ChildItem *.md | ForEach-Object {
    pandoc $_.Name -o ($_.BaseName + ".html") --standalone
}
```

## 表格範例

| 格式 | 副檔名 | 說明 |
|------|--------|------|
| Markdown | .md | 原始格式 |
| HTML | .html | 網頁格式 |
| Word | .docx | 文書處理 |
| PDF | .pdf | 可攜式文件 |

## 引用

> 模板讓文件製作更有效率，一次設定，永久使用。

## 結論

善用 YAML 元資料和自訂模板，可以大幅提升文件產出的效率和品質。

---

*本文件使用 Pandoc 產生*
