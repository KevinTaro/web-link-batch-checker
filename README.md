# 網頁批次抓取與網址檢查專案

## 專案簡介

本專案可批次抓取多個網頁的所有連結、標題、重要資訊，並自動匯出為 CSV 及 Excel 檔案。後續可自動檢查所有連結是否有效，並將檢查結果回寫到 Excel。

## 資料夾結構

```
output/
  ├── csv/      # 每個網頁的原始連結資料 CSV
  ├── xlsx/     # 合併後的 Excel 檔案
  └── checked/  # 檢查結果 Excel 檔案
```

## 主要程式

-   `網頁抓取工具.py`：輸入多個網址，批次抓取網頁連結，匯出 CSV 與合併 Excel。
-   `xlsx網址檢查工具.py`：自動檢查 Excel 檔案中的所有網址，回寫檢查結果。

## 使用方式

1. 執行 `xlsx網址檢查工具.py`，依指示輸入網址（或直接執行，會自動串聯抓取與檢查流程）。
2. 所有結果會自動分類存放於 `output` 資料夾。

## 依賴套件

-   requests
-   beautifulsoup4
-   pandas
-   openpyxl

請先安裝依賴：

```
pip install requests beautifulsoup4 pandas openpyxl
```

## 注意事項

-   請勿將 output 資料夾加入版本控制（已在 .gitignore 設定）。
-   程式支援中文檔名與內容。
-   若有特殊需求可自行擴充。
