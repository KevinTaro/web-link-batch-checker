# 網頁批次抓取與網址檢查專案

---

## 專案簡介

本專案可批次抓取多個網頁的所有連結、標題、重要資訊，並自動匯出為 CSV 及 Excel 檔案。
後續可自動檢查所有連結是否有效，並將檢查結果回寫到 Excel。
支援 CLI 與 GUI 介面，流程可重複執行，並即時顯示進度與結果。

## 特色與技術優化

-   全面採用 Python async/await 非同步技術，所有網頁抓取與網址檢查皆同時並行，效能大幅提升。
-   慢速網站自動延長 timeout 並多次重試，最大化檢查準確率。
-   支援大量網址高效處理，充分發揮現代電腦 I/O 與多核心效能。
-   GUI 支援一鍵完整流程、即時訊息、顯示最新產出檔案位置。
-   所有訊息 thread-safe，進度與狀態即時更新。

## 主要程式

-   `web_grab_tool.py`：輸入多個網址，非同步批次抓取網頁連結，匯出 CSV 與合併 Excel。
-   `xlsx_address_check_tool.py`：非同步批次檢查 Excel 檔案中的所有網址，回寫檢查結果。
-   `web_link_checker_gui.py`：Tkinter GUI，支援一鍵完整流程、重複執行、即時訊息、顯示最新產出檔案。

## 使用方式

### CLI 執行

1. 執行 `xlsx_address_check_tool.py` 或 `web_grab_tool.py`，依指示輸入網址（或選擇要檢查的 Excel 檔案）。
2. 所有結果會自動分類存放於 `output` 資料夾：
    - `csv/`：每個網頁的原始連結資料 CSV
    - `xlsx/`：合併後的 Excel 檔案
    - `checked/`：檢查結果 Excel 檔案

### GUI 執行

1. 執行 `web_link_checker_gui.py`，可用視窗操作：
    - 輸入網址 → 批次抓取
    - 選擇 Excel 檔案 → 批次檢查
    - 一鍵執行完整流程（抓取+檢查）
    - 顯示最新產出檔案位置
2. 所有進度與結果即時顯示於視窗。

## 安裝需求

請先安裝必要套件：

```bash
pip install pandas aiohttp openpyxl beautifulsoup4 requests
```

## 依賴套件

-   pandas
-   aiohttp
-   openpyxl
-   beautifulsoup4
-   requests

## 執行方式

-   CLI：直接執行 `xlsx_address_check_tool.py` 或 `web_grab_tool.py`，依指示操作。
-   GUI：執行 `web_link_checker_gui.py`，依視窗操作流程。

## 注意事項

-   請勿將 output 資料夾加入版本控制（已在 .gitignore 設定）。
-   程式支援中文檔名與內容。
-   若有特殊需求可自行擴充。

## 貢獻方式

歡迎提出 Pull Request 或 Issue，協助改進本工具。

## 問題回報

如有 bug 或建議，請至 [Issues](https://github.com/KevinTaro/web-link-batch-checker/issues) 回報。

## 聯絡方式

如需協助或有其他問題，請於 Issues 留言。

## 授權條款

本專案採用 MIT License 授權。詳見 [LICENSE](LICENSE) 檔案。

> 本專案公開於 GitHub Public Repository，請勿上傳敏感或私人資料。

## 最新優化與穩定性說明

-   網路請求（抓取/檢查）全面採用高彈性重試機制，timeout 會自動延長至 120 秒，重試次數最多 5 次。
-   特殊處理 [WinError 64] 指定的網路名稱無法使用，遇到此錯誤會自動重試並記錄。
-   降低同時執行的網址數（最大 8），大幅減少端口負載與 timeout 誤判率，提升穩定性。
-   所有異常皆有詳細 log 記錄，方便追蹤與除錯。
-   速度略降但穩定性顯著提升，適合大量網址或不穩定網路環境。

## 參數建議

-   若遇大量 timeout，可適度調高重試次數或延長 timeout，並減少同時執行數（limit）。
-   相關參數可於 `web_grab_tool.py` 及 `xlsx_address_check_tool.py` 內調整。
