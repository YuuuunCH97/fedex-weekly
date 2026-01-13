# fedex-weekly
抓取"FedEx package and air freight fuel surcharge rates"

自動化抓取 **FedEx Ground** 每週燃油附加費率 (Fuel Surcharge) 所設計。

FedEx 官網具有防爬蟲機制，程式使用 **Selenium** 模擬真實瀏覽器行為進行抓取，並自動將抓取到的最新費率更新至本地端的 Excel 運算表中。

##  主要功能
1. **自動抓取費率**：
   - 訪問 FedEx 官網，繞過 "System Down" 防火牆。
   - 使用正規表達式 (Regex) 精準提取當週最新的 `FedEx Ground` 費率 (例如: `21.25%`)。

2. **更新主計算表 (Fee Calculator)**：
   - 自動開啟指定的 `.xlsx` 檔案（例如 `123456.xlsx`）。
   - 鎖定特定分頁 (`FedEx Fee Calculator`)。
   - 自動將 **K 欄 (第 11 欄)** 的所有公式更新為最新的 `VLOOKUP` 設定。
   - *公式範例：* `=VLOOKUP(21.25%,$Q$1:$T$37,4,0)`

3. **歷史紀錄保存 (Log)**：
   - 每次執行成功後，會將「執行時間」與「抓取到的費率」寫入 `FedEx_Update_Log.xlsx`，方便追蹤歷史價格走勢。

## 🛠️ 環境需求
* **作業系統**：Windows 10 / 11
* **程式語言**：Python 3.x
* **瀏覽器**：Google Chrome (必須安裝)
* **Python 套件**：
  * `selenium` (瀏覽器自動化)
  * `webdriver-manager` (自動管理 Chrome 驅動)
  * `openpyxl` (讀寫 Excel)
  * `pandas` (資料處理)

## 🚀 安裝與設定

### 1. 安裝必要套件
在終端機 (Terminal) 執行以下指令：
```bash
pip install selenium webdriver-manager openpyxl pandas
