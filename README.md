# Acc. SE比對工具

Excel 業務負責人（Acc. SE）自動比對回填工具。取代 `SR104861_AL.xlsx` 現行用
XLOOKUP 巢狀公式做的「客戶 → Acc. SE」比對，改用統編 → 客戶名稱 → GroupName
模糊比對 → N/A 的四層優先序，並記錄每筆的命中路徑與失敗原因，方便追蹤與人工複核。

設計依據：`桌面\CR負責人比對工具_設計規劃書.md`

## 目前進度：Phase 1（CLI 核心邏輯）＋ Phase 2（Excel 輸出＋網頁介面，含 GitHub Pages 靜態版）

已完成：檔案讀取（含雙層標頭解析）、智慧欄位對應（含 mapping profile 記憶）、
欄位格式輪廓分析、13 種異常模式清洗、Pass1-4 比對引擎、統計摘要、單元測試；
回填主檔輸出（只覆寫 Acc. SE 欄的值，其餘欄位/格式/儲存格顏色完全維持原始檔案
不變）、異常明細表；三種操作介面並存：
- `cli.py`：命令列
- `webapp/`：本機 Flask 網頁介面（3 步驟：上傳→欄位對應確認→結果）
- `browser-app/`：純靜態網頁版，用 Pyodide（瀏覽器內的 Python 執行環境）直接
  執行 `core/` 既有邏輯，資料完全不經過任何伺服器，可部署到 GitHub Pages

尚未開始：Phase 3（易用性提升，視需求評估）。

## 使用方式

```bash
# 建立虛擬環境並安裝套件
python -m venv venv
venv\Scripts\pip install -r requirements.txt

# 跑單一目標檔
venv\Scripts\python cli.py --target "目標檔.xlsx" --target-sheet "工作表1" \
  --roster "名冊.xlsx" --roster-sheet "Special Acc list"

# 跑單元測試
venv\Scripts\python -m pytest tests/ -q

# 用真實樣本檔驗證（樣本檔本身不進版控，需自行放在專案根目錄）
venv\Scripts\python scripts/run_sample_validation.py \
  --roster "名冊.xlsx" --roster-sheet "Special Acc list" \
  --target "目標檔1.xlsx|工作表1" \
  --target "目標檔2.xlsx|工作表2|G"

# 啟動本機網頁介面（瀏覽器開 http://127.0.0.1:5000）
venv\Scripts\python -m webapp.app

# 本機測試 GitHub Pages 靜態版（一定要從 CR/ 根目錄啟動，讓相對路徑抓得到 core/）
venv\Scripts\python -m http.server 8000
# 瀏覽器開 http://127.0.0.1:8000/browser-app/
```

## 部署 GitHub Pages 靜態版

`browser-app/` 是純靜態網站，用 Pyodide 在瀏覽器裡直接執行 `core/` 的 Python
程式碼（`fetch` 原始 .py 檔文字寫進 Pyodide 虛擬檔案系統），不需要任何伺服器，
所以可以直接放上 GitHub Pages：

1. GitHub repo 頁面 → **Settings → Pages**
2. **Build and deployment → Source** 選 **Deploy from a branch**
3. **Branch** 選 `main`，資料夾選 **`/ (root)`**（不是 `/docs`，因為 `core/` 要能被
   `browser-app/` 用相對路徑 `../core/*.py` 抓到，兩者都要在同一個發布範圍內）
4. 存檔後幾分鐘內會給一個網址，格式類似
   `https://aerio1707-cmyk.github.io/ESS/browser-app/`
5. 之後每次 push 到 `main` 都會自動重新部署，不用手動操作

## 注意事項

- 三份樣本 Excel 檔含真實客戶資料（公司名稱、統一編號、負責人姓名/email/手機），
  依 `.gitignore` 規則不進版控，也不會出現在此 repo 的任何 commit 中。
- `output/`（CLI 比對結果）、`webapp/uploads/`／`webapp/outputs/`（網頁介面的上傳
  檔與比對結果）與 `config/mapping_profiles/`（欄位對應記憶，可能含真實檔案結構
  特徵）同樣被排除在版控之外。
