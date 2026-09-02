# Acc. SE比對工具

Excel 業務負責人（Acc. SE）自動比對回填工具。取代 `SR104861_AL.xlsx` 現行用
XLOOKUP 巢狀公式做的「客戶 → Acc. SE」比對，改用統編 → 客戶名稱 → GroupName
模糊比對 → N/A 的三層優先序（都沒中才算無法匹配），並記錄每筆的命中路徑與
失敗原因，方便追蹤與人工複核。比對前會先套用異體字對照表（`config/char_variants.json`，
例如台/臺）正規化，避免同一個字的不同寫法造成漏配。

設計依據：`桌面\CR負責人比對工具_設計規劃書.md`

## 目前進度：Phase 1（CLI 核心邏輯）＋ Phase 2（Excel 輸出＋網頁介面）＋ Phase 3（易用性調整）

已完成：檔案讀取（含雙層標頭解析）、智慧欄位對應（含 mapping profile 記憶；目標檔完全
沒有 Acc. SE 欄時，自動選資料結束後緊接的空白欄，表頭自動寫入「Acc. SE」文字＋黃色底色）、
欄位格式輪廓分析、13 種異常模式清洗、Pass1-4 比對引擎（含異體字正規化、僅一字之差
輔助診斷）、統計摘要、單元測試；回填主檔輸出（除了上述新增空白欄的例外，只覆寫 Acc. SE
欄的值，其餘欄位/格式/儲存格顏色完全維持原始檔案不變）、異常明細表；兩種操作介面並存，
輸出的檔案內容完全一致：
- `cli.py`：命令列，輸出已回填主檔＋異常明細表（跟 GitHub Pages 版相同）＋一份逐列核對
  用 CSV（CLI 沒有網頁預覽表格可看，這份 CSV 補這個用途）
- `browser-app/`：純靜態網頁版，用 Pyodide（瀏覽器內的 Python 執行環境）直接
  執行 `core/` 既有邏輯，資料完全不經過任何伺服器，部署在 GitHub Pages，目前
  實際使用的介面

（原本還有一份本機 Flask 網頁版 `webapp/`，因為跟 GitHub Pages 版功能重複、
使用者實際上只用 GitHub Pages，2026-09-01 已移除，避免兩份介面各自維護
造成邏輯漂移。）

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
- `output/`（CLI 比對結果）與 `config/mapping_profiles/`（欄位對應記憶，可能含
  真實檔案結構特徵）同樣被排除在版控之外。
