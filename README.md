# Acc. SE比對工具

Excel 業務負責人（Acc. SE）自動比對回填工具。取代 `SR104861_AL.xlsx` 現行用
XLOOKUP 巢狀公式做的「客戶 → Acc. SE」比對，改用統編 → 客戶名稱 → GroupName
模糊比對 → N/A 的四層優先序，並記錄每筆的命中路徑與失敗原因，方便追蹤與人工複核。

設計依據：`桌面\CR負責人比對工具_設計規劃書.md`

## 目前進度：Phase 1（CLI 核心邏輯）＋ Phase 2（Excel 美化輸出＋網頁介面）

已完成：檔案讀取（含雙層標頭解析）、智慧欄位對應（含 mapping profile 記憶）、
欄位格式輪廓分析、13 種異常模式清洗、Pass1-4 比對引擎、統計摘要、單元測試；
Excel 排版美化輸出（Match Status 欄、狀態底色、凍結窗格、自動欄寬列高）、異常
明細表、簡易網頁介面（`webapp/`，3 步驟：上傳→欄位對應確認→結果）。

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

# 啟動網頁介面（瀏覽器開 http://127.0.0.1:5000）
venv\Scripts\python -m webapp.app
```

## 注意事項

- 三份樣本 Excel 檔含真實客戶資料（公司名稱、統一編號、負責人姓名/email/手機），
  依 `.gitignore` 規則不進版控，也不會出現在此 repo 的任何 commit 中。
- `output/`（CLI 比對結果）、`webapp/uploads/`／`webapp/outputs/`（網頁介面的上傳
  檔與比對結果）與 `config/mapping_profiles/`（欄位對應記憶，可能含真實檔案結構
  特徵）同樣被排除在版控之外。
