"""本機單人小工具的簡易 run_id 暫存管理。

每次上傳建立一個 uuid4 資料夾，上傳檔案存 uploads/<run_id>/，比對輸出存
outputs/<run_id>/。單人本機工具，不做過期清理機制（Phase 2 範圍內先不做）。
"""

from __future__ import annotations

import json
import uuid
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_ROOT = BASE_DIR / "uploads"
OUTPUT_ROOT = BASE_DIR / "outputs"


def new_run_id() -> str:
    return uuid.uuid4().hex


def upload_dir(run_id: str) -> Path:
    path = UPLOAD_ROOT / run_id
    path.mkdir(parents=True, exist_ok=True)
    return path


def output_dir(run_id: str) -> Path:
    path = OUTPUT_ROOT / run_id
    path.mkdir(parents=True, exist_ok=True)
    return path


def target_path(run_id: str) -> Path:
    return upload_dir(run_id) / "target.xlsx"


def roster_path(run_id: str) -> Path:
    return upload_dir(run_id) / "roster.xlsx"


def save_original_names(run_id: str, target_name: str, roster_name: str) -> None:
    """記住使用者原始上傳的檔名（實際檔案仍固定存成 target.xlsx/roster.xlsx），
    供輸出檔命名（規劃書 8.2：{原檔名}_已回填_{YYYYMMDD}.xlsx）使用。"""
    meta = {"target_stem": Path(target_name).stem, "roster_stem": Path(roster_name).stem}
    with open(upload_dir(run_id) / "meta.json", "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False)


def load_original_stems(run_id: str) -> dict:
    meta_path = upload_dir(run_id) / "meta.json"
    if not meta_path.exists():
        return {"target_stem": "target", "roster_stem": "roster"}
    with open(meta_path, encoding="utf-8") as f:
        return json.load(f)
