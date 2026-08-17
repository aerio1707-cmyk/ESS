"""Phase 1 驗證腳本：用真實樣本檔跑一次完整流程，印出統計摘要供人工抽樣核對。

刻意不在程式碼中寫死真實檔名/工作表名（避免含真實資料的識別資訊進入 public repo），
所有路徑一律由執行時的參數傳入。

範例（於 CR/ 目錄下、啟用 venv 後執行）：
    python scripts/run_sample_validation.py \
        --roster "SA_LIST-20260720.xlsx" --roster-sheet "Special Acc list" \
        --target "20260522 NH F5_CR List_NH468-AR-ASR-2.xlsx|工作表1" \
        --target "SR104861_AL.xlsx|(有函式) SR104861_AL|G"
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl.utils import column_index_from_string  # noqa: E402

from core.pipeline import run_full_pipeline  # noqa: E402


def parse_target_spec(spec: str) -> tuple[str, str, int | None]:
    """"path|sheet[|formula_col]" -> (path, sheet, formula_col_index或None)"""
    parts = spec.split("|")
    if len(parts) not in (2, 3):
        raise ValueError(f"--target 格式錯誤，需為 path|sheet 或 path|sheet|formula_col：{spec}")
    path, sheet = parts[0], parts[1]
    formula_col = column_index_from_string(parts[2]) if len(parts) == 3 and parts[2] else None
    return path, sheet, formula_col


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="用真實樣本檔跑完整比對流程並輸出統計摘要")
    parser.add_argument("--roster", required=True)
    parser.add_argument("--roster-sheet", required=True)
    parser.add_argument(
        "--target",
        action="append",
        required=True,
        dest="targets",
        help='"path|sheet" 或 "path|sheet|formula_col"，可重複指定多個目標檔',
    )
    parser.add_argument("--output-dir", default="output")
    parser.add_argument("--non-interactive", action="store_true")
    return parser.parse_args()


def main() -> int:
    args = parse_args()

    for spec in args.targets:
        path, sheet, formula_col = parse_target_spec(spec)
        print("=" * 70)
        print(f"目標檔：{path}（{sheet}）")
        print("=" * 70)

        stats, _results, review_path = run_full_pipeline(
            target_path=path,
            target_sheet=sheet,
            roster_path=args.roster,
            roster_sheet=args.roster_sheet,
            formula_col=formula_col,
            output_dir=args.output_dir,
            interactive=not args.non_interactive,
        )

        print(stats.render())
        print()
        print(f"核對用明細已輸出：{review_path}")
        print()

    return 0


if __name__ == "__main__":
    sys.exit(main())
