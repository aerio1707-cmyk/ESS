"""Acc. SE比對工具 — Phase 1 CLI 進入點。

範例：
    python cli.py --target "目標檔.xlsx" --target-sheet "工作表1" \
        --roster "名冊.xlsx" --roster-sheet "Special Acc list" \
        [--formula-col G] [--non-interactive]
"""

from __future__ import annotations

import argparse
import sys

from openpyxl.utils import column_index_from_string

from core.pipeline import run_full_pipeline


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Excel 業務負責人（Acc. SE）自動比對回填工具")
    parser.add_argument("--target", required=True, help="目標客戶清單檔路徑")
    parser.add_argument("--target-sheet", required=True, help="目標檔工作表名稱")
    parser.add_argument("--roster", required=True, help="負責人名冊檔路徑")
    parser.add_argument("--roster-sheet", required=True, help="名冊檔工作表名稱（雙層標頭）")
    parser.add_argument(
        "--formula-col", default=None, help="目標檔現行公式欄（如 G），用於與新工具結果對照，可省略"
    )
    parser.add_argument("--output-dir", default="output", help="核對用 CSV 輸出目錄，預設 output/")
    parser.add_argument(
        "--non-interactive", action="store_true", help="欄位對應信心不足時不互動詢問，直接留空"
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    formula_col = column_index_from_string(args.formula_col) if args.formula_col else None

    result = run_full_pipeline(
        target_path=args.target,
        target_sheet=args.target_sheet,
        roster_path=args.roster,
        roster_sheet=args.roster_sheet,
        formula_col=formula_col,
        output_dir=args.output_dir,
        interactive=not args.non_interactive,
    )

    print(result.stats.render())
    print()
    print(f"已回填主檔：{result.primary_output_path}")
    print(f"異常明細表：{result.unmatched_report_path}")
    print(f"逐列核對用 CSV：{result.review_csv_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
