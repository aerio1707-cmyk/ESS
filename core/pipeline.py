"""串起 excel_io → field_mapping → profiling → sanitize → matching → stats 的完整流程。

供 cli.py（單一目標檔）與 scripts/run_sample_validation.py（多檔驗證）共用。
"""

from __future__ import annotations

import csv
from dataclasses import dataclass
from datetime import date
from pathlib import Path

from core.excel_io import build_merged_value_map, cell_value, load_worksheet
from core.field_mapping import build_field_mapping
from core.matching import RosterEntry, RosterIndex, build_roster_index, match_target_row
from core.models import FieldMapping, MatchResult, MatchStats
from core.profiling import profile_column, stringify
from core.sanitize import clean_customer_name, clean_group_name, clean_tax_id, normalize_text
from core.stats import build_stats


def _iter_data_rows(ws, data_start_row: int):
    for row in range(data_start_row, ws.max_row + 1):
        yield row


def load_roster_entries(
    path: str | Path, sheet_name: str, interactive: bool = True
) -> tuple[list[RosterEntry], FieldMapping]:
    mapping = build_field_mapping(path, sheet_name, is_double_header=True, interactive=interactive)
    ws = load_worksheet(path, mapping.sheet_name)
    merged_map = build_merged_value_map(ws)

    raw_tax_ids = [
        cell_value(ws, row, mapping.tax_id_col, merged_map)
        for row in _iter_data_rows(ws, mapping.data_start_row)
    ]
    tax_id_profile = profile_column(raw_tax_ids)

    entries: list[RosterEntry] = []
    for row in _iter_data_rows(ws, mapping.data_start_row):
        raw_tax = cell_value(ws, row, mapping.tax_id_col, merged_map)
        raw_customer = cell_value(ws, row, mapping.customer_name_col, merged_map)
        raw_group = cell_value(ws, row, mapping.group_name_col, merged_map) if mapping.group_name_col else None
        raw_acc_se = cell_value(ws, row, mapping.acc_se_output_col, merged_map)

        clean_tax, _ = clean_tax_id(raw_tax, tax_id_profile)
        clean_customer, _, _ = clean_customer_name(raw_customer)
        clean_group = clean_group_name(raw_group)
        clean_acc_se = normalize_text(raw_acc_se)

        if not any([clean_tax, clean_customer, clean_group, clean_acc_se]):
            continue  # 完全空白列（Excel 匯出常見尾端空白列）

        entries.append(
            RosterEntry(
                row_index=row,
                tax_id=clean_tax,
                customer_name=clean_customer,
                group_name=clean_group,
                acc_se_name=clean_acc_se,
            )
        )
    return entries, mapping


@dataclass
class TargetLoadResult:
    results: list[MatchResult]
    mapping: FieldMapping
    formula_na_baseline: int | None


def run_target_matching(
    path: str | Path,
    sheet_name: str,
    roster_index: RosterIndex,
    formula_col: int | None = None,
    interactive: bool = True,
) -> TargetLoadResult:
    mapping = build_field_mapping(path, sheet_name, is_double_header=False, interactive=interactive)
    ws = load_worksheet(path, mapping.sheet_name)
    merged_map = build_merged_value_map(ws)

    raw_tax_ids = [
        cell_value(ws, row, mapping.tax_id_col, merged_map)
        for row in _iter_data_rows(ws, mapping.data_start_row)
    ]
    tax_id_profile = profile_column(raw_tax_ids)

    results: list[MatchResult] = []
    baseline_na = 0 if formula_col else None

    for row in _iter_data_rows(ws, mapping.data_start_row):
        raw_tax = cell_value(ws, row, mapping.tax_id_col, merged_map)
        raw_customer = cell_value(ws, row, mapping.customer_name_col, merged_map)

        if stringify(raw_tax) == "" and stringify(raw_customer) == "":
            continue  # 完全空白列

        clean_tax, tax_anomalies = clean_tax_id(raw_tax, tax_id_profile)
        clean_customer, main_name, customer_anomalies = clean_customer_name(raw_customer)

        result = match_target_row(
            row_index=row,
            raw_tax_id=stringify(raw_tax),
            raw_customer_name=stringify(raw_customer),
            tax_id_clean=clean_tax,
            customer_name_clean=clean_customer,
            customer_main_name_clean=main_name,
            roster=roster_index,
            pre_anomalies=tax_anomalies + customer_anomalies,
        )
        results.append(result)

        if formula_col:
            raw_formula = cell_value(ws, row, formula_col, merged_map)
            if normalize_text(raw_formula).upper() in ("N/A", ""):
                baseline_na += 1

    return TargetLoadResult(results=results, mapping=mapping, formula_na_baseline=baseline_na)


def export_review_csv(
    results: list[MatchResult], target_path: str | Path, output_dir: str | Path = "output"
) -> Path:
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    stem = Path(target_path).stem
    out_path = output_dir / f"{stem}_review_{date.today():%Y%m%d}.csv"

    with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(
            [
                "原始列號",
                "客戶名稱",
                "統編(原始值)",
                "統編(清洗後)",
                "回填Acc.SE",
                "比對狀態",
                "異常標記",
                "重複候選",
            ]
        )
        for r in results:
            writer.writerow(
                [
                    r.row_index,
                    r.raw_customer_name,
                    r.raw_tax_id,
                    r.clean_tax_id,
                    r.matched_acc_se,
                    r.match_pass.value,
                    "; ".join(a.value for a in r.anomalies),
                    "; ".join(r.duplicate_candidates),
                ]
            )
    return out_path


def run_full_pipeline(
    target_path: str | Path,
    target_sheet: str,
    roster_path: str | Path,
    roster_sheet: str,
    formula_col: int | None = None,
    output_dir: str | Path = "output",
    interactive: bool = True,
) -> tuple[MatchStats, list[MatchResult], Path]:
    roster_entries, _roster_mapping = load_roster_entries(roster_path, roster_sheet, interactive=interactive)
    roster_index = build_roster_index(roster_entries)

    target_result = run_target_matching(
        target_path, target_sheet, roster_index, formula_col=formula_col, interactive=interactive
    )
    stats = build_stats(target_result.results, roster_index, target_result.formula_na_baseline)
    review_path = export_review_csv(target_result.results, target_path, output_dir)
    return stats, target_result.results, review_path
