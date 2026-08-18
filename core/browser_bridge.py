"""瀏覽器端（Pyodide）薄橋接層。

把 `webapp/app.py` 對應的三支路由邏輯改寫成「接受 JSON 字串、回傳 JSON 字串」的
函式，供 `browser-app/static/js/pyodide-runner.js` 透過 Pyodide 呼叫。刻意用單一
JSON 字串進、單一 JSON 字串出，避免處理 Pyodide 的 JS↔Python proxy 物件轉換細節。
不依賴 Flask、不管理 run_id/session——檔案路徑一律由呼叫端在 Pyodide 虛擬檔案系統
裡直接指定並傳入。
"""

from __future__ import annotations

import json

from core.excel_io import list_sheets
from core.excel_writer import write_primary_output, write_unmatched_report
from core.field_mapping import detect_mapping_candidates, save_profile
from core.matching import build_roster_index
from core.models import FieldMapping
from core.pipeline import load_roster_entries_with_mapping, run_target_matching_with_mapping
from core.stats import build_stats

PREVIEW_LIMIT = 200


def _ok(payload: dict) -> str:
    return json.dumps(payload, ensure_ascii=False)


def _err(message: str) -> str:
    return json.dumps({"error": message}, ensure_ascii=False)


def list_sheets_json(payload_json: str) -> str:
    payload = json.loads(payload_json)
    try:
        target_sheets = list_sheets(payload["target_path"])
        roster_sheets = list_sheets(payload["roster_path"])
    except Exception as exc:  # noqa: BLE001
        return _err(f"無法讀取 Excel 檔案，請確認檔案未損毀：{exc}")
    return _ok({"target_sheets": target_sheets, "roster_sheets": roster_sheets})


def _serialize_detection(detection) -> dict:
    return {
        "sheet_name": detection.sheet_name,
        "header_row": detection.header_row,
        "data_start_row": detection.data_start_row,
        "signature": detection.signature,
        "header_options": [{"text": text, "col": col} for text, col in detection.header_options],
        "best_guess": detection.best_guess,
        "saved_mapping": (
            {
                "tax_id_col": detection.saved_mapping.tax_id_col,
                "customer_name_col": detection.saved_mapping.customer_name_col,
                "group_name_col": detection.saved_mapping.group_name_col,
                "acc_se_col": detection.saved_mapping.acc_se_output_col,
            }
            if detection.saved_mapping
            else None
        ),
    }


def mapping_preview_json(payload_json: str) -> str:
    payload = json.loads(payload_json)
    try:
        target_detection = detect_mapping_candidates(
            payload["target_path"], payload["target_sheet"], is_double_header=False
        )
        roster_detection = detect_mapping_candidates(
            payload["roster_path"], payload["roster_sheet"], is_double_header=True
        )
    except Exception as exc:  # noqa: BLE001
        return _err(f"欄位偵測失敗：{exc}")
    return _ok(
        {
            "target": _serialize_detection(target_detection),
            "roster": _serialize_detection(roster_detection),
        }
    )


def _build_mapping(side: dict, is_double_header: bool) -> FieldMapping | None:
    required = ["sheet_name", "header_row", "data_start_row", "tax_id_col", "customer_name_col", "acc_se_col"]
    if any(side.get(key) in (None, "") for key in required):
        return None
    return FieldMapping(
        sheet_name=side["sheet_name"],
        header_row=int(side["header_row"]),
        data_start_row=int(side["data_start_row"]),
        tax_id_col=int(side["tax_id_col"]),
        customer_name_col=int(side["customer_name_col"]),
        acc_se_output_col=int(side["acc_se_col"]),
        group_name_col=int(side["group_name_col"]) if side.get("group_name_col") not in (None, "") else None,
        is_double_header=is_double_header,
    )


def run_json(payload_json: str) -> str:
    payload = json.loads(payload_json)
    target_path = payload["target_path"]
    roster_path = payload["roster_path"]
    target_form = payload.get("target") or {}
    roster_form = payload.get("roster") or {}
    formula_col_raw = payload.get("formula_col")
    output_dir = payload["output_dir"]
    source_stem = payload["source_stem"]

    target_mapping = _build_mapping(target_form, is_double_header=False)
    roster_mapping = _build_mapping(roster_form, is_double_header=True)
    if target_mapping is None:
        return _err("目標檔的統編／客戶名稱／Acc. SE 欄位對應尚未完整選擇")
    if roster_mapping is None:
        return _err("名冊檔的統編／客戶名稱／Acc. SE(Name) 欄位對應尚未完整選擇")

    if target_form.get("signature"):
        save_profile(target_form["signature"], target_mapping)
    if roster_form.get("signature"):
        save_profile(roster_form["signature"], roster_mapping)

    formula_col = int(formula_col_raw) if formula_col_raw not in (None, "") else None

    try:
        roster_entries = load_roster_entries_with_mapping(roster_path, roster_mapping)
        roster_index = build_roster_index(roster_entries)
        target_result = run_target_matching_with_mapping(
            target_path, target_mapping, roster_index, formula_col=formula_col
        )
    except Exception as exc:  # noqa: BLE001
        return _err(f"比對執行失敗，請確認欄位對應是否選對：{exc}")

    stats = build_stats(target_result.results, roster_index, target_result.formula_na_baseline)

    primary_path = write_primary_output(
        target_path, target_mapping, target_result.results, output_dir=output_dir, source_stem=source_stem
    )
    unmatched_path = write_unmatched_report(target_result.results, source_stem, output_dir=output_dir)

    preview_rows = [
        {
            "row_index": r.row_index,
            "customer_name": r.raw_customer_name,
            "tax_id_raw": r.raw_tax_id,
            "tax_id_clean": r.clean_tax_id,
            "matched_acc_se": r.matched_acc_se,
            "match_pass": r.match_pass.value,
            "anomalies": [a.value for a in r.anomalies],
        }
        for r in target_result.results[:PREVIEW_LIMIT]
    ]

    return _ok(
        {
            "stats": {
                "total": stats.total,
                "pass_counts": {p.value: c for p, c in stats.pass_counts.items()},
                "anomaly_counts": {a.value: c for a, c in stats.anomaly_counts.items()},
                "formula_na_baseline": stats.formula_na_baseline,
                "roster_missing_acc_se_count": stats.roster_missing_acc_se_count,
                "roster_group_name_skipped_count": stats.roster_group_name_skipped_count,
            },
            "preview_rows": preview_rows,
            "preview_truncated": len(target_result.results) > PREVIEW_LIMIT,
            "downloads": {"primary": str(primary_path), "unmatched": str(unmatched_path)},
        }
    )
