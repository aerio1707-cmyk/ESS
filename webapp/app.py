"""CR 負責人比對工具 — Phase 2 簡易網頁介面（本機單人小工具）。

3 步驟流程：① 上傳目標檔+名冊檔 ② 欄位對應確認 ③ 統計摘要＋核對預覽＋下載。
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from flask import Flask, jsonify, request, send_file  # noqa: E402

from core.excel_io import list_sheets  # noqa: E402
from core.excel_writer import write_primary_output, write_unmatched_report  # noqa: E402
from core.field_mapping import detect_mapping_candidates, save_profile  # noqa: E402
from core.matching import build_roster_index  # noqa: E402
from core.models import FieldMapping  # noqa: E402
from core.pipeline import load_roster_entries_with_mapping, run_target_matching_with_mapping  # noqa: E402
from core.stats import build_stats  # noqa: E402
from webapp import run_store  # noqa: E402

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB，足夠涵蓋目前樣本檔的數十倍

ALLOWED_EXT = {".xlsx"}
PREVIEW_LIMIT = 200


def _error(message: str, status: int = 400):
    return jsonify({"error": message}), status


def _validate_upload(file_storage) -> str | None:
    if file_storage is None or file_storage.filename == "":
        return "請選擇檔案"
    ext = Path(file_storage.filename).suffix.lower()
    if ext not in ALLOWED_EXT:
        return f"目前只支援 .xlsx 檔案（收到 {ext or '未知格式'}）"
    return None


@app.route("/")
def index():
    from flask import render_template

    return render_template("index.html")


@app.route("/api/upload", methods=["POST"])
def api_upload():
    target_file = request.files.get("target")
    roster_file = request.files.get("roster")

    for label, f in (("目標檔", target_file), ("名冊檔", roster_file)):
        err = _validate_upload(f)
        if err:
            return _error(f"{label}：{err}")

    run_id = run_store.new_run_id()
    target_file.save(run_store.target_path(run_id))
    roster_file.save(run_store.roster_path(run_id))
    run_store.save_original_names(run_id, target_file.filename, roster_file.filename)

    try:
        target_sheets = list_sheets(run_store.target_path(run_id))
        roster_sheets = list_sheets(run_store.roster_path(run_id))
    except Exception as exc:  # noqa: BLE001
        return _error(f"無法讀取 Excel 檔案，請確認檔案未損毀：{exc}", 400)

    return jsonify({"run_id": run_id, "target_sheets": target_sheets, "roster_sheets": roster_sheets})


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


@app.route("/api/mapping-preview", methods=["POST"])
def api_mapping_preview():
    data = request.get_json(force=True)
    run_id = data.get("run_id")
    target_sheet = data.get("target_sheet")
    roster_sheet = data.get("roster_sheet")
    if not (run_id and target_sheet and roster_sheet):
        return _error("缺少 run_id / target_sheet / roster_sheet")

    target_path = run_store.target_path(run_id)
    roster_path = run_store.roster_path(run_id)
    if not target_path.exists() or not roster_path.exists():
        return _error("找不到這個 run_id 對應的上傳檔案，請重新上傳", 404)

    try:
        target_detection = detect_mapping_candidates(target_path, target_sheet, is_double_header=False)
        roster_detection = detect_mapping_candidates(roster_path, roster_sheet, is_double_header=True)
    except Exception as exc:  # noqa: BLE001
        return _error(f"欄位偵測失敗：{exc}", 400)

    return jsonify(
        {
            "target": _serialize_detection(target_detection),
            "roster": _serialize_detection(roster_detection),
        }
    )


def _build_mapping_from_form(side: dict, is_double_header: bool) -> FieldMapping | None:
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


@app.route("/api/run", methods=["POST"])
def api_run():
    data = request.get_json(force=True)
    run_id = data.get("run_id")
    target_form = data.get("target") or {}
    roster_form = data.get("roster") or {}
    formula_col_raw = data.get("formula_col")

    if not run_id:
        return _error("缺少 run_id")

    target_path = run_store.target_path(run_id)
    roster_path = run_store.roster_path(run_id)
    if not target_path.exists() or not roster_path.exists():
        return _error("找不到這個 run_id 對應的上傳檔案，請重新上傳", 404)

    target_mapping = _build_mapping_from_form(target_form, is_double_header=False)
    roster_mapping = _build_mapping_from_form(roster_form, is_double_header=True)
    if target_mapping is None:
        return _error("目標檔的統編／客戶名稱／Acc. SE 欄位對應尚未完整選擇")
    if roster_mapping is None:
        return _error("名冊檔的統編／客戶名稱／Acc. SE(Name) 欄位對應尚未完整選擇")

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
        return _error(f"比對執行失敗，請確認欄位對應是否選對：{exc}", 400)

    stats = build_stats(target_result.results, roster_index, target_result.formula_na_baseline)

    stems = run_store.load_original_stems(run_id)
    out_dir = run_store.output_dir(run_id)
    primary_path = write_primary_output(
        target_path, target_mapping, target_result.results, output_dir=out_dir, source_stem=stems["target_stem"]
    )
    unmatched_path = write_unmatched_report(
        target_result.results, stems["target_stem"], output_dir=out_dir
    )

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

    return jsonify(
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
            "downloads": {
                "primary": primary_path.name,
                "unmatched": unmatched_path.name,
            },
        }
    )


@app.route("/api/download/<run_id>/<filename>")
def api_download(run_id: str, filename: str):
    if "/" in filename or "\\" in filename or ".." in filename:
        return _error("非法檔名", 400)
    out_dir = run_store.output_dir(run_id)
    file_path = out_dir / filename
    if not file_path.exists() or file_path.resolve().parent != out_dir.resolve():
        return _error("找不到檔案", 404)
    return send_file(file_path, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True, port=5000)
