"""單元測試：core/excel_writer.py（合成假資料，不含真實客戶資訊）。"""

import openpyxl

from core.excel_writer import (
    MATCH_STATUS_HEADER,
    write_primary_output,
    write_unmatched_report,
)
from core.models import AnomalyCode, FieldMapping, MatchPass, MatchResult


def _make_target_workbook(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["客戶名稱", "統一編號", "Acc. SE"])
    ws.append(["測試公司A", "10000001", "=VLOOKUP(B2,X:X,1,0)"])  # 模擬既有公式欄
    ws.append(["測試公司B", "10000002", None])
    ws.append(["測試公司C", "10000003", None])
    wb.save(path)


def _results():
    return [
        MatchResult(
            row_index=2,
            raw_tax_id="10000001",
            raw_customer_name="測試公司A",
            clean_tax_id="10000001",
            clean_customer_name="測試公司A",
            matched_acc_se="王小明",
            match_pass=MatchPass.TAX_ID,
        ),
        MatchResult(
            row_index=3,
            raw_tax_id="10000002",
            raw_customer_name="測試公司B",
            clean_tax_id="10000002",
            clean_customer_name="測試公司B",
            matched_acc_se="甲負責人",
            match_pass=MatchPass.TAX_ID,
            anomalies=[AnomalyCode.ROSTER_DUPLICATE_CONFLICT],
            duplicate_candidates=["甲負責人", "乙負責人"],
        ),
        MatchResult(
            row_index=4,
            raw_tax_id="10000003",
            raw_customer_name="測試公司C",
            clean_tax_id="10000003",
            clean_customer_name="測試公司C",
            matched_acc_se="N/A",
            match_pass=MatchPass.UNMATCHED,
        ),
    ]


def test_write_primary_output_adds_match_status_column_and_overwrites_formula(tmp_path):
    target_path = tmp_path / "target.xlsx"
    _make_target_workbook(target_path)

    mapping = FieldMapping(
        sheet_name="Sheet",
        header_row=1,
        data_start_row=2,
        tax_id_col=2,
        customer_name_col=1,
        acc_se_output_col=3,
    )

    out_path = write_primary_output(target_path, mapping, _results(), output_dir=tmp_path / "out")

    wb = openpyxl.load_workbook(out_path)
    ws = wb.active

    assert ws.cell(row=1, column=4).value == MATCH_STATUS_HEADER
    # 原本是公式，回填後應為純值，不再是公式字串
    assert ws.cell(row=2, column=3).value == "王小明"
    assert ws.cell(row=3, column=3).value == "甲負責人"
    assert ws.cell(row=4, column=3).value == "N/A"
    assert ws.cell(row=4, column=4).value == MatchPass.UNMATCHED.value
    assert ws.freeze_panes == "A2"


def test_write_primary_output_colors_unmatched_and_duplicate_rows(tmp_path):
    target_path = tmp_path / "target.xlsx"
    _make_target_workbook(target_path)
    mapping = FieldMapping(
        sheet_name="Sheet", header_row=1, data_start_row=2, tax_id_col=2, customer_name_col=1, acc_se_output_col=3
    )

    out_path = write_primary_output(target_path, mapping, _results(), output_dir=tmp_path / "out")
    wb = openpyxl.load_workbook(out_path)
    ws = wb.active

    assert ws.cell(row=4, column=3).fill.fgColor.rgb == "00F8CBAD"  # 無法匹配＝淡紅
    assert ws.cell(row=3, column=3).fill.fgColor.rgb == "00FFE699"  # 名冊重複衝突＝淡黃
    assert ws.cell(row=2, column=3).fill.fgColor.rgb in (None, "00000000")


def test_write_unmatched_report_only_includes_unmatched_and_duplicate(tmp_path):
    out_path = write_unmatched_report(_results(), "target", output_dir=tmp_path / "out")
    wb = openpyxl.load_workbook(out_path)
    ws = wb.active

    rows = [ws.cell(row=r, column=2).value for r in range(2, ws.max_row + 1)]
    assert "測試公司B" in rows  # 重複衝突
    assert "測試公司C" in rows  # 無法匹配
    assert "測試公司A" not in rows  # 正常命中，不應出現
