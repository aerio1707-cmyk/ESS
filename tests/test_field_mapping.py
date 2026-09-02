"""單元測試：core/field_mapping.py 的欄位偵測，重點是「目標檔沒有 Acc. SE 欄時自動
選用資料結束後緊接的空白欄」這條 2026-09-01 新規則（合成假資料，不含真實客戶資訊）。
"""

import openpyxl

from core.field_mapping import detect_mapping_candidates


def _write_target(path, headers):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "工作表1"
    ws.append(headers)
    ws.append(["測試公司A", "10000001"] + [None] * (len(headers) - 2))
    wb.save(path)


def test_falls_back_to_blank_column_when_no_acc_se_column_at_all(tmp_path):
    path = tmp_path / "target.xlsx"
    _write_target(path, ["客戶名稱", "統一編號", "circuit no"])

    detection = detect_mapping_candidates(path, "工作表1", is_double_header=False)

    assert detection.acc_se_new_column == 4
    assert detection.best_guess["acc_se"] == 4


def test_no_fallback_when_acc_se_column_exists(tmp_path):
    path = tmp_path / "target.xlsx"
    _write_target(path, ["客戶名稱", "統一編號", "Acc. SE"])

    detection = detect_mapping_candidates(path, "工作表1", is_double_header=False)

    assert detection.acc_se_new_column is None
    assert detection.best_guess["acc_se"] == 3


def test_no_fallback_when_acc_se_candidates_are_ambiguous(tmp_path):
    """兩個欄位都符合 Acc. SE 關鍵字（例如「Acc. SE」跟「負責人」），代表欄位其實存在，
    只是不確定選哪個，這種情況要讓使用者手動選，不能悄悄改成自動新增空白欄。
    """
    path = tmp_path / "target.xlsx"
    _write_target(path, ["客戶名稱", "統一編號", "Acc. SE", "負責人"])

    detection = detect_mapping_candidates(path, "工作表1", is_double_header=False)

    assert detection.acc_se_new_column is None
    assert detection.best_guess["acc_se"] is None
    assert len(detection.candidates["acc_se"]) == 2


def test_no_fallback_applied_for_roster_double_header(tmp_path):
    """名冊（雙層標頭）本來就必須有 Acc. SE / Name 欄，這個新規則不套用在名冊身上。"""
    path = tmp_path / "roster.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Special Acc list"
    ws.append(["統編", "客戶", "Group Name", ""])
    ws.append(["", "", "", ""])
    ws.append(["10000001", "測試公司A", "測試集團", ""])
    wb.save(path)

    detection = detect_mapping_candidates(path, "Special Acc list", is_double_header=True)

    assert detection.acc_se_new_column is None
    assert detection.best_guess["acc_se"] is None
