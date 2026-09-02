"""單元測試：用合成假資料（不含真實客戶資訊）驗證清洗規則與 Pass1-4 比對邏輯。"""

from core.matching import (
    RosterEntry,
    build_roster_index,
    group_name_passes_threshold,
    match_target_row,
)
from core.models import AnomalyCode, MatchPass
from core.profiling import profile_column
from core.sanitize import clean_customer_name, clean_tax_id, is_formula_error, normalize_text


# ---------------------------------------------------------------------------
# profiling / sanitize
# ---------------------------------------------------------------------------


def test_profile_column_mode_length():
    profile = profile_column(["12345678", "23456789", 7654321, None, "  "])
    assert profile.total == 5
    assert profile.empty == 2
    assert profile.mode_length == 8  # 兩筆8碼 vs 一筆7碼


def test_clean_tax_id_pads_leading_zero_when_short_by_one():
    profile = profile_column(["12345678", "23456789", 7724424])
    clean, anomalies = clean_tax_id(7724424, profile)
    assert clean == "07724424"
    assert AnomalyCode.LEADING_ZERO_PADDED in anomalies


def test_clean_tax_id_does_not_pad_when_length_matches_mode():
    profile = profile_column(["12345678", "23456789"])
    clean, anomalies = clean_tax_id("23456789", profile)
    assert clean == "23456789"
    assert anomalies == []


def test_clean_tax_id_personal_id_not_padded():
    profile = profile_column(["12345678"])
    clean, anomalies = clean_tax_id("A104250525", profile)
    assert clean == "A104250525"
    assert AnomalyCode.PERSONAL_ID in anomalies


def test_clean_tax_id_hash_prefix_code_kept_as_is():
    profile = profile_column(["12345678"])
    clean, anomalies = clean_tax_id("#N24020003", profile)
    assert clean == "#N24020003"
    assert AnomalyCode.NON_STANDARD_TAX_ID_CODE in anomalies


def test_clean_tax_id_missing_value():
    profile = profile_column(["12345678"])
    clean, anomalies = clean_tax_id(None, profile)
    assert clean == ""
    assert AnomalyCode.MISSING_TAX_ID in anomalies


def test_formula_error_detected_and_treated_as_missing():
    profile = profile_column(["12345678"])
    assert is_formula_error("#N/A")
    clean, anomalies = clean_tax_id("#N/A", profile)
    assert clean == ""
    assert AnomalyCode.FORMULA_ERROR in anomalies


def test_clean_customer_name_extracts_main_name_before_annotation():
    # NFKC 正規化會把全形括號轉半形，因此清洗後的全名括號也會是半形
    clean, main_name, anomalies = clean_customer_name("測試股份有限公司（內湖辦公室）")
    assert clean == "測試股份有限公司(內湖辦公室)"
    assert main_name == "測試股份有限公司"
    assert AnomalyCode.NAME_ANNOTATION_STRIPPED in anomalies


def test_clean_customer_name_without_annotation_has_no_main_name():
    clean, main_name, anomalies = clean_customer_name("測試股份有限公司")
    assert clean == "測試股份有限公司"
    assert main_name is None
    assert anomalies == []


def test_normalize_text_fullwidth_and_invisible_chars():
    assert normalize_text("Ａ１２３　​") == "A123"


# ---------------------------------------------------------------------------
# group name threshold
# ---------------------------------------------------------------------------


def test_group_name_threshold_short_chinese_rejected():
    assert group_name_passes_threshold("大同") is False


def test_group_name_threshold_three_chinese_accepted():
    assert group_name_passes_threshold("大同集團") is True


def test_group_name_threshold_alnum_needs_four():
    assert group_name_passes_threshold("ABC") is False
    assert group_name_passes_threshold("ABCD") is True


# ---------------------------------------------------------------------------
# matching passes
# ---------------------------------------------------------------------------


def _roster():
    entries = [
        RosterEntry(row_index=3, tax_id="70461075", customer_name="測試股份有限公司", group_name="測試集團", acc_se_name="王小明"),
        RosterEntry(row_index=4, tax_id="", customer_name="全家國際餐飲@FamilyMart", group_name="全家國際餐飲@FamilyMart", acc_se_name="徐世武"),
        RosterEntry(row_index=9, tax_id="", customer_name="", group_name="全家國際餐飲", acc_se_name="徐世武"),
        RosterEntry(row_index=5, tax_id="99999999", customer_name="重複衝突公司", group_name="", acc_se_name="甲負責人"),
        RosterEntry(row_index=6, tax_id="99999999", customer_name="重複衝突公司", group_name="", acc_se_name="乙負責人"),
        RosterEntry(row_index=7, tax_id="88888888", customer_name="缺漏公司", group_name="", acc_se_name=""),
        RosterEntry(row_index=8, tax_id="", customer_name="", group_name="同", acc_se_name="丙負責人"),  # 過短 GroupName
    ]
    return build_roster_index(entries)


def test_pass1_tax_id_exact_match():
    roster = _roster()
    result = match_target_row(
        row_index=2,
        raw_tax_id="70461075",
        raw_customer_name="測試股份有限公司",
        tax_id_clean="70461075",
        customer_name_clean="測試股份有限公司",
        customer_main_name_clean=None,
        roster=roster,
        pre_anomalies=[],
    )
    assert result.match_pass == MatchPass.TAX_ID
    assert result.matched_acc_se == "王小明"


def test_pass2_customer_name_exact_match_when_tax_id_missing():
    roster = _roster()
    result = match_target_row(
        row_index=2,
        raw_tax_id="",
        raw_customer_name="全家國際餐飲@FamilyMart",
        tax_id_clean="",
        customer_name_clean="全家國際餐飲@FamilyMart",
        customer_main_name_clean=None,
        roster=roster,
        pre_anomalies=[],
    )
    assert result.match_pass == MatchPass.CUSTOMER_NAME
    assert result.matched_acc_se == "徐世武"


def test_pass3_group_name_substring_match():
    roster = _roster()
    result = match_target_row(
        row_index=2,
        raw_tax_id="",
        raw_customer_name="全家國際餐飲股份有限公司",
        tax_id_clean="",
        customer_name_clean="全家國際餐飲股份有限公司",
        customer_main_name_clean=None,
        roster=roster,
        pre_anomalies=[],
    )
    assert result.match_pass == MatchPass.GROUP_NAME
    assert result.matched_acc_se == "徐世武"


def test_pass4_unmatched_fills_na():
    roster = _roster()
    result = match_target_row(
        row_index=2,
        raw_tax_id="00000001",
        raw_customer_name="完全查無此公司",
        tax_id_clean="00000001",
        customer_name_clean="完全查無此公司",
        customer_main_name_clean=None,
        roster=roster,
        pre_anomalies=[],
    )
    assert result.match_pass == MatchPass.UNMATCHED
    assert result.matched_acc_se == "N/A"


def test_roster_duplicate_conflict_takes_first_and_records_candidates():
    roster = _roster()
    result = match_target_row(
        row_index=2,
        raw_tax_id="99999999",
        raw_customer_name="重複衝突公司",
        tax_id_clean="99999999",
        customer_name_clean="重複衝突公司",
        customer_main_name_clean=None,
        roster=roster,
        pre_anomalies=[],
    )
    assert result.match_pass == MatchPass.TAX_ID
    assert result.matched_acc_se == "甲負責人"
    assert AnomalyCode.ROSTER_DUPLICATE_CONFLICT in result.anomalies
    assert result.duplicate_candidates == ["甲負責人", "乙負責人"]


def test_roster_missing_acc_se_counted_separately():
    roster = _roster()
    assert roster.missing_acc_se_count == 1


def test_roster_group_name_too_short_skipped():
    roster = _roster()
    assert roster.group_name_skipped_count == 1
    assert all(e.group_name != "同" for e in roster.group_entries)


# ---------------------------------------------------------------------------
# 異體字正規化（台/臺，config/char_variants.json）與「僅一字之差」輔助診斷
# ---------------------------------------------------------------------------


def test_pass2_matches_via_variant_char_normalization():
    entries = [
        RosterEntry(
            row_index=1, tax_id="", customer_name="臺灣恩益禧股份有限公司", group_name="臺灣恩益禧", acc_se_name="施柏屹"
        ),
    ]
    roster = build_roster_index(entries)
    result = match_target_row(
        row_index=2,
        raw_tax_id="",
        raw_customer_name="台灣恩益禧股份有限公司",
        tax_id_clean="",
        customer_name_clean="台灣恩益禧股份有限公司",
        customer_main_name_clean=None,
        roster=roster,
        pre_anomalies=[],
    )
    assert result.match_pass == MatchPass.CUSTOMER_NAME
    assert result.matched_acc_se == "施柏屹"
    assert AnomalyCode.VARIANT_CHAR_NORMALIZED in result.anomalies


def test_pass3_matches_via_variant_char_normalization_when_literal_substring_fails():
    entries = [
        RosterEntry(row_index=1, tax_id="", customer_name="", group_name="臺灣恩益禧", acc_se_name="施柏屹"),
    ]
    roster = build_roster_index(entries)
    result = match_target_row(
        row_index=2,
        raw_tax_id="",
        raw_customer_name="台灣恩益禧股份有限公司",
        tax_id_clean="",
        customer_name_clean="台灣恩益禧股份有限公司",
        customer_main_name_clean=None,
        roster=roster,
        pre_anomalies=[],
    )
    assert result.match_pass == MatchPass.GROUP_NAME
    assert result.matched_acc_se == "施柏屹"
    assert AnomalyCode.VARIANT_CHAR_NORMALIZED in result.anomalies


def test_near_miss_customer_name_candidate_when_unmatched():
    """揚/陽不在異體字對照表裡，比對不到，但應該被找出來當作候選讓人工判斷要不要收錄。"""
    entries = [
        RosterEntry(row_index=1, tax_id="", customer_name="民揚科技股份有限公司", group_name="", acc_se_name="王小明"),
    ]
    roster = build_roster_index(entries)
    result = match_target_row(
        row_index=2,
        raw_tax_id="",
        raw_customer_name="民陽科技股份有限公司",
        tax_id_clean="",
        customer_name_clean="民陽科技股份有限公司",
        customer_main_name_clean=None,
        roster=roster,
        pre_anomalies=[],
    )
    assert result.match_pass == MatchPass.UNMATCHED
    assert result.near_miss_customer_name == "民揚科技股份有限公司"
    assert result.near_miss_group_name is None


def test_near_miss_group_name_candidate_when_no_customer_name_candidate():
    entries = [
        RosterEntry(row_index=1, tax_id="", customer_name="", group_name="鑫和衛星電視", acc_se_name="楊敏裕"),
    ]
    roster = build_roster_index(entries)
    result = match_target_row(
        row_index=2,
        raw_tax_id="",
        raw_customer_name="鑫和衛星電祝股份有限公司",
        tax_id_clean="",
        customer_name_clean="鑫和衛星電祝股份有限公司",
        customer_main_name_clean=None,
        roster=roster,
        pre_anomalies=[],
    )
    assert result.match_pass == MatchPass.UNMATCHED
    assert result.near_miss_customer_name is None
    assert result.near_miss_group_name == "鑫和衛星電視"


def test_near_miss_none_when_nothing_close_enough():
    roster = _roster()
    result = match_target_row(
        row_index=2,
        raw_tax_id="",
        raw_customer_name="完全無關的公司名稱",
        tax_id_clean="",
        customer_name_clean="完全無關的公司名稱",
        customer_main_name_clean=None,
        roster=roster,
        pre_anomalies=[],
    )
    assert result.match_pass == MatchPass.UNMATCHED
    assert result.near_miss_customer_name is None
    assert result.near_miss_group_name is None
