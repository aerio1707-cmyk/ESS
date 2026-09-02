"""模組 2：資料清洗與異常防錯機制（規劃書 3.1、3.2）。"""

from __future__ import annotations

import functools
import json
import re
import unicodedata
from pathlib import Path

from core.models import AnomalyCode
from core.profiling import PERSONAL_ID_RE, ColumnProfile, stringify

INVISIBLE_CHARS = ("​", "‌", "‍", "﻿")  # 零寬空格/連接符、BOM
FORMULA_ERROR_TOKENS = {"#N/A", "#VALUE!", "#REF!", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!"}
ANNOTATION_RE = re.compile(r"^(.*?)[（(]")

CHAR_VARIANTS_PATH = Path(__file__).resolve().parent.parent / "config" / "char_variants.json"


@functools.lru_cache(maxsize=1)
def load_char_variants() -> dict[str, str]:
    """異體字對照表：{異體字: 正規化後的字}，只給比對用，不影響顯示用的原始文字。

    目前只收錄已在實際資料中確認過的組合（見 config/char_variants.json），刻意不預先
    塞一堆猜測的異體字組合——寧可漏抓、靠 find_near_miss_candidates 輔助人工逐步發現，
    也不要猜錯造成誤配。
    """
    with open(CHAR_VARIANTS_PATH, encoding="utf-8") as f:
        return json.load(f)


def variant_normalize(text: str) -> str:
    """把文字裡的異體字轉成對照表裡的正規化字，僅供比對使用（見 load_char_variants）。"""
    table = load_char_variants()
    if not table:
        return text
    return "".join(table.get(ch, ch) for ch in text)

# 補零規則僅套用於「全數字且長度為眾數-1」且落在 5~7 碼區間（規劃書 3.2 #1、#9）
PADDABLE_LENGTH_RANGE = range(5, 8)


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value)
    text = unicodedata.normalize("NFKC", text)
    for ch in INVISIBLE_CHARS:
        text = text.replace(ch, "")
    return text.strip()


def is_formula_error(raw_value: object) -> bool:
    return isinstance(raw_value, str) and raw_value.strip().upper() in FORMULA_ERROR_TOKENS


def clean_tax_id(raw_value: object, profile: ColumnProfile) -> tuple[str, list[AnomalyCode]]:
    """回傳 (清洗後比對用值, 觸發的異常代碼列表)。原始值由呼叫端另外保留。"""
    anomalies: list[AnomalyCode] = []

    if is_formula_error(raw_value):
        anomalies.append(AnomalyCode.FORMULA_ERROR)
        return "", anomalies

    text = normalize_text(raw_value)
    if not text:
        anomalies.append(AnomalyCode.MISSING_TAX_ID)
        return "", anomalies

    if PERSONAL_ID_RE.match(text):
        anomalies.append(AnomalyCode.PERSONAL_ID)
        return text, anomalies

    if not text.isdigit():
        # 含字母前綴的暫編代碼或海外客戶代碼（如 #N24020003、FUHK350000）
        anomalies.append(AnomalyCode.NON_STANDARD_TAX_ID_CODE)
        return text, anomalies

    if (
        profile.mode_length is not None
        and len(text) == profile.mode_length - 1
        and len(text) in PADDABLE_LENGTH_RANGE
    ):
        padded = text.zfill(profile.mode_length)
        anomalies.append(AnomalyCode.LEADING_ZERO_PADDED)
        return padded, anomalies

    return text, anomalies


def clean_customer_name(raw_value: object) -> tuple[str, str | None, list[AnomalyCode]]:
    """回傳 (清洗後全名, 抽取出的主公司名稱或None, 異常代碼列表)。

    主公司名稱：用於規劃書 3.2 #6「客戶名稱帶附註細節」的二次比對（括號前主名稱）。
    """
    anomalies: list[AnomalyCode] = []
    text = normalize_text(raw_value)
    if not text:
        anomalies.append(AnomalyCode.MISSING_CUSTOMER_NAME)
        return "", None, anomalies

    match = ANNOTATION_RE.match(text)
    main_name = None
    if match:
        candidate = match.group(1).strip()
        if candidate and candidate != text:
            main_name = candidate
            anomalies.append(AnomalyCode.NAME_ANNOTATION_STRIPPED)

    return text, main_name, anomalies


def clean_group_name(raw_value: object) -> str:
    return normalize_text(raw_value)
