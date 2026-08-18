"""智慧欄位對應：關鍵字字典 + 使用者手動覆蓋雙軌制（規劃書 2.3），
並將確認過的對應存成 mapping profile，供下次同結構檔案自動套用（規劃書十三.5，
Phase 1 提前納入）。

`detect_mapping_candidates` / `finalize_mapping` 為非阻塞式偵測+確認流程，供網頁介面
（Phase 2）使用；`build_field_mapping` 是 CLI 用的一站式版本（信心不足時用 input() 詢問）。
"""

from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass, field
from pathlib import Path

from core.excel_io import (
    build_merged_value_map,
    detect_header_row,
    load_worksheet,
    parse_double_header,
    parse_single_header,
)
from core.models import FieldMapping

CONFIG_DIR = Path(__file__).resolve().parent.parent / "config"
KEYWORDS_PATH = CONFIG_DIR / "field_keywords.json"
PROFILE_DIR = CONFIG_DIR / "mapping_profiles"

FIELD_KEYS = ["tax_id", "customer_name", "group_name", "acc_se"]


def load_keywords() -> dict:
    with open(KEYWORDS_PATH, encoding="utf-8") as f:
        return json.load(f)


def resolve_field(
    header_map: dict[str, int], exact_keywords: list[str], fuzzy_keywords: list[str]
) -> tuple[int | None, list[tuple[str, int]]]:
    """完全比對優先於模糊比對（規劃書 2.3）。回傳 (欄號或None, 候選清單)。"""
    exact_matches = [
        (text, col) for text, col in header_map.items() if text.strip() in exact_keywords
    ]
    if len(exact_matches) == 1:
        return exact_matches[0][1], exact_matches
    if len(exact_matches) > 1:
        return None, exact_matches

    fuzzy_matches = [
        (text, col)
        for text, col in header_map.items()
        if any(kw in text for kw in fuzzy_keywords)
    ]
    if len(fuzzy_matches) == 1:
        return fuzzy_matches[0][1], fuzzy_matches
    return None, fuzzy_matches


def prompt_manual_choice(field_label: str, candidates: list[tuple[str, int]]) -> int:
    if not candidates:
        prompt = f"找不到「{field_label}」欄位，請手動輸入欄號（1-indexed）："
        return int(input(prompt).strip())
    print(f"「{field_label}」欄位偵測到多個候選，請選擇：")
    for i, (text, col) in enumerate(candidates, start=1):
        print(f"  [{i}] {text}（第 {col} 欄）")
    print("  [0] 都不是，手動輸入欄號")
    choice = input("請輸入選項編號：").strip()
    idx = int(choice)
    if idx == 0:
        return int(input("請輸入欄號（1-indexed）：").strip())
    return candidates[idx - 1][1]


def structure_signature(sheet_name: str, header_texts: list[str]) -> str:
    raw = sheet_name + "|" + "|".join(sorted(header_texts))
    return hashlib.md5(raw.encode("utf-8")).hexdigest()


def load_saved_profile(signature: str) -> dict | None:
    profile_path = PROFILE_DIR / f"{signature}.json"
    if profile_path.exists():
        with open(profile_path, encoding="utf-8") as f:
            return json.load(f)
    return None


def _mapping_from_saved(saved: dict) -> FieldMapping:
    return FieldMapping(
        sheet_name=saved["sheet_name"],
        header_row=saved["header_row"],
        data_start_row=saved["data_start_row"],
        tax_id_col=saved["tax_id_col"],
        customer_name_col=saved["customer_name_col"],
        acc_se_output_col=saved["acc_se_output_col"],
        group_name_col=saved.get("group_name_col"),
        is_double_header=saved["is_double_header"],
    )


def save_profile(signature: str, mapping: FieldMapping) -> None:
    PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    profile_path = PROFILE_DIR / f"{signature}.json"
    data = {
        "sheet_name": mapping.sheet_name,
        "header_row": mapping.header_row,
        "data_start_row": mapping.data_start_row,
        "tax_id_col": mapping.tax_id_col,
        "customer_name_col": mapping.customer_name_col,
        "acc_se_output_col": mapping.acc_se_output_col,
        "group_name_col": mapping.group_name_col,
        "is_double_header": mapping.is_double_header,
    }
    with open(profile_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


@dataclass
class MappingDetection:
    """欄位偵測結果，供呼叫端（CLI 或網頁）決定如何確認/覆蓋後再呼叫 finalize_mapping。"""

    sheet_name: str
    header_row: int
    data_start_row: int
    is_double_header: bool
    signature: str
    header_options: list[tuple[str, int]]  # 全部標頭選項 (文字, 欄號)，供下拉選單使用
    best_guess: dict[str, int | None] = field(default_factory=dict)  # key in FIELD_KEYS
    candidates: dict[str, list[tuple[str, int]]] = field(default_factory=dict)
    saved_mapping: FieldMapping | None = None  # 若命中既有 profile，直接可用，無需使用者確認


def detect_mapping_candidates(
    path: str | Path, sheet_name: str, *, is_double_header: bool = False
) -> MappingDetection:
    """讀取檔案、偵測標頭、猜測各欄位最佳欄號，不阻塞、不寫入 profile。"""
    ws = load_worksheet(path, sheet_name)
    merged_map = build_merged_value_map(ws)
    keywords = load_keywords()

    header_row = detect_header_row(ws, merged_map)

    if is_double_header:
        sub_row = header_row + 1
        header_map = parse_double_header(ws, header_row, sub_row, merged_map)
        data_start_row = sub_row + 1
    else:
        header_map = parse_single_header(ws, header_row, merged_map)
        data_start_row = header_row + 1

    header_texts = list(header_map.keys())
    signature = structure_signature(sheet_name, header_texts)
    header_options = sorted(header_map.items(), key=lambda kv: kv[1])

    saved = load_saved_profile(signature)
    saved_mapping = _mapping_from_saved(saved) if saved else None

    best_guess: dict[str, int | None] = {}
    candidates: dict[str, list[tuple[str, int]]] = {}

    for key in ("tax_id", "customer_name"):
        col, cand = resolve_field(header_map, keywords[key]["exact"], keywords[key]["fuzzy"])
        best_guess[key] = col
        candidates[key] = cand

    if is_double_header:
        best_guess["acc_se"] = header_map.get("Acc. SE__Name")
        best_guess["group_name"] = header_map.get("Group Name")
        candidates["acc_se"] = []
        candidates["group_name"] = []
    else:
        col, cand = resolve_field(header_map, keywords["acc_se"]["exact"], keywords["acc_se"]["fuzzy"])
        best_guess["acc_se"] = col
        candidates["acc_se"] = cand
        best_guess["group_name"] = None
        candidates["group_name"] = []

    return MappingDetection(
        sheet_name=sheet_name,
        header_row=header_row,
        data_start_row=data_start_row,
        is_double_header=is_double_header,
        signature=signature,
        header_options=header_options,
        best_guess=best_guess,
        candidates=candidates,
        saved_mapping=saved_mapping,
    )


def finalize_mapping(detection: MappingDetection, choices: dict[str, int | None]) -> FieldMapping:
    """依使用者確認/覆蓋後的欄號組成最終 FieldMapping，並記住這次選擇（規劃書十三.5）。"""
    mapping = FieldMapping(
        sheet_name=detection.sheet_name,
        header_row=detection.header_row,
        data_start_row=detection.data_start_row,
        tax_id_col=choices.get("tax_id"),
        customer_name_col=choices.get("customer_name"),
        acc_se_output_col=choices.get("acc_se"),
        group_name_col=choices.get("group_name"),
        is_double_header=detection.is_double_header,
    )
    save_profile(detection.signature, mapping)
    return mapping


def build_field_mapping(
    path: str | Path,
    sheet_name: str,
    *,
    is_double_header: bool = False,
    interactive: bool = True,
) -> FieldMapping:
    """CLI 用一站式版本：偵測 → (信心不足時 input() 詢問) → 存 profile。"""
    detection = detect_mapping_candidates(path, sheet_name, is_double_header=is_double_header)
    if detection.saved_mapping:
        return detection.saved_mapping

    labels = {
        "tax_id": "統編",
        "customer_name": "客戶名稱",
        "group_name": "Group Name",
        "acc_se": "Acc. SE / Name（負責人姓名）" if is_double_header else "Acc. SE（回填目標欄）",
    }

    choices: dict[str, int | None] = {}
    for key in FIELD_KEYS:
        guess = detection.best_guess.get(key)
        if guess is not None:
            choices[key] = guess
            continue
        if key == "group_name" and not is_double_header:
            choices[key] = None
            continue
        if not interactive:
            choices[key] = None
            continue
        choices[key] = prompt_manual_choice(labels[key], detection.candidates.get(key, []))

    return finalize_mapping(detection, choices)
