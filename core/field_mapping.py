"""智慧欄位對應：關鍵字字典 + 使用者手動覆蓋雙軌制（規劃書 2.3），
並將確認過的對應存成 mapping profile，供下次同結構檔案自動套用（規劃書十三.5，
Phase 1 提前納入）。
"""

from __future__ import annotations

import hashlib
import json
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
    print(f"  [0] 都不是，手動輸入欄號")
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


def build_field_mapping(
    path: str | Path,
    sheet_name: str,
    *,
    is_double_header: bool = False,
    interactive: bool = True,
) -> FieldMapping:
    """建立單一檔案的欄位對應。is_double_header=True 用於名冊檔（雙層標頭，規劃書 2.5）。"""
    ws = load_worksheet(path, sheet_name)
    merged_map = build_merged_value_map(ws)
    keywords = load_keywords()

    header_row = detect_header_row(ws, merged_map)

    if is_double_header:
        sub_row = header_row + 1
        header_map = parse_double_header(ws, header_row, sub_row, merged_map)
        data_start_row = sub_row + 1
        header_texts = list(header_map.keys())
    else:
        header_map = parse_single_header(ws, header_row, merged_map)
        data_start_row = header_row + 1
        header_texts = list(header_map.keys())

    signature = structure_signature(sheet_name, header_texts)
    saved = load_saved_profile(signature)
    if saved:
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

    def resolve_or_prompt(label: str, key: str) -> int | None:
        col, candidates = resolve_field(header_map, keywords[key]["exact"], keywords[key]["fuzzy"])
        if col is not None:
            return col
        if not interactive:
            return None
        return prompt_manual_choice(label, candidates)

    tax_id_col = resolve_or_prompt("統編", "tax_id")
    customer_name_col = resolve_or_prompt("客戶名稱", "customer_name")

    if is_double_header:
        # 名冊：鎖定 Acc. SE 群組的 Name 子欄，不誤抓 Backup Acc. SE / Manager（規劃書 2.5）
        acc_se_col = header_map.get("Acc. SE__Name")
        if acc_se_col is None and interactive:
            acc_se_col = prompt_manual_choice("Acc. SE / Name（負責人姓名）", [])
        group_name_col = header_map.get("Group Name")
        if group_name_col is None and interactive:
            group_name_col = prompt_manual_choice("Group Name", [])
    else:
        acc_se_col = resolve_or_prompt("Acc. SE（回填目標欄）", "acc_se")
        group_name_col = None

    mapping = FieldMapping(
        sheet_name=sheet_name,
        header_row=header_row,
        data_start_row=data_start_row,
        tax_id_col=tax_id_col,
        customer_name_col=customer_name_col,
        acc_se_output_col=acc_se_col,
        group_name_col=group_name_col,
        is_double_header=is_double_header,
    )
    save_profile(signature, mapping)
    return mapping
