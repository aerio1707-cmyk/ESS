"""模組 3：欄位資料格式自動偵測（Profiling）。

不假設固定規則（如統一補到 8 碼），而是先統計型態分布、長度分布、格式模式，
再依統計結果動態決定正規化規則（規劃書第四章）。
"""

from __future__ import annotations

import re

from core.models import ColumnProfile

PERSONAL_ID_RE = re.compile(r"^[A-Za-z][0-9]{9}$")


def stringify(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def profile_column(values: list[object]) -> ColumnProfile:
    profile = ColumnProfile(total=len(values))
    length_counts: dict[int, int] = {}

    for raw in values:
        text = stringify(raw)
        if not text:
            profile.empty += 1
            continue

        if isinstance(raw, (int, float)):
            profile.numeric_count += 1
        else:
            profile.text_count += 1

        length_counts[len(text)] = length_counts.get(len(text), 0) + 1

        if text.startswith("#"):
            profile.hash_prefix_count += 1
        if PERSONAL_ID_RE.match(text):
            profile.personal_id_count += 1

    profile.length_distribution = length_counts
    if length_counts:
        profile.mode_length = max(length_counts, key=lambda k: length_counts[k])
    return profile
