"""模組 4：多層級核心匹配邏輯（規劃書第五章）。

Pass1 統編精準比對 → Pass2 客戶名稱精準比對 → Pass3 GroupName 模糊比對
→ Pass4 無法匹配（填 "N/A"）。純函式設計，不依賴 Excel，方便單元測試。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from core.models import AnomalyCode, MatchPass, MatchResult

# GroupName 最短比對長度門檻（規劃書 3.2 #13）：
# ≥3 個中文字，或 ≥4 個英數字，才允許進入 Pass 3 模糊比對
GROUP_NAME_MIN_CJK_CHARS = 3
GROUP_NAME_MIN_ALNUM_CHARS = 4

CJK_RE = re.compile(r"[一-鿿]")
ALNUM_RE = re.compile(r"[A-Za-z0-9]")


def group_name_passes_threshold(group_name: str) -> bool:
    if not group_name:
        return False
    cjk_count = len(CJK_RE.findall(group_name))
    if cjk_count >= GROUP_NAME_MIN_CJK_CHARS:
        return True
    alnum_count = len(ALNUM_RE.findall(group_name))
    return alnum_count >= GROUP_NAME_MIN_ALNUM_CHARS


@dataclass
class RosterEntry:
    row_index: int
    tax_id: str
    customer_name: str
    group_name: str
    acc_se_name: str


@dataclass
class RosterIndex:
    by_tax_id: dict[str, list[RosterEntry]] = field(default_factory=dict)
    by_customer_name: dict[str, list[RosterEntry]] = field(default_factory=dict)
    group_entries: list[RosterEntry] = field(default_factory=list)
    missing_acc_se_count: int = 0  # 規劃書 3.2 #11：名冊本身資料缺漏
    group_name_skipped_count: int = 0  # 規劃書 3.2 #13：GroupName 過短已跳過


def build_roster_index(entries: list[RosterEntry]) -> RosterIndex:
    index = RosterIndex()
    for e in entries:
        if not e.acc_se_name:
            index.missing_acc_se_count += 1
            continue
        if e.tax_id:
            index.by_tax_id.setdefault(e.tax_id, []).append(e)
        if e.customer_name:
            index.by_customer_name.setdefault(e.customer_name, []).append(e)
        if e.group_name:
            if group_name_passes_threshold(e.group_name):
                index.group_entries.append(e)
            else:
                index.group_name_skipped_count += 1
    return index


def _distinct_names(entries: list[RosterEntry]) -> list[str]:
    names: list[str] = []
    for e in entries:
        if e.acc_se_name not in names:
            names.append(e.acc_se_name)
    return names


def match_target_row(
    row_index: int,
    raw_tax_id: str,
    raw_customer_name: str,
    tax_id_clean: str,
    customer_name_clean: str,
    customer_main_name_clean: str | None,
    roster: RosterIndex,
    pre_anomalies: list[AnomalyCode],
) -> MatchResult:
    def base_result(
        acc_se: str,
        match_pass: MatchPass,
        extra_anomalies: list[AnomalyCode],
        duplicate_candidates: list[str],
        secondary_match: str | None,
    ) -> MatchResult:
        return MatchResult(
            row_index=row_index,
            raw_tax_id=raw_tax_id,
            raw_customer_name=raw_customer_name,
            clean_tax_id=tax_id_clean,
            clean_customer_name=customer_name_clean,
            matched_acc_se=acc_se,
            match_pass=match_pass,
            anomalies=pre_anomalies + extra_anomalies,
            secondary_match=secondary_match,
            duplicate_candidates=duplicate_candidates,
        )

    # Pass 1：統編精準比對
    if tax_id_clean and tax_id_clean in roster.by_tax_id:
        entries = roster.by_tax_id[tax_id_clean]
        candidates = _distinct_names(entries)
        extra = [AnomalyCode.ROSTER_DUPLICATE_CONFLICT] if len(candidates) > 1 else []
        secondary = None
        for name_candidate in filter(None, [customer_name_clean, customer_main_name_clean]):
            if name_candidate in roster.by_customer_name:
                alt = _distinct_names(roster.by_customer_name[name_candidate])[0]
                if alt != candidates[0]:
                    secondary = alt
                break
        return base_result(candidates[0], MatchPass.TAX_ID, extra, candidates if extra else [], secondary)

    # Pass 2：客戶名稱精準比對（含括號前主名稱二次比對，規劃書 3.2 #6）
    for name_candidate in filter(None, [customer_name_clean, customer_main_name_clean]):
        if name_candidate in roster.by_customer_name:
            entries = roster.by_customer_name[name_candidate]
            candidates = _distinct_names(entries)
            extra = [AnomalyCode.ROSTER_DUPLICATE_CONFLICT] if len(candidates) > 1 else []
            return base_result(candidates[0], MatchPass.CUSTOMER_NAME, extra, candidates if extra else [], None)

    # Pass 3：GroupName 模糊比對（名冊 GroupName 是否為客戶名稱的子字串）
    names_to_check = [n for n in [customer_name_clean, customer_main_name_clean] if n]
    matched_entries = [
        e
        for e in roster.group_entries
        if any(e.group_name in name for name in names_to_check)
    ]
    if matched_entries:
        candidates = _distinct_names(matched_entries)
        extra = [AnomalyCode.ROSTER_DUPLICATE_CONFLICT] if len(candidates) > 1 else []
        return base_result(candidates[0], MatchPass.GROUP_NAME, extra, candidates if extra else [], None)

    # Pass 4：無法匹配
    return base_result("N/A", MatchPass.UNMATCHED, [], [], None)
