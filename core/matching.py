"""模組 4：多層級核心匹配邏輯（規劃書第五章）。

Pass1 統編精準比對 → Pass2 客戶名稱精準比對 → Pass3 GroupName 模糊比對
→ Pass4 無法匹配（填 "N/A"）。純函式設計，不依賴 Excel，方便單元測試。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from core.models import AnomalyCode, MatchPass, MatchResult
from core.sanitize import variant_normalize

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
    by_customer_name: dict[str, list[RosterEntry]] = field(default_factory=dict)  # key 為異體字正規化後的客戶名稱
    group_entries: list[RosterEntry] = field(default_factory=list)
    group_entries_normalized: list[str] = field(default_factory=list)  # 與 group_entries 一一對應的異體字正規化結果，避免 Pass3 重複計算
    customer_names_by_length: dict[int, list[str]] = field(default_factory=dict)  # 去重後的原始客戶名稱，供近似一字之差偵測用
    group_names_by_length: dict[int, list[str]] = field(default_factory=dict)  # 去重後的原始GroupName，供近似一字之差偵測用
    missing_acc_se_count: int = 0  # 規劃書 3.2 #11：名冊本身資料缺漏
    group_name_skipped_count: int = 0  # 規劃書 3.2 #13：GroupName 過短已跳過


def build_roster_index(entries: list[RosterEntry]) -> RosterIndex:
    index = RosterIndex()
    seen_customer_names: set[str] = set()
    seen_group_names: set[str] = set()
    for e in entries:
        if not e.acc_se_name:
            index.missing_acc_se_count += 1
            continue
        if e.tax_id:
            index.by_tax_id.setdefault(e.tax_id, []).append(e)
        if e.customer_name:
            index.by_customer_name.setdefault(variant_normalize(e.customer_name), []).append(e)
            if e.customer_name not in seen_customer_names:
                seen_customer_names.add(e.customer_name)
                index.customer_names_by_length.setdefault(len(e.customer_name), []).append(e.customer_name)
        if e.group_name:
            if group_name_passes_threshold(e.group_name):
                index.group_entries.append(e)
                index.group_entries_normalized.append(variant_normalize(e.group_name))
                # 近似一字之差候選只從「本來就有資格參與Pass3」的GroupName裡找，避免過短字串
                # （例如單一個字）跟任何文字都能算「只差一個字」，找出一堆沒有意義的雜訊候選。
                if e.group_name not in seen_group_names:
                    seen_group_names.add(e.group_name)
                    index.group_names_by_length.setdefault(len(e.group_name), []).append(e.group_name)
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
        near_miss_customer_name: str | None = None,
        near_miss_group_name: str | None = None,
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
            near_miss_customer_name=near_miss_customer_name,
            near_miss_group_name=near_miss_group_name,
        )

    # Pass 1：統編精準比對
    if tax_id_clean and tax_id_clean in roster.by_tax_id:
        entries = roster.by_tax_id[tax_id_clean]
        candidates = _distinct_names(entries)
        extra = [AnomalyCode.ROSTER_DUPLICATE_CONFLICT] if len(candidates) > 1 else []
        secondary = None
        for name_candidate in filter(None, [customer_name_clean, customer_main_name_clean]):
            key = variant_normalize(name_candidate)
            if key in roster.by_customer_name:
                alt = _distinct_names(roster.by_customer_name[key])[0]
                if alt != candidates[0]:
                    secondary = alt
                break
        return base_result(candidates[0], MatchPass.TAX_ID, extra, candidates if extra else [], secondary)

    # Pass 2：客戶名稱精準比對（含括號前主名稱二次比對，規劃書 3.2 #6；比對前先套異體字正規化）
    for name_candidate in filter(None, [customer_name_clean, customer_main_name_clean]):
        key = variant_normalize(name_candidate)
        if key in roster.by_customer_name:
            entries = roster.by_customer_name[key]
            candidates = _distinct_names(entries)
            has_duplicate = len(candidates) > 1
            extra = [AnomalyCode.ROSTER_DUPLICATE_CONFLICT] if has_duplicate else []
            if any(e.customer_name != name_candidate for e in entries):
                extra = extra + [AnomalyCode.VARIANT_CHAR_NORMALIZED]
            return base_result(candidates[0], MatchPass.CUSTOMER_NAME, extra, candidates if has_duplicate else [], None)

    # Pass 3：GroupName 模糊比對（名冊 GroupName 是否為客戶名稱的子字串）
    # 先試字面比對（沿用原本邏輯，成本低）；完全沒中才退而用異體字正規化重試一次。
    names_to_check = [n for n in [customer_name_clean, customer_main_name_clean] if n]
    literal_matches = [e for e in roster.group_entries if any(e.group_name in name for name in names_to_check)]
    if literal_matches:
        matched_entries, via_variant = literal_matches, False
    else:
        normalized_names_to_check = [variant_normalize(n) for n in names_to_check]
        matched_entries = [
            e
            for e, norm_group in zip(roster.group_entries, roster.group_entries_normalized)
            if any(norm_group in norm_name for norm_name in normalized_names_to_check)
        ]
        via_variant = bool(matched_entries)
    if matched_entries:
        candidates = _distinct_names(matched_entries)
        has_duplicate = len(candidates) > 1
        extra = [AnomalyCode.ROSTER_DUPLICATE_CONFLICT] if has_duplicate else []
        if via_variant:
            extra = extra + [AnomalyCode.VARIANT_CHAR_NORMALIZED]
        return base_result(candidates[0], MatchPass.GROUP_NAME, extra, candidates if has_duplicate else [], None)

    # Pass 4：無法匹配（提供「僅一字之差」候選，協助人工判斷是否為異體字對照表尚未收錄的組合）
    near_customer, near_group = (
        find_near_miss_candidates(customer_name_clean, roster) if customer_name_clean else (None, None)
    )
    return base_result("N/A", MatchPass.UNMATCHED, [], [], None, near_customer, near_group)


def find_near_miss_candidates(name: str, roster: RosterIndex) -> tuple[str | None, str | None]:
    """僅供無法匹配時使用的輔助診斷：在名冊「客戶名稱」／「Group Name」裡找出跟 name 恰好
    只差一個字元的候選，協助人工判斷是否為異體字對照表（config/char_variants.json）尚未
    收錄的組合。純粹是給人看的線索，不會被當作比對成功採信、不影響正式比對結果。

    先找客戶名稱候選；若已找到就不再做成本較高的 GroupName 子字串滑動比對——同一筆資料的
    問題根源通常只有一個，客戶名稱候選已足夠當作人工複核的線索。
    """
    customer_candidate = _find_single_char_diff(name, roster.customer_names_by_length.get(len(name), []))
    if customer_candidate is not None:
        return customer_candidate, None
    group_candidate = _find_single_char_diff_substring(name, roster.group_names_by_length)
    return None, group_candidate


def _find_single_char_diff(name: str, candidates: list[str]) -> str | None:
    for candidate in candidates:
        if candidate != name and _hamming_distance_is_one(name, candidate):
            return candidate
    return None


def _find_single_char_diff_substring(name: str, group_names_by_length: dict[int, list[str]]) -> str | None:
    """把 name 依名冊裡實際出現過的 GroupName 長度切出等長片段，找出恰好差一個字元的候選。"""
    for length, group_names in group_names_by_length.items():
        if length > len(name):
            continue
        for start in range(len(name) - length + 1):
            window = name[start : start + length]
            for group_name in group_names:
                if window != group_name and _hamming_distance_is_one(window, group_name):
                    return group_name
    return None


def _hamming_distance_is_one(a: str, b: str) -> bool:
    if len(a) != len(b):
        return False
    diff = 0
    for ca, cb in zip(a, b):
        if ca != cb:
            diff += 1
            if diff > 1:
                return False
    return diff == 1
