"""模組 5：結果統計摘要組裝（規劃書 6.1）。"""

from __future__ import annotations

from core.matching import RosterIndex
from core.models import MatchResult, MatchStats


def build_stats(
    results: list[MatchResult],
    roster: RosterIndex,
    formula_na_baseline: int | None = None,
) -> MatchStats:
    stats = MatchStats(
        formula_na_baseline=formula_na_baseline,
        roster_missing_acc_se_count=roster.missing_acc_se_count,
        roster_group_name_skipped_count=roster.group_name_skipped_count,
    )
    for result in results:
        stats.add_result(result)
    return stats
