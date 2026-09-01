from dataclasses import dataclass, field
from enum import Enum


class MatchPass(str, Enum):
    TAX_ID = "統編比對成功"
    CUSTOMER_NAME = "名稱比對成功"
    GROUP_NAME = "GroupName模糊比對成功"
    UNMATCHED = "無法匹配"


class AnomalyCode(str, Enum):
    LEADING_ZERO_PADDED = "前導零補正"
    PERSONAL_ID = "個資身分證字號"
    FORMULA_ERROR = "來源公式錯誤"
    NAME_ANNOTATION_STRIPPED = "客戶名稱含附註已抽取主名稱"
    MISSING_TAX_ID = "統編欄缺失"
    MISSING_CUSTOMER_NAME = "客戶名稱欄缺失"
    NON_STANDARD_TAX_ID_CODE = "統編非標準格式代碼"
    ROSTER_DUPLICATE_CONFLICT = "名冊重複衝突"
    ROSTER_DATA_MISSING = "名冊資料缺失"
    GROUP_NAME_TOO_SHORT = "GroupName過短已跳過模糊比對"


@dataclass
class FieldMapping:
    """單一檔案的欄位對應結果。col_* 為 1-indexed 欄號。"""

    sheet_name: str
    header_row: int
    data_start_row: int
    tax_id_col: int | None  # 目標檔可擇一留空（見 core/browser_bridge.py 的必填規則）
    customer_name_col: int | None
    acc_se_output_col: int
    group_name_col: int | None = None  # 僅名冊需要
    is_double_header: bool = False


@dataclass
class ColumnProfile:
    """單一欄位的格式輪廓分析結果。"""

    total: int = 0
    empty: int = 0
    numeric_count: int = 0
    text_count: int = 0
    length_distribution: dict[int, int] = field(default_factory=dict)
    hash_prefix_count: int = 0  # 以 # 開頭的暫編代碼
    personal_id_count: int = 0
    mode_length: int | None = None

    def summary(self) -> str:
        lines = [f"總筆數：{self.total}（空白 {self.empty}）"]
        if self.total:
            lines.append(f"眾數長度：{self.mode_length}")
            lines.append(f"# 開頭暫編代碼：{self.hash_prefix_count}")
            lines.append(f"身分證字號格式：{self.personal_id_count}")
        return "\n".join(lines)


@dataclass
class MatchResult:
    row_index: int  # 原始檔案列號（1-indexed，含標頭）
    raw_tax_id: str
    raw_customer_name: str
    clean_tax_id: str
    clean_customer_name: str
    matched_acc_se: str
    match_pass: MatchPass
    anomalies: list[AnomalyCode] = field(default_factory=list)
    secondary_match: str | None = None  # 次要命中的 Acc. SE（供人工複核名冊矛盾）
    duplicate_candidates: list[str] = field(default_factory=list)


@dataclass
class MatchStats:
    total: int = 0
    pass_counts: dict[MatchPass, int] = field(default_factory=dict)
    anomaly_counts: dict[AnomalyCode, int] = field(default_factory=dict)
    formula_na_baseline: int | None = None
    roster_missing_acc_se_count: int = 0  # 名冊本身 Acc.SE 姓名缺漏（規劃書 3.2 #11）
    roster_group_name_skipped_count: int = 0  # 名冊 GroupName 過短已跳過（規劃書 3.2 #13）

    def add_result(self, result: MatchResult) -> None:
        self.total += 1
        self.pass_counts[result.match_pass] = self.pass_counts.get(result.match_pass, 0) + 1
        for code in result.anomalies:
            self.anomaly_counts[code] = self.anomaly_counts.get(code, 0) + 1

    def render(self) -> str:
        lines = [f"客戶資料總數：{self.total}", ""]
        lines.append("成功匹配數（依 Pass 細分）：")
        for p in MatchPass:
            count = self.pass_counts.get(p, 0)
            pct = f"{count / self.total * 100:.1f}%" if self.total else "0%"
            lines.append(f"  {p.value}：{count}（{pct}）")
        lines.append("")
        lines.append("異常/無效資料數：")
        if self.anomaly_counts:
            for code, count in self.anomaly_counts.items():
                lines.append(f"  {code.value}：{count}")
        else:
            lines.append("  （無）")
        lines.append("")
        lines.append("名冊端資料品質（與目標資料問題分開列示，規劃書 3.2 #11）：")
        lines.append(f"  名冊 Acc.SE 姓名缺漏：{self.roster_missing_acc_se_count}")
        lines.append(f"  名冊 GroupName 過短已跳過模糊比對：{self.roster_group_name_skipped_count}")
        if self.formula_na_baseline is not None:
            new_na = self.pass_counts.get(MatchPass.UNMATCHED, 0)
            lines.append("")
            lines.append(
                f"現行公式基準對照：原公式 N/A {self.formula_na_baseline} 筆 "
                f"vs 新工具 N/A {new_na} 筆（改善 {self.formula_na_baseline - new_na} 筆）"
            )
        return "\n".join(lines)
