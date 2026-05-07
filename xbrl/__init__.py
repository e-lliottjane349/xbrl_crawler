"""
xbrl — 公開資訊觀測站 XBRL 財報爬蟲套件

子模組：
- core       共用基礎（HTTP、解析、清理、合併、比率、匯出）
- annual     年度版（近 N 期全年資料）
- quarterly  季度版（近 N 季單季資料）

常用 API 範例：

    from xbrl.annual import get_five_periods_statements
    from xbrl.core import export_statements_to_excel

    stmts = get_five_periods_statements("6451", 2025, "Q4", periods=5)
    export_statements_to_excel(stmts, "2025_三表_6451.xlsx")

    from xbrl.quarterly import get_quarterly_statements

    stmts = get_quarterly_statements("6451", 2025, "Q3", n_quarters=8)
    export_statements_to_excel(
        stmts, "2025Q3_八季三表_6451.xlsx",
        days_per_period=91, quarterly_labels=True,
    )
"""
from xbrl.core import (
    KEY_COLS,
    SECTION_IDS,
    STATEMENT_NAMES,
    add_analysis_columns,
    add_common_size_columns,
    calculate_financial_ratios,
    drop_header_rows,
    export_statements_to_excel,
    get_group_value,
    get_value,
)
from xbrl.annual import (
    get_five_periods_statements,
    get_three_statements,
    merge_statement_periods,
)
from xbrl.quarterly import (
    extract_quarter,
    generate_quarters,
    get_quarterly_statements,
    prev_quarter,
    q_label,
)
from xbrl.batch import (
    batch_export_annual,
    batch_export_quarterly,
)

__all__ = [
    # core
    "KEY_COLS",
    "SECTION_IDS",
    "STATEMENT_NAMES",
    "add_analysis_columns",
    "add_common_size_columns",
    "calculate_financial_ratios",
    "drop_header_rows",
    "export_statements_to_excel",
    "get_group_value",
    "get_value",
    # annual
    "get_five_periods_statements",
    "get_three_statements",
    "merge_statement_periods",
    # quarterly
    "extract_quarter",
    "generate_quarters",
    "get_quarterly_statements",
    "prev_quarter",
    "q_label",
    # batch
    "batch_export_annual",
    "batch_export_quarterly",
]
