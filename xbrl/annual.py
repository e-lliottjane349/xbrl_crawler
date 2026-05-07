"""
xbrl.annual — 年度版 XBRL 爬蟲

資產負債表、綜合損益表、現金流量表均以『整年』為單位合併多期。
呼叫策略：從最新年度往回，每份年報提供本年與上年，逐年補齊最舊一期。
"""
from __future__ import annotations

import pandas as pd

from xbrl.core import (
    KEY_COLS,
    STATEMENT_NAMES,
    STATEMENT_TYPE_MAP,
    drop_header_rows,
    fetch_all_tables,
    normalize_statement,
    sort_period_cols,
)


def get_three_statements(stock_id: str, year: int, quarter: str = "Q4") -> dict[str, pd.DataFrame]:
    """抓取單一年度的三大報表（本年 + 上年）。"""
    raw_tables = fetch_all_tables(stock_id, year, quarter)
    return {
        name: normalize_statement(raw_tables[name], STATEMENT_TYPE_MAP[name])
        for name in STATEMENT_NAMES
    }


def merge_statement_periods(dfs: list[pd.DataFrame], periods: int = 5) -> pd.DataFrame:
    """
    將多年同一張報表合併成近 N 期寬表。

    合併策略（與原邏輯一致）：
    - 第一份 df 保留兩欄（本年、上年）
    - 後續每份 df 僅補最舊一欄進來
    - 最終依年份由新到舊排序，保留前 N 期
    """
    if not dfs:
        raise ValueError("dfs 不可為空")

    merged = dfs[0].copy()

    for df in dfs[1:]:
        value_cols = [c for c in df.columns if c not in KEY_COLS]
        if len(value_cols) < 2:
            raise ValueError(f"日期欄不足 2 欄，無法補期數: {value_cols}")
        oldest_col = value_cols[-1]
        merged = merged.merge(df[KEY_COLS + [oldest_col]], on=KEY_COLS, how="left")

    value_cols = [c for c in merged.columns if c not in KEY_COLS]
    value_cols = sort_period_cols(value_cols, reverse=True)
    merged = merged[KEY_COLS + value_cols[:periods]]

    return drop_header_rows(merged)


def get_five_periods_statements(
    stock_id: str,
    year: int,
    quarter: str = "Q4",
    periods: int = 5,
) -> dict[str, pd.DataFrame]:
    """
    抓近 N 期三大報表。

    例：year=2025, periods=5
    - 2025 報表提供：2025、2024
    - 2024 報表補：2023
    - 2023 報表補：2022
    - 2022 報表補：2021
    需要抓的年數 = periods − 1。
    """
    num_years = max(periods - 1, 1)
    yearly_data = [
        get_three_statements(stock_id, year - offset, quarter)
        for offset in range(num_years)
    ]

    result = {}
    for stmt_name in STATEMENT_NAMES:
        stmt_dfs = [d[stmt_name] for d in yearly_data]
        result[stmt_name] = merge_statement_periods(stmt_dfs, periods=periods)
    return result
