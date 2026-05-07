"""
xbrl.quarterly — 季度版 XBRL 爬蟲

■ 資產負債表：每季季末快照（取各季報 col 0）
■ 綜合損益表：每季單季數字
    - Q1 / Q2 / Q3：季報 col 0 即為單季
    - Q4          ：年報 col 0（全年）− Q3 YTD
■ 現金流量表：每季單季（季報只有 YTD 累計，逐季相減推算）
    - Q1：YTD = 單季
    - Q2/Q3/Q4：當季 YTD − 前季 YTD
"""
from __future__ import annotations

import time

import pandas as pd

from xbrl.core import (
    drop_header_rows,
    extract_col,
    fetch_all_tables,
    flatten_columns,
    identify_columns,
    outer_merge,
    subtract_dfs,
)


# ════════════════════════════════════════════════════════════════════════
#  季度工具
# ════════════════════════════════════════════════════════════════════════

def prev_quarter(year: int, quarter: str) -> tuple[int, str]:
    """回傳前一季的 (year, quarter)。"""
    q = quarter.upper()
    mapping = {
        "Q1": (year - 1, "Q4"),
        "Q2": (year, "Q1"),
        "Q3": (year, "Q2"),
        "Q4": (year, "Q3"),
    }
    if q not in mapping:
        raise ValueError(f"Invalid quarter: {quarter}")
    return mapping[q]


def generate_quarters(year: int, quarter: str, n: int = 8) -> list[tuple[int, str]]:
    """產生最近 n 季的 (year, quarter) 清單，由新到舊。"""
    result, y, q = [], year, quarter.upper()
    for _ in range(n):
        result.append((y, q))
        y, q = prev_quarter(y, q)
    return result


def q_label(year: int, quarter: str) -> str:
    """(2025, 'Q3') → '2025Q3'"""
    return f"{year}{quarter.upper()}"


# ════════════════════════════════════════════════════════════════════════
#  單季資料擷取
# ════════════════════════════════════════════════════════════════════════

def extract_quarter(stock_id: str, year: int, quarter: str) -> dict:
    """
    抓取單一季度並擷取需要的欄位，回傳：
      lbl    : 季度標籤，如 '2025Q3'
      bs     : 資產負債表季末快照（col 0），欄名 = lbl
      is_sq  : 損益表 col 0
                 Q2/Q3 → 單季；Q1/Q4 → 全期/全年
      is_ytd : 損益表 YTD 欄
                 Q2/Q3 → col 2；Q1/Q4 → col 0（與 is_sq 同欄）
      cf_ytd : 現金流量表 col 0（YTD 累計）
    """
    q = quarter.upper()
    lbl = q_label(year, q)
    tables = fetch_all_tables(stock_id, year, q)

    bs = extract_col(tables["資產負債表"], 0, lbl)

    # 確認 IS 有幾欄數值欄，避免 Q1/Q4 只有兩欄時越界
    is_flat = flatten_columns(tables["綜合損益表"])
    _, _, is_vcols = identify_columns(is_flat)

    is_sq = extract_col(tables["綜合損益表"], 0, lbl)
    ytd_idx = 2 if q in ("Q2", "Q3") and len(is_vcols) >= 3 else 0
    is_ytd = extract_col(tables["綜合損益表"], ytd_idx, f"{lbl}_ytd")

    cf_ytd = extract_col(tables["現金流量表"], 0, f"{lbl}_ytd")

    return {
        "lbl": lbl,
        "bs": bs,
        "is_sq": is_sq,
        "is_ytd": is_ytd,
        "cf_ytd": cf_ytd,
    }


# ════════════════════════════════════════════════════════════════════════
#  近 N 季三大報表（主函式）
# ════════════════════════════════════════════════════════════════════════

def get_quarterly_statements(
    stock_id: str,
    year: int,
    quarter: str,
    n_quarters: int = 8,
    sleep_secs: float = 0.5,
) -> dict[str, pd.DataFrame]:
    """
    抓取近 n_quarters 季的三大報表，回傳：
      '資產負債表' → 每季季末快照
      '綜合損益表' → 每季單季
      '現金流量表' → 每季單季（YTD 相減）

    Parameters
    ----------
    stock_id   : 股票代號，如 '6451'
    year       : 最新季報年份
    quarter    : 最新季報季別，如 'Q3'
    n_quarters : 要抓幾季（預設 8）
    sleep_secs : 每次請求之間的等待秒數
    """
    quarters = generate_quarters(year, quarter, n=n_quarters)

    # 額外抓最舊季的前一季，供 YTD 相減使用
    oldest_y, oldest_q = quarters[-1]
    extra_y, extra_q = prev_quarter(oldest_y, oldest_q)
    all_to_fetch = quarters + [(extra_y, extra_q)]
    cache: dict[str, dict | None] = {}

    print(f"\n▶ 開始抓取 [{stock_id}] 近 {n_quarters} 季財報")
    print(f"  範圍：{q_label(*quarters[-1])} ～ {q_label(*quarters[0])}\n")

    for y, q in all_to_fetch:
        lbl = q_label(y, q)
        print(f"  [{lbl}] 抓取中... ", end="", flush=True)
        try:
            cache[lbl] = extract_quarter(stock_id, y, q)
            print("✓")
        except Exception as exc:
            print(f"✗  ({exc})")
            cache[lbl] = None
        time.sleep(sleep_secs)
    print()

    # ── 資產負債表：每季 col 0 ──
    bs_list = [
        cache[q_label(y, q)]["bs"]
        for y, q in quarters
        if cache.get(q_label(y, q))
    ]
    bs_merged = outer_merge(bs_list)

    # ── 綜合損益表：單季 ──
    is_list = []
    for y, q in quarters:
        lbl = q_label(y, q)
        data = cache.get(lbl)
        if not data:
            continue

        if q.upper() == "Q4":
            q3_lbl = q_label(y, "Q3")
            q3_data = cache.get(q3_lbl)
            if q3_data:
                sq_df = subtract_dfs(
                    data["is_sq"], lbl,
                    q3_data["is_ytd"], f"{q3_lbl}_ytd",
                    lbl,
                )
            else:
                print(f"  ⚠  {lbl} 損益表：找不到 {q3_lbl} YTD，使用全年數代替單季")
                sq_df = data["is_sq"]
        else:
            sq_df = data["is_sq"]

        is_list.append(sq_df)

    is_merged = outer_merge(is_list)

    # ── 現金流量表：單季（YTD 相減）──
    cf_list = []
    for y, q in quarters:
        lbl = q_label(y, q)
        data = cache.get(lbl)
        if not data:
            continue

        if q.upper() == "Q1":
            sq_df = data["cf_ytd"].rename(columns={f"{lbl}_ytd": lbl})
        else:
            prev_y, prev_q = prev_quarter(y, q)
            prev_lbl = q_label(prev_y, prev_q)
            prev_data = cache.get(prev_lbl)
            if prev_data:
                curr_renamed = data["cf_ytd"].rename(columns={f"{lbl}_ytd": lbl})
                sq_df = subtract_dfs(
                    curr_renamed, lbl,
                    prev_data["cf_ytd"], f"{prev_lbl}_ytd",
                    lbl,
                )
            else:
                print(f"  ⚠  {lbl} 現金流量表：找不到 {prev_lbl} YTD，使用 YTD 代替單季")
                sq_df = data["cf_ytd"].rename(columns={f"{lbl}_ytd": lbl})

        cf_list.append(sq_df)

    cf_merged = outer_merge(cf_list)

    return {
        "資產負債表": drop_header_rows(bs_merged),
        "綜合損益表": drop_header_rows(is_merged),
        "現金流量表": drop_header_rows(cf_merged),
    }
