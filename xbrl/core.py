"""
xbrl.core — 共用基礎層

涵蓋範圍：
- 公開資訊觀測站 HTTP 擷取
- HTML table 解析、欄位辨識、數值清理
- 欄位命名規則（日期標籤、代號、中英文會計項目）
- 分類標題列移除
- 多期合併工具
- 共同比（common-size）分析
- 財務比率計算
- Excel 匯出
"""
from __future__ import annotations

import re
from io import StringIO

import pandas as pd
import requests
from bs4 import BeautifulSoup


# ════════════════════════════════════════════════════════════════════════
#  常數
# ════════════════════════════════════════════════════════════════════════

KEY_COLS = ["代號", "中文會計項目", "英文會計項目"]

SECTION_IDS = {
    "資產負債表": "BalanceSheet",
    "綜合損益表": "StatementOfComprehensiveIncome",
    "現金流量表": "StatementsOfCashFlows",
}

STATEMENT_NAMES = tuple(SECTION_IDS.keys())

STATEMENT_TYPE_MAP = {
    "資產負債表": "balance",
    "綜合損益表": "income",
    "現金流量表": "cashflow",
}


# ════════════════════════════════════════════════════════════════════════
#  HTTP 層
# ════════════════════════════════════════════════════════════════════════

def build_mops_url(
    stock_id: str,
    year: int,
    quarter: str,
    report_id: str = "C",
) -> str:
    """
    建立公開資訊觀測站新版 XBRL 查詢 URL。

    Parameters
    ----------
    stock_id  : 股票代號，例如 '5439'
    year      : 西元年，例如 2024
    quarter   : 季別，接受 'Q3' 或 '3' 兩種格式，自動轉換
    report_id : 'C' = 合併報表；'A' = 個別報表
    """
    # 相容 'Q3' 或 '3' 兩種輸入格式
    sseason = int(str(quarter).upper().lstrip("Q"))
    return (
        f"https://mopsov.twse.com.tw/server-java/t164sb01"
        f"?step=1&CO_ID={stock_id}"
        f"&SYEAR={year}&SSEASON={sseason}"
        f"&REPORT_ID={report_id}"
    )


def fetch_soup(
    stock_id: str,
    year: int,
    quarter: str,
    report_id: str = "C",
) -> BeautifulSoup:
    """
    抓取單季財務報告頁面並回傳 BeautifulSoup。

    先嘗試指定的 report_id，若頁面中找不到 BalanceSheet 錨點，
    自動 fallback 到另一種報表類型（合併↔個別）。
    """
    fallback = "A" if report_id == "C" else "C"
    for rid in (report_id, fallback):
        url = build_mops_url(stock_id, year, quarter, rid)
        resp = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=30)
        if resp.status_code != 200:
            continue
        soup = BeautifulSoup(
            resp.content.decode("big5", errors="replace"), "html.parser"
        )
        if soup.find(id="BalanceSheet") is not None:
            return soup
    raise ValueError(
        f"找不到 section ID: BalanceSheet"
    )


# ════════════════════════════════════════════════════════════════════════
#  欄位與資料清理
# ════════════════════════════════════════════════════════════════════════

def flatten_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [
            "".join(str(x).strip() for x in col if pd.notna(x) and str(x).strip() != "")
            for col in df.columns
        ]
    else:
        df.columns = [str(c).strip() for c in df.columns]
    return df


def normalize_code(code) -> str:
    if pd.isna(code):
        return ""
    s = str(code).strip()
    return re.sub(r"\.0+$", "", s)


def split_zh_en(text):
    """將『會計項目』拆分為中文與英文兩欄。"""
    if pd.isna(text):
        return pd.Series([None, None])
    s = re.sub(r"\s+", " ", str(text).replace("\u3000", " ").strip())
    m = re.search(r"[A-Za-z]", s)
    if not m:
        return pd.Series([s, None])
    zh = s[: m.start()].strip(" -–—")
    en = s[m.start():].strip()
    return pd.Series([zh or None, en or None])


def identify_columns(df: pd.DataFrame):
    """
    回傳 (code_col, title_col, value_cols)。
    value_cols 含所有含 4 位數年份的欄位（未裁切）。
    """
    cols = list(df.columns)
    code_col = None
    title_col = None
    value_cols = []

    for col in cols:
        col_str = str(col)
        if code_col is None and ("代號" in col_str or "Code" in col_str):
            code_col = col
        if title_col is None and ("會計項目" in col_str or "Accounting Title" in col_str):
            title_col = col

    for col in cols:
        if col == code_col or col == title_col:
            continue
        if re.search(r"\d{4}", str(col)):
            value_cols.append(col)

    return code_col, title_col, value_cols


def parse_numeric(series: pd.Series) -> pd.Series:
    """清理數值欄：去千分位、負號括號、空白值標記後轉 numeric。"""
    return (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("(", "-", regex=False)
        .str.replace(")", "", regex=False)
        .replace({"nan": None, "None": None, "": None, "—": None, "-": None})
        .pipe(pd.to_numeric, errors="coerce")
    )


def extract_date_label(col_name, statement_type: str) -> str:
    """
    年度報表的日期欄名標準化：
    - balance   → '2025/12/31'
    - income    → '2025/1/1 - 2025/12/31'
    - cashflow  → '2025/1/1 - 2025/12/31'
    """
    text = str(col_name)
    dates = re.findall(r"\d{4}/\d{1,2}/\d{1,2}", text)
    if not dates:
        return text
    date = dates[0]
    year = date.split("/")[0]
    if statement_type == "balance":
        return date
    if statement_type in ("income", "cashflow"):
        return f"{year}/1/1 - {year}/12/31"
    return date


def drop_header_rows(df: pd.DataFrame) -> pd.DataFrame:
    """
    移除純分類標題列（代號為空 且 所有數值欄都為 NaN），
    例如「資產」「流動資產」「營業收入」這些節點。
    """
    value_cols = [
        c for c in df.columns
        if c not in KEY_COLS and not str(c).endswith(" %")
    ]
    if not value_cols or "代號" not in df.columns:
        return df.reset_index(drop=True)
    mask_empty_code = df["代號"].astype(str).str.strip() == ""
    mask_all_na = df[value_cols].isna().all(axis=1)
    return df[~(mask_empty_code & mask_all_na)].reset_index(drop=True)


# ════════════════════════════════════════════════════════════════════════
#  報表解析
# ════════════════════════════════════════════════════════════════════════

def parse_section_table(soup: BeautifulSoup, section_id: str) -> pd.DataFrame:
    anchor = soup.find(id=section_id)
    if anchor is None:
        raise ValueError(f"找不到 section ID: {section_id}")
    table = anchor.find_next("table")
    if table is None:
        raise ValueError(f"找不到 table: {section_id}")
    return pd.read_html(StringIO(str(table)))[0]


def fetch_all_tables(stock_id: str, year: int, quarter: str) -> dict[str, pd.DataFrame]:
    """抓取單季/單年的三大報表原始 DataFrame（未經清理）。"""
    soup = fetch_soup(stock_id, year, quarter)
    return {
        name: parse_section_table(soup, sec_id)
        for name, sec_id in SECTION_IDS.items()
    }


def normalize_statement(raw_df: pd.DataFrame, statement_type: str) -> pd.DataFrame:
    """
    年度版整理：
    - 取最後兩欄數值欄（本期 + 上期），以日期標籤命名
    - 拆中英文會計項目
    - 清理數值
    - 移除分類標題列
    """
    df = flatten_columns(raw_df).copy()
    code_col, title_col, value_cols = identify_columns(df)

    if code_col is None or title_col is None or len(value_cols) < 2:
        raise ValueError(f"無法辨識欄位，現有欄位為: {df.columns.tolist()}")

    value_cols = value_cols[-2:]  # 取最後兩欄（本期、上期）
    curr_col, prev_col = value_cols[0], value_cols[1]
    curr_label = extract_date_label(curr_col, statement_type)
    prev_label = extract_date_label(prev_col, statement_type)

    out = df[[code_col, title_col, curr_col, prev_col]].copy()
    out.columns = ["代號", "會計項目", curr_label, prev_label]
    out["代號"] = out["代號"].apply(normalize_code)
    out[["中文會計項目", "英文會計項目"]] = out["會計項目"].apply(split_zh_en)
    out = out[KEY_COLS + [curr_label, prev_label]]

    for col in (curr_label, prev_label):
        out[col] = parse_numeric(out[col])

    out = out.dropna(how="all")
    return drop_header_rows(out)


def extract_col(raw_df: pd.DataFrame, col_idx: int, col_name: str) -> pd.DataFrame:
    """
    季度版整理：從原始 df 取指定 col_idx 的數值欄，用 col_name 命名。
    不做 drop_header_rows（由外部合併後再統一清理）。
    """
    df = flatten_columns(raw_df)
    code_col, title_col, value_cols = identify_columns(df)

    if code_col is None or title_col is None:
        raise ValueError(f"無法辨識代號/會計項目欄：{df.columns.tolist()}")
    if col_idx >= len(value_cols):
        raise ValueError(f"col_idx={col_idx} 超出範圍，可用欄位：{value_cols}")

    out = df[[code_col, title_col, value_cols[col_idx]]].copy()
    out.columns = ["代號", "會計項目", col_name]
    out["代號"] = out["代號"].apply(normalize_code)
    out[["中文會計項目", "英文會計項目"]] = out["會計項目"].apply(split_zh_en)
    out[col_name] = parse_numeric(out[col_name])
    return out[KEY_COLS + [col_name]].dropna(how="all")


# ════════════════════════════════════════════════════════════════════════
#  合併與期間排序
# ════════════════════════════════════════════════════════════════════════

def outer_merge(dfs: list[pd.DataFrame]) -> pd.DataFrame:
    """用 outer join 把多個單期 df 合併（季度版用）。"""
    if not dfs:
        return pd.DataFrame(columns=KEY_COLS)
    base = dfs[0]
    for df in dfs[1:]:
        base = base.merge(df, on=KEY_COLS, how="outer")
    return base


def subtract_dfs(
    curr_df: pd.DataFrame, curr_col: str,
    prev_df: pd.DataFrame, prev_col: str,
    result_col: str,
) -> pd.DataFrame:
    """用於 CF 的 YTD 相減，或 Q4 損益 = 全年 − Q3 YTD。"""
    m = curr_df.merge(prev_df[KEY_COLS + [prev_col]], on=KEY_COLS, how="left")
    m[result_col] = m[curr_col] - m[prev_col]
    return m[KEY_COLS + [result_col]]


def sort_period_cols(cols: list, reverse: bool = True) -> list:
    """
    對期間欄名做字串排序（由新到舊）。
    適用 '2025/12/31'、'2025/1/1 - 2025/12/31'、'2025Q3' 等格式，
    因為它們都以 4 位數年份開頭，字典序與時間序一致。
    """
    return sorted(cols, reverse=reverse)


# ════════════════════════════════════════════════════════════════════════
#  共同比分析
# ════════════════════════════════════════════════════════════════════════

def add_common_size_columns(
    df: pd.DataFrame,
    base_code: str,
    round_digits: int = 2,
) -> pd.DataFrame:
    """
    在每個期間欄後面插入百分比欄（百分比 = 該科目 / base_code 科目 × 100）。
    """
    out = df.copy()
    value_cols = [c for c in out.columns if c not in KEY_COLS]
    base_row = out[out["代號"].apply(normalize_code) == normalize_code(base_code)]

    if base_row.empty:
        raise ValueError(f"找不到基準代號 {base_code}")

    final_cols = []
    for col in out.columns:
        final_cols.append(col)
        if col in value_cols:
            base_value = pd.to_numeric(base_row.iloc[0][col], errors="coerce")
            pct_col = f"{col} %"
            if pd.isna(base_value) or base_value == 0:
                out[pct_col] = pd.NA
            else:
                out[pct_col] = (
                    pd.to_numeric(out[col], errors="coerce") / base_value * 100
                ).round(round_digits)
            final_cols.append(pct_col)

    return out[final_cols]


def add_analysis_columns(statements: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
    """
    對三大報表加入共同比欄：
    - 資產負債表：÷ 資產總計 (1XXX)
    - 綜合損益表：÷ 營業收入 (4000)
    - 現金流量表：維持原樣
    """
    analyzed = {}
    try:
        analyzed["資產負債表"] = add_common_size_columns(statements["資產負債表"], "1XXX")
    except Exception as e:
        print(f"  ⚠  資產負債表共同比失敗：{e}")
        analyzed["資產負債表"] = statements["資產負債表"].copy()

    try:
        analyzed["綜合損益表"] = add_common_size_columns(statements["綜合損益表"], "4000")
    except Exception as e:
        print(f"  ⚠  綜合損益表共同比失敗：{e}")
        analyzed["綜合損益表"] = statements["綜合損益表"].copy()

    analyzed["現金流量表"] = statements["現金流量表"].copy()
    return analyzed


# ════════════════════════════════════════════════════════════════════════
#  財務比率
# ════════════════════════════════════════════════════════════════════════

def get_value(df: pd.DataFrame, code: str, col: str):
    row = df[df["代號"].apply(normalize_code) == normalize_code(code)]
    return pd.to_numeric(row.iloc[0][col], errors="coerce") if not row.empty else None


def get_group_value(df: pd.DataFrame, prefix: str, col: str):
    codes = df["代號"].apply(normalize_code)
    rows = df[codes.str.startswith(prefix, na=False)]
    return pd.to_numeric(rows[col], errors="coerce").sum() if not rows.empty else None


def avg(v1, v2):
    if v1 is None or v2 is None or pd.isna(v1) or pd.isna(v2):
        return None
    return (v1 + v2) / 2


def safe_div(a, b):
    if a is None or b is None or pd.isna(a) or pd.isna(b) or b == 0:
        return None
    return a / b


def _period_match_key(label) -> str:
    """
    將期間標籤轉成可供 BS↔IS 匹配的 key：
    - '2025Q3' → '2025Q3'
    - '2025/12/31' → '2025'
    - '2025/1/1 - 2025/12/31' → '2025'
    """
    s = str(label)
    m = re.match(r"(\d{4}Q[1-4])", s)
    if m:
        return m.group(1)
    m = re.search(r"(\d{4})", s)
    return m.group(1) if m else s


def _compute_ratio_inputs(bs, is_, bs_p, is_p, prev_bs_p, prev_is_p) -> dict:
    """把計算比率所需的所有原始數值一次抓齊。"""
    v = dict(
        assets=get_group_value(bs, "1", bs_p),
        liabilities=get_group_value(bs, "2", bs_p),
        equity=get_group_value(bs, "3", bs_p),
        current_assets=get_value(bs, "11XX", bs_p),
        current_liabilities=get_value(bs, "21XX", bs_p),
        inventory=get_value(bs, "130X", bs_p),
        ar=get_value(bs, "1170", bs_p),
        ap=get_value(bs, "2170", bs_p),
        cash=get_value(bs, "1100", bs_p),
        ppe=get_value(bs, "1600", bs_p),
        long_debt=get_group_value(bs, "25", bs_p),

        revenue=get_value(is_, "4000", is_p),
        cogs=get_value(is_, "5000", is_p),
        interest=get_value(is_, "7050", is_p),
        op_profit=get_value(is_, "7900", is_p),
        net_income=get_value(is_, "8200", is_p),
        operating_expense=get_value(is_, "6000", is_p),
        # 獲利率用
        gross_profit=get_value(is_, "5950", is_p),       # 營業毛利(毛損)淨額
        operating_income=get_value(is_, "6900", is_p),   # 營業利益
        ni_continuing=get_value(is_, "8000", is_p),      # 繼續營業單位本期淨利(使用者指定)
        eps=get_value(is_, "9750", is_p),                # 基本每股盈餘
    )
    v["fixed_expense"] = (v["operating_expense"] or 0) + (v["interest"] or 0)

    if prev_bs_p and prev_is_p:
        v["avg_assets"] = avg(v["assets"], get_group_value(bs, "1", prev_bs_p))
        v["avg_inventory"] = avg(v["inventory"], get_value(bs, "130X", prev_bs_p))
        v["avg_ar"] = avg(v["ar"], get_value(bs, "1170", prev_bs_p))
        v["avg_ap"] = avg(v["ap"], get_value(bs, "2170", prev_bs_p))
        v["avg_cash"] = avg(v["cash"], get_value(bs, "1100", prev_bs_p))
        v["avg_equity"] = avg(v["equity"], get_group_value(bs, "3", prev_bs_p))
    else:
        for k in ("avg_assets", "avg_inventory", "avg_ar", "avg_ap", "avg_cash", "avg_equity"):
            v[k] = None
    return v


def _build_ratio_items(v: dict, days_per_period: int, quarterly_labels: bool):
    """由原始數值字典產生 (類別, 名稱, 值) 三元組清單。"""
    tax_rate = 0.2

    inv_to = safe_div(v["cogs"], v["avg_inventory"])
    ar_to = safe_div(v["revenue"], v["avg_ar"])
    ap_to = safe_div(v["cogs"], v["avg_ap"])
    inv_days = safe_div(days_per_period, inv_to)
    ar_days = safe_div(days_per_period, ar_to)
    ap_days = safe_div(days_per_period, ap_to)
    roa = safe_div(v["net_income"], v["avg_assets"])
    roe = safe_div(v["net_income"], v["avg_equity"])

    ccc = None
    if all(x is not None for x in (inv_days, ar_days, ap_days)):
        ccc = inv_days + ar_days - ap_days

    op_cycle = None
    if inv_days is not None and ar_days is not None:
        op_cycle = inv_days + ar_days

    working_capital = None
    if v["current_assets"] is not None and v["current_liabilities"] is not None:
        working_capital = v["current_assets"] - v["current_liabilities"]

    # 依模式調整標籤
    if quarterly_labels:
        suf = "（季）"
        profit_cat = "獲利能力(季度,非年化)"
        roa_lbl = "季度 ROA"
        roe_lbl = "季度 ROE"
        lcap_lbl = "長期資本報酬率(季度)"
    else:
        suf = ""
        profit_cat = "獲利能力(報酬率)"
        roa_lbl = "ROA(總資產報酬率)"
        roe_lbl = "ROE(權益報酬率)"
        lcap_lbl = "長期資本報酬率"

    return [
        ("流動性與償債能力", "流動比率", safe_div(v["current_assets"], v["current_liabilities"])),
        ("流動性與償債能力", "速動比率", safe_div((v["current_assets"] or 0) - (v["inventory"] or 0), v["current_liabilities"])),
        ("流動性與償債能力", "營運資金", working_capital),
        ("流動性與償債能力", "利息保障倍數", safe_div((v["op_profit"] or 0) + (v["interest"] or 0), v["interest"])),
        ("流動性與償債能力", "固定支出保障倍數", safe_div((v["op_profit"] or 0) + v["fixed_expense"], v["fixed_expense"])),

        ("經營效率(週轉率)", f"總資產週轉率{suf}", safe_div(v["revenue"], v["avg_assets"])),
        ("經營效率(週轉率)", f"存貨週轉率{suf}", inv_to),
        ("經營效率(週轉率)", "存貨週轉天數", inv_days),
        ("經營效率(週轉率)", f"應收帳款週轉率{suf}", ar_to),
        ("經營效率(週轉率)", "應收帳款週轉天數", ar_days),
        ("經營效率(週轉率)", f"應付帳款週轉率{suf}", ap_to),
        ("經營效率(週轉率)", "應付帳款週轉天數", ap_days),
        ("經營效率(週轉率)", "營業週期", op_cycle),
        ("經營效率(週轉率)", f"現金週轉率{suf}", safe_div(v["revenue"], v["avg_cash"])),
        ("經營效率(週轉率)", "現金轉換週期(CCC)", ccc),

        (profit_cat, "毛利率", safe_div(v["gross_profit"], v["revenue"])),
        (profit_cat, "營業利益率", safe_div(v["operating_income"], v["revenue"])),
        (profit_cat, "淨利率", safe_div(v["ni_continuing"], v["revenue"])),
        (profit_cat, "EPS(每股盈餘)", v["eps"]),
        (profit_cat, roa_lbl, roa),
        (profit_cat, roe_lbl, roe),
        (profit_cat, lcap_lbl, safe_div(
            (v["net_income"] or 0) + (v["interest"] or 0) * (1 - tax_rate),
            (v["long_debt"] or 0) + (v["equity"] or 0),
        )),

        ("資本結構", "負債比率", safe_div(v["liabilities"], v["assets"])),
        ("資本結構", "權益比率", safe_div(v["equity"], v["assets"])),

        ("財務槓桿", "權益乘數(財務槓桿)", safe_div(v["assets"], v["equity"])),
        ("財務槓桿", "財務槓桿指數", safe_div(roe, roa)),

        ("資產結構", "長期資金對PPE比率", safe_div((v["long_debt"] or 0) + (v["equity"] or 0), v["ppe"])),
    ]


def calculate_financial_ratios(
    statements: dict[str, pd.DataFrame],
    *,
    days_per_period: int = 365,
    quarterly_labels: bool = False,
) -> pd.DataFrame:
    """
    計算多期財務比率，回傳橫向寬表（種類、名稱、各期）。

    Parameters
    ----------
    days_per_period : 週轉天數的分母。年度版用 365，季度版用 91。
    quarterly_labels : True 時使用季度版標籤（例：「季度 ROA」「總資產週轉率(季)」）。
    """
    bs = statements["資產負債表"]
    is_ = statements["綜合損益表"]

    bs_periods = [c for c in bs.columns if c not in KEY_COLS and not str(c).endswith(" %")]
    is_periods = [c for c in is_.columns if c not in KEY_COLS and not str(c).endswith(" %")]
    bs_periods = sort_period_cols(bs_periods, reverse=True)
    is_periods = sort_period_cols(is_periods, reverse=True)

    is_key_map = {_period_match_key(c): c for c in is_periods}

    rows = []
    for i, bs_p in enumerate(bs_periods):
        is_p = is_key_map.get(_period_match_key(bs_p))
        if is_p is None:
            continue

        prev_bs_p = bs_periods[i + 1] if i + 1 < len(bs_periods) else None
        prev_is_p = (
            is_key_map.get(_period_match_key(prev_bs_p)) if prev_bs_p else None
        )

        v = _compute_ratio_inputs(bs, is_, bs_p, is_p, prev_bs_p, prev_is_p)
        for category, name, value in _build_ratio_items(v, days_per_period, quarterly_labels):
            rows.append({"種類": category, "名稱": name, "期間": bs_p, "計算結果": value})

    if not rows:
        return pd.DataFrame(columns=["種類", "名稱"])

    ratio_df = pd.DataFrame(rows)

    # 手動 pivot：保留 (種類, 名稱) 的原始插入順序，避開 pandas pivot 的欄索引名污染
    unique_pairs = ratio_df[["種類", "名稱"]].drop_duplicates().reset_index(drop=True)
    unique_periods = sort_period_cols(ratio_df["期間"].unique().tolist(), reverse=True)

    result = unique_pairs.copy()
    for period in unique_periods:
        sub = (
            ratio_df.loc[ratio_df["期間"] == period, ["種類", "名稱", "計算結果"]]
            .rename(columns={"計算結果": period})
        )
        result = result.merge(sub, on=["種類", "名稱"], how="left")

    return result


# ════════════════════════════════════════════════════════════════════════
#  Excel 匯出
# ════════════════════════════════════════════════════════════════════════

def export_statements_to_excel(
    statements: dict[str, pd.DataFrame],
    output_filename: str,
    *,
    days_per_period: int = 365,
    quarterly_labels: bool = False,
    include_common_size: bool = True,
    include_ratios: bool = True,
) -> str:
    """
    將三大報表（含共同比欄與財務比率分析頁）匯出為 Excel 檔。

    Parameters
    ----------
    statements          : 原始三大報表（會在內部加共同比與計算比率）
    output_filename     : 輸出檔名，例如 '2025_三表分析_6451.xlsx'
    days_per_period     : 週轉天數分母，年度 365、季度 91
    quarterly_labels    : 財務比率是否使用季度版標籤
    include_common_size : 是否在 BS/IS 加入百分比欄
    include_ratios      : 是否產出『財務比率分析』分頁
    """
    analyzed = add_analysis_columns(statements) if include_common_size else statements

    ratio_df = None
    if include_ratios:
        ratio_df = calculate_financial_ratios(
            statements,
            days_per_period=days_per_period,
            quarterly_labels=quarterly_labels,
        )

    with pd.ExcelWriter(output_filename, engine="openpyxl") as writer:
        analyzed["資產負債表"].to_excel(writer, sheet_name="資產負債表", index=False)
        analyzed["綜合損益表"].to_excel(writer, sheet_name="綜合損益表", index=False)
        analyzed["現金流量表"].to_excel(writer, sheet_name="現金流量表", index=False)
        if ratio_df is not None and not ratio_df.empty:
            ratio_df.to_excel(writer, sheet_name="財務比率分析", index=False)

    return output_filename