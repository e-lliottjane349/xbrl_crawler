"""
xbrl.batch — 批次處理多家公司

提供兩個主要函式：
- batch_export_annual    : 批次抓年度三表並輸出 Excel
- batch_export_quarterly : 批次抓季度三表並輸出 Excel

兩者皆支援：
- stock_ids 傳單一字串或列表
- 單家公司失敗不會中斷整個批次
- 可指定 output_dir 將所有 Excel 集中輸出
- 結尾會印出成功/失敗統計
"""
from __future__ import annotations

import time
from pathlib import Path
from typing import Union

from xbrl.annual import get_five_periods_statements
from xbrl.core import export_statements_to_excel
from xbrl.quarterly import get_quarterly_statements


StockIds = Union[str, list[str]]


def _normalize_stock_ids(stock_ids: StockIds) -> list[str]:
    if isinstance(stock_ids, str):
        return [stock_ids]
    return list(stock_ids)


def _print_summary(results: dict[str, dict]) -> None:
    total = len(results)
    ok = sum(1 for r in results.values() if r["status"] == "ok")
    fail = total - ok

    print("\n" + "═" * 60)
    print(f"批次處理完成：成功 {ok} / {total}，失敗 {fail}")
    if fail > 0:
        print("\n失敗清單：")
        for sid, r in results.items():
            if r["status"] == "failed":
                print(f"  [{sid}] {r['error']}")
    print("═" * 60)


def batch_export_annual(
    stock_ids: StockIds,
    year: int,
    quarter: str = "Q4",
    periods: int = 5,
    output_dir: str | Path = ".",
    sleep_between_stocks: float = 1.0,
) -> dict[str, dict]:
    """
    批次抓多家公司的近 N 期年度三大報表，每家一個 Excel 檔。

    Parameters
    ----------
    stock_ids : 股票代號，可傳單一字串或列表
    year      : 最新年度
    quarter   : 最新季別（年報通常為 Q4）
    periods   : 抓取期數（預設 5）
    output_dir : 輸出資料夾，不存在會自動建立
    sleep_between_stocks : 公司之間的等待秒數

    Returns
    -------
    {stock_id: {"status": "ok"|"failed", "file": 路徑或 None, "error": 錯誤字串或 None}}
    """
    stock_ids = _normalize_stock_ids(stock_ids)
    out_path = Path(output_dir)
    out_path.mkdir(parents=True, exist_ok=True)

    total = len(stock_ids)
    print(f"▶ 批次處理（年度）：共 {total} 家，近 {periods} 期")
    print(f"  基準：{year}{quarter.upper()}")
    print(f"  輸出：{out_path.resolve()}\n")

    results: dict[str, dict] = {}
    for i, sid in enumerate(stock_ids, 1):
        print(f"\n[{i}/{total}] ═══ 股票 {sid} ═══")
        try:
            statements = get_five_periods_statements(
                stock_id=sid,
                year=year,
                quarter=quarter,
                periods=periods,
            )
            filename = str(out_path / f"{year}_三表分析_{sid}_近{periods}期.xlsx")
            export_statements_to_excel(
                statements,
                output_filename=filename,
                days_per_period=365,
                quarterly_labels=False,
            )
            results[sid] = {"status": "ok", "file": filename, "error": None}
            print(f"  ✅ 已輸出：{filename}")
        except Exception as exc:
            results[sid] = {"status": "failed", "file": None, "error": str(exc)}
            print(f"  ❌ 失敗：{exc}")

        if i < total and sleep_between_stocks > 0:
            time.sleep(sleep_between_stocks)

    _print_summary(results)
    return results


def batch_export_quarterly(
    stock_ids: StockIds,
    year: int,
    quarter: str,
    n_quarters: int = 8,
    output_dir: str | Path = ".",
    sleep_between_stocks: float = 1.0,
) -> dict[str, dict]:
    """
    批次抓多家公司的近 N 季三大報表（單季化），每家一個 Excel 檔。

    Parameters
    ----------
    stock_ids  : 股票代號，可傳單一字串或列表
    year       : 最新季報年份
    quarter    : 最新季別，如 'Q3'
    n_quarters : 抓取季數（預設 8）
    output_dir : 輸出資料夾，不存在會自動建立
    sleep_between_stocks : 公司之間的等待秒數

    Returns
    -------
    {stock_id: {"status": "ok"|"failed", "file": 路徑或 None, "error": 錯誤字串或 None}}
    """
    stock_ids = _normalize_stock_ids(stock_ids)
    out_path = Path(output_dir)
    out_path.mkdir(parents=True, exist_ok=True)

    total = len(stock_ids)
    print(f"▶ 批次處理(季度)：共 {total} 家，近 {n_quarters} 季")
    print(f"  基準：{year}{quarter.upper()}")
    print(f"  輸出：{out_path.resolve()}\n")

    results: dict[str, dict] = {}
    for i, sid in enumerate(stock_ids, 1):
        print(f"\n[{i}/{total}] ═══ 股票 {sid} ═══")
        try:
            statements = get_quarterly_statements(
                stock_id=sid,
                year=year,
                quarter=quarter,
                n_quarters=n_quarters,
            )
            filename = str(
                out_path / f"{year}{quarter.upper()}_三表分析_{sid}_近{n_quarters}季.xlsx"
            )
            export_statements_to_excel(
                statements,
                output_filename=filename,
                days_per_period=91,
                quarterly_labels=True,
            )
            results[sid] = {"status": "ok", "file": filename, "error": None}
            print(f"  ✅ 已輸出：{filename}")
        except Exception as exc:
            results[sid] = {"status": "failed", "file": None, "error": str(exc)}
            print(f"  ❌ 失敗：{exc}")

        if i < total and sleep_between_stocks > 0:
            time.sleep(sleep_between_stocks)

    _print_summary(results)
    return results
