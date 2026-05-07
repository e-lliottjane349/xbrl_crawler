"""
xbrl_crawler_8q_FA.py — 近 N 季單季版執行入口

■ 資產負債表：每季季末快照
■ 綜合損益表：每季單季數字 (Q4 = 全年 - Q3 YTD)
■ 現金流量表：每季單季數字 (YTD 相減 ;Q1 即 YTD)
■ Note: 上市櫃公司現金流量表除年報外為單季數字，興櫃公司為ytd數字

支援：
- 單家公司: STOCK_IDS = "6451"
- 多家公司: STOCK_IDS = ["6451", "2330", "2454"]
- 季數可調: N_QUARTERS = 4、8、12 皆可

將本檔案與 xbrl/ 資料夾放在同一層即可執行：

    python xbrl_crawler_8q_FA.py
"""
from xbrl.batch import batch_export_quarterly


# ════════════════════════════════════════════════════════════════════════
#  ⚙️  參數設定
# ════════════════════════════════════════════════════════════════════════

# 股票代號：單一字串或列表皆可
STOCK_IDS = "3081"
# STOCK_IDS = ["6451", "2330", "2454", "2317"]   # ← 批次處理多家

YEAR = 2026          # 最新季報年份
QUARTER = "Q1"       # 最新季別
N_QUARTERS = 12       # 抓取季數（可改 4、8、12 等）
OUTPUT_DIR = "output_quarterly"     # 輸出資料夾；批次建議改為 "output_quarterly"

SLEEP_BETWEEN_STOCKS = 1.0   # 批次時每家公司之間的等待秒數


# ════════════════════════════════════════════════════════════════════════
#  🚀  主程式
# ════════════════════════════════════════════════════════════════════════

def main():
    return batch_export_quarterly(
        stock_ids=STOCK_IDS,
        year=YEAR,
        quarter=QUARTER,
        n_quarters=N_QUARTERS,
        output_dir=OUTPUT_DIR,
        sleep_between_stocks=SLEEP_BETWEEN_STOCKS,
    )


if __name__ == "__main__":
    main()
