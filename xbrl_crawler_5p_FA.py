"""
xbrl_crawler_5p_FA.py — 近 N 期年度版執行入口

支援：
- 單家公司：STOCK_IDS = "6451"
- 多家公司：STOCK_IDS = ["6451", "2330", "2454"]
- 期數可調：PERIODS = 3、5、7 皆可

將本檔案與 xbrl/ 資料夾放在同一層即可執行：

    python xbrl_crawler_5p_FA.py
"""
from xbrl.batch import batch_export_annual


# ════════════════════════════════════════════════════════════════════════
#  ⚙️  參數設定
# ════════════════════════════════════════════════════════════════════════

# 股票代號：單一字串或列表皆可
STOCK_IDS = "3081"
# STOCK_IDS = ["6451", "2330", "2454", "2317"]   # ← 批次處理多家

YEAR = 2025          # 最新年度
QUARTER = "Q4"       # 年報通常為 Q4
PERIODS = 7          # 抓取期數（可改 3、5、7 等）
OUTPUT_DIR = "output_annual"     # 輸出資料夾；批次建議改為 "output_annual"

SLEEP_BETWEEN_STOCKS = 1.0   # 批次時每家公司之間的等待秒數


# ════════════════════════════════════════════════════════════════════════
#  🚀  主程式
# ════════════════════════════════════════════════════════════════════════

def main():
    return batch_export_annual(
        stock_ids=STOCK_IDS,
        year=YEAR,
        quarter=QUARTER,
        periods=PERIODS,
        output_dir=OUTPUT_DIR,
        sleep_between_stocks=SLEEP_BETWEEN_STOCKS,
    )


if __name__ == "__main__":
    main()
