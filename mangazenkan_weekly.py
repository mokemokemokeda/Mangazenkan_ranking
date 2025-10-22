# -*- coding: utf-8 -*-
# pip install requests beautifulsoup4 pandas openpyxl
import requests
from bs4 import BeautifulSoup
import pandas as pd
import datetime
from openpyxl import load_workbook
import re

URL = "https://www.mangazenkan.com/r/weekly/ebook/"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; mangazenkan-scraper/1.2)"
}

def fetch_weekly_ranking(url=URL):
    """漫画全巻ドットコムの週間ランキングを取得してDataFrameで返す"""
    res = requests.get(url, headers=HEADERS, timeout=20)
    res.raise_for_status()
    soup = BeautifulSoup(res.text, "html.parser")

    items = soup.select("div.col-4.col-sm-4.col-md-3.col-lg-2")
    results = []

    for item in items:
        # --- ランク ---
        rank_elem = item.select_one("p.rank-number-small")
        rank = rank_elem.get_text(strip=True) if rank_elem else None

        # --- タイトル ---
        title_elem = item.select_one("div.product-name")
        title = title_elem.get_text(strip=True) if title_elem else None

        # --- 巻数（spanタグ構造も考慮）---
        volume = ""
        volume_elem = item.select_one("div.purchase-button-small")
        if volume_elem:
            # 例: "<span>4<small>巻</small></span>" にも対応
            vol_text = volume_elem.get_text(strip=True)
            m = re.search(r"(\d+)", vol_text)
            if m:
                volume = m.group(1)

        # --- 出版社（存在する場合）---
        publisher_elem = item.select_one("div.publisher")
        publisher = publisher_elem.get_text(strip=True) if publisher_elem else None

        if rank and title:
            results.append({
                "rank": int(rank),
                "title": title,
                "volume": volume,
                "publisher": publisher
            })

    return pd.DataFrame(results)


def save_to_excel(df, file_path="weekly_ranking.xlsx"):
    """取得結果をExcelに追記（シート名は実行日）"""
    sheet_name = datetime.datetime.now().strftime("%Y-%m-%d")
    try:
        # 既存ファイルに追記（同日シートがあれば置き換え）
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    except FileNotFoundError:
        # ファイルがない場合は新規作成
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"✅ '{file_path}' にシート '{sheet_name}' を追加しました ({len(df)}件)")


def main():
    df = fetch_weekly_ranking()
    save_to_excel(df)


if __name__ == "__main__":
    main()
