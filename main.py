import time
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import requests
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from tqdm import tqdm

@dataclass
class AOinfo:
    fullname: str
    title: str
    rating: int | None
    author: str | None
    created_at: int

def get_AOdata(session: requests.Session, page: int, rating: int = 0) -> tuple[list[AOinfo], int]:
    r = session.get(
        "https://data.scp-jp.com/api/pages",
        params={
            "include_all_categories": "true",
            "tags": "+jp +剪定対象-子",
            "q": "anomalous-jp:",
            "rating_min": rating,
            "per_page": 100,
            "page": page,
        },
        timeout=30,
    )
    r.raise_for_status()
    data = r.json()
    if data.get("status") != "ok":
        raise RuntimeError(f"API error: {data}")

    total_pages = data["pagination"]["totalPages"]
    infos = data.get("data", [])
    ao_infos = [AOinfo(
        fullname = i.get("fullname", ""),
        title = i.get("title", ""),
        rating = i.get("rating"),
        author = i.get("createdByName"),
        created_at = i.get("createdAt", 0)
    )
    for i in infos]   
    return (ao_infos, total_pages) 
    
def fetch_all(session: requests.Session, rating: int = 0) -> list[AOinfo]:
    items, total_pages = get_AOdata(session, 1, rating)

    with tqdm(total=total_pages, desc="取得中", unit="page") as bar:
        bar.update(1)  
        bar.set_postfix(件数=len(items))

        for page in range(2, total_pages + 1):
            time.sleep(0.3)
            more, _ = get_AOdata(session, page, rating)
            items.extend(more)
            bar.update(1)
            bar.set_postfix(件数=len(items))

    return items

def write_ranking(ws: Worksheet, infos: list[AOinfo]) -> int:
    """ランキングシートの2行目以降に書き込む。書いた人数を返す"""
    by_author: dict[str, list[AOinfo]] = defaultdict(list)
    for i in infos:
        by_author[i.author or "(不明)"].append(i)

    def score(items: list[AOinfo]) -> tuple[int, int]:
        return len(items), sum(i.rating or 0 for i in items)

    # 件数の多い順、同数なら合計評価の高い順
    rows = sorted(by_author.items(), key=lambda kv: score(kv[1]), reverse=True)

    if ws.max_row > 1:  # 前のデータが残っていたら消す
        ws.delete_rows(2, ws.max_row - 1)

    rank, prev = 0, None
    for n, (author, items) in enumerate(rows, start=1):
        count, total = score(items)
        if count != prev:  # 件数が同じなら同じ順位
            rank, prev = n, count
        best = max(items, key=lambda i: i.rating or 0)
        ws.append([
            rank,
            author,
            count,
            total,
            round(total / count, 1),
            best.rating or 0,
            best.fullname.split(":", 1)[-1],  # anomalous-jp:07094 → 07094
        ])
        ws.cell(row=ws.max_row, column=5).number_format = "0.0"
    return len(rows)

if __name__ == "__main__":
    with requests.Session() as session:
        infos = fetch_all(session, rating=1)

    wb = load_workbook("templete.xlsx")
    n = write_ranking(wb["ランキング"], infos)
    out = Path("output")
    out.mkdir(exist_ok=True)
    wb.save(out / "anomalous.xlsx")
    print(f"{len(infos)}件 / {n}人")
