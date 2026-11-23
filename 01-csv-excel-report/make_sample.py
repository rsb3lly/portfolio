"""架空の花屋の売上データ（CSV・100行）を作るスクリプト。

実在のお店・人物とは関係ありません。乱数の種を固定しているので、
何度実行しても同じデータができます。
"""
from __future__ import annotations

import argparse
import csv
import random
from datetime import date, timedelta
from pathlib import Path

# 商品名と単価（円・税込の例）
PRODUCTS = {
    "切り花ブーケ": 1650,
    "鉢花": 1980,
    "観葉植物": 3300,
    "多肉植物": 880,
    "ギフトセット": 5500,
}

START = date(2026, 7, 1)
END = date(2026, 12, 31)


def make_rows(n: int, seed: int) -> list[dict]:
    rng = random.Random(seed)
    span = (END - START).days
    rows = []
    for _ in range(n):
        day = START + timedelta(days=rng.randint(0, span))
        name = rng.choice(list(PRODUCTS))
        qty = rng.randint(1, 6)
        price = PRODUCTS[name]
        rows.append({"日付": day.isoformat(), "商品名": name, "数量": qty,
                     "単価": price, "売上": qty * price})
    rows.sort(key=lambda r: r["日付"])
    return rows


def main() -> None:
    parser = argparse.ArgumentParser(description="架空の売上CSVを作ります")
    parser.add_argument("-o", "--output", default="sample_data/sales.csv")
    parser.add_argument("-n", "--rows", type=int, default=100)
    parser.add_argument("--seed", type=int, default=2026)
    args = parser.parse_args()

    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    rows = make_rows(args.rows, args.seed)
    # Excelで開いても文字化けしないよう、BOM付きUTF-8で保存
    with out.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)
    print(f"{len(rows)}行を書き出しました: {out}")


if __name__ == "__main__":
    main()
