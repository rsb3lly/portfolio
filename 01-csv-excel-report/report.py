"""売上CSVを月別・商品別に集計し、グラフ付きのExcelに書き出すスクリプト。

使い方:
    python report.py sample_data/sales.csv -o output/sales_report.xlsx

CSVに必要な列: 日付, 商品名, 数量, 売上
（列名が違うときは --date-col などで指定できます）
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HEADER_FILL = PatternFill("solid", fgColor="2F6B4F")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TOTAL_FILL = PatternFill("solid", fgColor="EEF3EA")
THIN = Side(style="thin", color="C9D3CC")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
YEN = '#,##0"円"'


def load(path: Path, date_col: str, item_col: str, qty_col: str, sales_col: str) -> pd.DataFrame:
    try:
        df = pd.read_csv(path, encoding="utf-8-sig")
    except FileNotFoundError:
        sys.exit(f"CSVが見つかりません: {path}")
    missing = [c for c in (date_col, item_col, qty_col, sales_col) if c not in df.columns]
    if missing:
        sys.exit(f"CSVに必要な列がありません: {', '.join(missing)}（今ある列: {', '.join(df.columns)}）")
    df = df.rename(columns={date_col: "日付", item_col: "商品名", qty_col: "数量", sales_col: "売上"})
    df["日付"] = pd.to_datetime(df["日付"], errors="coerce")
    bad = df["日付"].isna().sum()
    if bad:
        print(f"注意: 日付として読めない行が{bad}行あったので、集計から外しました")
        df = df.dropna(subset=["日付"])
    df["月"] = df["日付"].dt.strftime("%Y-%m")
    return df


def write_table(ws, df: pd.DataFrame, start_row: int = 1, money_cols: tuple[str, ...] = ()) -> None:
    """表を書き、見出しの色・罫線・金額の書式をそろえる。最終行は合計行として色を付ける。"""
    for j, col in enumerate(df.columns, start=1):
        c = ws.cell(row=start_row, column=j, value=col)
        c.fill, c.font, c.border = HEADER_FILL, HEADER_FONT, BORDER
        c.alignment = Alignment(horizontal="center")
    for i, row in enumerate(df.itertuples(index=False), start=start_row + 1):
        for j, val in enumerate(row, start=1):
            c = ws.cell(row=i, column=j, value=val)
            c.border = BORDER
            if df.columns[j - 1] in money_cols:
                c.number_format = YEN
            if i == start_row + len(df):
                c.fill = TOTAL_FILL
                c.font = Font(bold=True)
    for j, col in enumerate(df.columns, start=1):
        width = max(len(str(col)) * 2 + 2, *(len(f"{v:,}") + 3 if isinstance(v, (int, float)) else len(str(v)) * 2 + 2
                                             for v in df.iloc[:, j - 1]))
        ws.column_dimensions[get_column_letter(j)].width = min(width, 30)
    ws.freeze_panes = ws.cell(row=start_row + 1, column=1)


def with_total(df: pd.DataFrame, label_col: str) -> pd.DataFrame:
    total = {c: (df[c].sum() if c != label_col else "合計") for c in df.columns}
    return pd.concat([df, pd.DataFrame([total])], ignore_index=True)


def build(df: pd.DataFrame, out: Path) -> dict:
    monthly = df.groupby("月", as_index=False)[["数量", "売上"]].sum()
    by_item = (df.groupby("商品名", as_index=False)[["数量", "売上"]].sum()
               .sort_values("売上", ascending=False, ignore_index=True))
    cross = df.pivot_table(index="月", columns="商品名", values="売上", aggfunc="sum", fill_value=0)
    cross = cross[by_item["商品名"].tolist()].reset_index()
    cross["合計"] = cross.drop(columns="月").sum(axis=1)

    wb = Workbook()
    ws1 = wb.active
    ws1.title = "月別集計"
    write_table(ws1, with_total(monthly, "月"), money_cols=("売上",))

    ws2 = wb.create_sheet("商品別集計")
    item_tbl = by_item.copy()
    item_tbl["構成比"] = item_tbl["売上"] / item_tbl["売上"].sum()
    item_tbl = with_total(item_tbl, "商品名")
    write_table(ws2, item_tbl, money_cols=("売上",))
    for r in range(2, len(item_tbl) + 2):
        ws2.cell(row=r, column=4).number_format = "0.0%"

    ws3 = wb.create_sheet("月×商品")
    write_table(ws3, with_total(cross, "月"), money_cols=tuple(cross.columns[1:]))

    ws4 = wb.create_sheet("グラフ")
    n_month = len(monthly)
    bar = BarChart()
    bar.title, bar.y_axis.title, bar.x_axis.title = "月別の売上", "売上（円）", "月"
    bar.add_data(Reference(ws1, min_col=3, min_row=1, max_row=n_month + 1), titles_from_data=True)
    bar.set_categories(Reference(ws1, min_col=1, min_row=2, max_row=n_month + 1))
    bar.height, bar.width, bar.legend = 8, 16, None
    ws4.add_chart(bar, "A1")

    line = LineChart()
    line.title, line.y_axis.title, line.x_axis.title = "商品別の売上の推移", "売上（円）", "月"
    n_items = len(by_item)
    line.add_data(Reference(ws3, min_col=2, max_col=n_items + 1, min_row=1, max_row=n_month + 1),
                  titles_from_data=True)
    line.set_categories(Reference(ws3, min_col=1, min_row=2, max_row=n_month + 1))
    for s in line.series:  # 曲線にすると実際にない値まで線が膨らむので、直線で結ぶ
        s.smooth = False
    line.height, line.width = 8, 16
    ws4.add_chart(line, "A18")

    ws5 = wb.create_sheet("元データ")
    raw = df[["日付", "商品名", "数量", "売上"]].copy()
    raw["日付"] = raw["日付"].dt.strftime("%Y-%m-%d")
    write_table(ws5, with_total(raw, "日付"), money_cols=("売上",))

    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return {"行数": len(df), "売上合計": int(df["売上"].sum()), "月数": n_month, "商品数": n_items}


def main() -> None:
    p = argparse.ArgumentParser(description="売上CSVを集計してExcelに書き出します")
    p.add_argument("csv", type=Path, help="入力するCSVファイル")
    p.add_argument("-o", "--output", type=Path, default=Path("output/sales_report.xlsx"))
    p.add_argument("--date-col", default="日付")
    p.add_argument("--item-col", default="商品名")
    p.add_argument("--qty-col", default="数量")
    p.add_argument("--sales-col", default="売上")
    a = p.parse_args()
    df = load(a.csv, a.date_col, a.item_col, a.qty_col, a.sales_col)
    if df.empty:
        sys.exit("集計できる行がありません")
    info = build(df, a.output)
    print(f"書き出しました: {a.output}")
    print(f"  {info['行数']}行／{info['月数']}か月／{info['商品数']}商品／売上合計 {info['売上合計']:,}円")


if __name__ == "__main__":
    main()
