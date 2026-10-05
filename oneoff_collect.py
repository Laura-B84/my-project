"""
Разовый сбор отзывов Halyk Kazakhstan за N дней (по умолчанию 92) — только для
ручного анализа. Не пишет в otzyvy_halyk.xlsx, не отправляет письма.
Результат: out/otzyvy_halyk_<N>d.xlsx (загружается как artifact).
"""

import argparse
from datetime import datetime, timedelta, timezone
from pathlib import Path

from openpyxl import Workbook

import collect_reviews as cr

parser = argparse.ArgumentParser()
parser.add_argument("--days", type=int, default=92)
args = parser.parse_args()

NOW = datetime.now(timezone.utc).replace(tzinfo=None)
CUTOFF = NOW - timedelta(days=args.days)
OUT = Path("out")
OUT.mkdir(exist_ok=True)

apple_items = []
for country in cr.COUNTRIES:
    apple_items.extend(cr.fetch_app_store_reviews(country, set(), cutoff_date=CUTOFF))
android_items = []
for country in cr.COUNTRIES:
    android_items.extend(cr.fetch_google_play_reviews(country, set(), cutoff_date=CUTOFF))

apple_rows = [cr.build_row(i)[0] for i in apple_items]
android_rows = [cr.build_row(i)[0] for i in android_items]
apple_rows = [r for r in apple_rows if r[cr.DATE_IDX] and r[cr.DATE_IDX] >= CUTOFF]
android_rows = [r for r in android_rows if r[cr.DATE_IDX] and r[cr.DATE_IDX] >= CUTOFF]
neg = sum(1 for r in apple_rows + android_rows if r[cr.NEG_IDX] == "да")

wb = Workbook()
ws = wb.active
ws.title = "Отзывы"
cr.setup_headers(ws)
cr.ensure_header_merges(ws)
cr.style_header_cells(ws)
banner = (f"Halyk Kazakhstan — отзывы за {args.days} дней (разовый сбор {NOW:%d.%m.%Y}, kz). "
          f"App Store: {len(apple_rows)}, Google Play: {len(android_rows)}, негативных: {neg}")
cr._merged_banner_row(ws, banner, cr.SECTION_FILL, cr.SECTION_FONT)
cr._write_section(ws, android_rows, apple_rows)
path = OUT / f"otzyvy_halyk_{args.days}d.xlsx"
wb.save(path)
print(f"saved {path} apple={len(apple_rows)} android={len(android_rows)} negative={neg}")
