"""
Сбор отзывов о приложении Halyk Kazakhstan (App Store + Google Play).

Логика соответствует prd.md (раздел 5): полностью локальный скрипт,
запускается ежедневно через Windows Task Scheduler, без облака.

- App Store: публичный RSS-фид (без ключей/авторизации).
- Google Play: публичный парсинг через google-play-scraper (без ключей).
- Дедупликация: App Store — по ID отзыва; Google Play — по хэшу
  автор+дата+текст (публичная страница не отдаёт стабильный ID).
- Дописывает только новые строки в otzyvy_halyk.xlsx, ничего не перезаписывает.
- При появлении новых отзывов с рейтингом ≤3★ показывает Windows toast.
"""

import hashlib
import logging
import time
from datetime import datetime, timezone
from pathlib import Path

import requests
from google_play_scraper import Sort, reviews as gp_reviews
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_FILE = BASE_DIR / "otzyvy_halyk.xlsx"
LOG_FILE = BASE_DIR / "collect_reviews.log"

APP_STORE_ID = "440635615"
GOOGLE_PLAY_ID = "kz.kkb.homebank"

HEADERS = [
    "Дата отзыва", "Стор", "Страна/сторфронт", "Рейтинг",
    "Заголовок отзыва", "Текст отзыва", "Автор", "Версия приложения",
    "ID/хэш отзыва", "Негативный",
]

NEGATIVE_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

COUNTRIES = ["kz"]

# Ограничение по просьбе пользователя: не тянуть всю историю, а только
# последние N отзывов по каждому источнику при первом запуске.
REVIEW_LIMIT = 30

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    encoding="utf-8",
)


def hash_review(author, date, text):
    raw = f"{author}|{date}|{text}".encode("utf-8")
    return hashlib.sha256(raw).hexdigest()[:24]


def _label(value, default=""):
    """RSS-поля Apple обычно {"label": "..."}, но иногда приходят голой строкой."""
    if isinstance(value, dict):
        return value.get("label", default)
    if value is None:
        return default
    return value


def fetch_app_store_reviews(country, seen_ids, limit=REVIEW_LIMIT):
    """Берёт top-`limit` самых свежих отзывов сторфронта (сортировка
    mostRecent) и оставляет только те, что ещё не сохранены. Окно всегда
    строго top-`limit` — скрипт не «уезжает» вглубь истории на повторных
    запусках, даже если часть окна уже была сохранена раньше."""
    raw_entries = []
    for page in range(1, 11):  # Apple отдаёт максимум 10 страниц по 50 отзывов
        url = (
            f"https://itunes.apple.com/{country}/rss/customerreviews/"
            f"id={APP_STORE_ID}/sortBy=mostRecent/page={page}/json"
        )
        try:
            resp = requests.get(url, timeout=15)
            resp.raise_for_status()
            data = resp.json()
        except Exception as e:
            logging.warning("App Store %s page %s: %s", country, page, e)
            break

        entries = data.get("feed", {}).get("entry", [])
        reviews_on_page = [e for e in entries if "im:rating" in e]
        if not reviews_on_page:
            break
        raw_entries.extend(reviews_on_page)
        if len(raw_entries) >= limit:
            break
        time.sleep(0.1)

    results = []
    for e in raw_entries[:limit]:
        try:
            rid = _label(e.get("id"))
            if not rid or rid in seen_ids:
                continue
            try:
                dt = datetime.fromisoformat(_label(e.get("updated")))
                dt = dt.astimezone(timezone.utc).replace(tzinfo=None)
            except Exception:
                dt = None
            rating_raw = _label(e.get("im:rating"))
            author_field = e.get("author")
            author = _label(author_field.get("name")) if isinstance(author_field, dict) else _label(author_field)
            results.append({
                "id": rid,
                "date": dt,
                "rating": int(rating_raw) if rating_raw else None,
                "title": _label(e.get("title")),
                "text": _label(e.get("content")),
                "author": author,
                "version": _label(e.get("im:version")),
                "country": country,
            })
        except Exception as exc:
            logging.warning("App Store %s: пропущен нечитаемый отзыв: %s", country, exc)
    return results


def fetch_google_play_reviews(country, seen_hashes, limit=REVIEW_LIMIT):
    """Забирает последние `limit` отзывов Google Play (сортировка по новизне)
    и оставляет только те, что ещё не сохранены."""
    results = []
    try:
        batch, _ = gp_reviews(
            GOOGLE_PLAY_ID,
            lang="ru",
            country=country,
            sort=Sort.NEWEST,
            count=limit,
        )
        for r in batch:
            author = r.get("userName", "") or ""
            text = r.get("content", "") or ""
            at = r.get("at")
            h = hash_review(author, at, text)
            if h in seen_hashes:
                continue
            results.append({
                "id": h,
                "date": at,
                "rating": r.get("score"),
                "title": "",
                "text": text,
                "author": author,
                "version": r.get("reviewCreatedVersion", "") or "",
                "country": country,
            })
    except Exception as e:
        logging.warning("Google Play %s: %s", country, e)
    return results


def load_or_create_workbook():
    if OUTPUT_FILE.exists():
        wb = load_workbook(OUTPUT_FILE)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Отзывы"
        ws.append(HEADERS)
    return wb, ws


def load_seen_ids(ws):
    app_store_ids, gp_hashes = set(), set()
    store_idx = HEADERS.index("Стор")
    id_idx = HEADERS.index("ID/хэш отзыва")
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or len(row) <= id_idx:
            continue
        store, rid = row[store_idx], row[id_idx]
        if not rid:
            continue
        if store == "App Store":
            app_store_ids.add(str(rid))
        elif store == "Google Play":
            gp_hashes.add(str(rid))
    return app_store_ids, gp_hashes


def append_review_row(ws, store, item):
    rating = item.get("rating")
    is_negative = rating is not None and rating <= 3
    ws.append([
        item.get("date"), store, item.get("country"), rating,
        item.get("title", ""), item.get("text", ""), item.get("author", ""),
        item.get("version", ""), item["id"], "да" if is_negative else "нет",
    ])
    if is_negative:
        last_row = ws.max_row
        for col in range(1, len(HEADERS) + 1):
            ws.cell(row=last_row, column=col).fill = NEGATIVE_FILL
    return is_negative


def notify_negative(count):
    try:
        from winotify import Notification
        toast = Notification(
            app_id="Отзывы Halyk",
            title="Новые негативные отзывы",
            msg=f"Найдено {count} новых отзывов с рейтингом ≤3★",
            duration="long",
        )
        toast.show()
    except Exception:
        logging.exception("Не удалось показать toast-уведомление")


def main():
    wb, ws = load_or_create_workbook()
    seen_app_ids, seen_gp_hashes = load_seen_ids(ws)

    new_total = 0
    new_negative = 0

    for country in COUNTRIES:
        try:
            for item in fetch_app_store_reviews(country, seen_app_ids):
                seen_app_ids.add(item["id"])
                new_total += 1
                if append_review_row(ws, "App Store", item):
                    new_negative += 1
        except Exception:
            logging.exception("App Store %s: сбор не удался", country)

    for country in COUNTRIES:
        try:
            for item in fetch_google_play_reviews(country, seen_gp_hashes):
                seen_gp_hashes.add(item["id"])
                new_total += 1
                if append_review_row(ws, "Google Play", item):
                    new_negative += 1
        except Exception:
            logging.exception("Google Play %s: сбор не удался", country)

    wb.save(OUTPUT_FILE)
    logging.info("Готово. Новых отзывов: %s, из них негативных: %s", new_total, new_negative)

    if new_negative > 0:
        notify_negative(new_negative)


if __name__ == "__main__":
    main()
