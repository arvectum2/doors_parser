from __future__ import annotations

import re
from pathlib import Path
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup


def norm(text: str) -> str:
    text = (text or "").lower().replace("ё", "е")
    text = re.sub(r"[^a-zа-я0-9]+", " ", text, flags=re.I)
    return re.sub(r"\s+", " ", text).strip()


def clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "")).strip()


def is_noise_text(text: str) -> bool:
    t = norm(text)
    if not t:
        return True
    bad_exact = {
        "где купить",
        "купить",
        "каталог",
        "главная",
        "контакты",
        "доставка",
        "оплата",
        "модельный ряд",
        "двери в интерьере",
    }
    if t in bad_exact:
        return True
    bad_contains = (
        "политик",
        "персональн",
        "заполните форму",
        "коммерческого предложения",
    )
    return any(x in t for x in bad_contains)


def bad_image_url(url: str) -> bool:
    low = (url or "").lower()
    if not low:
        return True
    bad = (
        "/empty/",
        "logo",
        "icon",
        "badge",
        "sticker",
        "label",
        "sale",
        "promo",
        "banner",
        "search",
        "arrow",
        "placeholder",
        "no_image",
        "no-image",
        "pixel",
        "open.png",
        "catalog.svg",
        "sprite",
        "loader",
        "favicon",
        "tl.png",
        "tg.png",
        "max.png",
        "whatsapp",
        "telegram",
        "/resize/20x/",
        "/cover/240x180/",
    )
    return any(x in low for x in bad)


def abs_url(page_url: str, raw: str) -> str:
    raw = (raw or "").strip().strip("'\"")
    if not raw:
        return ""
    if raw.startswith("//"):
        raw = "https:" + raw
    return urljoin(page_url, raw)


def tilda_restore(url: str) -> str:
    u = url or ""
    u = re.sub(r"/-/resizeb?/\d+x/", "/", u, flags=re.I)
    u = re.sub(r"/-/cover/\d+x\d+/", "/", u, flags=re.I)
    if "thb.tildacdn.com" in u:
        u = u.replace("thb.tildacdn.com", "static.tildacdn.com")
    return u


def filename_caption(url: str) -> str:
    path = urlparse(url).path
    stem = Path(path).stem
    stem = re.sub(r"[_\-]+", " ", stem)
    stem = clean_text(stem)
    low = norm(stem)
    if not stem:
        return ""
    if low in {"photo", "image", "img", "close", "big", "small"}:
        return ""
    if re.fullmatch(r"[_\-\s\d]+", stem):
        return ""
    if len(re.sub(r"\d", "", low)) < 2:
        return ""
    return stem


def fetch_html(session: requests.Session, url: str) -> str:
    try:
        r = session.get(url, timeout=(10, 30), allow_redirects=True)
        if r.status_code >= 400:
            return ""
        return r.text
    except Exception:
        return ""


def extract_hausdoors_candidates(source_url: str, html: str) -> list[dict]:
    soup = BeautifulSoup(html, "lxml")
    out: list[dict] = []
    seen: set[tuple[str, str]] = set()

    # Strong signal block on many Hausdoors pages
    for item in soup.select(".t1148__item"):
        img = item.select_one("img")
        if not img:
            continue
        src = img.get("data-original") or img.get("data-src") or img.get("src") or ""
        image_url = tilda_restore(abs_url(source_url, src))
        if not image_url or bad_image_url(image_url):
            continue
        caption = clean_text(item.get_text(" ", strip=True))
        if is_noise_text(caption):
            caption = filename_caption(image_url)
        key = (norm(caption), image_url)
        if key in seen:
            continue
        seen.add(key)
        out.append(
            {
                "model_name": caption,
                "image_url": image_url,
                "source_url": source_url,
                "comment": "visible_card_caption" if caption else "no_visible_model_caption",
                "has_caption": bool(caption),
            }
        )

    # Fallback generic cards
    selectors = [
        ".t-store__card",
        ".t-card",
        ".t-slds__item",
        ".tn-elem",
    ]
    for sel in selectors:
        for card in soup.select(sel):
            img = card.select_one("img")
            if not img:
                continue
            src = img.get("data-original") or img.get("data-src") or img.get("src") or ""
            image_url = tilda_restore(abs_url(source_url, src))
            if not image_url or bad_image_url(image_url):
                continue
            title = ""
            for tsel in (".t-card__title", ".t-store__card__title", ".t-name", ".t-title", ".tn-atom"):
                node = card.select_one(tsel)
                if node:
                    title = clean_text(node.get_text(" ", strip=True))
                    if not is_noise_text(title):
                        break
            if is_noise_text(title):
                title = filename_caption(image_url)
            key = (norm(title), image_url)
            if key in seen:
                continue
            seen.add(key)
            out.append(
                {
                    "model_name": title,
                    "image_url": image_url,
                    "source_url": source_url,
                    "comment": "visible_card_caption" if title else "no_visible_model_caption",
                    "has_caption": bool(title),
                }
            )

    return out


def rebuild_hausdoors_rows(old_rows: list[dict]) -> tuple[list[dict], dict]:
    by_line_url: dict[tuple[str, str], list[dict]] = {}
    for row in old_rows:
        key = (row.get("target_line", ""), row.get("source_url", ""))
        by_line_url.setdefault(key, []).append(row)

    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36",
            "Accept-Language": "ru,en;q=0.9",
        }
    )

    new_rows: list[dict] = []
    real_model_rows = 0
    generated_model_rows = 0
    variant_rows_old = sum(
        1
        for r in old_rows
        if re.search(r"\bvariant\s*\d+\b|\bвариант\s*\d+\b", r.get("model_name", ""), flags=re.I)
    )

    for (line, source_url), rows in by_line_url.items():
        defaults = rows[0]
        html = fetch_html(session, source_url)
        candidates = extract_hausdoors_candidates(source_url, html) if html else []
        uniq: dict[tuple[str, str], dict] = {}
        for c in candidates:
            key = (norm(c.get("model_name", "")), c.get("image_url", ""))
            if key not in uniq:
                uniq[key] = c
        final = list(uniq.values())

        if not final:
            fallback_img = ""
            for r in rows:
                img = r.get("image_url", "")
                if img and not bad_image_url(img):
                    fallback_img = tilda_restore(img)
                    break
            new_rows.append(
                {
                    "door_type": defaults.get("door_type", "межкомнатная"),
                    "producer": "Хаусдорс",
                    "target_line": line,
                    "model_name": line,
                    "description": defaults.get("description", ""),
                    "material": defaults.get("material", ""),
                    "coating": defaults.get("coating", ""),
                    "image_url": fallback_img,
                    "source_url": source_url,
                    "status": "needs_review",
                    "comment": "collection_image_only_no_model_cards",
                }
            )
            generated_model_rows += 1
            continue

        idx = 1
        for c in final:
            model_name = clean_text(c.get("model_name", ""))
            comment = c.get("comment", "")
            status = "found"
            if not model_name:
                model_name = f"{line} — модель {idx}"
                comment = "generated_model_name_no_visible_caption"
                status = "needs_review"
                generated_model_rows += 1
            else:
                real_model_rows += 1
            new_rows.append(
                {
                    "door_type": defaults.get("door_type", "межкомнатная"),
                    "producer": "Хаусдорс",
                    "target_line": line,
                    "model_name": model_name,
                    "description": defaults.get("description", ""),
                    "material": defaults.get("material", ""),
                    "coating": defaults.get("coating", ""),
                    "image_url": c.get("image_url", ""),
                    "source_url": source_url,
                    "status": status,
                    "comment": comment,
                }
            )
            idx += 1

    # Repeated image handling
    from collections import Counter

    img_counts = Counter(r.get("image_url", "") for r in new_rows if r.get("image_url", ""))
    repeated_gt10 = 0
    for r in new_rows:
        img = r.get("image_url", "")
        if img and img_counts.get(img, 0) > 10:
            repeated_gt10 += 1
            r["status"] = "needs_review"
            c = r.get("comment", "")
            r["comment"] = f"{c}; collection_image_used".strip("; ")

    stats = {
        "old_rows": len(old_rows),
        "new_rows": len(new_rows),
        "old_variant_rows": variant_rows_old,
        "new_variant_rows": sum(
            1
            for r in new_rows
            if re.search(r"\bvariant\s*\d+\b|\bвариант\s*\d+\b", r.get("model_name", ""), flags=re.I)
        ),
        "real_model_rows": real_model_rows,
        "generated_model_rows": generated_model_rows,
        "rows_with_image": sum(1 for r in new_rows if r.get("image_url", "")),
        "rows_without_image": sum(1 for r in new_rows if not r.get("image_url", "")),
        "repeated_image_gt_10_count": repeated_gt10,
        "needs_review_count": sum(1 for r in new_rows if r.get("status") == "needs_review"),
    }
    return new_rows, stats

