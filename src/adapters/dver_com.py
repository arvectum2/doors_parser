from __future__ import annotations

import json
import re
import time
from urllib.parse import parse_qsl, quote_plus, urlencode, urljoin, urlsplit, urlunsplit

from bs4 import BeautifulSoup

from src.adapters.base import (
    BaseAdapter,
    DoorItem,
    description_of,
    fetch,
    normalize,
    normalize_dver_model_name,
    now_str,
    title_of,
)

NON_DOOR_PREFIXES = (
    "коробка",
    "наличник",
    "добор",
    "погонаж",
    "телескоп",
    "притвор",
    "капитель",
    "карниз",
    "розетка",
    "плинтус",
    "подпятник",
    "профиль",
    "комплект",
    "разрез",
    "сечение",
    "образец",
    "цвет",
    "покрытие",
)

COMPONENT_IMAGE_HINTS = (
    "korob",
    "nalichnik",
    "dobor",
    "profil",
    "pogonazh",
    "komplekt",
    "razrez",
    "sechen",
    "kapitel",
    "karniz",
    "plintus",
    "podpyatnik",
)

PLACEHOLDER_IMAGE_HINTS = (
    "logo",
    "icon",
    "sale",
    "new",
    "hit",
    "catalog.svg",
    "search_loop",
    "pixel.gif",
    "loader",
    "placeholder",
    "no_image",
    "cart",
    "arrow",
)

COLOR_IMAGE_HINTS = (
    "/xml/colors_images/",
    "colors_images",
)

INTERIOR_HINTS = (
    "interior",
    "room",
    "stena",
    "wall",
)

DOOR_HINTS = (
    "door",
    "dver",
    "polotno",
    "/xml/images/",
)


def clean_scambio_model_name(name: str) -> str:
    cleaned, _ = normalize_dver_model_name(name or "")
    cleaned = re.sub(r"\.{2,}", " ", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" \t\r\n.,;:-")
    return cleaned


def is_non_door_model_name(name: str) -> bool:
    n = normalize(name)
    if not n:
        return True
    if n.startswith("дверь "):
        return False
    return any(n == p or n.startswith(p + " ") for p in NON_DOOR_PREFIXES)


def scambio_collection_queries(target_line: str) -> list[str]:
    nline = normalize(target_line)
    mapping = {
        "смальта деко": ["Смальта-Деко"],
        "смальта белла": ["Смальта-Бэлла", "Смальта-Белла"],
        "смальта галант": ["Смальта-Галант"],
        "смальта лайн": ["Смальта-Лайн"],
        "смальта нова": ["Смальта-Нова"],
        "смальта риф": ["Смальта-Риф"],
        "смальта шарм": ["Смальта-Шарм"],
        "пс": ["Серия ПС"],
        "псл": ["Серия ПСЛ"],
        "пск": ["Серия ПСК"],
        "псц модерн": ["Серия ПСЦ модерн"],
        "псц классика": ["Серия ПСЦ классика"],
        "псу": ["Серия ПСУ"],
        "пх алл кромка черный": ["Серия ПХ AL кромка Черный", "Серия ПХ Алл кромка Черный"],
        "пх алл кромка хром": ["Серия ПХ AL кромка Хром", "Серия ПХ Алл кромка Хром"],
        "пст": ["Серия ПСТ"],
        "псе": ["Серия ПСЕ"],
        "пде": ["Серия ПДЕ"],
        "п абс кромка черный": ["Серия П АБС кромка Черный"],
        "п": ["Серия П"],
        "инвизибл": ["Инвизибл", "Invisible"],
    }
    return mapping.get(nline, [target_line])


def _collection_paths_for_line(target_line: str) -> list[str]:
    nline = normalize(target_line)
    if nline.startswith("смальта "):
        return ["https://dver.com/mezhkomnatnye-dveri/dveri-krashenye-emal/"]
    if nline == "инвизибл":
        return ["https://dver.com/mezhkomnatnye-dveri/dveri-pod-pokrasku/"]
    return ["https://dver.com/mezhkomnatnye-dveri/dveri-ekoshpon/"]


def _build_paged_url(base_url: str, page: int) -> str:
    if page <= 1:
        return base_url
    parts = urlsplit(base_url)
    query = dict(parse_qsl(parts.query, keep_blank_values=True))
    query["page"] = str(page)
    query["PAGEN_1"] = str(page)
    new_query = urlencode(query, doseq=True)
    return urlunsplit((parts.scheme, parts.netloc, parts.path, new_query, ""))


def _scambio_image_score(url: str, source: str) -> int:
    u = (url or "").lower()
    score = 0
    if "/xml/images/" in u:
        score += 300
    if source == "jsonld":
        score += 180
    elif source == "img":
        score += 80
    elif source == "meta":
        score += 60
    if u.endswith(".jpg") or u.endswith(".jpeg"):
        score += 25
    if u.endswith(".png"):
        score += 10
    if "resize" in u:
        score -= 40
    return score


def _classify_image_rejection(url: str, alt_text: str = "") -> str | None:
    url_low = (url or "").lower().strip()
    alt_low = (alt_text or "").lower().strip()
    hay = f"{url_low} {alt_low}".strip()
    if any(h in hay for h in COLOR_IMAGE_HINTS):
        return "rejected_color_sample"
    if any(h in hay for h in PLACEHOLDER_IMAGE_HINTS):
        return "rejected_placeholder"
    if any(h in hay for h in COMPONENT_IMAGE_HINTS):
        return "rejected_component_image"
    if any(h in hay for h in INTERIOR_HINTS) and not any(h in hay for h in DOOR_HINTS):
        return "rejected_interior_only"
    if not re.search(r"\.(jpg|jpeg|png|webp)(\?|$)", url_low):
        return "rejected_placeholder"
    return None


class DverComAdapter(BaseAdapter):
    """Scambio-specific parser: parse collection pages one-by-one and collect real models."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.image_audit: list[dict] = []
        self.collection_checks: dict[str, list[str]] = {}

    def _collection_urls(self, target_line: str) -> list[str]:
        urls: list[str] = []
        for path in _collection_paths_for_line(target_line):
            for query in scambio_collection_queries(target_line):
                urls.append(f"{path}?collection={quote_plus(query)}")
        return urls

    def _extract_collection_product_links(self, html: str, page_url: str) -> set[str]:
        soup = BeautifulSoup(html, "lxml")
        links: set[str] = set()
        for a in soup.find_all("a", href=True):
            href = a.get("href", "").strip()
            if not href:
                continue
            full = urljoin(page_url, href)
            if "/mezhkomnatnye-dveri/all/" in full:
                links.add(full)
        return links

    def _collect_product_urls_from_collection(self, collection_url: str) -> list[str]:
        product_urls: set[str] = set()
        stagnant_pages = 0
        max_pages = int(self.cfg.get("collection_max_pages", 10))

        for page in range(1, max_pages + 1):
            page_url = _build_paged_url(collection_url, page)
            html = fetch(self.session, page_url, timeout=(8, 25), retries=6)
            if not html:
                if page > 1:
                    break
                continue
            page_links = self._extract_collection_product_links(html, page_url)
            if not page_links:
                stagnant_pages += 1
                if stagnant_pages >= 2:
                    break
                continue
            new_links = page_links - product_urls
            product_urls.update(page_links)
            if not new_links:
                stagnant_pages += 1
                if stagnant_pages >= 2:
                    break
            else:
                stagnant_pages = 0
        return sorted(product_urls)

    def _extract_image_candidates(self, soup: BeautifulSoup, html: str, page_url: str) -> list[tuple[str, str, str]]:
        candidates: list[tuple[str, str, str]] = []
        seen: set[str] = set()

        # JSON-LD is usually the best source on dver.com.
        for script in soup.find_all("script", attrs={"type": "application/ld+json"}):
            raw = (script.string or script.get_text() or "").strip()
            if not raw:
                continue
            try:
                data = json.loads(raw)
            except Exception:
                continue
            objects = data if isinstance(data, list) else [data]
            for obj in objects:
                image_value = obj.get("image") if isinstance(obj, dict) else None
                if not image_value:
                    continue
                image_values = image_value if isinstance(image_value, list) else [image_value]
                for image_url in image_values:
                    if not isinstance(image_url, str):
                        continue
                    full = urljoin(page_url, image_url.strip())
                    if full and full not in seen:
                        seen.add(full)
                        candidates.append((full, "", "jsonld"))

        for prop in ("og:image", "og:image:secure_url"):
            meta = soup.find("meta", property=prop)
            if meta and meta.get("content"):
                full = urljoin(page_url, meta.get("content", "").strip())
                if full and full not in seen:
                    seen.add(full)
                    candidates.append((full, "", "meta"))

        for match in re.findall(r"/xml/images/[A-Za-z0-9_.-]+\.(?:jpg|jpeg|png|webp)", html, flags=re.I):
            full = urljoin(page_url, match)
            if full not in seen:
                seen.add(full)
                candidates.append((full, "", "regex_xml"))

        for img in soup.find_all("img"):
            src = img.get("data-src") or img.get("src") or img.get("data-original") or ""
            if not src:
                continue
            full = urljoin(page_url, src.strip())
            if not full or full in seen:
                continue
            seen.add(full)
            alt = (img.get("alt") or img.get("title") or "").strip()
            candidates.append((full, alt, "img"))

        return candidates

    def _select_product_image(
        self,
        target_line: str,
        model_name: str,
        candidates: list[tuple[str, str, str]],
    ) -> tuple[str, str, list[str], str]:
        old_image_url = candidates[0][0] if candidates else ""
        rejected_urls: list[str] = []
        rejected_reason = ""
        scored: list[tuple[int, str]] = []

        for url, alt, source in candidates:
            reason = _classify_image_rejection(url, alt)
            if reason:
                rejected_urls.append(url)
                if not rejected_reason:
                    rejected_reason = reason
                continue
            scored.append((_scambio_image_score(url, source), url))

        if scored:
            scored.sort(reverse=True)
            return old_image_url, scored[0][1], rejected_urls, "selected_product_image"

        if rejected_reason:
            return old_image_url, "", rejected_urls, rejected_reason
        return old_image_url, "", rejected_urls, "no_good_image_found"

    def _parse_product(self, url: str, target) -> DoorItem | None:
        html = fetch(self.session, url, timeout=(8, 25), retries=5)
        if not html:
            return None

        soup = BeautifulSoup(html, "lxml")
        raw_name = ""
        h1 = soup.find("h1")
        if h1:
            raw_name = h1.get_text(" ", strip=True)
        if not raw_name:
            raw_name = title_of(soup)
        model_name = clean_scambio_model_name(raw_name)
        if not model_name or is_non_door_model_name(model_name):
            return None

        text = soup.get_text(" ", strip=True)
        desc = description_of(soup) or re.sub(r"\s+", " ", text)[:700]
        candidates = self._extract_image_candidates(soup, html, url)
        old_image_url, selected_image_url, rejected, reason = self._select_product_image(
            target.target_line, model_name, candidates
        )

        status = "found" if selected_image_url else "needs_review"
        comment_parts = [f"collection_page_source: {target.target_line}"]
        if status != "found":
            comment_parts.append(reason)

        self.image_audit.append(
            {
                "target_line": target.target_line,
                "model_name": model_name,
                "old_image_url": old_image_url,
                "selected_image_url": selected_image_url,
                "rejected_image_urls": " | ".join(rejected),
                "image_status": "found" if selected_image_url else "needs_review",
                "reason": reason,
            }
        )

        material = target.covering_from_source or target.material_manual or ""
        return DoorItem(
            door_type=target.door_type if target.door_type in ("входная", "межкомнатная") else self.door_type_default,
            manufacturer=target.manufacturer,
            line=target.target_line,
            model_name=model_name,
            description=desc,
            material=material,
            coating="",
            image_url=selected_image_url,
            source_url=url,
            status=status,
            comment="; ".join(comment_parts),
        )

    def run(self) -> tuple[list[DoorItem], dict]:
        items: list[DoorItem] = []
        log = {
            "base_url": self.base_url,
            "candidate_pages": 0,
            "pages_processed": 0,
            "models_found": 0,
            "targets": {},
            "errors": [],
            "collection_checks": {},
            "image_audit": self.image_audit,
        }

        print(f"  [{now_str()}] Scambio collection-driven parsing...")
        for target in self.targets:
            checked_urls = self._collection_urls(target.target_line)
            self.collection_checks[target.target_line] = checked_urls
            log["collection_checks"][target.target_line] = checked_urls

            target_links: set[str] = set()
            for collection_url in checked_urls:
                try:
                    links = self._collect_product_urls_from_collection(collection_url)
                    target_links.update(links)
                except Exception as exc:
                    log["errors"].append(f"{target.target_line}: {collection_url}: {exc}")
                time.sleep(0.1)

            log["candidate_pages"] += len(target_links)
            if self.debug:
                print(f"    [{now_str()}] {target.target_line}: {len(target_links)} product URLs")

            count_before = len(items)
            for product_url in sorted(target_links):
                item = self._parse_product(product_url, target)
                log["pages_processed"] += 1
                if item is not None:
                    items.append(item)
                time.sleep(0.02)
            log["targets"][target.target_line] = len(items) - count_before
            time.sleep(0.2)

        # Keep strict uniqueness by manufacturer + line + model + source URL.
        seen: set[tuple[str, str, str, str]] = set()
        unique: list[DoorItem] = []
        for item in items:
            key = (
                normalize(item.manufacturer),
                normalize(item.line),
                normalize(item.model_name),
                normalize(item.source_url),
            )
            if key in seen:
                continue
            seen.add(key)
            unique.append(item)

        log["models_found"] = len(unique)
        found_count = sum(1 for line, cnt in log["targets"].items() if cnt > 0)
        print(f"  [{now_str()}] done: {len(unique)} models, {found_count}/{len(self.targets)} lines found")
        return unique, log
