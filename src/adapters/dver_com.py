from __future__ import annotations

import json
import re
import time
import xml.etree.ElementTree as ET
from collections import Counter
from urllib.parse import parse_qsl, quote_plus, urlencode, urljoin, urlsplit, urlunsplit

from bs4 import BeautifulSoup

from src.adapters.base import (
    BaseAdapter,
    DoorItem,
    TargetLine,
    description_of,
    fetch,
    normalize,
    normalize_dver_model_name,
    now_str,
    title_of,
)

NON_DOOR_PREFIXES = (
    "комплект",
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
    "badge",
    "label",
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
    "open.png",
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

MODEL_MARKERS = (
    "дверь",
    "смальта",
    "пс",
    "пск",
    "псл",
    "псц",
    "псу",
    "пх",
    "пст",
    "псе",
    "пде",
    "п абс",
    "инвиз",
    "invisible",
    "скрыт",
)

COLOR_WORDS = (
    "дуб",
    "ясень",
    "венге",
    "белый",
    "беж",
    "серый",
    "графит",
    "жемчуг",
    "слоновая кость",
    "жасмин",
    "пралине",
    "магнолия",
    "агат",
    "зефир",
    "шеллгрей",
    "матовый",
    "ral",
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


def _extract_code(text: str, prefix: str) -> str | None:
    m = re.search(rf"(?<![a-zа-я0-9]){prefix}\s*[-–—]?\s*([0-9]+(?:\.[0-9]+)?)\b", text, re.I)
    return m.group(1) if m else None


def _extract_smalta_number(text: str) -> str | None:
    m = re.search(
        r"смальта\s*[-–—]?\s*(?:деко|б[еэ]лла|галант|лайн|нова|риф|шарм)?\s*[-–—]?\s*([0-9]+(?:\.[0-9]+)?)\b",
        text,
        re.I,
    )
    return m.group(1) if m else None


def _strict_has(pattern: str, text: str) -> bool:
    return re.search(pattern, text, re.I) is not None


def infer_scambio_target_line(model: str, title: str, desc: str, source_url: str) -> tuple[str | None, bool, str]:
    text = f"{model} {title} {desc} {source_url}".replace("ё", "е")
    low = text.lower().replace("xром", "хром")

    direct_map = [
        (r"смальта\s*[-–—]?\s*деко", "Смальта Деко"),
        (r"смальта\s*[-–—]?\s*б[еэ]лла", "Смальта Белла"),
        (r"смальта\s*[-–—]?\s*галант", "Смальта Галант"),
        (r"смальта\s*[-–—]?\s*лайн", "Смальта Лайн"),
        (r"смальта\s*[-–—]?\s*нова", "Смальта Нова"),
        (r"смальта\s*[-–—]?\s*риф", "Смальта Риф"),
        (r"смальта\s*[-–—]?\s*шарм", "Смальта Шарм"),
    ]
    for pat, line in direct_map:
        if _strict_has(pat, low):
            return line, True, "name_pattern_match"

    if _strict_has(r"(?<![a-zа-я0-9])инвиз|invisible|скрыт|под\s+покраск", low):
        return "Инвизибл", True, "invisible_pattern_match"

    if _strict_has(r"(?<![a-zа-я0-9])псц(?![a-zа-я0-9])", low):
        if _strict_has(r"модерн", low):
            return "ПСЦ Модерн", True, "psc_modern_pattern_match"
        if _strict_has(r"классик", low):
            return "ПСЦ Классика", True, "psc_classic_pattern_match"
        psc_num = _extract_code(low, "псц")
        if psc_num:
            try:
                num_val = float(psc_num)
            except ValueError:
                num_val = 0.0
            return ("ПСЦ Классика" if num_val < 20 else "ПСЦ Модерн"), False, "psc_numeric_heuristic"

    if _strict_has(r"п\s*абс.*черн|абс\s*кромк.*черн|кромк[аи]\s*абс.*черн", low):
        return "П АБС кромка Черный", True, "p_abs_black_pattern_match"

    if _strict_has(r"(?<![a-zа-я0-9])пх(?![a-zа-я0-9])", low):
        if _strict_has(r"хром", low):
            return "ПХ Алл кромка Хром", True, "ph_chrome_pattern_match"
        if _strict_has(r"черн", low):
            return "ПХ Алл кромка Черный", True, "ph_black_pattern_match"
        if _strict_has(r"\bal\b|\bалл\b|ал\.\s*кромк|кромк", low):
            return "ПХ Алл кромка Хром", True, "ph_chrome_default_pattern_match"
        return "ПХ Алл кромка Хром", False, "ph_default_heuristic"

    strict_series_patterns = [
        (r"(?<![a-zа-я0-9])псл(?![a-zа-я0-9])", "ПСЛ"),
        (r"(?<![a-zа-я0-9])пск(?![a-zа-я0-9])", "ПСК"),
        (r"(?<![a-zа-я0-9])псу(?![a-zа-я0-9])", "ПСУ"),
        (r"(?<![a-zа-я0-9])пст(?![a-zа-я0-9])", "ПСТ"),
        (r"(?<![a-zа-я0-9])псе(?![a-zа-я0-9])", "ПСЕ"),
        (r"(?<![a-zа-я0-9])пде(?![a-zа-я0-9])", "ПДЕ"),
    ]
    for pat, line in strict_series_patterns:
        if _strict_has(pat, low):
            return line, True, "series_pattern_match"

    if _strict_has(r"(?<![a-zа-я0-9])пс(?![a-zа-я0-9лкцует])", low):
        return "ПС", True, "ps_pattern_match"

    if _strict_has(r"(?<![a-zа-я0-9])п(?=\s*[-–—]?\s*\d)", low):
        return "П", True, "p_pattern_match"

    sm_num = _extract_smalta_number(low)
    if sm_num:
        if sm_num in {"07", "08", "7", "8"}:
            return "Смальта Лайн", False, "smalta_numeric_heuristic"
        if sm_num in {"201", "202", "203", "204", "208.1", "208.2", "209.2", "210"}:
            return "Смальта Риф", False, "smalta_numeric_heuristic"
        if sm_num in {"1", "01", "2", "02"}:
            return "Смальта Шарм", False, "smalta_numeric_heuristic"
        if sm_num in {"12"}:
            return "Смальта Белла", False, "smalta_numeric_heuristic"
        if sm_num in {"104"}:
            return "Смальта Галант", False, "smalta_numeric_heuristic"
        if sm_num in {"14"}:
            return "Смальта Нова", False, "smalta_numeric_heuristic"
        return "Смальта Деко", False, "smalta_numeric_heuristic"

    return None, False, ""


def _model_has_identity(model_name: str) -> bool:
    low = normalize(model_name)
    return any(marker in low for marker in MODEL_MARKERS) or bool(re.search(r"\b[пps]{1,3}\s*[-–—]?\s*\d", low, re.I))


def _looks_color_only(model_name: str) -> bool:
    low = normalize(model_name)
    if not low:
        return True
    if _model_has_identity(low):
        return False
    if len(low.split()) > 6:
        return False
    return any(word in low for word in COLOR_WORDS)


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
    elif source == "regex_xml":
        score += 40
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
    """Scambio parser that scans dver.com sitemap/product pages and maps lines by robust rules."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.image_audit: list[dict] = []
        self.collection_checks: dict[str, list[str]] = {}
        self.removed_components_count = 0
        self.rejected_bad_images_count = 0

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
            href = (a.get("href") or "").strip()
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

    def _collect_sitemap_product_urls(self) -> list[str]:
        sitemap_url = "https://dver.com/xml/sitemap.xml"
        text = fetch(self.session, sitemap_url, timeout=(8, 30), retries=6)
        if not text:
            return []
        try:
            root = ET.fromstring(text.encode("utf-8"))
        except Exception:
            return []
        urls: list[str] = []
        for node in root.iter():
            if node.tag.endswith("loc"):
                u = (node.text or "").strip()
                if not u:
                    continue
                if "/mezhkomnatnye-dveri/all/" not in u:
                    continue
                urls.append(u)
        return sorted(set(urls))

    def _extract_jsonld_names(self, soup: BeautifulSoup) -> list[str]:
        names: list[str] = []
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
                if isinstance(obj, dict):
                    nm = obj.get("name")
                    if isinstance(nm, str) and nm.strip():
                        names.append(nm.strip())
        return names

    def _extract_image_candidates(self, soup: BeautifulSoup, html: str, page_url: str) -> list[tuple[str, str, str]]:
        candidates: list[tuple[str, str, str]] = []
        seen: set[str] = set()

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
                self.rejected_bad_images_count += 1
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

    def _collect_candidate_product_urls(self, targets: list[TargetLine]) -> tuple[list[str], dict[str, list[str]]]:
        checks: dict[str, list[str]] = {}
        for t in targets:
            checks[t.target_line] = ["https://dver.com/xml/sitemap.xml"] + self._collection_urls(t.target_line)

        urls: set[str] = set(self._collect_sitemap_product_urls())
        for t in targets:
            for collection_url in self._collection_urls(t.target_line):
                try:
                    urls.update(self._collect_product_urls_from_collection(collection_url))
                except Exception:
                    continue
                time.sleep(0.05)
        return sorted(urls), checks

    def _parse_product(self, url: str, target_map: dict[str, TargetLine]) -> DoorItem | None:
        html = fetch(self.session, url, timeout=(8, 25), retries=6)
        if not html:
            return None

        soup = BeautifulSoup(html, "lxml")
        title = title_of(soup) or ""
        h1 = soup.find("h1")
        h1_text = h1.get_text(" ", strip=True) if h1 else ""
        jsonld_names = self._extract_jsonld_names(soup)
        jsonld_name = jsonld_names[0] if jsonld_names else ""
        desc = description_of(soup)
        if not desc:
            desc = re.sub(r"\s+", " ", soup.get_text(" ", strip=True))[:1200]

        base_name = h1_text or jsonld_name or title
        model_name = clean_scambio_model_name(base_name)
        if is_non_door_model_name(model_name):
            self.removed_components_count += 1
            return None

        line, strong, infer_reason = infer_scambio_target_line(model_name, title, desc, url)
        if not line:
            return None

        target = target_map.get(normalize(line))
        if not target:
            return None

        status = "found"
        comment_parts = [f"line_infer={infer_reason or 'n/a'}"]

        if not model_name:
            model_name = f"{line} — вариант"
            status = "needs_review"
            comment_parts.append("model_name_generated_from_collection_variant")
        elif not _model_has_identity(model_name):
            if _looks_color_only(model_name):
                model_name = f"{line} — {model_name}"
                status = "needs_review"
                comment_parts.append("model_name_generated_from_collection_variant")
            else:
                status = "needs_review"

        candidates = self._extract_image_candidates(soup, html, url)
        old_image_url, selected_image_url, rejected, reason = self._select_product_image(candidates)
        if not selected_image_url:
            status = "needs_review"
            comment_parts.append("image_not_found")
            comment_parts.append(reason)
        elif reason != "selected_product_image":
            status = "needs_review"
            comment_parts.append(reason)

        if not strong:
            status = "needs_review"
            comment_parts.append("heuristic_line_assignment")

        self.image_audit.append(
            {
                "target_line": line,
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
            line=line,
            model_name=model_name,
            description=desc,
            material=material,
            coating="",
            image_url=selected_image_url,
            source_url=url,
            status=status,
            comment="; ".join(dict.fromkeys(comment_parts)),
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
            "removed_components_count": 0,
            "rejected_bad_images_count": 0,
            "line_infer_stats": {},
        }

        target_map = {normalize(t.target_line): t for t in self.targets}
        for t in self.targets:
            log["targets"][t.target_line] = 0

        print(f"  [{now_str()}] Scambio sitemap/product parsing...")
        candidate_urls, checks = self._collect_candidate_product_urls(self.targets)
        self.collection_checks = checks
        log["collection_checks"] = checks
        log["candidate_pages"] = len(candidate_urls)
        print(f"  [{now_str()}] candidate product URLs: {len(candidate_urls)}")

        infer_counter = Counter()
        for idx, product_url in enumerate(candidate_urls, 1):
            try:
                item = self._parse_product(product_url, target_map)
            except Exception as exc:
                log["errors"].append(f"{product_url}: {exc}")
                continue
            log["pages_processed"] += 1
            if item is None:
                continue
            items.append(item)
            log["targets"][item.line] = log["targets"].get(item.line, 0) + 1
            line_infer = ""
            for part in (item.comment or "").split(";"):
                part = part.strip()
                if part.startswith("line_infer="):
                    line_infer = part.split("=", 1)[1].strip()
                    break
            infer_counter[line_infer or "n/a"] += 1
            if self.debug and idx % 150 == 0:
                print(f"    [{now_str()}] processed {idx}/{len(candidate_urls)}")
            time.sleep(0.01)

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
        log["removed_components_count"] = self.removed_components_count
        log["rejected_bad_images_count"] = self.rejected_bad_images_count
        log["line_infer_stats"] = dict(infer_counter)
        found_count = sum(1 for _, cnt in log["targets"].items() if cnt > 0)
        print(f"  [{now_str()}] done: {len(unique)} models, {found_count}/{len(self.targets)} lines found")
        return unique, log
