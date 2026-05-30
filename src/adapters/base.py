from __future__ import annotations

import re
import time
import xml.etree.ElementTree as ET
from dataclasses import dataclass, asdict
from datetime import datetime
from typing import Optional
from urllib.parse import urljoin, urlparse, urldefrag, unquote_plus

import requests
from bs4 import BeautifulSoup
from rapidfuzz import fuzz


@dataclass
class TargetLine:
    door_type: str
    manufacturer: str
    target_line: str
    material_manual: str = ""
    covering_from_source: str = ""
    source_site: str = ""
    source_note: str = ""


@dataclass
class DoorItem:
    door_type: str
    manufacturer: str
    line: str
    model_name: str
    description: str
    material: str
    coating: str
    image_url: str
    source_url: str
    status: str
    comment: str


# ── Shared constants ──────────────────────────────────────────────────────────

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "ru,en;q=0.9",
}

SKIP_URL_PARTS = [
    "/cart", "/basket", "/compare", "/login", "/auth", "/search",
    "/news", "/blog", "/contacts", "/about", "/delivery", "/pay",
    "/policy", "/agreement", "/review", "/reviews",
    "/zvonok", "/gde_kupit", "/kompaniya", "/partneram", "/furnitura",
    "/mramor", "/plintusa", "/tekhnicheskoe", "/dizajneram", "/prorabov",
    "/dileram", "/zastroyshchikam", "/komplektovshchikam",
    "?", "#",
]

CATALOG_URL_MARKERS = [
    "catalog", "katalog", "category", "product", "tovar", "dver",
    "door", "mezh", "vhod", "collection", "model", "seria", "series",
]

SKIP_TITLE_PATTERNS = [
    "заказать обратный звонок", "где купить", "о компании",
    "фурнитура", "партнерам", "стать дилером", "акции и скидки",
    "корзина", "оформление заказа", "доставка", "оплата",
    "контакты", "отзывы", "сертификат",
]

SITEMAP_ROOTS = [
    "/sitemap.xml", "/sitemap_index.xml", "/wp-sitemap.xml",
    "/sitemap_index.xml.gz", "/sitemap.xml.gz",
]

BAD_IMAGE_KEYWORDS = [
    "logo", "icon", "pixel", "placeholder", "sprite", "loader",
    "blank", "no-photo", "nophoto", "favicon", "banner",
    "payment", "social", "rutub", "capcha", "captcha",
    "open.png", "close.png", "play.png",
    "arrow_up", "arrow", "catalog.svg", "search_loop", "colors_images",
    "star", "novinka", "printest", "tele2", "icon_ft_soc",
    ".svg", "yandex", "mc.yandex", "watch/", "no_image_glass",
]

GENERIC_TITLES = [
    "межкомнатные двери", "входные двери", "каталог",
    "главная", "главная страница",
]


# ── Utilities ─────────────────────────────────────────────────────────────────

_CYRILLIC_TO_LATIN = {
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e",
    "ё": "e", "ж": "zh", "з": "z", "и": "i", "й": "i", "к": "k",
    "л": "l", "м": "m", "н": "n", "о": "o", "п": "p", "р": "r",
    "с": "s", "т": "t", "у": "u", "ф": "f", "х": "kh", "ц": "ts",
    "ч": "ch", "ш": "sh", "щ": "shch", "ъ": "", "ы": "y", "ь": "",
    "э": "e", "ю": "yu", "я": "ya",
}


def transliterate(text: str) -> str:
    result = ""
    for c in (text or "").lower():
        result += _CYRILLIC_TO_LATIN.get(c, c)
    return result


def normalize(text: str) -> str:
    text = (text or "").lower().replace("ё", "е")
    text = re.sub(r"[^a-zа-я0-9]+", " ", text, flags=re.I)
    return re.sub(r"\s+", " ", text).strip()


def is_same_domain(url: str, base: str) -> bool:
    a = urlparse(url).netloc.replace("www.", "")
    b = urlparse(base).netloc.replace("www.", "")
    return a == b or a.endswith("." + b)


def clean_url(base_url: str, href: str) -> str:
    url = urljoin(base_url, href)
    url, _ = urldefrag(url)
    return url.strip()


def now_str() -> str:
    return datetime.now().strftime("%H:%M:%S")


def fetch(session: requests.Session, url: str, timeout: tuple = (5, 15), retries: int = 3) -> str | None:
    last_err = None
    for attempt in range(retries):
        try:
            r = session.get(url, headers=HEADERS, timeout=timeout, allow_redirects=True)
            if r.status_code in {429, 502, 503, 504}:
                if attempt < retries - 1:
                    time.sleep(2 * (attempt + 1))
                    continue
                return None
            if r.status_code >= 400:
                return None
            ctype = r.headers.get("Content-Type", "")
            if "text" not in ctype and "html" not in ctype and "xml" not in ctype and not url.endswith(".xml"):
                return None
            if not r.encoding or r.encoding.lower() in {"iso-8859-1", "ascii"}:
                r.encoding = r.apparent_encoding
            return r.text
        except (requests.ConnectionError, requests.Timeout, requests.RequestException) as e:
            last_err = e
            if attempt < retries - 1:
                time.sleep(1 * (attempt + 1))
    return None


def text_of(soup: BeautifulSoup) -> str:
    for tag in soup(["script", "style", "noscript", "svg"]):
        tag.decompose()
    return soup.get_text(" ", strip=True)


def title_of(soup: BeautifulSoup) -> str:
    for sel in ["h1", ".product_title", ".product-title",
                ".catalog-element-title", "[itemprop=name]"]:
        n = soup.select_one(sel)
        if n:
            t = n.get_text(" ", strip=True)
            if t:
                return t
    og = soup.find("meta", property="og:title")
    if og and og.get("content"):
        return og["content"].strip()
    return soup.title.get_text(" ", strip=True) if soup.title else ""


def description_of(soup: BeautifulSoup) -> str:
    for selector in [
        "[itemprop=description]", ".product-description", ".product__description",
        ".description", ".detail_text", ".tab-content", ".product-tabs",
        ".element-description", ".prod-text", ".catalog-element-description",
    ]:
        node = soup.select_one(selector)
        if node:
            txt = node.get_text(" ", strip=True)
            txt = re.sub(r"\s+", " ", txt)
            if len(txt) > 40:
                return txt[:3000]
    meta = soup.find("meta", attrs={"name": "description"})
    if meta and meta.get("content"):
        return meta["content"].strip()
    og = soup.find("meta", property="og:description")
    if og and og.get("content"):
        return og["content"].strip()
    return ""


def image_of(soup: BeautifulSoup, page_url: str) -> tuple[str, str]:
    """Return (image_url, status). status = 'found'/'missing'/'bad'."""
    for prop in ["og:image", "og:image:secure_url"]:
        og = soup.find("meta", property=prop)
        if og and og.get("content"):
            url = urljoin(page_url, og["content"].strip())
            if not _is_bad_image(url):
                return url, "found"
    candidates: list[tuple[int, str]] = []
    for img in soup.find_all("img"):
        src = img.get("src") or img.get("data-src") or img.get("data-original") or img.get("data-lazy")
        if not src:
            continue
        if "," in src and " " in src:
            src = src.split(",")[0].split()[0]
        url = urljoin(page_url, src)
        low = normalize(url)
        alt = normalize(img.get("alt", ""))
        klass = normalize(" ".join(img.get("class", [])))
        if _is_bad_image(url, alt, klass, low):
            continue
        score = 0
        if any(x in low for x in ["product", "catalog", "dver", "door", "images", "upload", "resize_cache"]):
            score += 4
        if any(x in klass for x in ["product", "catalog", "photo", "image", "picture", "detail"]):
            score += 3
        if any(x in alt for x in ["двер", "door", "модель"]):
            score += 2
        if re.search(r"\.(jpg|jpeg|png|webp)(\?|$)", url.lower()):
            score += 1
        if score > 0:
            candidates.append((score, url))
    candidates.sort(reverse=True)
    if candidates:
        return candidates[0][1], "found"
    return "", "missing"


def _is_bad_image(url: str = "", alt: str = "", klass: str = "", low: str = "") -> bool:
    hay = (url + " " + alt + " " + klass + " " + low).lower()
    return any(k in hay for k in BAD_IMAGE_KEYWORDS)


def looks_productish(url: str, title: str, text: str, desc: str = "") -> bool:
    h = normalize(title + " " + desc[:1500] + " " + text[:3000] + " " + url)
    markers = ["описание", "характерист", "полотно", "двер", "модель",
               "размер", "цвет", "покрытие", "короб", "наличник", "толщина",
               "серия", "коллекция", "купить"]
    return sum(1 for m in markers if m in h) >= 3


def is_generic_page(title: str, url: str) -> bool:
    ntitle = normalize(title)
    path = urlparse(url).path.strip("/")
    # Homepage or empty path
    if not path:
        return True
    # Title matches a generic category name exactly or as full title
    for g in GENERIC_TITLES:
        if normalize(g) == ntitle:
            return True
    return False


def normalize_dver_model_name(name: str) -> tuple[str, str]:
    """Clean dver.com model name: remove sizes and 'Распродажа'.
    Returns (cleaned_name, original_name).
    """
    orig = name
    t = name
    # Remove "Распродажа" (any case, with/without punctuation)
    t = re.sub(r"\s*[–—-]?\s*распродажа\s*[.,;:!?]*", " ", t, flags=re.IGNORECASE)

    # Normalize multiplication symbols and spacing for dimensions
    t = t.replace("×", "x").replace("Х", "x").replace("х", "x").replace("*", "x")

    # Remove "размер 200.7 x 71.3", "размер 200 x 80", "размер: 200x80"
    t = re.sub(
        r"\bразмер\s*[:\-]?\s*\d[\d.]*\s*x\s*\d[\d.]*(\s*x\s*\d[\d.]*)?\b",
        " ",
        t,
        flags=re.IGNORECASE,
    )

    # Remove standalone dimensions: "200 x 80", "200x80", "2000x800", "200.7x71.3"
    t = re.sub(r"\b\d[\d.]*\s*x\s*\d[\d.]*(\s*x\s*\d[\d.]*)?\b", " ", t, flags=re.IGNORECASE)

    # Remove leftover tokens like "x80" or "200x"
    t = re.sub(r"\b\d[\d.]*\s*x\b|\bx\s*\d[\d.]*\b", " ", t, flags=re.IGNORECASE)

    # Remove "цвет ..." after dimension removal (цвет белое матовое)
    t = re.sub(r"\s*цвет\s+[\w\s]+$", " ", t, flags=re.IGNORECASE)

    # Remove duplicated separators after cleanup
    t = re.sub(r"\s*[|/\\]+\s*", " ", t)
    # Clean up multiple spaces, trailing separators
    t = re.sub(r"\s+", " ", t).strip().rstrip(",;:-")
    return t, orig


# ── Short / generic line names that need strict matching ──────────────────────

SHORT_LINE_NAMES = {
    "п", "пг", "по", "пс", "пст", "псе", "пде", "псц", "псу",
    "пх", "псл", "пск", "пабс", "инвизибл", "invisible",
    "стиль", "классика", "лофт", "модерн", "экошпон", "эмаль",
    "new line", "slope", "true", "gift", "art", "premier",
    "malta", "neo", "molding", "turin", "венир", "книжки",
    "реально",
}


def _word_boundary(needle: str, hay: str) -> bool:
    """Check if needle appears as a whole word in hay (splitting on non-alpha chars)."""
    hay_lower = hay.lower()
    for sep in ("/", "_", "-", " ", "\\", ".", "+", "&"):
        hay_lower = hay_lower.replace(sep, " ")
    words = hay_lower.split()
    return needle in words


def matches_line(line: str, title: str, text: str, url: str) -> tuple[bool, str]:
    """Match line against page title/URL.
    
    Strategy:
    1. Short names (≤4 chars / SHORT_LINE_NAMES): word-boundary in title or URL (both
       original Cyrillic and Latin transliteration), plus fuzzy fallback → needs_review
    2. Longer names: substring in title/URL → found, else fuzzy ≥85 → needs_review
    """
    nline = normalize(line)
    if not nline:
        return False, ""

    ntitle = normalize(title)
    nurl = normalize(url)
    tline = transliterate(line)

    # ── SHORT NAMES (≤4 chars or in SHORT_LINE_NAMES) ──
    uurl = unquote_plus(url)
    if len(nline) <= 4 or nline in SHORT_LINE_NAMES:
        # Exact word boundary in original title or decoded URL (Cyrillic)
        if _word_boundary(nline, title) or _word_boundary(nline, uurl):
            return True, "found"
        # Transliterated word boundary in decoded URL or title (2+ chars only)
        if len(tline) >= 2:
            if _word_boundary(tline, uurl) or _word_boundary(tline, transliterate(title or "")):
                return True, "found"
        # 1-char: substring in ntitle
        if len(nline) == 1 and nline in ntitle:
            return True, "found"
        # Fuzzy fallback for short names → needs_review (catches Latin/Cyrillic mismatches)
        hay = ntitle + " " + nurl
        if fuzz.partial_ratio(nline, hay) >= 85:
            return True, "needs_review"
        # Transliterated fuzzy (for 2+ char lines that use Latin in title/URL)
        if len(tline) >= 2:
            thay = normalize(transliterate(title or "") + " " + uurl.lower())
            if fuzz.partial_ratio(tline, thay) >= 85:
                return True, "needs_review"
        return False, ""

    # ── LONGER NAMES ──
    # Substring in normalized title or URL
    if nline in ntitle or nline in nurl:
        return True, "found"
    # Multi-word: all words in title
    lwords = [normalize(w) for w in line.split() if normalize(w)]
    twords = [normalize(w) for w in title.split() if normalize(w)]
    if lwords and all(w in twords for w in lwords):
        return True, "found"

    # Fuzzy on title+URL
    hay = ntitle + " " + nurl
    ratio = fuzz.partial_ratio(nline, hay)
    if ratio >= 85:
        return True, "needs_review"
    # Transliterated fuzzy (for mixed Latin/Cyrillic cases)
    thay = normalize(transliterate(title or "") + " " + uurl.lower())
    if fuzz.partial_ratio(tline, thay) >= 85:
        return True, "needs_review"

    return False, ""


# ── Sitemap ───────────────────────────────────────────────────────────────────

def parse_sitemap(session: requests.Session, sitemap_url: str, base_url: str, limit: int, seen: set[str]) -> list[str]:
    if sitemap_url in seen or len(seen) > 200:
        return []
    seen.add(sitemap_url)
    text = fetch(session, sitemap_url)
    if not text:
        return []
    urls: list[str] = []
    try:
        root = ET.fromstring(text.encode("utf-8"))
    except Exception:
        return []
    ns = ""
    if root.tag.startswith("{"):
        ns = root.tag.split("}")[0] + "}"
    for loc in root.iter(f"{ns}loc"):
        u = (loc.text or "").strip()
        if not u:
            continue
        if u.endswith(".xml") and "sitemap" in u.lower():
            urls.extend(parse_sitemap(session, u, base_url, limit, seen))
        elif is_same_domain(u, base_url):
            if not any(x in u.lower() for x in SKIP_URL_PARTS):
                urls.append(u)
        if len(urls) >= limit:
            break
    return list(dict.fromkeys(urls))[:limit]


def sitemap_urls(session: requests.Session, base_url: str, limit: int) -> list[str]:
    roots = [urljoin(base_url, p) for p in SITEMAP_ROOTS]
    found: list[str] = []
    seen = set()
    for sm in roots:
        found.extend(parse_sitemap(session, sm, base_url, limit, seen))
    return list(dict.fromkeys(found))[:limit]


# ── Adapter base class ────────────────────────────────────────────────────────

class BaseAdapter:
    def __init__(
        self,
        session: requests.Session,
        cfg: dict,
        targets: list[TargetLine],
        door_type_default: str = "межкомнатная",
        max_pages: int = 200,
        debug: bool = False,
    ):
        self.session = session
        self.cfg = cfg
        self.targets = targets
        self.door_type_default = door_type_default
        self.max_pages = max_pages
        self.debug = debug
        self.base_url = cfg["base_url"]

    def log(self, msg: str) -> None:
        if self.debug:
            print(msg)

    def collect_urls(self) -> list[str]:
        urls = []
        seeds = self.cfg.get("seeds", [self.base_url])
        use_sitemap = self.cfg.get("sitemap", True)
        max_seed_pages = self.cfg.get("max_seed_pages", 10)

        direct = self.cfg.get("direct_urls", [])
        if direct:
            urls.extend(direct)

        if use_sitemap:
            print(f"    [{now_str()}] collecting sitemap URLs...")
            sm = sitemap_urls(self.session, self.base_url, self.max_pages)
            if self.debug:
                print(f"    [{now_str()}] sitemap: {len(sm)} URLs")
            urls.extend(sm)

        print(f"    [{now_str()}] scanning seed pages for catalog/product links...")
        seed_links = set()

        def scrape_links(page_url: str) -> set[str]:
            links = set()
            html = fetch(self.session, page_url)
            if not html:
                return links
            soup = BeautifulSoup(html, "lxml")
            for a in soup.find_all("a", href=True):
                href = clean_url(page_url, a.get("href", ""))
                if not is_same_domain(href, self.base_url):
                    continue
                low = href.lower()
                if any(x in low for x in SKIP_URL_PARTS):
                    continue
                if any(x in low for x in CATALOG_URL_MARKERS):
                    links.add(href)
            return links

        for seed in seeds:
            seed_links.update(scrape_links(seed))
            # Follow pagination: ?page=2, ?page=3, etc.
            base_seed = seed.split("?")[0]
            for page_n in range(2, max_seed_pages + 1):
                page_url = f"{base_seed}?page={page_n}" if "?" not in seed else f"{seed}&page={page_n}"
                page_links = scrape_links(page_url)
                if not page_links:
                    break  # no more pages
                seed_links.update(page_links)

        if self.debug:
            print(f"    [{now_str()}] seed links: {len(seed_links)}")
        # Prioritise: seeds first, then seed links, then sitemap
        all_urls = list(dict.fromkeys(list(seeds) + list(seed_links) + urls))
        all_urls = [u for u in all_urls if is_same_domain(u, self.base_url)]
        all_urls = [u for u in all_urls if not any(x in u.lower() for x in SKIP_URL_PARTS)]
        if not all_urls:
            all_urls = list(seeds)
        urls = all_urls[:self.max_pages]
        print(f"    [{now_str()}] total candidate URLs: {len(urls)}")
        return urls

    def extract_gallery_images(self, soup: BeautifulSoup, url: str) -> list[tuple[str, str]]:
        """Extract multiple door image URLs and alt texts from a gallery/collection page.
        Returns list of (image_url, alt_text) for each distinct door variant image."""
        gallery_images = []
        seen_urls = set()

        # Priority 1: known gallery/collection class patterns
        for img in soup.find_all("img"):
            src = img.get("data-src") or img.get("src") or ""
            alt = img.get("alt", "").strip()
            klass = " ".join(img.get("class", []))
            parent = img.parent
            parent_klass = " ".join(parent.get("class", [])) if parent else ""
            low_src = src.lower()
            if not src or _is_bad_image(src, alt, klass, normalize(url)):
                continue
            full_url = urljoin(url, src)
            if full_url in seen_urls:
                continue

            is_door_image = (
                "max-h-[300px]" in klass                           # Aelita
                or "colections" in low_src                         # Aelita
                or "radio-box" in klass or "radio-box__img" in klass  # Lord
                or "t776__img" in klass or "t760__img" in klass    # Hausdoors
                or "img-full-gallery" in klass                     # Entro gallery
                or "em-click" in klass                             # Entro thumbnails
                or "clore" in parent_klass                         # Entro thumbnail grid
                or "product_img" in low_src                        # Generic
                or ("upload" in low_src and any(x in low_src for x in ["dver", "door", "catalog", "product", "model"]))
                or ("media" in low_src and any(x in low_src for x in ["dver", "door", "catalog", "product"]))
            )
            if not is_door_image:
                continue
            seen_urls.add(full_url)
            gallery_images.append((full_url, alt))

        # Priority 2: fallback — any images from upload/media/images directories
        # Always run to catch what Priority 1 might have missed
        for img in soup.find_all("img"):
            src = img.get("data-src") or img.get("src") or ""
            alt = img.get("alt", "").strip()
            if not src or _is_bad_image(src, alt, " ".join(img.get("class", [])), ""):
                continue
            full_url = urljoin(url, src)
            if full_url in seen_urls:
                continue
            low_src = src.lower()
            if any(x in low_src for x in ["upload", "media", "images", "product", "dver", "doors", "resize_cache"]):
                seen_urls.add(full_url)
                gallery_images.append((full_url, alt))

        return gallery_images

    def extract_from_page(self, html: str, url: str) -> list[DoorItem]:
        soup = BeautifulSoup(html, "lxml")
        title = title_of(soup)
        text = text_of(soup)
        if not title:
            return []

        ntitle = normalize(title)
        for skip in SKIP_TITLE_PATTERNS:
            if skip in ntitle:
                return []

        desc = description_of(soup)
        if not desc:
            desc = re.sub(r"\s+", " ", text)[:700]

        found_items = []
        line_url_map = self.cfg.get("line_url_map", None)
        matched_url_lines: set[str] = set()
        is_mapped = line_url_map and url in line_url_map
        if is_mapped:
            matched_url_lines.add(normalize(line_url_map[url]))

        if not looks_productish(url, title, text, desc):
            if not is_mapped:
                return []

        if is_generic_page(title, url):
            if not is_mapped:
                return []

        strict_match = self.cfg.get("strict_match", False)

        # Check if this page has a gallery — extract multiple images
        gallery_images = self.extract_gallery_images(soup, url)

        for t in self.targets:
            n_line = normalize(t.target_line)
            from_map = n_line in matched_url_lines
            if from_map:
                matched, status = True, "found"
            else:
                matched, status = matches_line(t.target_line, title, text, url)
            if not matched:
                continue

            # strict_match: line must appear in URL or title (or transliterated)
            if strict_match and not from_map:
                nline = n_line
                nurl = normalize(unquote_plus(url))
                ntitle2 = normalize(title)
                if nline in nurl or nline in ntitle2:
                    pass
                else:
                    tline = transliterate(t.target_line)
                    if tline in url.lower() or tline in (title or "").lower():
                        pass
                    else:
                        continue

            dt = t.door_type if t.door_type in ("входная", "межкомнатная") else self.door_type_default

            if gallery_images:
                # Gallery mode: one DoorItem per gallery image
                for idx, (img_url, alt_text) in enumerate(gallery_images, 1):
                    model_name = alt_text or f"{t.target_line} — вариант {idx}"
                    final_status = "needs_review"
                    comments = ["model extracted from collection gallery"]
                    if alt_text:
                        comments.append(f"alt: {alt_text[:60]}")
                    if is_mapped:
                        pass
                    elif status == "needs_review":
                        comments.append("fuzzy match")
                    if is_generic_page(title, url):
                        comments.append("generic page")

                    found_items.append(DoorItem(
                        door_type=dt,
                        manufacturer=t.manufacturer,
                        line=t.target_line,
                        model_name=model_name,
                        description=desc,
                        material="",
                        coating="",
                        image_url=img_url,
                        source_url=url,
                        status=final_status,
                        comment="; ".join(comments),
                    ))
            else:
                # Single model mode (original logic)
                img_url, img_status = image_of(soup, url)
                final_status = status
                comments = []

                if from_map:
                    pass
                elif status == "needs_review":
                    comments.append("fuzzy match")
                if is_generic_page(title, url):
                    final_status = "needs_review"
                    comments.append("generic page")
                if img_status == "missing":
                    if final_status == "found":
                        final_status = "needs_review"
                    comments.append("no image")

                found_items.append(DoorItem(
                    door_type=dt,
                    manufacturer=t.manufacturer,
                    line=t.target_line,
                    model_name=title,
                    description=desc,
                    material="",
                    coating="",
                    image_url=img_url,
                    source_url=url,
                    status=final_status,
                    comment="; ".join(comments),
                ))

        return found_items

    def run(self) -> tuple[list[DoorItem], dict]:
        print(f"  [{now_str()}] collecting candidate URLs (max {self.max_pages})...")
        urls = self.collect_urls()

        log = {
            "base_url": self.base_url,
            "candidate_pages": len(urls),
            "pages_processed": 0,
            "models_found": 0,
            "targets": {},
            "errors": [],
        }

        detail_pattern = self.cfg.get("detail_url_pattern", "")
        if detail_pattern:
            matched = [u for u in urls if detail_pattern in u]
            if matched:
                urls = matched[:self.max_pages]

        items: list[DoorItem] = []
        processed = 0

        for url in urls:
            if processed >= self.max_pages:
                break
            html = fetch(self.session, url)
            if not html:
                continue
            processed += 1
            page_items = self.extract_from_page(html, url)
            items.extend(page_items)
            if page_items:
                self.log(f"    [{now_str()}] {url} -> {len(page_items)} items")

        log["pages_processed"] = processed
        log["models_found"] = len(items)
        unique = dedup_items(items)

        for t in self.targets:
            log["targets"][t.target_line] = sum(
                1 for i in unique if normalize(i.line) == normalize(t.target_line)
            )

        found_count = sum(1 for v in log["targets"].values() if v > 0)
        print(f"  [{now_str()}] done: {len(unique)} models, {found_count}/{len(self.targets)} lines found, {processed} pages")
        return unique, log


def dedup_items(items: list[DoorItem]) -> list[DoorItem]:
    seen: set[tuple[str, str, str, str, str]] = set()
    result = []
    for it in items:
        key = (normalize(it.manufacturer), normalize(it.line), normalize(it.model_name),
               normalize(it.image_url), normalize(it.source_url))
        if key in seen:
            continue
        seen.add(key)
        result.append(it)
    return result
