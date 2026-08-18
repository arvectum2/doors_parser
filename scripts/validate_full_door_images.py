from __future__ import annotations

import math
import re
from dataclasses import dataclass
from io import BytesIO
from typing import Iterable
from urllib.parse import urljoin, urlparse

import requests
from PIL import Image, UnidentifiedImageError


NEG_INTERIOR = (
    "interior",
    "room",
    "living",
    "kitchen",
    "bedroom",
    "hall",
    "герой",
    "hero",
)
NEG_TEXTURE = (
    "texture",
    "color",
    "sample",
    "pattern",
    "swatch",
    "покрыт",
    "цвет",
    "texture",
)
NEG_FRAGMENT = (
    "fragment",
    "closeup",
    "close-up",
    "crop",
    "detail",
    "macro",
    "фрагмент",
    "крупный",
)
NEG_COMPONENT = (
    "handle",
    "lock",
    "petlya",
    "петля",
    "ручка",
    "замок",
    "korob",
    "короб",
    "коробка",
    "наличник",
    "добор",
    "погонаж",
    "profile",
    "profil",
    "razrez",
    "section",
    "торец",
    "сечение",
)
NEG_BANNER = (
    "banner",
    "promo",
    "logo",
    "icon",
    "badge",
    "sticker",
    "hit",
    "sale",
    "акция",
    "хит",
)
NEG_PLACEHOLDER = (
    "open.png",
    "catalog.svg",
    "placeholder",
    "no_image",
    "no-image",
    "empty",
    "sprite",
    "pixel",
    "loader",
    "favicon",
    "/resize/20x/",
    "resize/20x",
)
POS_HINTS = ("door", "dver", "dveri", "model", "product", "polotno", "полотно", "фасад", "модель")


@dataclass
class ImageDecision:
    url: str
    decision: str
    reason: str
    http_status: int
    content_type: str
    width: int
    height: int
    aspect_ratio: float
    score: float


def _clean_url(url: str) -> str:
    u = (url or "").strip()
    if not u:
        return ""
    if u.startswith("//"):
        u = "https:" + u
    u = u.replace("\\/", "/")
    u = re.sub(r"/-/resizeb?/\d+x/", "/", u, flags=re.I)
    u = re.sub(r"/-/cover/\d+x\d+/", "/", u, flags=re.I)
    u = re.sub(r"/-/format/webp/", "/", u, flags=re.I)
    u = u.replace("optim.tildacdn.com", "static.tildacdn.com").replace("thb.tildacdn.com", "static.tildacdn.com")
    return u


def _entropy(img: Image.Image) -> float:
    gray = img.convert("L")
    hist = gray.histogram()
    total = sum(hist) or 1
    ent = 0.0
    for h in hist:
        if h:
            p = h / total
            ent -= p * math.log2(p)
    return ent


def _find_reason_by_url(url_l: str) -> str | None:
    if any(k in url_l for k in NEG_PLACEHOLDER):
        return "placeholder_rejected"
    if any(k in url_l for k in NEG_BANNER):
        return "banner_rejected"
    if any(k in url_l for k in NEG_INTERIOR):
        return "interior_rejected"
    if any(k in url_l for k in NEG_FRAGMENT):
        return "fragment_rejected"
    if any(k in url_l for k in NEG_COMPONENT):
        return "component_rejected"
    if any(k in url_l for k in NEG_TEXTURE):
        return "texture_rejected"
    return None


def validate_image_url(
    session: requests.Session,
    url: str,
    cache: dict[str, ImageDecision],
    min_w: int = 260,
    min_h: int = 420,
) -> ImageDecision:
    clean = _clean_url(url)
    if not clean:
        return ImageDecision("", "reject", "no_url", 0, "", 0, 0, 0.0, -999.0)
    if clean in cache:
        return cache[clean]

    low = clean.lower()
    early_reason = _find_reason_by_url(low)

    status = 0
    ctype = ""
    w = h = 0
    aspect = 0.0
    score = -10.0
    decision = "reject"
    reason = early_reason or "no_full_door_candidate_found"
    try:
        r = session.get(clean, timeout=(8, 30), allow_redirects=True)
        status = r.status_code
        ctype = (r.headers.get("Content-Type") or "").lower()
        if status != 200:
            reason = "http_not_200"
        elif not ctype.startswith("image/"):
            reason = "not_image_content"
        else:
            try:
                img = Image.open(BytesIO(r.content))
                w, h = img.size
                aspect = (h / w) if w else 0.0
                ent = _entropy(img.resize((min(256, w), min(256, h))))
                wide = (w / h) if h else 99.0

                if w < min_w or h < min_h:
                    reason = "too_small"
                elif wide > 1.6:
                    reason = "not_vertical_product_image"
                elif aspect < 1.0:
                    reason = "not_vertical_product_image"
                elif early_reason:
                    reason = early_reason
                else:
                    # Low-entropy images can still be valid product renders on white background.
                    if ent < 2.2 and aspect < 1.12:
                        reason = "texture_rejected"
                    else:
                        reason = "full_door_image_ok"
                        decision = "accept"

                score = 0.0
                if decision == "accept":
                    score += min(5.0, max(0.0, aspect - 1.1) * 2.0)
                    if 1.45 <= aspect <= 3.2:
                        score += 3.0
                    if any(k in low for k in POS_HINTS):
                        score += 1.5
                    if "product" in low or "model" in low:
                        score += 1.0
                    score += min(2.0, (w * h) / 1_200_000)
                else:
                    score -= 3.0
            except UnidentifiedImageError:
                reason = "not_image_content"
            except Exception:
                reason = "image_decode_error"
    except Exception:
        reason = "fetch_error"

    out = ImageDecision(clean, decision, reason, status, ctype, w, h, round(aspect, 4), round(score, 3))
    cache[clean] = out
    return out


def extract_image_candidates(source_url: str, html: str) -> list[str]:
    if not html:
        return []
    urls: list[str] = []

    # Common URL capture from attributes and JSON-like payloads.
    raw = re.findall(
        r"""(?:src|data-src|data-original|data-lazy|data-image|data-img|href|content)\s*=\s*["']([^"']+)["']""",
        html,
        flags=re.I,
    )
    raw += re.findall(r"""url\((['"]?)([^'")]+)\1\)""", html, flags=re.I)
    raw += re.findall(r"""https?://[^"'\s>]+""", html, flags=re.I)
    raw += [u.replace("\\/", "/") for u in re.findall(r"""https?:\\\\/\\\\/[^"'\s>]+""", html, flags=re.I)]

    for item in raw:
        u = item[1] if isinstance(item, tuple) else item
        if not isinstance(u, str):
            continue
        if "data:image/" in u.lower():
            continue
        if not re.search(r"\.(jpg|jpeg|png|webp|gif)(\?|$)", u, flags=re.I):
            continue
        urls.append(urljoin(source_url, u))

    # srcset candidates
    for srcset in re.findall(r"""srcset\s*=\s*["']([^"']+)["']""", html, flags=re.I):
        for part in srcset.split(","):
            cand = part.strip().split(" ")[0].strip()
            if re.search(r"\.(jpg|jpeg|png|webp|gif)(\?|$)", cand, flags=re.I):
                urls.append(urljoin(source_url, cand))

    # De-dup and normalize.
    seen = set()
    out = []
    for u in urls:
        c = _clean_url(u)
        if not c:
            continue
        if urlparse(c).scheme not in {"http", "https"}:
            continue
        if c in seen:
            continue
        seen.add(c)
        out.append(c)
    return out


def pick_best_full_door(
    session: requests.Session,
    candidates: Iterable[str],
    cache: dict[str, ImageDecision],
) -> ImageDecision | None:
    best: ImageDecision | None = None
    for cand in candidates:
        dec = validate_image_url(session, cand, cache)
        if dec.decision != "accept":
            continue
        if best is None or dec.score > best.score:
            best = dec
    return best
