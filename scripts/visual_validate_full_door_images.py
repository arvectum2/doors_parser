from __future__ import annotations

import re
from dataclasses import dataclass
from io import BytesIO
from typing import Iterable
from urllib.parse import urljoin, urlparse

import numpy as np
import requests
from PIL import Image, ImageFilter, UnidentifiedImageError


NEG_URL_TOKENS = (
    "ral",
    "ncs",
    "color",
    "colour",
    "texture",
    "sample",
    "pattern",
    "decor",
    "finish",
    "material",
    "room",
    "interior",
    "hero",
    "banner",
    "promo",
    "badge",
    "logo",
    "icon",
    "handle",
    "lock",
    "profile",
    "section",
    "trim",
    "frame",
    "nalichnik",
    "doborn",
    "korob",
    "empty",
    "placeholder",
    "sprite",
    "pixel",
    "open.png",
    "catalog.svg",
)

POS_URL_TOKENS = ("door", "dver", "dveri", "model", "product", "polotno", "полотно", "модель")

COLORISH_TOKENS = (
    "матовый",
    "глянец",
    "цвет",
    "покрытие",
    "эмаль",
    "лак",
    "шпон",
    "дуб",
    "орех",
    "венге",
    "ясень",
    "ral",
    "ncs",
)


@dataclass
class VisualCheck:
    url: str
    status_code: int
    content_type: str
    width: int
    height: int
    bbox_width: int
    bbox_height: int
    bbox_aspect: float
    decision: str
    reason: str
    score: float


def normalize_url(url: str) -> str:
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


def clean_model_name(name: str) -> str:
    return re.sub(r"\s+", " ", (name or "").strip().lower().replace("ё", "е"))


def is_probable_color_row(model_name: str) -> bool:
    n = clean_model_name(model_name)
    if not n:
        return False
    if re.search(r"\bral[-_\s]?\d", n, flags=re.I) or re.search(r"\bncs[-_\s]?[a-z0-9]", n, flags=re.I):
        return True
    if n.startswith("матовый "):
        return True
    token_hits = sum(1 for t in COLORISH_TOKENS if t in n)
    # likely finish/color entry when short and colorish
    return token_hits >= 2 and len(n.split()) <= 6


def _find_bbox_alpha(img: Image.Image) -> tuple[int, int]:
    arr = np.array(img.convert("RGBA"))
    alpha = arr[:, :, 3]
    ys, xs = np.where(alpha > 18)
    if len(xs) == 0:
        return 0, 0
    return int(xs.max() - xs.min() + 1), int(ys.max() - ys.min() + 1)


def _find_bbox_edges(img: Image.Image) -> tuple[int, int]:
    gray = img.convert("L")
    # suppress small texture noise
    small = gray.resize((max(64, gray.width // 3), max(96, gray.height // 3)))
    edge = small.filter(ImageFilter.FIND_EDGES)
    arr = np.array(edge, dtype=np.uint8)
    thr = max(18, int(np.percentile(arr, 82)))
    mask = arr >= thr
    ys, xs = np.where(mask)
    if len(xs) == 0:
        return 0, 0
    bw = int(xs.max() - xs.min() + 1)
    bh = int(ys.max() - ys.min() + 1)
    # restore to original scale
    sx = gray.width / small.width
    sy = gray.height / small.height
    return int(bw * sx), int(bh * sy)


def _entropy_like(img: Image.Image) -> float:
    gray = img.convert("L").resize((128, 128))
    arr = np.array(gray, dtype=np.uint8)
    hist = np.bincount(arr.flatten(), minlength=256).astype(np.float64)
    p = hist / max(1.0, hist.sum())
    p = p[p > 0]
    return float(-(p * np.log2(p)).sum())


def validate_full_door_image(
    session: requests.Session,
    url: str,
    cache: dict[str, VisualCheck],
) -> VisualCheck:
    u = normalize_url(url)
    if not u:
        return VisualCheck("", 0, "", 0, 0, 0, 0, 0.0, "no_image", "no_image", -999.0)
    if u in cache:
        return cache[u]

    low = u.lower()
    early_reason = next((t for t in NEG_URL_TOKENS if t in low), "")
    try:
        resp = session.get(u, timeout=(8, 35), allow_redirects=True)
        status = resp.status_code
        ctype = (resp.headers.get("Content-Type") or "").lower()
        if status != 200:
            out = VisualCheck(u, status, ctype, 0, 0, 0, 0, 0.0, "invalid_image", "http_not_200", -999.0)
            cache[u] = out
            return out
        if not ctype.startswith("image/"):
            out = VisualCheck(u, status, ctype, 0, 0, 0, 0, 0.0, "invalid_image", "not_image_content", -999.0)
            cache[u] = out
            return out
        try:
            img = Image.open(BytesIO(resp.content))
        except UnidentifiedImageError:
            out = VisualCheck(u, status, ctype, 0, 0, 0, 0, 0.0, "invalid_image", "decode_failed", -999.0)
            cache[u] = out
            return out
        w, h = img.size
        if w < 180 or h < 260:
            out = VisualCheck(u, status, ctype, w, h, 0, 0, 0.0, "invalid_image", "too_small", -50.0)
            cache[u] = out
            return out

        # bbox
        if "A" in img.getbands():
            bw, bh = _find_bbox_alpha(img)
        else:
            bw, bh = _find_bbox_edges(img)
        if bw <= 0 or bh <= 0:
            bw, bh = w, h
        aspect_img = h / max(1, w)
        aspect_obj = bh / max(1, bw)
        cover_ratio = (bw * bh) / max(1, w * h)
        ent = _entropy_like(img)

        decision = "full_door"
        reason = "full_door_image_ok"
        score = 0.0

        if early_reason:
            if any(k in early_reason for k in ("banner", "promo", "badge", "logo", "icon")):
                decision, reason = "banner", "banner_rejected"
            elif any(k in early_reason for k in ("interior", "room", "hero")):
                decision, reason = "interior", "interior_rejected"
            elif any(k in early_reason for k in ("texture", "sample", "pattern", "decor", "finish", "material", "color", "colour", "ral", "ncs")):
                decision, reason = "color_sample", "color_sample_rejected"
            elif any(k in early_reason for k in ("handle", "lock", "profile", "section", "trim", "frame", "nalichnik", "doborn", "korob")):
                decision, reason = "component", "component_rejected"
            else:
                decision, reason = "invalid_image", "url_negative_filter_rejected"

        # geometry and texture checks (applied for not already hard-rejected token cases too)
        if decision == "full_door":
            if w / max(1, h) > 1.45:
                decision, reason = "interior", "wide_layout_rejected"
            elif aspect_img < 1.0:
                decision, reason = "fragment", "not_vertical_image_rejected"
            elif aspect_obj < 1.45:
                decision, reason = "fragment", "object_not_vertical_enough"
            elif cover_ratio < 0.12:
                decision, reason = "fragment", "object_too_small_in_frame"
            elif cover_ratio > 0.94 and aspect_img < 1.3:
                decision, reason = "texture", "texture_like_full_frame"
            elif ent < 2.0:
                decision, reason = "texture", "low_entropy_texture"

        if decision == "full_door":
            score += min(4.0, max(0.0, aspect_obj - 1.2) * 2.0)
            score += min(2.0, max(0.0, aspect_img - 1.0) * 1.2)
            score += min(2.0, cover_ratio * 2.5)
            if any(t in low for t in POS_URL_TOKENS):
                score += 1.0
        else:
            score = -3.0

        out = VisualCheck(u, status, ctype, w, h, bw, bh, round(aspect_obj, 4), decision, reason, round(score, 3))
        cache[u] = out
        return out
    except Exception:
        out = VisualCheck(u, 0, "", 0, 0, 0, 0, 0.0, "invalid_image", "fetch_error", -999.0)
        cache[u] = out
        return out


def extract_candidates_from_html(source_url: str, html: str) -> list[str]:
    if not html:
        return []
    urls: list[str] = []
    raw = re.findall(
        r"""(?:src|data-src|data-original|data-lazy|data-image|data-img|href|content)\s*=\s*["']([^"']+)["']""",
        html,
        flags=re.I,
    )
    raw += [m[1] for m in re.findall(r"""url\((['"]?)([^'")]+)\1\)""", html, flags=re.I)]
    raw += re.findall(r"""https?://[^"'\s>]+""", html, flags=re.I)
    raw += [u.replace("\\/", "/") for u in re.findall(r"""https?:\\\\/\\\\/[^"'\s>]+""", html, flags=re.I)]

    for u in raw:
        if not isinstance(u, str):
            continue
        if "data:image/" in u.lower():
            continue
        if not re.search(r"\.(jpg|jpeg|png|webp|gif)(\?|$)", u, flags=re.I):
            continue
        urls.append(urljoin(source_url, u))

    for srcset in re.findall(r"""srcset\s*=\s*["']([^"']+)["']""", html, flags=re.I):
        for part in srcset.split(","):
            cand = part.strip().split(" ")[0].strip()
            if re.search(r"\.(jpg|jpeg|png|webp|gif)(\?|$)", cand, flags=re.I):
                urls.append(urljoin(source_url, cand))

    seen = set()
    out = []
    for u in urls:
        n = normalize_url(u)
        if not n:
            continue
        if urlparse(n).scheme not in {"http", "https"}:
            continue
        if n in seen:
            continue
        seen.add(n)
        out.append(n)
    return out


def pick_best_visual_full_door(
    session: requests.Session,
    candidates: Iterable[str],
    cache: dict[str, VisualCheck],
) -> VisualCheck | None:
    best = None
    for c in candidates:
        chk = validate_full_door_image(session, c, cache)
        if chk.decision != "full_door":
            continue
        if best is None or chk.score > best.score:
            best = chk
    return best

