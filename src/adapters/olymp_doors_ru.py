import re
from urllib.parse import urljoin

from bs4 import BeautifulSoup

from src.adapters.base import (
    BaseAdapter,
    _is_bad_image,
    normalize,
)


class OlympAdapter(BaseAdapter):
    """Adapter for olymp-doors.ru with broader gallery extraction."""

    def extract_gallery_images(self, soup: BeautifulSoup, url: str) -> list[tuple[str, str]]:
        base_result = super().extract_gallery_images(soup, url)
        seen_urls = set()
        for _, u in base_result:
            seen_urls.add(u)
        gallery = list(base_result)

        for img in soup.find_all("img"):
            src = (
                img.get("src")
                or img.get("data-src")
                or img.get("data-lazy")
                or img.get("data-original")
                or ""
            )
            alt = img.get("alt", "").strip()
            if not src:
                continue
            klass = " ".join(img.get("class", []))
            parent = img.parent
            parent_klass = " ".join(parent.get("class", [])) if parent else ""
            grandparent = parent.parent if parent else None
            grandparent_klass = " ".join(grandparent.get("class", [])) if grandparent else ""

            if _is_bad_image(src, alt, klass, normalize(url)):
                continue
            full_url = urljoin(url, src)
            if full_url in seen_urls:
                continue

            is_door = False
            if any(x in klass for x in ["product", "img-full", "gallery", "preview", "thumb", "item", "card"]):
                is_door = True
            if any(x in parent_klass for x in ["product", "catalog", "gallery", "item", "card", "grid", "door", "dver"]):
                is_door = True
            if any(x in grandparent_klass for x in ["product", "catalog", "gallery", "item", "card", "grid", "door", "dver"]):
                is_door = True
            if any(x in src.lower() for x in ["dver", "door", "product", "catalog", "img", "gallery", "photo"]):
                is_door = True
            if re.search(r"\.(jpg|jpeg|png|webp)(\?|$)", full_url.lower()):
                is_door = True

            if is_door:
                seen_urls.add(full_url)
                gallery.append((full_url, alt))

        return gallery
