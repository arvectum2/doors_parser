"""KomfortAdapter: extract product cards and assign to correct lines by matching model_name."""
from bs4 import BeautifulSoup

from src.adapters.base import (
    BaseAdapter,
    DoorItem,
    description_of,
    is_generic_page,
    matches_line,
    normalize,
    text_of,
    title_of,
)


class KomfortAdapter(BaseAdapter):
    """Adapter for komfortdoors.ru with per-model line matching."""

    def extract_from_page(self, html: str, url: str) -> list[DoorItem]:
        soup = BeautifulSoup(html, "lxml")
        title = title_of(soup)
        if not title:
            return []
        text = text_of(soup)
        ntitle = normalize(title)

        if is_generic_page(title, url):
            return []

        desc = description_of(soup)
        if not desc:
            desc = " ".join(text.split())[:700]

        # Check line_url_map for exact URL->line mapping
        line_url_map = self.cfg.get("line_url_map", None)
        page_line = normalize(line_url_map[url]) if (line_url_map and url in line_url_map) else None

        gallery = self.extract_gallery_images(soup, url)
        items: list[DoorItem] = []

        for t in self.targets:
            n_line = normalize(t.target_line)
            if page_line and n_line == page_line:
                matched, status = True, "found"
            else:
                matched, status = matches_line(t.target_line, title, text, url)
            if not matched:
                continue

            dt = t.door_type if t.door_type in ("входная", "межкомнатная") else self.door_type_default

            if gallery:
                for idx, (img_url, alt_text) in enumerate(gallery, 1):
                    model_name = alt_text or f"{t.target_line} — вариант {idx}"
                    # Determine which target_line this model truly belongs to
                    actual_line = self._match_model_to_line(model_name, t.target_line)
                    final_status = "needs_review"
                    comments = ["model extracted from collection gallery"]
                    if alt_text:
                        comments.append(f"alt: {alt_text[:60]}")
                    if actual_line != t.target_line:
                        comments.append(f"reassigned from {t.target_line} to {actual_line}")

                    items.append(DoorItem(
                        door_type=dt,
                        manufacturer=t.manufacturer,
                        line=actual_line,
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
                img_url, img_status = self.image_of_komfort(soup, url)
                items.append(DoorItem(
                    door_type=dt,
                    manufacturer=t.manufacturer,
                    line=t.target_line,
                    model_name=title,
                    description=desc,
                    material="",
                    coating="",
                    image_url=img_url,
                    source_url=url,
                    status=status,
                    comment="",
                ))

        return items

    def _match_model_to_line(self, model_name: str, default_line: str) -> str:
        """Determine which target line a model belongs to by its name."""
        mn = model_name.strip().upper()
        # Skip color swatches (single-word color names)
        color_words = {"БЕЛЫЙ", "КРЕМОВО-СЕРЫЙ", "КАМЕННЫЙ КРЕМ", "МОККО",
                       "СЕРЫЙ КАМЕНЬ", "ТЁМНЫЙ ШОКОЛАД", "ТЕМНЫЙ ШОКОЛАД"}
        if mn in color_words:
            return default_line

        # Match by model name prefix
        if mn.startswith("NL-"):
            return "New line"
        if mn.startswith("SL "):
            return "Slope"
        if mn.startswith("TRUE"):
            return "True"
        if mn.startswith("GIFT"):
            return "Gift"
        if mn.startswith("ART") and not mn.startswith("ART.") and not mn.startswith("ART "):
            return "Art"
        if mn.startswith("PREMIER"):
            return "Premier"
        if mn.startswith("L "):  # Malta / L series
            return "Malta"
        if mn.startswith("NEO"):
            return "Neo"
        if mn.startswith("M "):  # Molding / M series
            return "Molding"
        if mn.startswith("TURIN"):
            return "Turin"
        if mn.startswith("ЭКОШПОН") or mn.startswith("EKOSHPON"):
            return "Экошпон"

        return default_line

    def image_of_komfort(self, soup: BeautifulSoup, page_url: str) -> tuple[str, str]:
        """Extract first product image from page."""
        for img in soup.find_all("img"):
            src = img.get("src") or img.get("data-src") or ""
            if not src:
                continue
            low = src.lower()
            if "product" in low or "upload" in low or "cache" in low:
                from urllib.parse import urljoin
                return urljoin(page_url, src), "found"
        return "", "missing"
