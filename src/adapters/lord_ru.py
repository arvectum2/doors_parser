from bs4 import BeautifulSoup

from src.adapters.base import (
    BaseAdapter,
    DoorItem,
    _is_bad_image,
    description_of,
    is_generic_page,
    matches_line,
    normalize,
    text_of,
    title_of,
    urljoin,
)


class LordAdapter(BaseAdapter):
    """Adapter for dveri-lord.ru to expand collection radio variants."""

    def extract_from_page(self, html: str, url: str) -> list[DoorItem]:
        soup = BeautifulSoup(html, "lxml")
        title = title_of(soup)
        if not title:
            return []
        text = text_of(soup)
        ntitle = normalize(title)

        for skip in self.cfg.get("skip_title_patterns", []):
            if skip in ntitle:
                return []

        desc = description_of(soup)
        if not desc:
            desc = " ".join(text.split())[:700]

        line_url_map = self.cfg.get("line_url_map", None)
        matched_url_lines: set[str] = set()
        if line_url_map and url in line_url_map:
            matched_url_lines.add(normalize(line_url_map[url]))

        if is_generic_page(title, url) and not matched_url_lines:
            return []

        gallery = self.extract_gallery_images(soup, url)

        # Lord-specific fallback: radio-box and product choice nodes
        seen = {u for u, _ in gallery}
        for img in soup.select("img.radio-box__img, .radio-box img, .product-variations img"):
            src = img.get("src") or img.get("data-src") or ""
            if not src:
                continue
            full = urljoin(url, src)
            if full in seen:
                continue
            alt = (img.get("alt") or "").strip()
            klass = " ".join(img.get("class", []))
            if _is_bad_image(full, alt, klass, ""):
                continue
            seen.add(full)
            gallery.append((full, alt))

        items: list[DoorItem] = []
        for t in self.targets:
            nline = normalize(t.target_line)
            from_map = nline in matched_url_lines
            if from_map:
                matched, status = True, "found"
            else:
                matched, status = matches_line(t.target_line, title, text, url)
            if not matched:
                continue

            dt = t.door_type if t.door_type in ("входная", "межкомнатная") else self.door_type_default
            if gallery:
                for idx, (img_url, alt_text) in enumerate(gallery, 1):
                    model_name = alt_text or f"{t.target_line} - variant {idx}"
                    comments = ["model extracted from collection gallery"]
                    if not from_map and status == "needs_review":
                        comments.append("fuzzy match")
                    items.append(DoorItem(
                        door_type=dt,
                        manufacturer=t.manufacturer,
                        line=t.target_line,
                        model_name=model_name,
                        description=desc,
                        material="",
                        coating="",
                        image_url=img_url,
                        source_url=url,
                        status="needs_review",
                        comment="; ".join(comments),
                    ))
            else:
                items.append(DoorItem(
                    door_type=dt,
                    manufacturer=t.manufacturer,
                    line=t.target_line,
                    model_name=title,
                    description=desc,
                    material="",
                    coating="",
                    image_url="",
                    source_url=url,
                    status="needs_review",
                    comment="no gallery images",
                ))

        return items
