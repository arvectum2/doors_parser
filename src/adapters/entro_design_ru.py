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


class EntroAdapter(BaseAdapter):
    """Adapter for entro-design.ru collection pages with gallery variants."""

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

        if is_generic_page(title, url):
            return []

        desc = description_of(soup)
        if not desc:
            desc = " ".join(text.split())[:700]

        line_url_map = self.cfg.get("line_url_map", None)
        matched_url_lines: set[str] = set()
        if line_url_map and url in line_url_map:
            matched_url_lines.add(normalize(line_url_map[url]))

        gallery = self.extract_gallery_images(soup, url)

        # Entro specific fallback: data-fancybox / lightbox thumbs
        seen = {u for u, _ in gallery}
        for node in soup.select("a[data-fancybox], a[href*='upload'], a[href*='media']"):
            href = node.get("href", "").strip()
            if not href:
                continue
            full = urljoin(url, href)
            if full in seen:
                continue
            if _is_bad_image(full, "", "", ""):
                continue
            seen.add(full)
            caption = (node.get("title") or "").strip()
            gallery.append((full, caption))

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
