from bs4 import BeautifulSoup

from src.adapters.base import (
    BaseAdapter, DoorItem,
    normalize, now_str, fetch, text_of, title_of, description_of,
    image_of, looks_productish, is_generic_page, matches_line,
    urljoin, _is_bad_image,
)


class AelitaAdapter(BaseAdapter):
    """JS-heavy site. Uses line_url_map and extracts all gallery images from collection pages."""

    def extract_from_page(self, html: str, url: str) -> list[DoorItem]:
        soup = BeautifulSoup(html, "lxml")
        title = title_of(soup)
        text = text_of(soup)
        if not title:
            return []

        ntitle = normalize(title)
        for skip in self.cfg.get("skip_title_patterns", []):
            if skip in ntitle:
                return []

        desc = description_of(soup)
        if not desc:
            desc = " ".join(text.split())[:700]

        found_items = []
        line_url_map = self.cfg.get("line_url_map", None)
        matched_url_lines: set[str] = set()
        is_mapped = line_url_map and url in line_url_map
        if is_mapped:
            matched_url_lines.add(normalize(line_url_map[url]))

        if is_generic_page(title, url) and not is_mapped:
            return []

        # Extract gallery images for this collection page
        gallery_images = self.extract_gallery_images(soup, url)

        for t in self.targets:
            n_line = normalize(t.target_line)
            from_map = n_line in matched_url_lines
            if from_map:
                matched, status_val = True, "found"
            else:
                matched, status_val = matches_line(t.target_line, title, text, url)
            if not matched:
                continue

            dt = t.door_type if t.door_type in ("входная", "межкомнатная") else self.door_type_default

            if gallery_images:
                # Gallery mode: one item per image
                for idx, (img_url, alt_text) in enumerate(gallery_images, 1):
                    model_name = alt_text or f"{t.target_line} — вариант {idx}"
                    final_status = "needs_review"
                    comments = ["model extracted from collection gallery"]
                    if alt_text:
                        comments.append(f"alt: {alt_text[:60]}")
                    if not from_map and status_val == "needs_review":
                        comments.append("fuzzy match")

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
                # Single mode (original logic)
                img_url, img_status = image_of(soup, url)
                final_status = status_val
                comments = []
                if not from_map and status_val == "needs_review":
                    comments.append("fuzzy match")
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
