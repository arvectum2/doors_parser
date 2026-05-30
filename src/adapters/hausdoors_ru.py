"""Adapter for hausdoors.ru (Tilda-based site). Each collection page
contains multiple product blocks (t776, t760). We parse each product
block as a separate DoorItem."""

from src.adapters.base import (
    BaseAdapter, DoorItem, fetch, normalize, now_str, _is_bad_image,
    urljoin, text_of, title_of, description_of, looks_productish,
    is_generic_page, matches_line,
)
from bs4 import BeautifulSoup


class HausdoorsAdapter(BaseAdapter):
    """Extracts multiple products from Tilda collection pages."""

    IRRELEVANT_WORDS = ["филиал", "перегородк", "раздвижн"]

    def _is_irrelevant(self, title: str, url: str) -> bool:
        hay = (title + " " + url).lower()
        if any(w in hay for w in self.IRRELEVANT_WORDS):
            return True
        # Skip pages whose sole purpose is listing branches
        if "branches" in url.lower() or "company" in url.lower():
            title_low = title.lower()
            if "филиал" in title_low or "компани" in title_low:
                return True
        return False

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

        line_url_map = self.cfg.get("line_url_map", None)
        matched_url_lines: set[str] = set()
        is_mapped = line_url_map and url in line_url_map
        if is_mapped:
            matched_url_lines.add(normalize(line_url_map[url]))

        # Only check irrelevance and generic for non-mapped pages
        if not is_mapped:
            if self._is_irrelevant(title, url):
                if self.debug:
                    self.log(f"    skip irrelevant: {title}")
                return []
            if is_generic_page(title, url):
                return []

        found_items = []

        # Find Tilda product blocks (t776 and t760 are product listings)
        product_blocks = soup.select("[class*=t776]:not([class*=t776__col]):not([class*=t776__content]):not([class*=t776__scroll]), [class*=t760]:not([class*=t760__col]):not([class*=t760__content])")

        for t in self.targets:
            n_line = normalize(t.target_line)
            from_map = n_line in matched_url_lines
            if from_map:
                matched, status = True, "found"
            else:
                matched, status = matches_line(t.target_line, title, text, url)
            if not matched:
                continue

            dt = t.door_type if t.door_type in ("входная", "межкомнатная") else self.door_type_default

            if product_blocks:
                # Extract each product block as a separate DoorItem
                for idx, block in enumerate(product_blocks, 1):
                    name_el = block.select_one("[class*=title], [class*=name], h3, h4")
                    name = name_el.get_text(" ", strip=True) if name_el else ""
                    if not name or name == title:
                        name = f"{t.target_line} — вариант {idx}"

                    img_el = block.select_one("img")
                    img_url = ""
                    img_status = "missing"
                    if img_el:
                        src = img_el.get("src") or img_el.get("data-src") or ""
                        if src and not src.startswith("data:"):
                            img_url = urljoin(url, src)
                            if not _is_bad_image(img_url, img_el.get("alt",""), " ".join(img_el.get("class",[])), ""):
                                img_status = "found"

                    block_desc_el = block.select_one("[class*=descr]")
                    block_desc = block_desc_el.get_text(" ", strip=True) if block_desc_el else desc

                    final_status = "needs_review"
                    comments = ["model extracted from collection gallery"]
                    if img_status == "missing":
                        comments.append("no image")

                    found_items.append(DoorItem(
                        door_type=dt,
                        manufacturer=t.manufacturer,
                        line=t.target_line,
                        model_name=name,
                        description=block_desc[:700],
                        material="",
                        coating="",
                        image_url=img_url,
                        source_url=url,
                        status=final_status,
                        comment="; ".join(comments),
                    ))
            else:
                # Fallback: try gallery extraction from parent class
                gallery = self.extract_gallery_images(soup, url)
                if gallery:
                    for idx, (img_url, alt_text) in enumerate(gallery, 1):
                        model_name = alt_text or f"{t.target_line} - variant {idx}"
                        comments = ["model extracted from collection gallery"]
                        if not from_map and status == "needs_review":
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
                            status="needs_review",
                            comment="; ".join(comments),
                        ))
                else:
                    # Single fallback
                    img_url, img_status = self._single_image(soup, url)
                    final_status = status
                    comments = []
                    if from_map:
                        pass
                    elif status == "needs_review":
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

    @staticmethod
    def _single_image(soup: BeautifulSoup, page_url: str) -> tuple[str, str]:
        for img in soup.find_all("img"):
            src = img.get("src") or img.get("data-src") or ""
            if not src:
                continue
            full = urljoin(page_url, src)
            if _is_bad_image(full, img.get("alt",""), " ".join(img.get("class",[])), ""):
                continue
            return full, "found"
        return "", "missing"
