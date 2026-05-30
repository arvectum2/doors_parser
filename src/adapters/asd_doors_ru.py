from src.adapters.base import (
    BaseAdapter, DoorItem, fetch, normalize, now_str, _is_bad_image,
    clean_url, urljoin, urlparse, title_of, description_of, text_of, is_same_domain,
    sitemap_urls,
)
from bs4 import BeautifulSoup


class ASDoorsAdapter(BaseAdapter):
    """Adapter for as-doors.ru.
    
    ASD target lines are model landmarks, not collections.
    Collects ALL door model pages from the sitemap (onstock/ and doors/zakaznyie/)
    and extracts each individual door model.
    """

    def collect_urls(self) -> list[str]:
        """Collect all door model pages from onstock/ and zakaznyie/ sections."""
        print(f"    [{now_str()}] collecting ASD model URLs from sitemap...")
        all_sitemap = sitemap_urls(self.session, self.base_url, 500)
        model_urls = []
        for u in all_sitemap:
            if "onstock/" in u or "zakaznyie/" in u:
                # Exclude listing pages themselves
                if not u.rstrip("/").endswith(("/onstock", "/onstock/", "/zakaznyie.html")):
                    if not any(x in u for x in ["dealers/", "/blog/", "/sale/"]):
                        model_urls.append(u)
        # Deduplicate
        model_urls = list(dict.fromkeys(model_urls))

        # Fallback: crawl seed pages if sitemap is empty or incomplete
        if not model_urls:
            print(f"    [{now_str()}] sitemap empty, scanning ASD seeds...")
            seeds = self.cfg.get("seeds", [self.base_url])
            queue = list(seeds)
            seen = set(queue)
            while queue and len(model_urls) < self.max_pages:
                page = queue.pop(0)
                html = fetch(self.session, page)
                if not html:
                    continue
                soup = BeautifulSoup(html, "lxml")
                for a in soup.find_all("a", href=True):
                    href = clean_url(page, a.get("href", ""))
                    if not is_same_domain(href, self.base_url):
                        continue
                    low = href.lower()
                    if any(x in low for x in ["dealers/", "/blog/", "/sale/"]):
                        continue
                    if "onstock/" in low or "zakaznyie" in low:
                        if low.rstrip("/").endswith(("/onstock", "/zakaznyie.html")):
                            if href not in seen and len(seen) < 2000:
                                seen.add(href)
                                queue.append(href)
                        else:
                            model_urls.append(href)

        print(f"    [{now_str()}] ASD model pages: {len(model_urls)}")
        return model_urls[:self.max_pages]

    def extract_from_page(self, html: str, url: str) -> list[DoorItem]:
        """Extract a single door model from its individual page."""
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

        # Find the best product image
        img_url = ""
        img_status = "missing"
        all_images_bad = False
        has_any_image = False
        for img in soup.find_all("img"):
            src = img.get("src") or img.get("data-src") or ""
            if not src:
                continue
            has_any_image = True
            full = urljoin(url, src)
            alt = img.get("alt", "")
            klass = " ".join(img.get("class", []))
            if _is_bad_image(full, alt, klass, ""):
                all_images_bad = True
                continue
            low = (full + alt + klass).lower()
            if any(x in low for x in ["upload", "product", "dver", "door", "photo"]):
                if not img_url:
                    img_url = full
                    img_status = "found"
                # Prefer larger images
                if "800" in low or "1024" in low or "full" in low:
                    img_url = full
                    break

        found_items = []
        for t in self.targets:
            dt = t.door_type if t.door_type in ("входная", "межкомнатная") else self.door_type_default
            final_status = "found" if img_status == "found" else "needs_review"
            comments = []
            if img_status == "missing":
                if all_images_bad and has_any_image:
                    comments.append("bad_image_url_removed")
                else:
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
