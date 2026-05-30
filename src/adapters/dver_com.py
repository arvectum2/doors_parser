from bs4 import BeautifulSoup
from urllib.parse import urlparse

from src.adapters.base import (
    BaseAdapter,
    DoorItem,
    SKIP_URL_PARTS,
    description_of,
    fetch,
    is_generic_page,
    is_same_domain,
    matches_line,
    normalize,
    normalize_dver_model_name,
    now_str,
    text_of,
    title_of,
    clean_url,
)


class DverComAdapter(BaseAdapter):
    """Adapter for dver.com with strict model name cleanup."""

    def collect_urls(self) -> list[str]:
        """Override: allow ?-based collection URLs (dver.com uses ?collection= for series)."""
        urls = []
        seeds = self.cfg.get("seeds", [self.base_url])
        use_sitemap = self.cfg.get("sitemap", True)
        max_seed_pages = self.cfg.get("max_seed_pages", 10)

        direct = self.cfg.get("direct_urls", [])
        if direct:
            urls.extend(direct)

        if use_sitemap:
            print(f"    [{now_str()}] collecting sitemap URLs...")
            from src.adapters.base import sitemap_urls
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
                # Skip only non-? URL parts; keep ? for collection filter URLs
                skip_parts = [p for p in SKIP_URL_PARTS if p != "?"]
                if any(x in low for x in skip_parts):
                    continue
                if any(x in low for x in self.cfg.get("catalog_url_markers", ["catalog", "katalog", "category", "product", "tovar", "dver", "door", "mezh", "vhod", "collection", "model", "seria", "series"])):
                    links.add(href)
            return links

        for seed in seeds:
            seed_links.update(scrape_links(seed))
            base_seed = seed.split("?")[0]
            for page_n in range(2, max_seed_pages + 1):
                page_url = f"{base_seed}?page={page_n}" if "?" not in seed else f"{seed}&page={page_n}"
                page_links = scrape_links(page_url)
                if not page_links:
                    break
                seed_links.update(page_links)

        if self.debug:
            print(f"    [{now_str()}] seed links: {len(seed_links)}")
        # Prioritise: seeds first (filter URLs with ? are NOT in sitemap), then seed links, then sitemap
        all_urls = list(dict.fromkeys(list(seeds) + list(seed_links) + urls))
        all_urls = [u for u in all_urls if is_same_domain(u, self.base_url)]
        skip_parts = [p for p in SKIP_URL_PARTS if p != "?"]
        all_urls = [u for u in all_urls if not any(x in u.lower() for x in skip_parts)]
        if not all_urls:
            all_urls = list(seeds)
        urls = all_urls[:self.max_pages]
        print(f"    [{now_str()}] total candidate URLs: {len(urls)}")
        return urls

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

        gallery = self.extract_gallery_images(soup, url)
        items: list[DoorItem] = []

        for t in self.targets:
            matched, status = matches_line(t.target_line, title, text, url)
            if not matched:
                continue
            dt = t.door_type if t.door_type in ("входная", "межкомнатная") else self.door_type_default

            if gallery:
                for idx, (img_url, alt_text) in enumerate(gallery, 1):
                    raw_name = alt_text or title or f"{t.target_line} - variant {idx}"
                    cleaned, orig = normalize_dver_model_name(raw_name)
                    comments = ["model extracted from collection gallery"]
                    if cleaned != orig:
                        comments.insert(0, f"raw_name: {orig}")
                    if status == "needs_review":
                        comments.append("fuzzy match")
                    items.append(DoorItem(
                        door_type=dt,
                        manufacturer=t.manufacturer,
                        line=t.target_line,
                        model_name=cleaned or raw_name,
                        description=desc,
                        material="",
                        coating="",
                        image_url=img_url,
                        source_url=url,
                        status="needs_review",
                        comment="; ".join(comments),
                    ))
            else:
                cleaned, orig = normalize_dver_model_name(title)
                comment = ""
                if cleaned != orig:
                    comment = f"raw_name: {orig}"
                items.append(DoorItem(
                    door_type=dt,
                    manufacturer=t.manufacturer,
                    line=t.target_line,
                    model_name=cleaned or title,
                    description=desc,
                    material="",
                    coating="",
                    image_url="",
                    source_url=url,
                    status=status,
                    comment=comment,
                ))

        return items
