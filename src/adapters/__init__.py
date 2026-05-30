from src.adapters.base import (
    TargetLine, DoorItem,
    HEADERS, SKIP_URL_PARTS, CATALOG_URL_MARKERS, SKIP_TITLE_PATTERNS,
    BAD_IMAGE_KEYWORDS, GENERIC_TITLES, SHORT_LINE_NAMES,
    normalize, transliterate, is_same_domain, clean_url, now_str,
    fetch, text_of, title_of, description_of, image_of, looks_productish,
    is_generic_page, matches_line, dedup_items, normalize_dver_model_name,
    parse_sitemap, sitemap_urls, BaseAdapter,
)

from src.adapters.aelita_doors_ru import AelitaAdapter
from src.adapters.asd_doors_ru import ASDoorsAdapter
from src.adapters.dver_com import DverComAdapter
from src.adapters.entro_design_ru import EntroAdapter
from src.adapters.hausdoors_ru import HausdoorsAdapter
from src.adapters.lord_ru import LordAdapter
from src.adapters.olymp_doors_ru import OlympAdapter
from src.adapters.komfort_doors_ru import KomfortAdapter

ADAPTER_REGISTRY = {
    "aelita": AelitaAdapter,
    "asd": ASDoorsAdapter,
    "dver_com": DverComAdapter,
    "entro": EntroAdapter,
    "hausdoors": HausdoorsAdapter,
    "komfort": KomfortAdapter,
    "lord": LordAdapter,
    "olymp": OlympAdapter,
}


def get_adapter(key: str) -> type[BaseAdapter]:
    return ADAPTER_REGISTRY.get(key, BaseAdapter)
