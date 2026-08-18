from __future__ import annotations

import argparse
import csv
import re
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright
from rapidfuzz import fuzz


ROOT = Path(__file__).resolve().parent.parent
TARGET_PRODUCERS = {"Олимп", "АСД", "Энтро", "Хаусдорс"}

CANDIDATE_FIELDS = [
    "producer",
    "target_line",
    "model_name",
    "source_url",
    "current_image_url",
    "candidate_image_url",
    "candidate_index",
    "candidate_block",
    "candidate_text",
    "candidate_alt",
    "candidate_title",
    "candidate_width",
    "candidate_height",
    "is_current_image",
    "parser_decision",
    "reason",
    "expected_label",
    "nearby_text_before",
    "nearby_text_after",
    "dom_order_index",
    "selected_by_dom_order",
    "section",
    "repeated_count",
]


@dataclass
class Candidate:
    image_url: str
    block: str
    text: str
    alt: str
    title: str
    width: int
    height: int
    section: str
    dom_order_index: int
    nearby_text_before: str
    nearby_text_after: str


def normalize(text: str) -> str:
    text = (text or "").lower().replace("ё", "е")
    text = re.sub(r"[^a-zа-я0-9]+", " ", text, flags=re.I)
    return re.sub(r"\s+", " ", text).strip()


def read_csv(path: Path) -> list[dict]:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        return [{k: (v or "").strip() for k, v in row.items()} for row in csv.DictReader(f, delimiter=";")]


def write_csv(path: Path, rows: list[dict], fields: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields, delimiter=";")
        w.writeheader()
        for row in rows:
            w.writerow({k: row.get(k, "") for k in fields})


def bad_hint(url_or_text: str) -> bool:
    low = (url_or_text or "").lower()
    banned = (
        "logo",
        "icon",
        "badge",
        "sticker",
        "label",
        "hit",
        "sale",
        "promo",
        "banner",
        "akcii",
        "акц",
        "new",
        "novinka",
        "sample",
        "texture",
        "pattern",
        "swatch",
        "colors_images",
        "/empty/",
        "placeholder",
        "no_image",
        "no-image",
        "pixel",
        "open.png",
        "search_loop",
        "catalog.svg",
        "sprite",
        "loader",
        "favicon",
        "whatsapp",
        "telegram",
        "vk.com",
        "cover/240x180",
        "/resize/20x",
        "preview",
    )
    return any(b in low for b in banned)


def reason_for_reject(c: Candidate) -> str:
    s = f"{c.image_url} {c.text} {c.alt} {c.title} {c.nearby_text_before} {c.nearby_text_after}".lower()
    if c.section in {"hero"} or any(k in s for k in ("hero", "header", "banner", "cover")):
        return "hero_rejected"
    if any(k in s for k in ("badge", "hit", "sale", "sticker", "label")):
        return "badge_rejected"
    if any(k in s for k in ("texture", "sample", "pattern", "swatch", "материал", "покрытие")):
        return "texture_rejected"
    return "fallback_not_allowed"


def candidate_score(producer: str, model_name: str, c: Candidate) -> float:
    score = 0.0
    if c.width >= 500 and c.height >= 500:
        score += 20
    if c.width * c.height >= 300_000:
        score += 12
    ratio = c.width / max(c.height, 1)
    if ratio > 2.3:
        score -= 25
    if c.section == "model_row":
        score += 26
    if c.block in {"product_gallery", "model_card"}:
        score += 14
    if c.section in {"hero", "footer"}:
        score -= 40
    merged = normalize(f"{c.text} {c.alt} {c.title} {c.image_url} {c.nearby_text_before} {c.nearby_text_after}")
    model_norm = normalize(model_name)
    score += fuzz.token_set_ratio(model_norm, merged) / 5.0
    if bad_hint(merged):
        score -= 40
    if producer == "Энтро" and c.section != "model_row":
        score -= 80
    return score


def normalize_entro_model_name(model_name: str) -> str:
    # "Aria - variant 1" / "Aria — вариант 1" / "Aria 1" -> "aria 1"
    t = normalize(model_name)
    t = t.replace("variant", " ").replace("вариант", " ")
    t = re.sub(r"\s+", " ", t).strip()
    # keep prefix + trailing number
    m = re.search(r"^([a-zа-я0-9 ]+?)\s+(\d{1,2})$", t)
    if m:
        return f"{m.group(1).strip()} {m.group(2)}"
    return t


def label_like(text: str) -> bool:
    t = normalize(text)
    # Aria 1 / Zero 8 / Alliance 3
    return bool(re.search(r"^[a-zа-я][a-zа-я0-9\- ]{1,30}\s\d{1,2}$", t))


def extract_candidates_js() -> str:
    return r"""
() => {
  const out = [];
  const seen = new Set();

  const norm = (s) => (s || '').toLowerCase().replace(/\s+/g, ' ').trim();
  const absUrl = (u) => { try { return new URL(u, location.href).href; } catch { return ''; } };
  const getBgUrl = (el) => {
    const st = window.getComputedStyle(el);
    const bg = st && st.backgroundImage ? st.backgroundImage : '';
    const m = bg.match(/url\((['"]?)(.*?)\1\)/i);
    return m ? m[2] : '';
  };
  const textShort = (el) => (el && el.innerText ? el.innerText.replace(/\s+/g,' ').trim().slice(0,240) : '');
  const isStopTitle = (txt) => {
    const t = norm(txt);
    return t.includes('где купить') || t.includes('другие коллекции') || t.includes('where to buy') || t.includes('other collections');
  };
  const isModelRowTitle = (txt) => norm(txt).includes('модельный ряд');

  // lazy load for sliders/images
  window.scrollTo(0, 0);
  for (let i=0;i<8;i++) window.scrollBy(0, Math.max(450, window.innerHeight));
  window.scrollTo(0, 0);

  const all = [...document.querySelectorAll('body *')];
  const idxMap = new Map();
  all.forEach((el, i) => idxMap.set(el, i));

  // find model-row heading and stop heading by DOM order
  let modelRowIdx = -1;
  let stopIdx = -1;
  for (let i=0;i<all.length;i++){
    const el = all[i];
    const txt = textShort(el);
    if (modelRowIdx < 0 && isModelRowTitle(txt)) {
      modelRowIdx = i;
      continue;
    }
    if (modelRowIdx >= 0 && isStopTitle(txt)) {
      stopIdx = i;
      break;
    }
  }
  if (modelRowIdx >= 0 && stopIdx < 0) stopIdx = all.length;

  const sectionByIndex = (ix) => {
    if (modelRowIdx >= 0 && ix > modelRowIdx && ix < stopIdx) return 'model_row';
    if (modelRowIdx >= 0 && ix <= modelRowIdx) return 'hero';
    if (stopIdx >= 0 && ix >= stopIdx) return 'footer';
    return 'other';
  };

  const nearbyText = (el, dir) => {
    let cur = el;
    for (let depth=0; depth<6 && cur; depth++){
      let sib = dir < 0 ? cur.previousElementSibling : cur.nextElementSibling;
      while (sib){
        const txt = textShort(sib);
        if (txt && txt.length > 1) return txt;
        sib = dir < 0 ? sib.previousElementSibling : sib.nextElementSibling;
      }
      cur = cur.parentElement;
    }
    return '';
  };

  const classifyBlock = (el, imgUrl, section) => {
    if (section === 'model_row') return 'model_row';
    let cur = el;
    for(let i=0;i<7 && cur;i++){
      const idc = `${cur.id||''} ${(cur.className||'')}`.toLowerCase();
      if (idc.includes('gallery') || idc.includes('product') || idc.includes('model') || idc.includes('card') || idc.includes('catalog') || idc.includes('store') || idc.includes('swiper') || idc.includes('slick')) return 'product_gallery';
      if (idc.includes('hero') || idc.includes('header') || idc.includes('banner') || idc.includes('cover')) return 'hero';
      cur = cur.parentElement;
    }
    const low = (imgUrl || '').toLowerCase();
    if (low.includes('hero') || low.includes('banner') || low.includes('cover')) return 'hero';
    return 'other';
  };

  const elems = [...document.querySelectorAll('img, [style*="background-image"]')];
  for (const el of elems) {
    const ix = idxMap.has(el) ? idxMap.get(el) : -1;
    const section = sectionByIndex(ix);
    const blockText = textShort(el.closest('div,section,article,li') || el.parentElement || el);
    const text = blockText || textShort(el.parentElement || el);
    const before = nearbyText(el, -1);
    const after = nearbyText(el, 1);
    const alt = (el.getAttribute && el.getAttribute('alt')) || '';
    const title = (el.getAttribute && el.getAttribute('title')) || '';
    let w = 0, h = 0;
    let urls = [];
    if (el.tagName.toLowerCase() === 'img'){
      w = Number(el.naturalWidth || el.width || 0);
      h = Number(el.naturalHeight || el.height || 0);
      const srcset = el.getAttribute('srcset') || el.getAttribute('data-srcset') || '';
      if (srcset){
        srcset.split(',').map(x=>x.trim().split(' ')[0]).filter(Boolean).forEach(u=>urls.push(u));
      }
      urls.push(el.getAttribute('src') || '', el.getAttribute('data-src') || '', el.getAttribute('data-original') || '');
      // picture/source fallback
      const pic = el.closest('picture');
      if (pic){
        pic.querySelectorAll('source').forEach(s=>{
          const ss = s.getAttribute('srcset') || '';
          if (ss) ss.split(',').map(x=>x.trim().split(' ')[0]).filter(Boolean).forEach(u=>urls.push(u));
        });
      }
    } else {
      const r = el.getBoundingClientRect();
      w = Math.round(r.width || 0);
      h = Math.round(r.height || 0);
      urls.push(getBgUrl(el));
    }

    for (const raw of urls){
      const u = absUrl(raw);
      if (!u) continue;
      const low = u.toLowerCase();
      if (!(low.includes('.jpg') || low.includes('.jpeg') || low.includes('.png') || low.includes('.webp'))) continue;
      const block = classifyBlock(el, u, section);
      const key = `${u}|${ix}|${block}|${text}`;
      if (seen.has(key)) continue;
      seen.add(key);
      out.push({
        image_url: u,
        block,
        text,
        alt,
        title,
        width: w || 0,
        height: h || 0,
        section,
        dom_order_index: ix,
        nearby_text_before: before,
        nearby_text_after: after
      });
    }
  }

  return {
    candidates: out,
    model_row_found: modelRowIdx >= 0,
    model_row_start_index: modelRowIdx,
    model_row_stop_index: stopIdx
  };
}
"""


def entro_select(model_name: str, candidates: list[Candidate]) -> tuple[str, str]:
    expected = normalize_entro_model_name(model_name)
    row = [c for c in candidates if c.section == "model_row"]
    if not row:
        return "", "no_model_card_found"

    # 1) exact match by nearby label after image (DOM-order heuristic)
    for c in sorted(row, key=lambda x: x.dom_order_index):
        after = normalize(c.nearby_text_after)
        before = normalize(c.nearby_text_before)
        text = normalize(c.text)
        alt_title = normalize(f"{c.alt} {c.title}")
        if expected and any(expected == t or expected in t for t in (after, before, text, alt_title)):
            return c.image_url, "dom_order_model_row_match"

    # 2) fuzzy by combined nearby text
    best = None
    best_score = 0
    for c in row:
        merged = normalize(f"{c.nearby_text_before} {c.nearby_text_after} {c.text} {c.alt} {c.title}")
        score = fuzz.token_set_ratio(expected, merged)
        if score > best_score:
            best_score = score
            best = c
    if best and best_score >= 82:
        return best.image_url, "model_row_fuzzy_match"

    # 3) strict no-card
    return "", "no_model_card_found"


def choose_selected(producer: str, model_name: str, current: str, candidates: list[Candidate]) -> tuple[str, str]:
    if not candidates:
        return "", "no_model_card_found"
    if producer == "Энтро":
        return entro_select(model_name, candidates)

    valid = [
        c
        for c in candidates
        if not bad_hint(f"{c.image_url} {c.text} {c.alt} {c.title} {c.nearby_text_before} {c.nearby_text_after}")
        and c.width >= 260
        and c.height >= 260
        and c.section != "hero"
    ]
    if not valid:
        return "", "fallback_not_allowed"

    # keep current if valid and explicit model/gallery area
    for c in valid:
        if current and c.image_url == current and c.block in {"product_gallery", "model_row"}:
            return c.image_url, "product_gallery_image"

    scored = sorted(valid, key=lambda c: candidate_score(producer, model_name, c), reverse=True)
    top = scored[0]
    if candidate_score(producer, model_name, top) >= 18:
        return top.image_url, "product_gallery_image"
    return "", "fallback_not_allowed"


def write_candidates_xlsx(path: Path, by_producer_rows: dict[str, list[dict]]) -> None:
    wb = Workbook()
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]
    sheet_order = [("Энтро", "Entro"), ("Хаусдорс", "Hausdoors"), ("АСД", "ASD"), ("Олимп", "Olymp")]
    for producer_ru, title in sheet_order:
        ws = wb.create_sheet(title=title)
        ws.append(CANDIDATE_FIELDS)
        for row in by_producer_rows.get(producer_ru, []):
            ws.append([row.get(k, "") for k in CANDIDATE_FIELDS])
        widths = [14, 22, 28, 50, 46, 46, 8, 16, 28, 24, 24, 10, 10, 8, 18, 24, 18, 26, 26, 10, 10, 12, 10]
        for col, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(col)].width = w
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F4E78")
            c.alignment = Alignment(horizontal="center", vertical="center")
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, help="Path to doors_catalog.csv")
    parser.add_argument("--out", required=True, help="Output directory")
    args = parser.parse_args()

    input_path = Path(args.input)
    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)

    rows = read_csv(input_path)
    target_rows = [r for r in rows if r.get("producer") in TARGET_PRODUCERS]

    page_cache: dict[tuple[str, str], dict] = {}
    by_producer_output: dict[str, list[dict]] = defaultdict(list)

    stats_checked = Counter()
    stats_exact = Counter()
    stats_dom = Counter()
    stats_fuzzy = Counter()
    stats_no_card = Counter()
    stats_needs_manual = Counter()
    stats_will_clear = Counter()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={"width": 1440, "height": 2200}, ignore_https_errors=True)
        page = context.new_page()

        for idx, row in enumerate(target_rows, 1):
            producer = row.get("producer", "")
            source_url = row.get("source_url", "")
            model_name = row.get("model_name", "")
            current_img = row.get("image_url", "")
            target_line = row.get("target_line", "")
            stats_checked[producer] += 1

            cache_key = (producer, source_url)
            if cache_key not in page_cache:
                candidates: list[Candidate] = []
                page_info = {"model_row_found": False}
                if source_url:
                    try:
                        page.goto(source_url, wait_until="domcontentloaded", timeout=70000)
                        page.wait_for_timeout(2200)
                        payload = page.evaluate(extract_candidates_js())
                        page_info = {
                            "model_row_found": bool((payload or {}).get("model_row_found")),
                            "model_row_start_index": (payload or {}).get("model_row_start_index", -1),
                            "model_row_stop_index": (payload or {}).get("model_row_stop_index", -1),
                        }
                        for obj in (payload or {}).get("candidates", []) or []:
                            url = (obj.get("image_url") or "").strip()
                            if not url:
                                continue
                            candidates.append(
                                Candidate(
                                    image_url=url,
                                    block=(obj.get("block") or "other").strip() or "other",
                                    text=(obj.get("text") or "").strip(),
                                    alt=(obj.get("alt") or "").strip(),
                                    title=(obj.get("title") or "").strip(),
                                    width=int(obj.get("width") or 0),
                                    height=int(obj.get("height") or 0),
                                    section=(obj.get("section") or "other").strip() or "other",
                                    dom_order_index=int(obj.get("dom_order_index") or -1),
                                    nearby_text_before=(obj.get("nearby_text_before") or "").strip(),
                                    nearby_text_after=(obj.get("nearby_text_after") or "").strip(),
                                )
                            )
                    except PlaywrightTimeoutError:
                        candidates = []
                    except Exception:
                        candidates = []
                page_cache[cache_key] = {"candidates": candidates, **page_info}

            payload = page_cache[cache_key]
            candidates: list[Candidate] = payload["candidates"]
            selected_url, selected_reason = choose_selected(producer, model_name, current_img, candidates)

            if selected_reason == "model_row_exact_match":
                stats_exact[producer] += 1
            elif selected_reason == "dom_order_model_row_match":
                stats_dom[producer] += 1
            elif selected_reason == "model_row_fuzzy_match":
                stats_fuzzy[producer] += 1
            elif selected_reason == "no_model_card_found":
                stats_no_card[producer] += 1
            if not selected_url:
                stats_needs_manual[producer] += 1
                if current_img:
                    stats_will_clear[producer] += 1

            sorted_candidates = sorted(candidates, key=lambda c: candidate_score(producer, model_name, c), reverse=True)[:10]

            # ensure selected candidate is present in output rows
            if selected_url and not any(c.image_url == selected_url for c in sorted_candidates):
                for c in candidates:
                    if c.image_url == selected_url:
                        sorted_candidates.append(c)
                        break

            repeated_counter = Counter(c.image_url for c in candidates)
            expected_label = normalize_entro_model_name(model_name) if producer == "Энтро" else ""

            out_rows = []
            for cidx, c in enumerate(sorted_candidates, 1):
                if selected_url and c.image_url == selected_url:
                    decision = "selected"
                    reason = selected_reason
                else:
                    decision = "rejected"
                    reason = reason_for_reject(c)

                selected_by_dom = "true" if (decision == "selected" and reason == "dom_order_model_row_match") else "false"
                out_rows.append(
                    {
                        "producer": producer,
                        "target_line": target_line,
                        "model_name": model_name,
                        "source_url": source_url,
                        "current_image_url": current_img,
                        "candidate_image_url": c.image_url,
                        "candidate_index": cidx,
                        "candidate_block": c.block,
                        "candidate_text": c.text,
                        "candidate_alt": c.alt,
                        "candidate_title": c.title,
                        "candidate_width": c.width,
                        "candidate_height": c.height,
                        "is_current_image": "1" if current_img and c.image_url == current_img else "0",
                        "parser_decision": decision,
                        "reason": reason,
                        "expected_label": expected_label,
                        "nearby_text_before": c.nearby_text_before,
                        "nearby_text_after": c.nearby_text_after,
                        "dom_order_index": c.dom_order_index,
                        "selected_by_dom_order": selected_by_dom,
                        "section": c.section,
                        "repeated_count": repeated_counter.get(c.image_url, 1),
                    }
                )

            if not selected_url:
                out_rows.append(
                    {
                        "producer": producer,
                        "target_line": target_line,
                        "model_name": model_name,
                        "source_url": source_url,
                        "current_image_url": current_img,
                        "candidate_image_url": "",
                        "candidate_index": "",
                        "candidate_block": "",
                        "candidate_text": "",
                        "candidate_alt": "",
                        "candidate_title": "",
                        "candidate_width": "",
                        "candidate_height": "",
                        "is_current_image": "0",
                        "parser_decision": "needs_manual_choice",
                        "reason": selected_reason if selected_reason else "no_model_card_found",
                        "expected_label": expected_label,
                        "nearby_text_before": "",
                        "nearby_text_after": "",
                        "dom_order_index": "",
                        "selected_by_dom_order": "false",
                        "section": "",
                        "repeated_count": "",
                    }
                )

            by_producer_output[producer].extend(out_rows)

            if idx % 100 == 0:
                print(f"processed {idx}/{len(target_rows)}")

        browser.close()

    mapping = {
        "Энтро": "entro_candidates.csv",
        "Хаусдорс": "hausdoors_candidates.csv",
        "АСД": "asd_candidates.csv",
        "Олимп": "olymp_candidates.csv",
    }
    for producer, filename in mapping.items():
        write_csv(out_dir / filename, by_producer_output.get(producer, []), CANDIDATE_FIELDS)

    write_candidates_xlsx(out_dir / "image_candidates.xlsx", by_producer_output)

    manual_template_fields = ["producer", "target_line", "model_name", "source_url", "selected_image_url", "decision", "comment"]
    template_rows = [
        {
            "producer": r.get("producer", ""),
            "target_line": r.get("target_line", ""),
            "model_name": r.get("model_name", ""),
            "source_url": r.get("source_url", ""),
            "selected_image_url": "",
            "decision": "",
            "comment": "",
        }
        for r in target_rows
    ]
    write_csv(out_dir / "manual_selection_template.csv", template_rows, manual_template_fields)

    # summary requested by user
    print("candidate extraction completed")
    print(
        f"Entro checked={stats_checked['Энтро']} exact_model_row_match={stats_exact['Энтро']} "
        f"dom_order_model_row_match={stats_dom['Энтро']} no_model_card_found={stats_no_card['Энтро']} "
        f"will_clear_image={stats_will_clear['Энтро']}"
    )
    haus_rows = by_producer_output.get("Хаусдорс", [])
    repeated_gt10 = sum(
        1
        for r in haus_rows
        if r.get("candidate_image_url")
        and r.get("repeated_count")
        and str(r.get("repeated_count")).isdigit()
        and int(r.get("repeated_count")) > 10
        and r.get("parser_decision") == "selected"
    )
    print(f"Hausdoors repeated image_url > 10 count={repeated_gt10}")
    print(f"ASD needs_manual_choice={stats_needs_manual['АСД']}")
    print(f"Olymp needs_manual_choice={stats_needs_manual['Олимп']}")
    print(f"xlsx={out_dir / 'image_candidates.xlsx'}")
    print(f"manual_template={out_dir / 'manual_selection_template.csv'}")


if __name__ == "__main__":
    main()

