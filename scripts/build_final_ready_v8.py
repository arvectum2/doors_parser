from __future__ import annotations

import csv
import json
import re
import shutil
from collections import Counter, defaultdict
from datetime import datetime
from html import escape
from pathlib import Path
from urllib.parse import urlparse

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from validate_full_door_images import ImageDecision, extract_image_candidates, pick_best_full_door, validate_image_url


ROOT = Path(__file__).resolve().parent.parent
BASE_DIR = ROOT / "output" / "final_ready_v7"
BASE_CSV = BASE_DIR / "doors_catalog.csv"
TARGETS_CSV = ROOT / "data" / "target_lines.csv"
OUT_DIR = ROOT / "output" / "final_ready_v8"
CHECK_DIR = OUT_DIR / "check_by_producer"
PROBLEM = {"Олимп", "АСД", "Энтро", "Хаусдорс"}
KEEP_UNCHANGED = {"Скамбио Порте", "Лорд", "Аэлита", "Комфорт", "Двери регионов", "Термодур"}

CATALOG_FIELDS = [
    "door_type",
    "producer",
    "target_line",
    "model_name",
    "description",
    "material",
    "coating",
    "image_url",
    "source_url",
    "status",
    "comment",
]
WP_FIELDS = [
    "post_title",
    "post_content",
    "producer",
    "line",
    "material",
    "coating",
    "image_url",
    "source_url",
    "door_type",
    "status",
]
QA_FIELDS = [
    "producer",
    "target_line",
    "expected_material",
    "rows_count",
    "unique_models_count",
    "with_image_count",
    "missing_image_count",
    "needs_review_count",
    "bad_image_count",
    "status",
    "comment",
]
NOT_FOUND_FIELDS = [
    "door_type",
    "manufacturer",
    "target_line",
    "source_site",
    "status",
    "comment",
]
REPARSE_FIELDS = [
    "producer",
    "target_line",
    "model_name",
    "source_url",
    "old_image_url",
    "new_image_url",
    "action",
    "reason",
    "http_status",
    "content_type",
    "width",
    "height",
    "status_after",
    "comment_after",
]
AUDIT_FIELDS = [
    "producer",
    "target_line",
    "model_name",
    "source_url",
    "old_image_url",
    "new_image_url",
    "old_decision",
    "new_decision",
    "reason",
    "width",
    "height",
    "aspect_ratio",
    "action",
    "comment",
]
REBUILD_FIELDS = [
    "producer",
    "old_rows",
    "new_rows",
    "old_variant_rows",
    "new_variant_rows",
    "real_model_rows",
    "generated_model_rows",
    "rows_with_image",
    "rows_without_image",
    "repeated_image_gt_10_count",
    "needs_review_count",
    "comment",
]


def norm(text: str) -> str:
    text = (text or "").lower().replace("ё", "е")
    text = re.sub(r"[^a-zа-я0-9]+", " ", text, flags=re.I)
    return re.sub(r"\s+", " ", text).strip()


def norm_model(text: str) -> str:
    text = norm(text)
    text = re.sub(r"\b(дверь|model|модель|variant|вариант|фото)\b", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def append_comment(comment: str, addon: str) -> str:
    comment = (comment or "").strip()
    addon = (addon or "").strip()
    if not addon:
        return comment
    if not comment:
        return addon
    if addon in comment:
        return comment
    return f"{comment}; {addon}"


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


def write_xlsx(path: Path, rows: list[dict], fields: list[str], title: str = "parsed_catalog") -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = title[:31]
    ws.append(fields)
    for row in rows:
        ws.append([row.get(k, "") for k in fields])
    widths = [14, 22, 28, 40, 70, 18, 16, 60, 60, 14, 42][: len(fields)]
    for col, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = width
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"
    wb.save(path)


def bad_image_url(url: str) -> bool:
    low = (url or "").lower()
    if not low:
        return True
    bad = (
        "logo",
        "icon",
        "badge",
        "sticker",
        "label",
        "promo",
        "search",
        "arrow",
        "placeholder",
        "no_image",
        "no-image",
        "open.png",
        "catalog.svg",
        "sprite",
        "loader",
        "favicon",
        "resize/20x",
        "/resize/20x/",
        "close.png",
        "free-icon",
        "instagram",
        "vk-social",
        "phone-call",
        "email-2152",
        "tl.png",
        "tg.png",
        "empty",
        "pixel",
    )
    return any(x in low for x in bad)


def recalc_qa(rows: list[dict], target_rows: list[dict]) -> tuple[list[dict], list[dict], dict]:
    grouped: dict[tuple[str, str], list[dict]] = defaultdict(list)
    for row in rows:
        grouped[(norm(row.get("producer", "")), norm(row.get("target_line", "")))].append(row)
    qa_rows, not_found_rows = [], []
    lt = lr = lp = lnf = 0
    for t in target_rows:
        p = (t.get("manufacturer") or "").strip()
        l = (t.get("target_line") or "").strip()
        if not p or not l:
            continue
        lt += 1
        key = (norm(p), norm(l))
        sub = grouped.get(key, [])
        if not sub and norm(p) == norm("АСД") and norm(l) == norm("Все модели"):
            sub = grouped.get((norm("АСД"), norm("Уточнить коллекцию")), [])
        with_img = sum(1 for r in sub if r.get("image_url", ""))
        miss = len(sub) - with_img
        nr = sum(1 for r in sub if r.get("status") == "needs_review")
        bad = sum(1 for r in sub if bad_image_url(r.get("image_url", "")))
        uniq = len({norm_model(r.get("model_name", "")) for r in sub if r.get("model_name", "")})
        placeholders = [r for r in sub if "collection_not_found_on_site" in (r.get("comment") or "").lower()]
        real = [r for r in sub if r not in placeholders]
        if not sub:
            st = "not_found"
            cm = "checked_catalog_no_rows_for_target_line"
            lnf += 1
            not_found_rows.append({"door_type": t.get("door_type", ""), "manufacturer": p, "target_line": l, "source_site": t.get("source_site", ""), "status": "not_found", "comment": cm})
        elif real:
            lr += 1
            st = "needs_review" if (miss or nr or bad) else "found"
            cm = f"with_image={with_img}; missing_image={miss}; needs_review={nr}; bad_image={bad}"
        else:
            lp += 1
            st = "needs_review"
            cm = "collection_not_found_on_site"
        qa_rows.append({"producer": p, "target_line": l, "expected_material": (t.get("covering_from_source") or t.get("material_manual") or "").strip(), "rows_count": len(sub), "unique_models_count": uniq, "with_image_count": with_img, "missing_image_count": miss, "needs_review_count": nr, "bad_image_count": bad, "status": st, "comment": cm})
    return qa_rows, not_found_rows, {"target_lines_total": lt, "target_lines_with_real_models": lr, "target_lines_as_placeholder": lp, "target_lines_not_found": lnf}


def olymp_alt_url(source_url: str) -> str:
    p = urlparse(source_url)
    path = p.path.strip("/")
    if not path:
        return source_url
    slug = path.split("/")[-1]
    slug2 = re.sub(r"-\d+$", "", slug)
    if slug2 == slug:
        return source_url
    return f"{p.scheme}://{p.netloc}/{'/'.join(path.split('/')[:-1] + [slug2])}"


def get_candidates_for_source(
    session: requests.Session,
    source_url: str,
    page_cache: dict[str, list[str]],
) -> list[str]:
    src = (source_url or "").strip()
    if not src:
        return []
    if src in page_cache:
        return page_cache[src]
    urls = [src]
    if "dveriolimp.ru" in src:
        alt = olymp_alt_url(src)
        if alt != src:
            urls.append(alt)
    all_candidates: list[str] = []
    seen = set()
    for u in urls:
        try:
            html = session.get(u, timeout=(8, 30), allow_redirects=True).text
        except Exception:
            continue
        for c in extract_image_candidates(u, html):
            if c in seen:
                continue
            seen.add(c)
            all_candidates.append(c)
    page_cache[src] = all_candidates
    return all_candidates


def collapse_olymp_empty_photon(rows: list[dict]) -> tuple[list[dict], int]:
    by_line = defaultdict(list)
    for r in rows:
        by_line[r.get("target_line", "")].append(r)
    out = []
    removed = 0
    for line, items in by_line.items():
        empty_photon = [r for r in items if re.search(r"(?i)\s[—-]\s*фото\s*\d+\s*$", r.get("model_name", "")) and not r.get("image_url", "").strip()]
        if not empty_photon:
            out.extend(items)
            continue
        keep = [r for r in items if r not in empty_photon]
        removed += len(empty_photon)
        if keep:
            out.extend(keep)
            continue
        placeholder = dict(empty_photon[0])
        placeholder["model_name"] = line or "Олимп"
        placeholder["image_url"] = ""
        placeholder["status"] = "needs_review"
        placeholder["comment"] = append_comment(placeholder.get("comment", ""), "no_full_door_image_found")
        out.append(placeholder)
    return out, removed


def preview_html(audit_rows: list[dict]) -> str:
    head = """<!doctype html><html><head><meta charset="utf-8"><title>Full Door Preview</title>
<style>
body{font-family:Arial,sans-serif;margin:20px} table{border-collapse:collapse;width:100%}
th,td{border:1px solid #ddd;padding:6px;vertical-align:top;font-size:12px}
th{background:#f4f6f8} img{max-width:130px;max-height:220px;object-fit:contain}
.c{color:#b00020}.r{color:#0d652d}
</style></head><body><h2>full_door_image_preview</h2>"""
    sections = []
    for producer in ("Олимп", "АСД", "Энтро", "Хаусдорс"):
        rows = [r for r in audit_rows if r.get("producer") == producer]
        if not rows:
            continue
        sections.append(f"<h3>{escape(producer)} ({len(rows)})</h3>")
        sections.append("<table><tr><th>target_line</th><th>model_name</th><th>decision</th><th>reason</th><th>image</th><th>url</th></tr>")
        for r in rows:
            dec = r.get("new_decision", "")
            cls = "r" if dec == "keep_full_door_image" or dec == "replace_with_full_door_image" else "c"
            url = r.get("new_image_url", "") or ""
            img = f'<img src="{escape(url)}">' if url else ""
            sections.append(
                "<tr>"
                f"<td>{escape(r.get('target_line',''))}</td>"
                f"<td>{escape(r.get('model_name',''))}</td>"
                f"<td class='{cls}'>{escape(dec)}</td>"
                f"<td>{escape(r.get('reason',''))}</td>"
                f"<td>{img}</td>"
                f"<td>{escape(url)}</td>"
                "</tr>"
            )
        sections.append("</table>")
    return head + "".join(sections) + "</body></html>"


def map_reason(reason: str) -> str:
    r = (reason or "").lower()
    if "interior" in r:
        return "interior_rejected"
    if "fragment" in r:
        return "fragment_rejected"
    if "texture" in r:
        return "texture_rejected"
    if "component" in r:
        return "component_rejected"
    if "banner" in r or "badge" in r:
        return "banner_rejected"
    if "not_vertical" in r:
        return "not_vertical_product_image"
    if "full_door_image_ok" in r:
        return "full_door_image_ok"
    return "no_full_door_candidate_found"


def asd_relaxed_accept(dec: ImageDecision) -> ImageDecision:
    if dec.decision == "accept":
        return dec
    low = (dec.url or "").lower()
    if (
        "as-doors.ru/assets/cache_image/" in low
        and dec.http_status == 200
        and dec.content_type.startswith("image/")
        and dec.width >= 700
        and dec.height >= 500
        and dec.aspect_ratio >= 0.5
        and not any(x in low for x in ("banner", "badge", "icon", "logo", "placeholder", "no_image", "empty"))
    ):
        return ImageDecision(
            dec.url,
            "accept",
            "full_door_image_ok",
            dec.http_status,
            dec.content_type,
            dec.width,
            dec.height,
            dec.aspect_ratio,
            max(dec.score, 1.2),
        )
    return dec


def main() -> None:
    if OUT_DIR.exists():
        shutil.rmtree(OUT_DIR)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    CHECK_DIR.mkdir(parents=True, exist_ok=True)

    base_rows = [{k: r.get(k, "") for k in CATALOG_FIELDS} for r in read_csv(BASE_CSV)]
    targets = read_csv(TARGETS_CSV)
    by_prod = defaultdict(list)
    for r in base_rows:
        by_prod[r.get("producer", "")].append(dict(r))

    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})
    decision_cache: dict[str, ImageDecision] = {}
    page_cache: dict[str, list[str]] = {}

    updated_by_prod: dict[str, list[dict]] = {}
    full_audit: list[dict] = []
    image_reparse_rows: list[dict] = []
    stats = defaultdict(lambda: {"kept": 0, "replaced": 0, "cleared": 0, "without": 0})
    rejection_counts = Counter()

    for producer in PROBLEM:
        rows = [dict(r) for r in by_prod.get(producer, [])]
        out_rows = []
        for row in rows:
            old_url = row.get("image_url", "")
            old_dec = validate_image_url(session, old_url, decision_cache) if old_url else ImageDecision("", "reject", "no_url", 0, "", 0, 0, 0.0, -999.0)
            if producer == "АСД":
                old_dec = asd_relaxed_accept(old_dec)

            chosen_dec = old_dec
            action = "kept"
            final_decision = "keep_full_door_image"
            reason = "full_door_image_ok"

            if old_dec.decision != "accept":
                cands = get_candidates_for_source(session, row.get("source_url", ""), page_cache)
                if producer == "АСД":
                    best = None
                    for cand in cands:
                        d = asd_relaxed_accept(validate_image_url(session, cand, decision_cache))
                        if d.decision != "accept":
                            continue
                        if best is None or d.score > best.score:
                            best = d
                else:
                    best = pick_best_full_door(session, cands, decision_cache)
                if best:
                    chosen_dec = best
                    row["image_url"] = best.url
                    row["status"] = "found"
                    row["comment"] = append_comment(row.get("comment", ""), "full_door_image_selected")
                    action = "replaced" if old_url and old_url != best.url else "kept"
                    final_decision = "replace_with_full_door_image" if action == "replaced" else "keep_full_door_image"
                    reason = "replaced_with_full_door" if action == "replaced" else "full_door_image_ok"
                    if action == "replaced":
                        stats[producer]["replaced"] += 1
                    else:
                        stats[producer]["kept"] += 1
                else:
                    row["image_url"] = ""
                    row["status"] = "needs_review"
                    row["comment"] = append_comment(row.get("comment", ""), "no_full_door_image_found")
                    action = "cleared" if old_url else "no_candidate_found"
                    final_decision = "clear_not_full_door"
                    reason = map_reason(old_dec.reason) if old_url else "no_full_door_candidate_found"
                    if action == "cleared":
                        stats[producer]["cleared"] += 1
                    rejection_counts[reason] += 1
            else:
                # keep accepted current image
                row["image_url"] = old_dec.url
                stats[producer]["kept"] += 1

            if not row.get("image_url", ""):
                row["status"] = "needs_review"
                row["comment"] = append_comment(row.get("comment", ""), "no_full_door_image_found")
                stats[producer]["without"] += 1

            full_audit.append(
                {
                    "producer": producer,
                    "target_line": row.get("target_line", ""),
                    "model_name": row.get("model_name", ""),
                    "source_url": row.get("source_url", ""),
                    "old_image_url": old_url,
                    "new_image_url": row.get("image_url", ""),
                    "old_decision": "keep_full_door_image" if old_dec.decision == "accept" else "clear_not_full_door",
                    "new_decision": final_decision,
                    "reason": reason,
                    "width": chosen_dec.width,
                    "height": chosen_dec.height,
                    "aspect_ratio": chosen_dec.aspect_ratio,
                    "action": action,
                    "comment": row.get("comment", ""),
                }
            )
            if old_url != row.get("image_url", "") or old_dec.decision != "accept":
                image_reparse_rows.append(
                    {
                        "producer": producer,
                        "target_line": row.get("target_line", ""),
                        "model_name": row.get("model_name", ""),
                        "source_url": row.get("source_url", ""),
                        "old_image_url": old_url,
                        "new_image_url": row.get("image_url", ""),
                        "action": action,
                        "reason": reason,
                        "http_status": chosen_dec.http_status,
                        "content_type": chosen_dec.content_type,
                        "width": chosen_dec.width,
                        "height": chosen_dec.height,
                        "status_after": row.get("status", ""),
                        "comment_after": row.get("comment", ""),
                    }
                )
            out_rows.append(row)

        if producer == "Олимп":
            out_rows, removed = collapse_olymp_empty_photon(out_rows)
            stats[producer]["cleared"] += 0
            stats[producer]["removed_empty_photon"] = removed
        updated_by_prod[producer] = out_rows

    final_rows: list[dict] = []
    for producer, rows in by_prod.items():
        if producer in PROBLEM:
            final_rows.extend(updated_by_prod.get(producer, []))
        else:
            final_rows.extend(rows)

    rows_sorted = sorted(final_rows, key=lambda r: (norm(r.get("producer", "")), norm(r.get("target_line", "")), norm(r.get("model_name", ""))))
    write_csv(OUT_DIR / "doors_catalog.csv", rows_sorted, CATALOG_FIELDS)
    write_xlsx(OUT_DIR / "doors_catalog.xlsx", rows_sorted, CATALOG_FIELDS)
    wp_rows = [{"post_title": r.get("model_name", ""), "post_content": r.get("description", ""), "producer": r.get("producer", ""), "line": r.get("target_line", ""), "material": r.get("material", ""), "coating": r.get("coating", ""), "image_url": r.get("image_url", ""), "source_url": r.get("source_url", ""), "door_type": r.get("door_type", ""), "status": r.get("status", "")} for r in rows_sorted]
    write_csv(OUT_DIR / "wordpress_import.csv", wp_rows, WP_FIELDS)

    qa_rows, not_found_rows, coverage = recalc_qa(rows_sorted, targets)
    write_csv(OUT_DIR / "qa_report.csv", qa_rows, QA_FIELDS)
    write_csv(OUT_DIR / "not_found.csv", not_found_rows, NOT_FOUND_FIELDS)
    write_csv(OUT_DIR / "image_reparse_report.csv", image_reparse_rows, REPARSE_FIELDS)
    write_csv(OUT_DIR / "full_door_image_audit.csv", full_audit, AUDIT_FIELDS)
    (OUT_DIR / "full_door_image_preview.html").write_text(preview_html(full_audit), encoding="utf-8")

    # rebuild report for 4 producers
    rebuild_rows = []
    for p in ("Олимп", "АСД", "Энтро", "Хаусдорс"):
        old_rows = by_prod.get(p, [])
        new_rows = updated_by_prod.get(p, [])
        cnt = Counter(r.get("image_url", "") for r in new_rows if r.get("image_url", ""))
        rebuild_rows.append(
            {
                "producer": p,
                "old_rows": len(old_rows),
                "new_rows": len(new_rows),
                "old_variant_rows": 0,
                "new_variant_rows": 0,
                "real_model_rows": sum(1 for r in new_rows if r.get("status") == "found"),
                "generated_model_rows": sum(1 for r in new_rows if re.search(r"(?i)\s[—-]\s*фото\s*\d+\s*$", r.get("model_name", ""))),
                "rows_with_image": sum(1 for r in new_rows if r.get("image_url", "")),
                "rows_without_image": sum(1 for r in new_rows if not r.get("image_url", "")),
                "repeated_image_gt_10_count": sum(1 for _, c in cnt.items() if c > 10),
                "needs_review_count": sum(1 for r in new_rows if r.get("status") == "needs_review"),
                "comment": "v8_full_door_image_validation",
            }
        )
    write_csv(OUT_DIR / "rebuild_report.csv", rebuild_rows, REBUILD_FIELDS)

    # per-producer files
    producer_file = {"Скамбио Порте": "scambio", "Олимп": "olymp", "Хаусдорс": "hausdoors", "Лорд": "lord", "Аэлита": "aelita", "Энтро": "entro", "Комфорт": "komfort", "Двери регионов": "dveriregionov", "Термодур": "termodur", "АСД": "asd"}
    producers = sorted({r.get("producer", "") for r in rows_sorted if r.get("producer", "")}, key=lambda x: norm(x))
    for p in producers:
        sub = [r for r in rows_sorted if r.get("producer") == p]
        stem = producer_file.get(p, norm(p).replace(" ", "_"))
        write_csv(CHECK_DIR / f"{stem}.csv", sub, CATALOG_FIELDS)
        write_xlsx(CHECK_DIR / f"{stem}.xlsx", sub, CATALOG_FIELDS, title=stem)

    # unchanged check for 6 stable producers
    old_by_prod = defaultdict(list)
    new_by_prod = defaultdict(list)
    for r in base_rows:
        old_by_prod[r.get("producer", "")].append(r)
    for r in rows_sorted:
        new_by_prod[r.get("producer", "")].append(r)
    unchanged = {}
    for p in KEEP_UNCHANGED:
        a = sorted(old_by_prod.get(p, []), key=lambda r: (r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", ""), r.get("image_url", ""), r.get("status", ""), r.get("comment", "")))
        b = sorted(new_by_prod.get(p, []), key=lambda r: (r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", ""), r.get("image_url", ""), r.get("status", ""), r.get("comment", "")))
        unchanged[p] = {"rows_before": len(a), "rows_after": len(b), "exact_match": a == b}

    rows_by_producer = dict(sorted(Counter(r.get("producer", "") for r in rows_sorted if r.get("producer", "")).items()))
    producer_stats = {}
    for p in PROBLEM:
        pr = [r for r in rows_sorted if r.get("producer") == p]
        audit = [a for a in full_audit if a.get("producer") == p]
        producer_stats[p] = {
            "rows_total": len(pr),
            "kept_full_door_images": sum(1 for a in audit if a.get("action") == "kept"),
            "replaced_images": sum(1 for a in audit if a.get("action") == "replaced"),
            "cleared_images": sum(1 for a in audit if a.get("action") == "cleared"),
            "rows_without_image": sum(1 for r in pr if not r.get("image_url", "")),
        }

    run_report = {
        "generated_at": datetime.now().isoformat(),
        "mode": "final_ready_v8",
        "source_catalog": str(BASE_CSV),
        "old_total_rows": len(base_rows),
        "total_rows": len(rows_sorted),
        "producers_count": len(producers),
        **coverage,
        "rows_with_image": sum(1 for r in rows_sorted if r.get("image_url", "")),
        "rows_without_image": sum(1 for r in rows_sorted if not r.get("image_url", "")),
        "needs_review_count": sum(1 for r in rows_sorted if r.get("status") == "needs_review"),
        "rows_by_producer": rows_by_producer,
        "problem_producers_stats": producer_stats,
        "rejection_counts": {
            "interior": rejection_counts.get("interior_rejected", 0),
            "fragment": rejection_counts.get("fragment_rejected", 0),
            "texture": rejection_counts.get("texture_rejected", 0),
            "component": rejection_counts.get("component_rejected", 0),
            "banner_badge": rejection_counts.get("banner_rejected", 0),
        },
        "olymp_empty_photoN_removed": stats["Олимп"].get("removed_empty_photon", 0),
        "unchanged_producers_check": unchanged,
    }
    (OUT_DIR / "run_report.json").write_text(json.dumps(run_report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"total_rows": len(rows_sorted), "rows_by_producer": rows_by_producer, "olymp_empty_photoN_removed": stats["Олимп"].get("removed_empty_photon", 0)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
