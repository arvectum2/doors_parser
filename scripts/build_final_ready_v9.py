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

from visual_validate_full_door_images import (
    VisualCheck,
    extract_candidates_from_html,
    is_probable_color_row,
    pick_best_visual_full_door,
    validate_full_door_image,
)


ROOT = Path(__file__).resolve().parent.parent
BASE_DIR = ROOT / "output" / "final_ready_v8"
BASE_CSV = BASE_DIR / "doors_catalog.csv"
TARGETS_CSV = ROOT / "data" / "target_lines.csv"
OUT_DIR = ROOT / "output" / "final_ready_v9"
CHECK_DIR = OUT_DIR / "check_by_producer"
SHEETS_DIR = OUT_DIR / "full_door_contact_sheets"
PROBLEM = ("Олимп", "АСД", "Энтро", "Хаусдорс")
UNCHANGED_SIX = ("Скамбио Порте", "Лорд", "Аэлита", "Комфорт", "Двери регионов", "Термодур")

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
VISUAL_AUDIT_FIELDS = [
    "producer",
    "target_line",
    "model_name",
    "old_image_url",
    "new_image_url",
    "image_width",
    "image_height",
    "object_bbox_width",
    "object_bbox_height",
    "object_aspect_ratio",
    "visual_decision",
    "action",
    "reason",
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
        "badge",
        "banner",
        "hero",
        "interior",
        "texture",
        "sample",
        "pattern",
        "handle",
        "lock",
        "profile",
        "section",
        "empty",
        "placeholder",
        "open.png",
        "catalog.svg",
        "resize/20x",
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


def contact_sheet_html(rows: list[dict], title: str) -> str:
    groups = ("kept", "replaced", "cleared", "deleted_row", "needs_review")
    out = [
        "<!doctype html><html><head><meta charset='utf-8'>",
        f"<title>{escape(title)}</title>",
        "<style>body{font-family:Arial,sans-serif;margin:20px} h2{margin-top:28px} table{border-collapse:collapse;width:100%} th,td{border:1px solid #ddd;padding:6px;font-size:12px;vertical-align:top} th{background:#f4f6f8} img{max-width:130px;max-height:220px;object-fit:contain}</style></head><body>",
        f"<h1>{escape(title)}</h1>",
    ]
    for g in groups:
        sub = [r for r in rows if r.get("action") == g]
        if not sub:
            continue
        out.append(f"<h2>{escape(g)} ({len(sub)})</h2><table><tr><th>producer</th><th>target_line</th><th>model_name</th><th>decision</th><th>reason</th><th>thumbnail</th><th>url</th></tr>")
        for r in sub:
            url = r.get("new_image_url", "") or r.get("old_image_url", "")
            img = f"<img src='{escape(url)}'>" if url else ""
            out.append(
                "<tr>"
                f"<td>{escape(r.get('producer',''))}</td>"
                f"<td>{escape(r.get('target_line',''))}</td>"
                f"<td>{escape(r.get('model_name',''))}</td>"
                f"<td>{escape(r.get('visual_decision',''))}</td>"
                f"<td>{escape(r.get('reason',''))}</td>"
                f"<td>{img}</td>"
                f"<td>{escape(url)}</td>"
                "</tr>"
            )
        out.append("</table>")
    out.append("</body></html>")
    return "".join(out)


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


def candidate_urls_for_source(session: requests.Session, source_url: str, page_cache: dict[str, list[str]]) -> list[str]:
    src = (source_url or "").strip()
    if not src:
        return []
    if src in page_cache:
        return page_cache[src]
    source_list = [src]
    if "dveriolimp.ru" in src:
        alt = olymp_alt_url(src)
        if alt != src:
            source_list.append(alt)
    out, seen = [], set()
    for s in source_list:
        try:
            html = session.get(s, timeout=(8, 35), allow_redirects=True).text
        except Exception:
            continue
        for c in extract_candidates_from_html(s, html):
            if c in seen:
                continue
            seen.add(c)
            out.append(c)
    page_cache[src] = out
    return out


def main() -> None:
    if OUT_DIR.exists():
        shutil.rmtree(OUT_DIR)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    CHECK_DIR.mkdir(parents=True, exist_ok=True)
    SHEETS_DIR.mkdir(parents=True, exist_ok=True)

    base_rows = [{k: r.get(k, "") for k in CATALOG_FIELDS} for r in read_csv(BASE_CSV)]
    targets = read_csv(TARGETS_CSV)

    by_prod = defaultdict(list)
    for r in base_rows:
        by_prod[r.get("producer", "")].append(dict(r))

    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})
    vis_cache: dict[str, VisualCheck] = {}
    page_cache: dict[str, list[str]] = {}

    updated_by_prod: dict[str, list[dict]] = {}
    audit_rows: list[dict] = []
    reparse_rows: list[dict] = []
    per_stats = defaultdict(lambda: defaultdict(int))
    reason_counter = Counter()

    for producer in PROBLEM:
        rows = [dict(r) for r in by_prod.get(producer, [])]
        new_rows: list[dict] = []
        for row in rows:
            old_url = row.get("image_url", "")
            model_name = row.get("model_name", "")

            # Hausdoors color/finish rows must be removed from door catalog.
            if producer == "Хаусдорс" and is_probable_color_row(model_name):
                per_stats[producer]["deleted"] += 1
                reason = "color_sample_model_row_deleted"
                reason_counter["color_sample"] += 1
                audit_rows.append(
                    {
                        "producer": producer,
                        "target_line": row.get("target_line", ""),
                        "model_name": model_name,
                        "old_image_url": old_url,
                        "new_image_url": "",
                        "image_width": "",
                        "image_height": "",
                        "object_bbox_width": "",
                        "object_bbox_height": "",
                        "object_aspect_ratio": "",
                        "visual_decision": "color_sample",
                        "action": "deleted_row",
                        "reason": reason,
                        "comment": row.get("comment", ""),
                    }
                )
                reparse_rows.append(
                    {
                        "producer": producer,
                        "target_line": row.get("target_line", ""),
                        "model_name": model_name,
                        "source_url": row.get("source_url", ""),
                        "old_image_url": old_url,
                        "new_image_url": "",
                        "action": "deleted_row",
                        "reason": reason,
                        "http_status": "",
                        "content_type": "",
                        "width": "",
                        "height": "",
                        "status_after": "",
                        "comment_after": row.get("comment", ""),
                    }
                )
                continue

            old_check = validate_full_door_image(session, old_url, vis_cache) if old_url else VisualCheck("", 0, "", 0, 0, 0, 0, 0.0, "no_image", "no_image", -999.0)

            action = "kept"
            visual_decision = old_check.decision
            reason = old_check.reason
            chosen = old_check

            if old_check.decision != "full_door":
                cands = candidate_urls_for_source(session, row.get("source_url", ""), page_cache)
                best = pick_best_visual_full_door(session, cands, vis_cache)
                if best:
                    row["image_url"] = best.url
                    row["status"] = "found"
                    row["comment"] = append_comment(row.get("comment", ""), "full_door_image_selected")
                    action = "replaced" if old_url and best.url != old_url else "kept"
                    visual_decision = "full_door"
                    reason = "replaced_with_full_door" if action == "replaced" else "full_door_image_ok"
                    chosen = best
                    if action == "replaced":
                        per_stats[producer]["replaced"] += 1
                    else:
                        per_stats[producer]["kept"] += 1
                else:
                    row["image_url"] = ""
                    row["status"] = "needs_review"
                    row["comment"] = append_comment(row.get("comment", ""), "no_full_door_image_found")
                    action = "cleared" if old_url else "needs_review"
                    visual_decision = old_check.decision if old_check.decision != "no_image" else "no_image"
                    reason = old_check.reason if old_check.decision != "no_image" else "no_full_door_candidate_found"
                    chosen = old_check
                    per_stats[producer]["cleared"] += 1 if action == "cleared" else 0
                    per_stats[producer]["needs_review"] += 1 if action == "needs_review" else 0
            else:
                row["image_url"] = old_check.url
                per_stats[producer]["kept"] += 1

            if not row.get("image_url", ""):
                row["status"] = "needs_review"
                row["comment"] = append_comment(row.get("comment", ""), "no_full_door_image_found")
                per_stats[producer]["without_image"] += 1

            # strict: for 4 producers keep only full_door images, otherwise empty
            if row.get("image_url", ""):
                final_check = validate_full_door_image(session, row.get("image_url", ""), vis_cache)
                if final_check.decision != "full_door":
                    row["image_url"] = ""
                    row["status"] = "needs_review"
                    row["comment"] = append_comment(row.get("comment", ""), "no_full_door_image_found")
                    action = "cleared"
                    visual_decision = final_check.decision
                    reason = final_check.reason
                    chosen = final_check
                    per_stats[producer]["cleared"] += 1
                    per_stats[producer]["without_image"] += 1

            if visual_decision in {"color_sample", "texture", "fragment", "interior", "component", "banner"}:
                reason_counter[visual_decision] += 1

            audit_rows.append(
                {
                    "producer": producer,
                    "target_line": row.get("target_line", ""),
                    "model_name": model_name,
                    "old_image_url": old_url,
                    "new_image_url": row.get("image_url", ""),
                    "image_width": chosen.width,
                    "image_height": chosen.height,
                    "object_bbox_width": chosen.bbox_width,
                    "object_bbox_height": chosen.bbox_height,
                    "object_aspect_ratio": chosen.bbox_aspect,
                    "visual_decision": visual_decision,
                    "action": action,
                    "reason": reason,
                    "comment": row.get("comment", ""),
                }
            )
            if old_url != row.get("image_url", "") or action in {"cleared", "replaced", "needs_review"}:
                reparse_rows.append(
                    {
                        "producer": producer,
                        "target_line": row.get("target_line", ""),
                        "model_name": model_name,
                        "source_url": row.get("source_url", ""),
                        "old_image_url": old_url,
                        "new_image_url": row.get("image_url", ""),
                        "action": action,
                        "reason": reason,
                        "http_status": chosen.status_code,
                        "content_type": chosen.content_type,
                        "width": chosen.width,
                        "height": chosen.height,
                        "status_after": row.get("status", ""),
                        "comment_after": row.get("comment", ""),
                    }
                )
            new_rows.append(row)
        updated_by_prod[producer] = new_rows

    # Olymp rule: max one empty photo-N per target_line
    olymp_rows = updated_by_prod.get("Олимп", [])
    by_line = defaultdict(list)
    for r in olymp_rows:
        by_line[r.get("target_line", "")].append(r)
    olymp_final = []
    for line, rows in by_line.items():
        empties = [r for r in rows if re.search(r"(?i)\s[—-]\s*фото\s*\d+\s*$", r.get("model_name", "")) and not r.get("image_url", "").strip()]
        non = [r for r in rows if r not in empties]
        if non:
            olymp_final.extend(non)
        elif empties:
            keep = dict(empties[0])
            keep["model_name"] = line or "Олимп"
            keep["status"] = "needs_review"
            keep["comment"] = append_comment(keep.get("comment", ""), "no_full_door_image_found")
            olymp_final.append(keep)
            for extra in empties[1:]:
                audit_rows.append(
                    {
                        "producer": "Олимп",
                        "target_line": extra.get("target_line", ""),
                        "model_name": extra.get("model_name", ""),
                        "old_image_url": extra.get("image_url", ""),
                        "new_image_url": "",
                        "image_width": "",
                        "image_height": "",
                        "object_bbox_width": "",
                        "object_bbox_height": "",
                        "object_aspect_ratio": "",
                        "visual_decision": "no_image",
                        "action": "deleted_row",
                        "reason": "extra_empty_photoN_removed",
                        "comment": extra.get("comment", ""),
                    }
                )
    updated_by_prod["Олимп"] = olymp_final

    final_rows: list[dict] = []
    for prod, rows in by_prod.items():
        if prod in PROBLEM:
            final_rows.extend(updated_by_prod.get(prod, []))
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
    write_csv(OUT_DIR / "image_reparse_report.csv", reparse_rows, REPARSE_FIELDS)
    write_csv(OUT_DIR / "full_door_visual_audit.csv", audit_rows, VISUAL_AUDIT_FIELDS)

    # contact sheets
    for producer in PROBLEM:
        prs = [r for r in audit_rows if r.get("producer") == producer]
        fname = {
            "Хаусдорс": "hausdoors_full_door_check.html",
            "Энтро": "entro_full_door_check.html",
            "Олимп": "olymp_full_door_check.html",
            "АСД": "asd_full_door_check.html",
        }[producer]
        (SHEETS_DIR / fname).write_text(contact_sheet_html(prs, f"{producer} full-door audit"), encoding="utf-8")

    # per producer exports
    producer_file = {"Скамбио Порте": "scambio", "Олимп": "olymp", "Хаусдорс": "hausdoors", "Лорд": "lord", "Аэлита": "aelita", "Энтро": "entro", "Комфорт": "komfort", "Двери регионов": "dveriregionov", "Термодур": "termodur", "АСД": "asd"}
    producers = sorted({r.get("producer", "") for r in rows_sorted if r.get("producer", "")}, key=lambda x: norm(x))
    for p in producers:
        sub = [r for r in rows_sorted if r.get("producer") == p]
        stem = producer_file.get(p, norm(p).replace(" ", "_"))
        write_csv(CHECK_DIR / f"{stem}.csv", sub, CATALOG_FIELDS)
        write_xlsx(CHECK_DIR / f"{stem}.xlsx", sub, CATALOG_FIELDS, title=stem)

    # unchanged six check
    old_by = defaultdict(list)
    new_by = defaultdict(list)
    for r in base_rows:
        old_by[r.get("producer", "")].append(r)
    for r in rows_sorted:
        new_by[r.get("producer", "")].append(r)
    unchanged_check = {}
    for p in UNCHANGED_SIX:
        a = sorted(old_by.get(p, []), key=lambda r: (r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", ""), r.get("image_url", ""), r.get("status", ""), r.get("comment", "")))
        b = sorted(new_by.get(p, []), key=lambda r: (r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", ""), r.get("image_url", ""), r.get("status", ""), r.get("comment", "")))
        unchanged_check[p] = {"rows_before": len(a), "rows_after": len(b), "exact_match": a == b}

    rows_by_producer = dict(sorted(Counter(r.get("producer", "") for r in rows_sorted if r.get("producer", "")).items()))
    before_by_producer = dict(sorted(Counter(r.get("producer", "") for r in base_rows if r.get("producer", "")).items()))
    problem_summary = {}
    for p in PROBLEM:
        before = before_by_producer.get(p, 0)
        after = rows_by_producer.get(p, 0)
        pr = [r for r in rows_sorted if r.get("producer") == p]
        problem_summary[p] = {
            "rows_before": before,
            "rows_after": after,
            "kept_full_door": per_stats[p].get("kept", 0),
            "replaced": per_stats[p].get("replaced", 0),
            "cleared": per_stats[p].get("cleared", 0),
            "deleted_rows": per_stats[p].get("deleted", 0),
            "without_image": sum(1 for r in pr if not r.get("image_url", "")),
        }

    run_report = {
        "generated_at": datetime.now().isoformat(),
        "mode": "final_ready_v9",
        "source_catalog": str(BASE_CSV),
        "old_total_rows": len(base_rows),
        "total_rows": len(rows_sorted),
        "producers_count": len(producers),
        **coverage,
        "rows_with_image": sum(1 for r in rows_sorted if r.get("image_url", "")),
        "rows_without_image": sum(1 for r in rows_sorted if not r.get("image_url", "")),
        "needs_review_count": sum(1 for r in rows_sorted if r.get("status") == "needs_review"),
        "rows_by_producer": rows_by_producer,
        "problem_producers_summary": problem_summary,
        "visual_decision_counts": {
            "color_sample": reason_counter.get("color_sample", 0),
            "texture": reason_counter.get("texture", 0),
            "fragment": reason_counter.get("fragment", 0),
            "interior": reason_counter.get("interior", 0),
            "component": reason_counter.get("component", 0),
            "banner": reason_counter.get("banner", 0),
        },
        "hausdoors_color_rows_deleted": per_stats["Хаусдорс"].get("deleted", 0),
        "unchanged_producers_check": unchanged_check,
    }
    (OUT_DIR / "run_report.json").write_text(json.dumps(run_report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"total_rows": len(rows_sorted), "rows_by_producer": rows_by_producer, "hausdoors_color_rows_deleted": per_stats["Хаусдорс"].get("deleted", 0)}, ensure_ascii=False))


if __name__ == "__main__":
    main()

