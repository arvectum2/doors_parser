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

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


ROOT = Path(__file__).resolve().parent.parent
BASE_CSV = ROOT / "output" / "final_ready_v11" / "doors_catalog.csv"
TARGETS_CSV = ROOT / "data" / "target_lines.csv"
OUT_DIR = ROOT / "output" / "final_ready_v13"
CHECK_DIR = OUT_DIR / "check_by_producer"
REVIEW_DIR = OUT_DIR / "full_visual_review"

TOUCH_PRODUCERS = {"Хаусдорс", "Энтро", "Олимп"}
UNCHANGED_SIX = {"Скамбио Порте", "Лорд", "Аэлита", "Комфорт", "Двери регионов", "Термодур"}

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
REPEATED_FIELDS = [
    "producer",
    "target_line",
    "image_url",
    "repeated_count",
    "affected_model_names",
    "action",
    "reason",
]


def norm(text: str) -> str:
    text = (text or "").lower().replace("ё", "е")
    text = re.sub(r"[^a-zа-я0-9]+", " ", text, flags=re.I)
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


def strip_comment_token(comment: str, token: str) -> str:
    c = (comment or "").strip()
    if not c:
        return c
    parts = [p.strip() for p in c.split(";") if p.strip()]
    parts = [p for p in parts if p != token]
    return "; ".join(parts)


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


def recalc_qa(rows: list[dict], target_rows: list[dict]) -> tuple[list[dict], dict]:
    grouped: dict[tuple[str, str], list[dict]] = defaultdict(list)
    for row in rows:
        grouped[(norm(row.get("producer", "")), norm(row.get("target_line", "")))].append(row)
    qa_rows = []
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
        nr = sum(1 for r in sub if r.get("status", "") == "needs_review")
        uniq = len({norm(r.get("model_name", "")) for r in sub if r.get("model_name", "")})
        placeholders = [r for r in sub if "collection_not_found_on_site" in (r.get("comment") or "").lower()]
        real = [r for r in sub if r not in placeholders]
        if not sub:
            st = "not_found"
            cm = "checked_catalog_no_rows_for_target_line"
            lnf += 1
        elif real:
            lr += 1
            st = "needs_review" if (miss or nr) else "found"
            cm = f"with_image={with_img}; missing_image={miss}; needs_review={nr}"
        else:
            lp += 1
            st = "needs_review"
            cm = "collection_not_found_on_site"
        qa_rows.append(
            {
                "producer": p,
                "target_line": l,
                "expected_material": (t.get("covering_from_source") or t.get("material_manual") or "").strip(),
                "rows_count": len(sub),
                "unique_models_count": uniq,
                "with_image_count": with_img,
                "missing_image_count": miss,
                "needs_review_count": nr,
                "bad_image_count": 0,
                "status": st,
                "comment": cm,
            }
        )
    coverage = {
        "target_lines_total": lt,
        "target_lines_with_real_models": lr,
        "target_lines_as_placeholder": lp,
        "target_lines_not_found": lnf,
    }
    return qa_rows, coverage


def repeated_groups(rows: list[dict], producer: str) -> dict[tuple[str, str], list[int]]:
    idxs = defaultdict(list)
    for i, r in enumerate(rows):
        if r.get("producer") != producer:
            continue
        img = (r.get("image_url") or "").strip()
        if not img:
            continue
        key = (r.get("target_line", ""), img)
        idxs[key].append(i)
    return {k: v for k, v in idxs.items() if len(v) > 1}


def is_photo_n(name: str) -> bool:
    return bool(re.search(r"(?i)\s[—-]\s*фото\s*\d+\s*$", name or ""))


def is_bad_olymp_name(name: str) -> bool:
    n = (name or "").strip().lower()
    if not n:
        return False
    if "ghizzi benatti rio z" in n:
        return True
    if "italyanskaya dver ra" in n:
        return True
    # generic latin fragments without clear model semantics
    return bool(re.fullmatch(r"[a-z\s]{8,}", n) and "dior" not in n and "turin" not in n and "malta" not in n)


def is_fallback_olymp_url(url: str) -> bool:
    fn = Path(urlparse(url).path).name.lower()
    return bool(re.fullmatch(r"_+[\d_]*\.png(\.webp)?", fn))


def is_hausdoors_finish_row(model_name: str) -> bool:
    n = (model_name or "").strip().lower().replace("ё", "е")
    if not n:
        return False
    if re.match(r"^(керамик|дуб|шпон|шеллак|кромка)\b", n):
        return True
    if re.search(r"\bral\b|\bncs\b", n):
        return True
    tokens = (
        "матовый",
        "глянец",
        "цвет",
        "покрытие",
        "эмаль",
        "лак",
        "орех",
        "венге",
        "ясень",
        "бриз",
        "грей",
        "гриджио",
        "какао",
        "мока",
        "нордик",
        "пуро",
        "пьетра",
        "табакко",
        "тортора",
    )
    return any(t in n for t in tokens)


def make_group_html(title: str, groups: list[dict]) -> str:
    css = """
    <style>
      body{font-family:Arial,sans-serif;margin:18px}
      .group{border:1px solid #ddd;margin:10px 0;padding:10px}
      .grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(320px,1fr));gap:8px}
      .card{border:1px solid #eee;padding:6px}
      .thumb{height:220px;display:flex;align-items:center;justify-content:center;background:#fafafa;border:1px solid #eee}
      img{max-height:210px;max-width:100%;object-fit:contain}
      .line{font-size:12px;line-height:1.35;margin-top:4px;word-break:break-word}
    </style>
    """
    out = [f"<!doctype html><html><head><meta charset='utf-8'><title>{escape(title)}</title>{css}</head><body><h1>{escape(title)}</h1>"]
    for g in groups:
        out.append("<div class='group'>")
        out.append(f"<div><b>target_line:</b> {escape(g['target_line'])}</div>")
        out.append(f"<div><b>image_url:</b> <a href='{escape(g['image_url'])}'>{escape(g['image_url'])}</a></div>")
        out.append(f"<div><b>repeated_count:</b> {g['repeated_count']} | <b>action:</b> {escape(g['action'])} | <b>reason:</b> {escape(g['reason'])}</div>")
        out.append("<div class='grid'>")
        for r in g["rows"]:
            img = r.get("image_url", "")
            thumb = f"<img src='{escape(img)}'>" if img else "<div>no image</div>"
            out.append(
                "<div class='card'>"
                f"<div class='thumb'>{thumb}</div>"
                f"<div class='line'><b>model:</b> {escape(r.get('model_name',''))}</div>"
                f"<div class='line'><b>status:</b> {escape(r.get('status',''))}</div>"
                f"<div class='line'><b>comment:</b> {escape(r.get('comment',''))}</div>"
                "</div>"
            )
        out.append("</div></div>")
    out.append("</body></html>")
    return "".join(out)


def main() -> None:
    if OUT_DIR.exists():
        shutil.rmtree(OUT_DIR)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    CHECK_DIR.mkdir(parents=True, exist_ok=True)
    REVIEW_DIR.mkdir(parents=True, exist_ok=True)

    base_rows = [{k: r.get(k, "") for k in CATALOG_FIELDS} for r in read_csv(BASE_CSV)]
    targets = read_csv(TARGETS_CSV)

    rows = [dict(r) for r in base_rows]
    repeated_review: list[dict] = []
    html_groups = {"Хаусдорс": [], "Энтро": [], "Олимп": []}
    finish_rows_deleted = 0
    repeated_groups_collapsed = 0
    entro_aria_mode = "unchanged"

    # 1) Hausdoors: remove finish/material rows.
    filtered = []
    removed_finish = []
    for r in rows:
        if r.get("producer") == "Хаусдорс" and is_hausdoors_finish_row(r.get("model_name", "")):
            removed_finish.append(r)
            continue
        filtered.append(r)
    rows = filtered
    finish_rows_deleted = len(removed_finish)
    if removed_finish:
        repeated_review.append(
            {
                "producer": "Хаусдорс",
                "target_line": "*",
                "image_url": "",
                "repeated_count": len(removed_finish),
                "affected_model_names": " | ".join(r.get("model_name", "") for r in removed_finish[:120]),
                "action": "deleted_bad_rows",
                "reason": "finish_material_rows_removed",
            }
        )

    # 2) Hausdoors repeated groups -> collapse to one collection row.
    groups = repeated_groups(rows, "Хаусдорс")
    drop_indices = set()
    add_rows = []
    for (line, img), idxs in groups.items():
        repeated_groups_collapsed += 1
        affected = [rows[i].get("model_name", "") for i in idxs]
        keeper = rows[idxs[0]]
        new_row = dict(keeper)
        new_row["model_name"] = line or keeper.get("model_name", "")
        if img:
            new_row["image_url"] = img
            new_row["status"] = "needs_review"
            new_row["comment"] = "collection_image_used_no_individual_model_photo_collapsed"
        else:
            new_row["image_url"] = ""
            new_row["status"] = "needs_review"
            new_row["comment"] = "no_individual_full_door_photo_found"
        add_rows.append(new_row)
        for i in idxs:
            drop_indices.add(i)
        repeated_review.append(
            {
                "producer": "Хаусдорс",
                "target_line": line,
                "image_url": img,
                "repeated_count": len(idxs),
                "affected_model_names": " | ".join(affected),
                "action": "collapsed_to_one_row",
                "reason": "collection_image_used_no_individual_model_photo_collapsed",
            }
        )
        html_groups["Хаусдорс"].append(
            {
                "target_line": line,
                "image_url": img,
                "repeated_count": len(idxs),
                "action": "collapsed_to_one_row",
                "reason": "collection_image_used_no_individual_model_photo_collapsed",
                "rows": [rows[i] for i in idxs],
            }
        )
    if drop_indices:
        rows = [r for i, r in enumerate(rows) if i not in drop_indices]
        rows.extend(add_rows)

    # 3) Entro Aria collapse if repeated image group remains.
    groups = repeated_groups(rows, "Энтро")
    aria_groups = []
    for (line, img), idxs in groups.items():
        names = " ".join(rows[i].get("model_name", "").lower() for i in idxs)
        srcs = " ".join(rows[i].get("source_url", "").lower() for i in idxs)
        if "aria" in line.lower() or "aria" in names or "/aria" in srcs:
            aria_groups.append(((line, img), idxs))
    if aria_groups:
        entro_aria_mode = "collapsed"
        drop = set()
        adds = []
        for (line, img), idxs in aria_groups:
            repeated_groups_collapsed += 1
            affected = [rows[i].get("model_name", "") for i in idxs]
            k = rows[idxs[0]]
            nr = dict(k)
            nr["model_name"] = "Aria"
            nr["target_line"] = line
            nr["image_url"] = img
            nr["status"] = "needs_review"
            nr["comment"] = "collection_image_used_no_individual_model_photo_collapsed"
            adds.append(nr)
            for i in idxs:
                drop.add(i)
            repeated_review.append(
                {
                    "producer": "Энтро",
                    "target_line": line,
                    "image_url": img,
                    "repeated_count": len(idxs),
                    "affected_model_names": " | ".join(affected),
                    "action": "collapsed_to_one_row",
                    "reason": "collection_image_used_no_individual_model_photo_collapsed",
                }
            )
            html_groups["Энтро"].append(
                {
                    "target_line": line,
                    "image_url": img,
                    "repeated_count": len(idxs),
                    "action": "collapsed_to_one_row",
                    "reason": "collection_image_used_no_individual_model_photo_collapsed",
                    "rows": [rows[i] for i in idxs],
                }
            )
        rows = [r for i, r in enumerate(rows) if i not in drop]
        rows.extend(adds)

    # 4) Olymp cleanup.
    filtered_rows = []
    deleted_bad = []
    for r in rows:
        if r.get("producer") == "Олимп" and is_bad_olymp_name(r.get("model_name", "")):
            deleted_bad.append(r)
            continue
        filtered_rows.append(r)
    rows = filtered_rows
    if deleted_bad:
        repeated_review.append(
            {
                "producer": "Олимп",
                "target_line": "*",
                "image_url": "",
                "repeated_count": len(deleted_bad),
                "affected_model_names": " | ".join(r.get("model_name", "") for r in deleted_bad),
                "action": "deleted_bad_rows",
                "reason": "garbage_model_name_removed",
            }
        )

    # photo N should be needs_review.
    for r in rows:
        if r.get("producer") != "Олимп":
            continue
        if is_photo_n(r.get("model_name", "")):
            r["status"] = "needs_review"
            r["comment"] = strip_comment_token(r.get("comment", ""), "full_door_image_selected")
            if r.get("image_url", ""):
                r["comment"] = append_comment(r.get("comment", ""), "no_visible_caption_full_door_photo")
        if r.get("image_url", "") and is_fallback_olymp_url(r.get("image_url", "")):
            r["status"] = "needs_review"
            r["comment"] = append_comment(r.get("comment", ""), "collection_image_used")

    # Olymp repeated groups: keep one image row, clear others.
    groups = repeated_groups(rows, "Олимп")
    for (line, img), idxs in groups.items():
        keeper = idxs[0]
        for i in idxs:
            if not is_photo_n(rows[i].get("model_name", "")):
                keeper = i
                break
        affected = [rows[i].get("model_name", "") for i in idxs]
        for i in idxs:
            rows[i]["status"] = "needs_review"
            rows[i]["comment"] = strip_comment_token(rows[i].get("comment", ""), "full_door_image_selected")
            if i != keeper:
                rows[i]["image_url"] = ""
                rows[i]["comment"] = append_comment(rows[i].get("comment", ""), "duplicate_photo_row")
        rows[keeper]["comment"] = append_comment(rows[keeper].get("comment", ""), "collection_image_used")
        repeated_review.append(
            {
                "producer": "Олимп",
                "target_line": line,
                "image_url": img,
                "repeated_count": len(idxs),
                "affected_model_names": " | ".join(affected),
                "action": "cleared_duplicate_rows",
                "reason": "repeated_image_url_inside_target_line",
            }
        )
        html_groups["Олимп"].append(
            {
                "target_line": line,
                "image_url": img,
                "repeated_count": len(idxs),
                "action": "cleared_duplicate_rows",
                "reason": "repeated_image_url_inside_target_line",
                "rows": [rows[i] for i in idxs],
            }
        )

    # enforce no repeated found in touched producers
    for producer in TOUCH_PRODUCERS:
        groups = repeated_groups(rows, producer)
        for (_, _), idxs in groups.items():
            for i in idxs:
                if rows[i].get("status", "") == "found":
                    rows[i]["status"] = "needs_review"
                    rows[i]["comment"] = append_comment(rows[i].get("comment", ""), "collection_image_used_no_individual_model_photo")

    # materialize outputs
    rows_sorted = sorted(rows, key=lambda r: (norm(r.get("producer", "")), norm(r.get("target_line", "")), norm(r.get("model_name", ""))))
    write_csv(OUT_DIR / "doors_catalog.csv", rows_sorted, CATALOG_FIELDS)
    write_xlsx(OUT_DIR / "doors_catalog.xlsx", rows_sorted, CATALOG_FIELDS)

    wp_rows = [
        {
            "post_title": r.get("model_name", ""),
            "post_content": r.get("description", ""),
            "producer": r.get("producer", ""),
            "line": r.get("target_line", ""),
            "material": r.get("material", ""),
            "coating": r.get("coating", ""),
            "image_url": r.get("image_url", ""),
            "source_url": r.get("source_url", ""),
            "door_type": r.get("door_type", ""),
            "status": r.get("status", ""),
        }
        for r in rows_sorted
    ]
    write_csv(OUT_DIR / "wordpress_import.csv", wp_rows, WP_FIELDS)

    qa_rows, coverage = recalc_qa(rows_sorted, targets)
    write_csv(OUT_DIR / "qa_report.csv", qa_rows, QA_FIELDS)
    write_csv(OUT_DIR / "repeated_image_review.csv", repeated_review, REPEATED_FIELDS)

    # build image reparse diff
    base_map = {(r.get("producer", ""), r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", "")): r for r in base_rows}
    new_map = {(r.get("producer", ""), r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", "")): r for r in rows_sorted}
    rep_rows = []
    for k in sorted(set(base_map) | set(new_map)):
        b = base_map.get(k)
        n = new_map.get(k)
        if b and not n:
            rep_rows.append(
                {
                    "producer": b.get("producer", ""),
                    "target_line": b.get("target_line", ""),
                    "model_name": b.get("model_name", ""),
                    "source_url": b.get("source_url", ""),
                    "old_image_url": b.get("image_url", ""),
                    "new_image_url": "",
                    "action": "deleted_bad_rows",
                    "reason": "garbage_model_name_removed",
                    "http_status": "",
                    "content_type": "",
                    "width": "",
                    "height": "",
                    "status_after": "",
                    "comment_after": "",
                }
            )
        elif b and n:
            if b.get("image_url", "") != n.get("image_url", "") or b.get("status", "") != n.get("status", "") or b.get("comment", "") != n.get("comment", ""):
                rep_rows.append(
                    {
                        "producer": n.get("producer", ""),
                        "target_line": n.get("target_line", ""),
                        "model_name": n.get("model_name", ""),
                        "source_url": n.get("source_url", ""),
                        "old_image_url": b.get("image_url", ""),
                        "new_image_url": n.get("image_url", ""),
                        "action": "updated_row",
                        "reason": "repeated_image_rules_applied",
                        "http_status": "",
                        "content_type": "",
                        "width": "",
                        "height": "",
                        "status_after": n.get("status", ""),
                        "comment_after": n.get("comment", ""),
                    }
                )
    write_csv(OUT_DIR / "image_reparse_report.csv", rep_rows, REPARSE_FIELDS)

    # review htmls
    (REVIEW_DIR / "hausdoors_repeated_images_review.html").write_text(
        make_group_html("Hausdoors repeated images review", html_groups["Хаусдорс"]), encoding="utf-8"
    )
    (REVIEW_DIR / "entro_repeated_images_review.html").write_text(
        make_group_html("Entro repeated images review", html_groups["Энтро"]), encoding="utf-8"
    )
    (REVIEW_DIR / "olymp_repeated_images_review.html").write_text(
        make_group_html("Olymp repeated images review", html_groups["Олимп"]), encoding="utf-8"
    )

    # check_by_producer
    producer_file = {
        "Скамбио Порте": "scambio",
        "Олимп": "olymp",
        "Хаусдорс": "hausdoors",
        "Лорд": "lord",
        "Аэлита": "aelita",
        "Энтро": "entro",
        "Комфорт": "komfort",
        "Двери регионов": "dveriregionov",
        "Термодур": "termodur",
        "АСД": "asd",
    }
    producers = sorted({r.get("producer", "") for r in rows_sorted if r.get("producer", "")}, key=norm)
    for p in producers:
        sub = [r for r in rows_sorted if r.get("producer") == p]
        stem = producer_file.get(p, norm(p).replace(" ", "_"))
        write_csv(CHECK_DIR / f"{stem}.csv", sub, CATALOG_FIELDS)
        write_xlsx(CHECK_DIR / f"{stem}.xlsx", sub, CATALOG_FIELDS, title=stem)

    # unchanged checks
    base_by = defaultdict(list)
    new_by = defaultdict(list)
    for r in base_rows:
        base_by[r.get("producer", "")].append(r)
    for r in rows_sorted:
        new_by[r.get("producer", "")].append(r)
    unchanged = {}
    for p in UNCHANGED_SIX:
        a = sorted(base_by.get(p, []), key=lambda r: (r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", ""), r.get("image_url", ""), r.get("status", ""), r.get("comment", "")))
        b = sorted(new_by.get(p, []), key=lambda r: (r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", ""), r.get("image_url", ""), r.get("status", ""), r.get("comment", "")))
        unchanged[p] = {"rows_before": len(a), "rows_after": len(b), "exact_match": a == b}

    # ASD not worsened
    asd_before = [r for r in base_rows if r.get("producer") == "АСД"]
    asd_after = [r for r in rows_sorted if r.get("producer") == "АСД"]
    asd_with_before = sum(1 for r in asd_before if r.get("image_url", ""))
    asd_with_after = sum(1 for r in asd_after if r.get("image_url", ""))

    run_report = {
        "generated_at": datetime.now().isoformat(),
        "mode": "final_ready_v12_repeated_image_rules",
        "source_catalog": str(BASE_CSV),
        "total_rows": len(rows_sorted),
        "producers_count": len(producers),
        "rows_by_producer": dict(sorted(Counter(r.get("producer", "") for r in rows_sorted if r.get("producer", "")).items())),
        "rows_with_image": sum(1 for r in rows_sorted if r.get("image_url", "")),
        "rows_without_image": sum(1 for r in rows_sorted if not r.get("image_url", "")),
        "needs_review_count": sum(1 for r in rows_sorted if r.get("status", "") == "needs_review"),
        **coverage,
        "asd_with_image_before": asd_with_before,
        "asd_with_image_after": asd_with_after,
        "asd_not_worsened": asd_with_after >= asd_with_before,
        "unchanged_six_check": unchanged,
        "repeated_groups_reviewed": len(repeated_review),
    }
    (OUT_DIR / "run_report.json").write_text(json.dumps(run_report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"total_rows": len(rows_sorted), "repeated_groups": len(repeated_review)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
