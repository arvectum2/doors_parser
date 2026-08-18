from __future__ import annotations

import argparse
import csv
import json
import re
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from rapidfuzz import fuzz


TARGET_PRODUCERS = {"Олимп", "АСД", "Энтро", "Хаусдорс"}

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

PHOTO_SUMMARY_FIELDS = [
    "producer",
    "rows_total",
    "rows_with_image_before",
    "rows_with_image_after",
    "replaced_images_count",
    "removed_bad_images_count",
    "image_not_found_count",
    "suspicious_images_remaining_count",
    "comment",
]


def normalize(text: str) -> str:
    text = (text or "").lower().replace("ё", "е")
    text = re.sub(r"[^a-zа-я0-9]+", " ", text, flags=re.I)
    return re.sub(r"\s+", " ", text).strip()


def normalize_model_key(text: str) -> str:
    t = normalize(text)
    t = re.sub(r"\b(дверь|door|model|модель|вариант|variant)\b", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def normalize_entro_model_label(model_name: str) -> str:
    # "Aria - variant 1" / "Aria — вариант 1" / "Aria 1" -> "aria 1"
    t = normalize(model_name)
    t = t.replace("variant", " ").replace("вариант", " ")
    t = re.sub(r"\s+", " ", t).strip()
    m = re.search(r"^([a-zа-я0-9 ]+?)\s+(\d{1,2})$", t)
    if m:
        return f"{m.group(1).strip()} {m.group(2).strip()}"
    return t


def append_comment(comment: str, addon: str) -> str:
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
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)


def bad_image_url(url: str) -> bool:
    low = (url or "").lower()
    if not low:
        return True
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
        "/resize/20x/",
        "resize/20x",
        "/cover/240x180/",
        "preview",
        "nalichnik",
        "наличник",
        "plintus",
        "плинтус",
        "korob",
        "короб",
        "dobor",
        "добор",
        "profil",
        "профил",
        "pritvorn",
        "притвор",
        "kapitel",
        "капител",
        "karniz",
        "карниз",
        "rozet",
        "розет",
        "podpyat",
        "подпят",
        "ral-",
        "ncs",
        "pantone",
    )
    if any(b in low for b in banned):
        return True
    # tiny service filenames like tl.png / tg.png
    base = low.rsplit("/", 1)[-1]
    if re.match(r"^[a-z]{1,3}\.(png|jpg|jpeg|webp)$", base):
        return True
    return False


def tilda_restore(url: str) -> str:
    u = url or ""
    u = re.sub(r"/-/resizeb?/\d+x/", "/", u, flags=re.I)
    u = re.sub(r"/-/cover/\d+x\d+/", "/", u, flags=re.I)
    if "thb.tildacdn.com" in u:
        u = u.replace("thb.tildacdn.com", "static.tildacdn.com")
    return u


def parse_int(x: str) -> int:
    try:
        return int(float(x or 0))
    except Exception:
        return 0


def recalc_qa(rows: list[dict], target_rows: list[dict]) -> tuple[list[dict], list[dict], dict]:
    grouped: dict[tuple[str, str], list[dict]] = defaultdict(list)
    for row in rows:
        grouped[(normalize(row.get("producer", "")), normalize(row.get("target_line", "")))].append(row)

    qa_rows: list[dict] = []
    not_found_rows: list[dict] = []
    lines_total = 0
    lines_real = 0
    lines_placeholder = 0
    lines_not_found = 0

    for t in target_rows:
        producer = (t.get("manufacturer") or "").strip()
        line = (t.get("target_line") or "").strip()
        if not producer or not line:
            continue
        lines_total += 1
        key = (normalize(producer), normalize(line))
        sub = grouped.get(key, [])
        if not sub and normalize(producer) == normalize("АСД") and normalize(line) == normalize("Все модели"):
            sub = grouped.get((normalize("АСД"), normalize("Уточнить коллекцию")), [])

        with_image = sum(1 for r in sub if r.get("image_url", "").strip())
        missing = len(sub) - with_image
        needs_review = sum(1 for r in sub if (r.get("status") or "").strip() == "needs_review")
        bad_img = sum(1 for r in sub if bad_image_url(r.get("image_url", "")))
        unique_models = len({normalize_model_key(r.get("model_name", "")) for r in sub if r.get("model_name", "").strip()})

        placeholders = [r for r in sub if "collection_not_found_on_site" in (r.get("comment") or "").lower()]
        real_rows = [r for r in sub if r not in placeholders]
        if not sub:
            status = "not_found"
            comment = "checked_catalog_no_rows_for_target_line"
            lines_not_found += 1
            not_found_rows.append(
                {
                    "door_type": t.get("door_type", ""),
                    "manufacturer": producer,
                    "target_line": line,
                    "source_site": t.get("source_site", ""),
                    "status": "not_found",
                    "comment": comment,
                }
            )
        elif real_rows:
            lines_real += 1
            status = "needs_review" if (missing > 0 or needs_review > 0 or bad_img > 0) else "found"
            comment = f"with_image={with_image}; missing_image={missing}; needs_review={needs_review}; bad_image={bad_img}"
        else:
            lines_placeholder += 1
            status = "needs_review"
            comment = "collection_not_found_on_site"

        qa_rows.append(
            {
                "producer": producer,
                "target_line": line,
                "expected_material": (t.get("covering_from_source") or t.get("material_manual") or "").strip(),
                "rows_count": len(sub),
                "unique_models_count": unique_models,
                "with_image_count": with_image,
                "missing_image_count": missing,
                "needs_review_count": needs_review,
                "bad_image_count": bad_img,
                "status": status,
                "comment": comment,
            }
        )

    coverage = {
        "target_lines_total": lines_total,
        "target_lines_with_real_models": lines_real,
        "target_lines_as_placeholder": lines_placeholder,
        "target_lines_not_found": lines_not_found,
    }
    return qa_rows, not_found_rows, coverage


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, help="Base doors_catalog.csv")
    parser.add_argument("--candidates-dir", required=True, help="output/image_candidates_v2")
    parser.add_argument("--manual-selection", default="", help="manual_selection_template.csv")
    parser.add_argument("--out", required=True, help="output/final_photo_fix_manual_rules_v2")
    parser.add_argument("--targets", default="data/target_lines.csv", help="target_lines.csv for qa_report")
    args = parser.parse_args()

    base_path = Path(args.input)
    cand_dir = Path(args.candidates_dir)
    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    check_dir = out_dir / "check_by_producer"
    check_dir.mkdir(parents=True, exist_ok=True)

    rows = [{k: (r.get(k, "") or "").strip() for k in CATALOG_FIELDS} for r in read_csv(base_path)]
    before = [{k: r.get(k, "") for k in CATALOG_FIELDS} for r in rows]

    # Load candidates grouped by model key.
    by_key: dict[tuple[str, str, str, str], list[dict]] = defaultdict(list)
    producer_files = [
        cand_dir / "entro_candidates.csv",
        cand_dir / "hausdoors_candidates.csv",
        cand_dir / "asd_candidates.csv",
        cand_dir / "olymp_candidates.csv",
    ]
    for fp in producer_files:
        if not fp.exists():
            continue
        for r in read_csv(fp):
            key = (r.get("producer", ""), r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", ""))
            by_key[key].append(r)

    # manual overrides
    manual_map: dict[tuple[str, str, str, str], tuple[str, str]] = {}
    if args.manual_selection:
        mpath = Path(args.manual_selection)
        if mpath.exists():
            for r in read_csv(mpath):
                key = (r.get("producer", ""), r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", ""))
                sel = r.get("selected_image_url", "")
                dec = (r.get("decision", "") or "").lower()
                cmt = r.get("comment", "")
                if sel:
                    manual_map[key] = (sel, cmt or "manual_selected")
                elif dec in {"reject", "needs_review", "clear"}:
                    manual_map[key] = ("", cmt or "manual_rejected")

    # old image counts for fallback analysis
    old_img_counts: dict[str, Counter] = defaultdict(Counter)
    for r in rows:
        p = r.get("producer", "")
        u = r.get("image_url", "").strip()
        if p and u:
            old_img_counts[p][u] += 1

    reparse_rows: list[dict] = []
    changed = Counter()
    replaced = Counter()
    cleared = Counter()
    entro_exact = 0
    entro_fuzzy = 0

    for row in rows:
        producer = row.get("producer", "")
        if producer not in TARGET_PRODUCERS:
            continue
        key = (producer, row.get("target_line", ""), row.get("model_name", ""), row.get("source_url", ""))
        cands = by_key.get(key, [])
        old_img = row.get("image_url", "")
        old_status = row.get("status", "")
        old_comment = row.get("comment", "")
        new_img = old_img
        new_status = old_status or "found"
        new_comment = old_comment
        action = "kept_old_image"
        reason = "kept_old_image"

        # manual override first
        if key in manual_map:
            m_img, m_reason = manual_map[key]
            if m_img:
                new_img = m_img
                new_status = "found"
                new_comment = append_comment(new_comment, f"manual_selected:{m_reason}")
                reason = "manual_selected"
                if new_img != old_img:
                    action = "replaced_image"
            else:
                new_img = ""
                new_status = "needs_review"
                new_comment = append_comment(new_comment, "manual_photo_check_required")
                reason = "manual_rejected"
                if old_img:
                    action = "removed_bad_image"

        else:
            if producer == "Энтро":
                expected = normalize_entro_model_label(row.get("model_name", ""))
                best_exact = None
                best_fuzzy = None
                best_fuzzy_score = 0.0
                for c in cands:
                    img = c.get("candidate_image_url", "")
                    if not img:
                        continue
                    # for Entro candidate_text dominates; block can be other
                    candidate_text = normalize(c.get("candidate_text", ""))
                    if not candidate_text:
                        candidate_text = normalize(c.get("nearby_text_after", "") or c.get("nearby_text_before", ""))
                    if not candidate_text:
                        continue
                    if bad_image_url(img):
                        continue
                    if expected and candidate_text == expected:
                        best_exact = c
                        break
                    # fuzzy: startswith or high token score
                    if expected and (candidate_text.startswith(expected) or expected.startswith(candidate_text)):
                        score = 95.0
                    else:
                        score = fuzz.token_set_ratio(expected, candidate_text) if expected else 0
                    if score > best_fuzzy_score:
                        best_fuzzy_score = score
                        best_fuzzy = c

                if best_exact is not None:
                    new_img = best_exact.get("candidate_image_url", "")
                    new_status = "found"
                    new_comment = append_comment(new_comment, "model_row_image_used")
                    reason = "candidate_text_exact_match"
                    entro_exact += 1
                    if new_img != old_img:
                        action = "replaced_image"
                elif best_fuzzy is not None and best_fuzzy_score >= 88:
                    new_img = best_fuzzy.get("candidate_image_url", "")
                    new_status = "found"
                    new_comment = append_comment(new_comment, "model_row_image_used_fuzzy")
                    reason = "candidate_text_fuzzy_match"
                    entro_fuzzy += 1
                    if new_img != old_img:
                        action = "replaced_image"
                else:
                    new_img = ""
                    new_status = "needs_review"
                    new_comment = append_comment(new_comment, "no_model_row_card_found")
                    reason = "no_model_row_card_found"
                    if old_img:
                        action = "removed_bad_image"

            elif producer == "Олимп":
                # choose best non-resize candidate. if none -> keep old if safe; else clear.
                chosen = None
                for c in cands:
                    img = c.get("candidate_image_url", "")
                    if not img:
                        continue
                    l = img.lower()
                    if "/resize/20x/" in l or "resize/20x" in l or "/cover/240x180/" in l or "preview" in l or "thumb" in l:
                        continue
                    if bad_image_url(img):
                        continue
                    chosen = c
                    break
                if chosen:
                    new_img = tilda_restore(chosen.get("candidate_image_url", ""))
                    if bad_image_url(new_img):
                        new_img = ""
                    if new_img:
                        new_status = "found"
                        new_comment = append_comment(new_comment, "product_gallery_image")
                        reason = "product_gallery_image"
                        if new_img != old_img:
                            action = "replaced_image"
                    else:
                        chosen = None

                if chosen is None:
                    old_fixed = tilda_restore(old_img) if old_img else ""
                    if old_fixed and not bad_image_url(old_fixed):
                        new_img = old_fixed
                        new_status = "found"
                        new_comment = append_comment(new_comment, "kept_existing_nonpromo_image")
                        reason = "kept_existing_nonpromo_image"
                        if new_img != old_img:
                            action = "replaced_image"
                    else:
                        new_img = ""
                        new_status = "needs_review"
                        new_comment = append_comment(new_comment, "no_model_image_found")
                        reason = "no_model_image_found"
                        if old_img:
                            action = "removed_bad_image"

            elif producer == "АСД":
                # prefer selected product_gallery_image, else keep safe old, else clear
                chosen = None
                for c in cands:
                    img = c.get("candidate_image_url", "")
                    if not img:
                        continue
                    reason_c = c.get("reason", "")
                    if bad_image_url(img):
                        continue
                    if reason_c in {"product_gallery_image", "model_row_exact_match", "model_row_fuzzy_match"} or c.get("parser_decision") == "selected":
                        chosen = c
                        break
                if chosen:
                    new_img = chosen.get("candidate_image_url", "")
                    new_status = "found"
                    new_comment = append_comment(new_comment, "product_gallery_image")
                    reason = "product_gallery_image"
                    if new_img != old_img:
                        action = "replaced_image"
                elif old_img and not bad_image_url(old_img):
                    new_img = old_img
                    new_status = old_status or "found"
                    new_comment = append_comment(new_comment, "kept_existing_image")
                    reason = "kept_existing_image"
                else:
                    new_img = ""
                    new_status = "needs_review"
                    new_comment = append_comment(new_comment, "no_model_image_found")
                    reason = "no_model_image_found"
                    if old_img:
                        action = "removed_bad_image"

            elif producer == "Хаусдорс":
                # choose first selected candidate that is not bad and not /empty/
                chosen = None
                for c in cands:
                    if c.get("parser_decision") != "selected":
                        continue
                    img = c.get("candidate_image_url", "")
                    if not img or bad_image_url(img):
                        continue
                    chosen = c
                    break
                if chosen:
                    new_img = tilda_restore(chosen.get("candidate_image_url", ""))
                    if bad_image_url(new_img):
                        new_img = ""
                else:
                    new_img = old_img if old_img and not bad_image_url(old_img) else ""

                if new_img:
                    new_status = "found"
                    new_comment = append_comment(new_comment, "image_verified")
                    reason = "image_verified"
                    if new_img != old_img:
                        action = "replaced_image"
                else:
                    new_status = "needs_review"
                    new_comment = append_comment(new_comment, "no_model_image_found")
                    reason = "no_model_image_found"
                    if old_img:
                        action = "removed_bad_image"

        row["image_url"] = new_img
        row["status"] = new_status
        row["comment"] = new_comment

        if new_img != old_img or new_status != old_status or new_comment != old_comment:
            changed[producer] += 1
        if old_img and new_img and old_img != new_img:
            replaced[producer] += 1
        if old_img and not new_img:
            cleared[producer] += 1

        reparse_rows.append(
            {
                "producer": producer,
                "target_line": row.get("target_line", ""),
                "model_name": row.get("model_name", ""),
                "source_url": row.get("source_url", ""),
                "old_image_url": old_img,
                "new_image_url": new_img,
                "action": action,
                "reason": reason,
                "http_status": "",
                "content_type": "",
                "width": "",
                "height": "",
                "status_after": new_status,
                "comment_after": new_comment,
            }
        )

    # Hausdoors repeated selected handling after first pass
    haus_rows = [r for r in rows if r.get("producer") == "Хаусдорс"]
    haus_count = Counter(r.get("image_url", "") for r in haus_rows if r.get("image_url", ""))
    for r in haus_rows:
        img = r.get("image_url", "")
        if not img:
            continue
        rep = haus_count.get(img, 0)
        if rep > 10:
            r["comment"] = append_comment(r.get("comment", ""), "repeated_collection_image")
            r["comment"] = append_comment(r.get("comment", ""), "collection_image_used")
            r["status"] = "needs_review"

    # sort and write core outputs
    rows_sorted = sorted(rows, key=lambda r: (normalize(r.get("producer", "")), normalize(r.get("target_line", "")), normalize(r.get("model_name", ""))))
    write_csv(out_dir / "doors_catalog.csv", rows_sorted, CATALOG_FIELDS)
    write_xlsx(out_dir / "doors_catalog.xlsx", rows_sorted, CATALOG_FIELDS)

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
    write_csv(out_dir / "wordpress_import.csv", wp_rows, WP_FIELDS)
    write_csv(out_dir / "image_reparse_report.csv", reparse_rows, REPARSE_FIELDS)

    # QA + not_found
    targets = read_csv(Path(args.targets))
    qa_rows, not_found_rows, coverage = recalc_qa(rows_sorted, targets)
    write_csv(out_dir / "qa_report.csv", qa_rows, QA_FIELDS)
    write_csv(out_dir / "not_found.csv", not_found_rows, NOT_FOUND_FIELDS)

    # check_by_producer outputs
    producer_file = {"Олимп": "olymp", "АСД": "asd", "Энтро": "entro", "Хаусдорс": "hausdoors"}
    by_prod: dict[str, list[dict]] = defaultdict(list)
    for r in rows_sorted:
        if r.get("producer", "") in producer_file:
            by_prod[r["producer"]].append(r)
    summary_rows = []
    for p in ("Олимп", "АСД", "Энтро", "Хаусдорс"):
        sub = by_prod.get(p, [])
        write_csv(check_dir / f"{producer_file[p]}.csv", sub, CATALOG_FIELDS)
        write_xlsx(check_dir / f"{producer_file[p]}.xlsx", sub, CATALOG_FIELDS, title=producer_file[p])
        suspicious = 0
        if p == "Олимп":
            suspicious = sum(1 for r in sub if "/resize/20x/" in (r.get("image_url") or "").lower() or "resize/20x" in (r.get("image_url") or "").lower())
        if p == "Хаусдорс":
            suspicious = sum(1 for r in sub if "/empty/" in (r.get("image_url") or "").lower())
        if p == "АСД":
            suspicious = sum(1 for r in sub if any(k in (r.get("image_url") or "").lower() for k in ("hit", "badge", "sticker", "label")))
        summary_rows.append(
            {
                "producer": p,
                "rows_total": len(sub),
                "rows_with_image_before": sum(1 for r in before if r.get("producer") == p and r.get("image_url", "")),
                "rows_with_image_after": sum(1 for r in sub if r.get("image_url", "")),
                "replaced_images_count": replaced.get(p, 0),
                "removed_bad_images_count": cleared.get(p, 0),
                "image_not_found_count": sum(1 for r in sub if not r.get("image_url", "")),
                "suspicious_images_remaining_count": suspicious,
                "comment": "ok" if suspicious == 0 else "needs_manual_visual_check",
            }
        )
    write_csv(check_dir / "photo_fix_audit_summary.csv", summary_rows, PHOTO_SUMMARY_FIELDS)

    # run report
    producer_counts = Counter(r.get("producer", "") for r in rows_sorted if r.get("producer", ""))
    entro_needs_review = sum(1 for r in rows_sorted if r.get("producer") == "Энтро" and r.get("status") == "needs_review")
    olymp_resize_left = sum(
        1
        for r in rows_sorted
        if r.get("producer") == "Олимп"
        and ("/resize/20x/" in (r.get("image_url") or "").lower() or "resize/20x" in (r.get("image_url") or "").lower())
    )
    haus_repeat_top = sorted(((k, v) for k, v in haus_count.items()), key=lambda x: x[1], reverse=True)[:20]
    run_report = {
        "generated_at": datetime.now().isoformat(),
        "mode": "apply_image_selection_v2",
        "input_catalog": str(base_path),
        "candidates_dir": str(cand_dir),
        "total_rows": len(rows_sorted),
        "producers_count": len(producer_counts),
        **coverage,
        "rows_with_image": sum(1 for r in rows_sorted if r.get("image_url", "")),
        "rows_without_image": sum(1 for r in rows_sorted if not r.get("image_url", "")),
        "needs_review_count": sum(1 for r in rows_sorted if r.get("status") == "needs_review"),
        "rows_by_producer": dict(sorted(producer_counts.items())),
        "changed_rows_by_producer": dict(changed),
        "replaced_images_by_producer": dict(replaced),
        "cleared_images_by_producer": dict(cleared),
        "entro_candidate_text_exact_match_selected": entro_exact,
        "entro_candidate_text_fuzzy_selected": entro_fuzzy,
        "entro_needs_review_count": entro_needs_review,
        "olymp_resize_20x_remaining": olymp_resize_left,
        "hausdoors_top_repeated_images": [{"image_url": u, "count": c} for u, c in haus_repeat_top],
    }
    (out_dir / "run_report.json").write_text(json.dumps(run_report, ensure_ascii=False, indent=2), encoding="utf-8")

    print("apply_image_selection v2 completed")
    print(json.dumps(run_report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
