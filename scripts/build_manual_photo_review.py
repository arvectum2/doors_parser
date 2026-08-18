from __future__ import annotations

import csv
import html
import re
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


ROOT = Path(__file__).resolve().parent.parent
CATALOG_PATH = ROOT / "output" / "final_photo_fix_manual_rules_v2" / "doors_catalog.csv"
CANDIDATES_DIR = ROOT / "output" / "image_candidates_v2"
OUT_DIR = ROOT / "output" / "manual_photo_review"

PRODUCER_FILES = {
    "Энтро": CANDIDATES_DIR / "entro_candidates.csv",
    "Хаусдорс": CANDIDATES_DIR / "hausdoors_candidates.csv",
    "Олимп": CANDIDATES_DIR / "olymp_candidates.csv",
    "АСД": CANDIDATES_DIR / "asd_candidates.csv",
}

SHEETS = [("Энтро", "Entro"), ("Хаусдорс", "Hausdoors"), ("Олимп", "Olymp"), ("АСД", "ASD")]

FIELDS = [
    "producer",
    "target_line",
    "model_name",
    "source_url",
    "current_image_url",
    "current_status",
    "current_comment",
    "problem_reason",
    "candidate_1_url",
    "candidate_1_text",
    "candidate_1_source_block",
    "candidate_2_url",
    "candidate_2_text",
    "candidate_2_source_block",
    "candidate_3_url",
    "candidate_3_text",
    "candidate_3_source_block",
    "candidate_4_url",
    "candidate_4_text",
    "candidate_4_source_block",
    "candidate_5_url",
    "candidate_5_text",
    "candidate_5_source_block",
    "selected_image_url",
    "decision",
    "reviewer_comment",
]


def normalize(text: str) -> str:
    text = (text or "").lower().replace("ё", "е")
    text = re.sub(r"[^a-zа-я0-9]+", " ", text, flags=re.I)
    return re.sub(r"\s+", " ", text).strip()


def normalize_entro_model_label(model_name: str) -> str:
    t = normalize(model_name)
    t = t.replace("variant", " ").replace("вариант", " ")
    t = re.sub(r"\s+", " ", t).strip()
    m = re.search(r"^([a-zа-я0-9 ]+?)\s+(\d{1,2})$", t)
    if m:
        return f"{m.group(1).strip()} {m.group(2).strip()}"
    return t


def read_csv(path: Path) -> list[dict]:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        return [{k: (v or "").strip() for k, v in row.items()} for row in csv.DictReader(f, delimiter=";")]


def bad_candidate(c: dict) -> bool:
    merged = " ".join(
        [
            c.get("candidate_image_url", ""),
            c.get("candidate_text", ""),
            c.get("candidate_alt", ""),
            c.get("candidate_title", ""),
            c.get("reason", ""),
            c.get("candidate_block", ""),
        ]
    ).lower()
    banned = (
        "hero",
        "badge",
        "logo",
        "icon",
        "texture",
        "fragment",
        "empty",
        "promo",
        "resize/20x",
        "/resize/20x/",
        "cover/240x180",
        "preview",
        "sticker",
        "label",
        "hit",
        "sale",
        "open.png",
        "catalog.svg",
        "search_loop",
    )
    return any(b in merged for b in banned)


def candidate_sort_key(c: dict, expected_label: str = "") -> tuple:
    dec = c.get("parser_decision", "")
    reason = c.get("reason", "")
    block = c.get("candidate_block", "")
    text = normalize(c.get("candidate_text", ""))
    priority_dec = 0 if dec == "selected" else 1
    priority_block = 0 if block in {"model_row", "product_gallery", "model_card"} else 1
    # Entro label match bonus
    label_rank = 2
    if expected_label:
        if text == expected_label:
            label_rank = 0
        elif text.startswith(expected_label) or expected_label.startswith(text):
            label_rank = 1
    reason_rank = 0 if reason in {"model_row_exact_match", "dom_order_model_row_match", "model_row_fuzzy_match", "product_gallery_image"} else 1
    area = int(c.get("candidate_width") or 0) * int(c.get("candidate_height") or 0)
    return (priority_dec, label_rank, priority_block, reason_rank, -area)


def build_problem_reason(row: dict, repeated_count: int) -> str:
    reasons = []
    if not row.get("image_url", ""):
        reasons.append("missing_image")
    if row.get("status") == "needs_review":
        reasons.append("status_needs_review")
    comment = (row.get("comment", "") or "").lower()
    if "no_model_row_card_found" in comment:
        reasons.append("no_model_row_card_found")
    if "repeated_collection_image" in comment:
        reasons.append("repeated_collection_image")
    if repeated_count > 10:
        reasons.append("repeated_image_url_gt_10")
    if not reasons:
        reasons.append("manual_check")
    return "; ".join(dict.fromkeys(reasons))


def write_xlsx(path: Path, sheet_rows: dict[str, list[dict]]) -> None:
    wb = Workbook()
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]
    for producer, title in SHEETS:
        ws = wb.create_sheet(title=title)
        ws.append(FIELDS)
        for row in sheet_rows.get(producer, []):
            ws.append([row.get(k, "") for k in FIELDS])
        widths = [16, 22, 30, 50, 50, 14, 36, 30, 44, 24, 18, 44, 24, 18, 44, 24, 18, 44, 24, 18, 44, 24, 18, 50, 18, 32]
        for col, width in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(col)].width = width
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


def write_html(path: Path, sheet_rows: dict[str, list[dict]]) -> None:
    def img_cell(url: str) -> str:
        if not url:
            return "<div class='na'>—</div>"
        u = html.escape(url, quote=True)
        return f"<a href='{u}' target='_blank'><img src='{u}' loading='lazy' /></a><div class='url'>{u}</div>"

    parts = [
        "<!doctype html><html><head><meta charset='utf-8'><title>manual_photo_review</title>",
        "<style>body{font-family:Arial,sans-serif;font-size:12px}table{border-collapse:collapse;width:100%;margin:20px 0}"
        "th,td{border:1px solid #ddd;padding:6px;vertical-align:top}th{background:#1f4e78;color:#fff;position:sticky;top:0}"
        "img{max-width:140px;max-height:180px;display:block}.url{max-width:160px;word-break:break-all;color:#666;font-size:10px}"
        ".group{margin-top:30px}.na{color:#999}</style></head><body>",
        "<h1>manual_photo_review</h1>",
    ]
    for producer, _title in SHEETS:
        rows = sheet_rows.get(producer, [])
        parts.append(f"<div class='group'><h2>{html.escape(producer)} ({len(rows)})</h2>")
        parts.append("<table><thead><tr><th>target_line</th><th>model_name</th><th>problem_reason</th><th>current</th>"
                     "<th>cand1</th><th>cand2</th><th>cand3</th><th>cand4</th><th>cand5</th></tr></thead><tbody>")
        for r in rows:
            parts.append("<tr>")
            parts.append(f"<td>{html.escape(r.get('target_line',''))}</td>")
            parts.append(f"<td>{html.escape(r.get('model_name',''))}</td>")
            parts.append(f"<td>{html.escape(r.get('problem_reason',''))}</td>")
            parts.append(f"<td>{img_cell(r.get('current_image_url',''))}</td>")
            for i in range(1, 6):
                parts.append(f"<td>{img_cell(r.get(f'candidate_{i}_url',''))}<div>{html.escape(r.get(f'candidate_{i}_text',''))}</div>"
                             f"<div>{html.escape(r.get(f'candidate_{i}_source_block',''))}</div></td>")
            parts.append("</tr>")
        parts.append("</tbody></table></div>")
    parts.append("</body></html>")
    path.write_text("".join(parts), encoding="utf-8")


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    rows = read_csv(CATALOG_PATH)

    # candidate map by model key
    cands_by_key: dict[tuple[str, str, str, str], list[dict]] = defaultdict(list)
    for producer, fp in PRODUCER_FILES.items():
        if not fp.exists():
            continue
        for c in read_csv(fp):
            key = (c.get("producer", ""), c.get("target_line", ""), c.get("model_name", ""), c.get("source_url", ""))
            cands_by_key[key].append(c)

    # repeated current image counts per producer
    img_count_by_producer: dict[str, Counter] = defaultdict(Counter)
    for r in rows:
        p = r.get("producer", "")
        u = r.get("image_url", "")
        if p and u:
            img_count_by_producer[p][u] += 1

    sheet_rows: dict[str, list[dict]] = defaultdict(list)
    stats = Counter()
    entro_no_photo = 0
    haus_repeated = 0
    olymp_repeated = 0

    for r in rows:
        p = r.get("producer", "")
        if p not in PRODUCER_FILES:
            continue
        key = (p, r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", ""))
        repeated_count = img_count_by_producer[p].get(r.get("image_url", ""), 0) if r.get("image_url") else 0
        comment_l = (r.get("comment", "") or "").lower()
        need = False

        if p == "Энтро":
            if (not r.get("image_url")) or r.get("status") == "needs_review" or "no_model_row_card_found" in comment_l or "model_row_image_used" not in comment_l:
                need = True
            if not r.get("image_url"):
                entro_no_photo += 1
        elif p == "Хаусдорс":
            if r.get("status") == "needs_review" or (not r.get("image_url")) or repeated_count > 10 or "repeated_collection_image" in comment_l:
                need = True
            if repeated_count > 10:
                haus_repeated += 1
        elif p == "Олимп":
            if repeated_count > 10 or "collection_image_used" in comment_l or "repeated_collection_image" in comment_l:
                need = True
            if repeated_count > 10:
                olymp_repeated += 1
        elif p == "АСД":
            if repeated_count > 10 or any(x in comment_l for x in ("no_model_image_found", "false_positive", "image_not_found")):
                need = True

        if not need:
            continue

        stats[p] += 1
        cand_rows = cands_by_key.get(key, [])
        expected_label = normalize_entro_model_label(r.get("model_name", "")) if p == "Энтро" else ""
        filtered = [c for c in cand_rows if c.get("candidate_image_url") and not bad_candidate(c)]
        filtered.sort(key=lambda c: candidate_sort_key(c, expected_label))
        top = filtered[:5]

        row_out = {
            "producer": p,
            "target_line": r.get("target_line", ""),
            "model_name": r.get("model_name", ""),
            "source_url": r.get("source_url", ""),
            "current_image_url": r.get("image_url", ""),
            "current_status": r.get("status", ""),
            "current_comment": r.get("comment", ""),
            "problem_reason": build_problem_reason(r, repeated_count),
            "selected_image_url": top[0].get("candidate_image_url", "") if top else "",
            "decision": "",
            "reviewer_comment": "",
        }
        for i in range(1, 6):
            if i <= len(top):
                c = top[i - 1]
                row_out[f"candidate_{i}_url"] = c.get("candidate_image_url", "")
                row_out[f"candidate_{i}_text"] = c.get("candidate_text", "")
                row_out[f"candidate_{i}_source_block"] = c.get("candidate_block", "")
            else:
                row_out[f"candidate_{i}_url"] = ""
                row_out[f"candidate_{i}_text"] = ""
                row_out[f"candidate_{i}_source_block"] = ""

        sheet_rows[p].append(row_out)

    # sort for readability
    for p in list(sheet_rows):
        sheet_rows[p] = sorted(sheet_rows[p], key=lambda x: (normalize(x.get("target_line", "")), normalize(x.get("model_name", ""))))

    xlsx_path = OUT_DIR / "manual_photo_review.xlsx"
    html_path = OUT_DIR / "manual_photo_review.html"
    write_xlsx(xlsx_path, sheet_rows)
    write_html(html_path, sheet_rows)

    print("manual photo review built")
    for p, _title in SHEETS:
        print(f"{p}: {len(sheet_rows.get(p, []))}")
    print(f"entro_without_image: {entro_no_photo}")
    print(f"hausdoors_repeated_gt_10_rows: {haus_repeated}")
    print(f"olymp_repeated_gt_10_rows: {olymp_repeated}")
    print(f"xlsx={xlsx_path}")
    print(f"html={html_path}")


if __name__ == "__main__":
    main()

