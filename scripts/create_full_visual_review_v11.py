from __future__ import annotations

import csv
from collections import defaultdict
from html import escape
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent
INPUT_CSV = ROOT / "output" / "final_ready_v11" / "doors_catalog.csv"
OUT_DIR = ROOT / "output" / "final_ready_v11" / "full_visual_review"
PRODUCERS = ("АСД", "Энтро", "Хаусдорс", "Олимп")

FIELDS = [
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

DECISION_FIELDS = [
    "producer",
    "target_line",
    "model_name",
    "source_url",
    "current_image_url",
    "decision",
    "selected_image_url",
    "reviewer_comment",
]


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


def card_html(row: dict) -> str:
    img = row.get("image_url", "")
    thumb = (
        f"<a href='{escape(img)}' target='_blank'><img src='{escape(img)}' loading='lazy' alt='thumb'></a>"
        if img
        else "<div class='noimg'>no image</div>"
    )
    return (
        "<div class='card'>"
        f"<div class='thumb'>{thumb}</div>"
        f"<div class='line'><b>{escape(row.get('model_name', ''))}</b></div>"
        f"<div class='line muted'>{escape(row.get('target_line', ''))}</div>"
        f"<div class='line'><b>status:</b> {escape(row.get('status', ''))}</div>"
        f"<div class='line'><b>comment:</b> {escape(row.get('comment', ''))}</div>"
        f"<div class='line'><b>image_url:</b> <a href='{escape(img)}' target='_blank'>{escape(img)}</a></div>"
        f"<div class='line'><b>source_url:</b> <a href='{escape(row.get('source_url', ''))}' target='_blank'>{escape(row.get('source_url', ''))}</a></div>"
        "</div>"
    )


def producer_html(producer: str, rows: list[dict]) -> str:
    with_img = [r for r in rows if r.get("image_url", "")]
    no_img = [r for r in rows if not r.get("image_url", "")]

    def grouped_section(title: str, items: list[dict]) -> str:
        by_line: dict[str, list[dict]] = defaultdict(list)
        for r in items:
            by_line[r.get("target_line", "")].append(r)
        parts = [f"<h2>{escape(title)} ({len(items)})</h2>"]
        for line in sorted(by_line):
            parts.append(f"<h3>{escape(line or '(no target_line)')} ({len(by_line[line])})</h3>")
            parts.append("<div class='grid'>")
            for r in by_line[line]:
                parts.append(card_html(r))
            parts.append("</div>")
        return "".join(parts)

    css = """
    <style>
      body{font-family:Arial,sans-serif;margin:18px}
      h1,h2,h3{margin:10px 0}
      .grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(340px,1fr));gap:10px}
      .card{border:1px solid #ddd;border-radius:6px;padding:8px;background:#fff}
      .thumb{height:230px;display:flex;align-items:center;justify-content:center;border:1px solid #eee;background:#fafafa;margin-bottom:8px}
      img{max-height:220px;max-width:100%;object-fit:contain}
      .noimg{color:#888}
      .line{font-size:12px;line-height:1.35;word-break:break-word;margin-top:4px}
      .muted{color:#666}
    </style>
    """
    return (
        "<!doctype html><html><head><meta charset='utf-8'>"
        f"<title>{escape(producer)} all current images</title>{css}</head><body>"
        f"<h1>{escape(producer)}: current selected images</h1>"
        f"{grouped_section('Rows with image_url', with_img)}"
        f"{grouped_section('Rows without image_url', no_img)}"
        "</body></html>"
    )


def all_sites_html(rows_by_producer: dict[str, list[dict]]) -> str:
    css = """
    <style>
      body{font-family:Arial,sans-serif;margin:18px}
      table{border-collapse:collapse;width:100%}
      th,td{border:1px solid #ddd;padding:6px;text-align:left}
      th{background:#f4f6f8}
    </style>
    """
    parts = [
        "<!doctype html><html><head><meta charset='utf-8'><title>All problem sites current images</title>",
        css,
        "</head><body><h1>All Problem Sites: current selected images summary</h1>",
        "<table><tr><th>producer</th><th>rows_total</th><th>with_image</th><th>without_image</th></tr>",
    ]
    for p in PRODUCERS:
        rows = rows_by_producer.get(p, [])
        with_img = sum(1 for r in rows if r.get("image_url", ""))
        parts.append(f"<tr><td>{escape(p)}</td><td>{len(rows)}</td><td>{with_img}</td><td>{len(rows)-with_img}</td></tr>")
    parts.append("</table><p>Open per-producer HTML files for full visual review.</p></body></html>")
    return "".join(parts)


def main() -> None:
    rows = [r for r in read_csv(INPUT_CSV) if r.get("producer") in PRODUCERS]
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    rows_by_producer: dict[str, list[dict]] = {p: [] for p in PRODUCERS}
    for r in rows:
        rows_by_producer[r.get("producer", "")].append(r)

    filename_by_producer = {
        "АСД": "asd_all_current_images.html",
        "Энтро": "entro_all_current_images.html",
        "Хаусдорс": "hausdoors_all_current_images.html",
        "Олимп": "olymp_all_current_images.html",
    }
    for producer in PRODUCERS:
        html = producer_html(producer, rows_by_producer.get(producer, []))
        (OUT_DIR / filename_by_producer[producer]).write_text(html, encoding="utf-8")

    (OUT_DIR / "all_problem_sites_current_images.html").write_text(all_sites_html(rows_by_producer), encoding="utf-8")

    decisions = [
        {
            "producer": r.get("producer", ""),
            "target_line": r.get("target_line", ""),
            "model_name": r.get("model_name", ""),
            "source_url": r.get("source_url", ""),
            "current_image_url": r.get("image_url", ""),
            "decision": "",
            "selected_image_url": "",
            "reviewer_comment": "",
        }
        for r in rows
    ]
    write_csv(OUT_DIR / "manual_full_door_decisions.csv", decisions, DECISION_FIELDS)

    print("full_visual_review generated")


if __name__ == "__main__":
    main()

