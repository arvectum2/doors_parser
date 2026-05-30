#!/usr/bin/env bash
set -e
cd "$(dirname "$0")"
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python src/doors_parser.py --targets data/target_lines.csv --sites config/sites.yaml --out output --max-pages-per-site 3500
