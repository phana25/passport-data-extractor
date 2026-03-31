from __future__ import annotations

import csv
import json
from pathlib import Path


def export_json(data: dict, path: str) -> None:
    Path(path).write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")


def export_csv(data: dict, path: str) -> None:
    # Single-row CSV with headers from keys
    p = Path(path)
    keys = list(data.keys())
    with p.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=keys)
        w.writeheader()
        w.writerow(data)

