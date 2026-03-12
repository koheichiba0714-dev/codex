#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path
from typing import Any
from urllib.parse import quote_plus


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="大阪市B型のホームページ / Instagram 未登録事業所について、確認用CSVを作る。"
    )
    parser.add_argument(
        "--dashboard-input",
        default="apps/r6kouchin-dashboard/data/dashboard-data.json",
        help="dashboard manifest json path",
    )
    parser.add_argument(
        "--output",
        default="data/exports/r6kouchinjissekib/integrated/osakashi_missing_office_links_review.csv",
        help="review csv output path",
    )
    return parser.parse_args()


def repo_root() -> Path:
    return Path(__file__).resolve().parents[1]


def resolve_path(path_str: str) -> Path:
    path = Path(path_str).expanduser()
    if path.is_absolute():
        return path.resolve()
    return (repo_root() / path).resolve()


def load_records(manifest_path: Path) -> list[dict[str, Any]]:
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    app_root = manifest_path.parent.parent
    records: list[dict[str, Any]] = []
    for rel_path in manifest.get("data_files", {}).get("record_chunks", []):
        records.extend(json.loads((app_root / rel_path).read_text(encoding="utf-8")))
    return records


def build_search_url(parts: list[str]) -> str:
    query = " ".join(part for part in parts if part).strip()
    return f"https://www.google.com/search?q={quote_plus(query)}"


def website_search_url(record: dict[str, Any]) -> str:
    return build_search_url(
        [
            str(record.get("office_name") or ""),
            str(record.get("municipality") or ""),
            str(record.get("corporation_name") or ""),
            "就労継続支援B型",
        ]
    )


def instagram_search_url(record: dict[str, Any]) -> str:
    return build_search_url(
        [
            str(record.get("office_name") or ""),
            str(record.get("municipality") or ""),
            str(record.get("corporation_name") or ""),
            "site:instagram.com",
        ]
    )


def main() -> None:
    args = parse_args()
    dashboard_input = resolve_path(args.dashboard_input)
    output_path = resolve_path(args.output)

    records = [
        record
        for record in load_records(dashboard_input)
        if record.get("municipality") == "大阪市"
        and (not record.get("homepage_url") or not record.get("instagram_url"))
    ]
    records.sort(
        key=lambda row: (
            0 if not row.get("homepage_url") and not row.get("instagram_url") else 1,
            str(row.get("office_no") or ""),
        )
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", newline="", encoding="utf-8") as file_obj:
        writer = csv.DictWriter(
            file_obj,
            fieldnames=[
                "office_no",
                "municipality",
                "corporation_name",
                "office_name",
                "homepage_url",
                "instagram_url",
                "homepage_missing",
                "instagram_missing",
                "website_search_url",
                "instagram_search_url",
            ],
        )
        writer.writeheader()
        for record in records:
            writer.writerow(
                {
                    "office_no": record.get("office_no") or "",
                    "municipality": record.get("municipality") or "",
                    "corporation_name": record.get("corporation_name") or "",
                    "office_name": record.get("office_name") or "",
                    "homepage_url": record.get("homepage_url") or "",
                    "instagram_url": record.get("instagram_url") or "",
                    "homepage_missing": "あり" if not record.get("homepage_url") else "",
                    "instagram_missing": "あり" if not record.get("instagram_url") else "",
                    "website_search_url": website_search_url(record),
                    "instagram_search_url": instagram_search_url(record),
                }
            )

    print(
        json.dumps(
            {
                "record_count": len(records),
                "output": str(output_path),
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
