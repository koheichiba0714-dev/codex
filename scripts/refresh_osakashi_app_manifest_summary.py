#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from build_r6kouchinjissekib_dashboard_dataset import build_app_summary


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="大阪市向け app manifest の summary / analytics を実データから再計算する。")
    parser.add_argument(
        "--manifest",
        default="apps/r6kouchin-dashboard/data/dashboard-data.json",
        help="app manifest json path",
    )
    parser.add_argument(
        "--integrated-dashboard",
        default="data/exports/r6kouchinjissekib/integrated/shuro_b_dashboard_integrated.json",
        help="integrated dashboard json path",
    )
    return parser.parse_args()


def repo_root() -> Path:
    return Path(__file__).resolve().parents[1]


def resolve_path(path_str: str) -> Path:
    path = Path(path_str).expanduser()
    if path.is_absolute():
        return path.resolve()
    return (repo_root() / path).resolve()


def load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path: Path, payload: Any) -> None:
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def read_app_records(manifest_path: Path, manifest: dict[str, Any]) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    for relative_path in manifest.get("data_files", {}).get("record_chunks", []) or []:
        chunk_path = (manifest_path.parent.parent / str(relative_path)).resolve()
        records.extend(load_json(chunk_path))
    return records


def office_key(record: dict[str, Any]) -> str:
    return str(record.get("office_no") or "").strip()


def main() -> None:
    args = parse_args()
    manifest_path = resolve_path(args.manifest)
    integrated_path = resolve_path(args.integrated_dashboard)

    manifest = load_json(manifest_path)
    integrated = load_json(integrated_path)
    app_records = read_app_records(manifest_path, manifest)
    app_office_keys = {office_key(record) for record in app_records if office_key(record)}
    integrated_records = integrated.get("records", [])
    scoped_records = [record for record in integrated_records if office_key(record) in app_office_keys]

    if len(scoped_records) != len(app_records):
        missing_keys = sorted(app_office_keys - {office_key(record) for record in scoped_records})
        raise SystemExit(
            json.dumps(
                {
                    "message": "app manifest records could not be fully matched against integrated records",
                    "app_record_count": len(app_records),
                    "integrated_scope_count": len(scoped_records),
                    "missing_office_keys": missing_keys[:20],
                },
                ensure_ascii=False,
                indent=2,
            )
        )

    summary = build_app_summary(
        scoped_records,
        manifest.get("summary", {}),
        manifest.get("summary", {}).get("wam_match_summary", {}),
        manifest.get("summary", {}).get("wam_staffing_summary", {}),
    )

    manifest.setdefault("meta", {})
    manifest["meta"]["record_count"] = len(app_records)
    manifest["meta"]["total_records"] = len(app_records)
    manifest.setdefault("data_files", {})
    manifest["data_files"]["record_count"] = len(app_records)
    manifest["summary"] = summary
    manifest["analytics"] = {
        "overall_wage_stats": summary.get("overall_wage_stats"),
        "overall_utilization_stats": summary.get("overall_utilization_stats"),
        "wam_match_summary": summary.get("wam_match_summary"),
    }

    write_json(manifest_path, manifest)
    print(
        json.dumps(
            {
                "manifest": str(manifest_path),
                "record_count": len(app_records),
                "summary_record_count": summary.get("record_count"),
                "summary_response_breakdown": summary.get("response_breakdown"),
                "summary_municipality_count": summary.get("municipality_count"),
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
