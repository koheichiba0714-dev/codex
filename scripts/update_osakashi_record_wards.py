#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import io
import json
import re
import unicodedata
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.request import Request, urlopen


OFFICIAL_CSV_URL = "https://www.city.osaka.lg.jp/fukushi/cmsfiles/contents/0000603/603679/2026.3.1-15.csv"
WARD_RE = re.compile(r"大阪市([^\s0-9]+?区)")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="大阪市B型ダッシュボードの区未取得レコードへ区情報を補完する。")
    parser.add_argument(
        "--manifest",
        default="apps/r6kouchin-dashboard/data/dashboard-data.json",
        help="dashboard manifest path",
    )
    parser.add_argument(
        "--official-csv-cache",
        default="data/inputs/web_links/osaka_city_service_lists/osakashi_shuro_b_2026_03_01.csv",
        help="official Osaka City B-type CSV cache path",
    )
    parser.add_argument(
        "--override-input",
        default="data/inputs/web_links/osakashi_shuro_b_ward_overrides.csv",
        help="manual ward override CSV path",
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


def write_json(path: Path, payload: Any, pretty: bool = False) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if pretty:
        text = json.dumps(payload, ensure_ascii=False, indent=2)
    else:
        text = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    path.write_text(text, encoding="utf-8")


def load_csv_rows(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []
    with path.open(newline="", encoding="utf-8") as file_obj:
        return list(csv.DictReader(file_obj))


def normalize_name(value: Any, relaxed: bool = False) -> str:
    text = unicodedata.normalize("NFKC", str(value or "")).lower().replace("　", " ")
    cleaned_chars: list[str] = []
    for char in text:
        category = unicodedata.category(char)
        if category.startswith(("P", "S")) or char.isspace():
            continue
        cleaned_chars.append(char)
    cleaned = "".join(cleaned_chars)
    if not relaxed:
        return cleaned
    for token in [
        "就労継続支援b型事業所",
        "就労継続支援事業所b型",
        "就労継続支援b型",
        "就労継続支援b",
        "就労支援継続b型",
        "事業所",
        "就労支援施設",
        "就労支援",
        "b型作業所",
        "作業所",
        "オフィス",
        "センター",
    ]:
        cleaned = cleaned.replace(token, "")
    return cleaned


def ward_from_address(*parts: Any) -> str | None:
    text = "".join(str(part or "") for part in parts)
    matched = WARD_RE.search(text)
    return matched.group(1) if matched else None


def read_text_with_fallback(path: Path) -> str:
    raw = path.read_bytes()
    for encoding in ("utf-8", "cp932", "shift_jis"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def ensure_official_csv(cache_path: Path) -> Path:
    if cache_path.exists():
        return cache_path
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    request = Request(
        OFFICIAL_CSV_URL,
        headers={
            "User-Agent": "Mozilla/5.0",
        },
    )
    with urlopen(request, timeout=30) as response:
        raw = response.read()
    text = raw.decode("cp932", errors="replace")
    cache_path.write_text(text, encoding="utf-8")
    return cache_path


def load_official_rows(path: Path) -> list[dict[str, str]]:
    text = read_text_with_fallback(path)
    rows = list(csv.reader(io.StringIO(text)))
    if len(rows) < 3:
        return []
    header = rows[1]
    return [dict(zip(header, row)) for row in rows[2:] if any(cell.strip() for cell in row)]


def build_official_indexes(rows: list[dict[str, str]]) -> tuple[dict[str, list[dict[str, str]]], dict[str, list[dict[str, str]]]]:
    strict_index: dict[str, list[dict[str, str]]] = defaultdict(list)
    relaxed_index: dict[str, list[dict[str, str]]] = defaultdict(list)
    for row in rows:
        row = dict(row)
        row["ward"] = ward_from_address(row.get("事業所所在地"))
        if not row["ward"]:
            continue
        strict_key = normalize_name(row.get("事業所名称"))
        relaxed_key = normalize_name(row.get("事業所名称"), relaxed=True)
        if strict_key:
            strict_index[strict_key].append(row)
        if relaxed_key:
            relaxed_index[relaxed_key].append(row)
    return strict_index, relaxed_index


def load_chunk_records(manifest_path: Path) -> tuple[dict[str, Any], list[tuple[Path, list[dict[str, Any]]]]]:
    manifest = load_json(manifest_path)
    chunk_paths = manifest.get("data_files", {}).get("record_chunks", [])
    app_root_dir = manifest_path.parent.parent
    chunk_records: list[tuple[Path, list[dict[str, Any]]]] = []
    for relative_path in chunk_paths:
        chunk_path = (app_root_dir / relative_path).resolve()
        chunk_records.append((chunk_path, load_json(chunk_path)))
    return manifest, chunk_records


def current_area_label(record: dict[str, Any]) -> str | None:
    if record.get("osaka_city_ward"):
        return str(record.get("osaka_city_ward"))
    ward = ward_from_address(record.get("wam_office_address_city"), record.get("wam_office_address_line"))
    if ward:
        return ward
    city = str(record.get("wam_office_address_city") or "")
    if city:
        normalized = city.replace("大阪府", "").strip()
        return "大阪市（区未取得）" if normalized == "大阪市" else normalized
    return "大阪市（区未取得）" if record.get("municipality") == "大阪市" else record.get("municipality")


def choose_unique_ward(rows: list[dict[str, str]]) -> str | None:
    wards = {str(row.get("ward") or "").strip() for row in rows if row.get("ward")}
    if len(wards) == 1:
        return next(iter(wards))
    return None


def official_match(record: dict[str, Any], strict_index: dict[str, list[dict[str, str]]], relaxed_index: dict[str, list[dict[str, str]]]) -> tuple[str, str, str] | None:
    strict_key = normalize_name(record.get("office_name"))
    strict_rows = strict_index.get(strict_key, [])
    strict_ward = choose_unique_ward(strict_rows)
    if strict_ward:
        return strict_ward, "official_name_strict", str(strict_rows[0].get("事業所所在地") or "")

    relaxed_key = normalize_name(record.get("office_name"), relaxed=True)
    relaxed_rows = relaxed_index.get(relaxed_key, [])
    relaxed_ward = choose_unique_ward(relaxed_rows)
    if relaxed_ward:
        return relaxed_ward, "official_name_relaxed", str(relaxed_rows[0].get("事業所所在地") or "")
    return None


def override_lookup(rows: list[dict[str, str]]) -> dict[str, dict[str, str]]:
    return {
        str(row.get("office_no") or "").strip(): row
        for row in rows
        if str(row.get("office_no") or "").strip()
    }


def enrich_record(
    record: dict[str, Any],
    strict_index: dict[str, list[dict[str, str]]],
    relaxed_index: dict[str, list[dict[str, str]]],
    overrides: dict[str, dict[str, str]],
) -> tuple[dict[str, Any], str]:
    enriched = dict(record)
    office_no = str(record.get("office_no") or "").strip()
    ward = ward_from_address(record.get("wam_office_address_city"), record.get("wam_office_address_line"))
    if ward:
        enriched["osaka_city_ward"] = ward
        enriched["osaka_city_ward_source"] = "wam_address"
        enriched["osaka_city_ward_note"] = None
        return enriched, "wam_address"

    override = overrides.get(office_no)
    if override:
        enriched["osaka_city_ward"] = override.get("osaka_city_ward")
        enriched["osaka_city_ward_source"] = override.get("source_type") or "manual_override"
        enriched["osaka_city_ward_note"] = override.get("source_note") or None
        return enriched, "manual_override"

    matched = official_match(record, strict_index, relaxed_index)
    if matched:
        matched_ward, matched_source, matched_note = matched
        enriched["osaka_city_ward"] = matched_ward
        enriched["osaka_city_ward_source"] = matched_source
        enriched["osaka_city_ward_note"] = matched_note or None
        return enriched, matched_source

    enriched["osaka_city_ward"] = None
    enriched["osaka_city_ward_source"] = None
    enriched["osaka_city_ward_note"] = None
    return enriched, "unresolved"


def main() -> None:
    args = parse_args()
    manifest_path = resolve_path(args.manifest)
    official_csv_path = ensure_official_csv(resolve_path(args.official_csv_cache))
    override_path = resolve_path(args.override_input)

    manifest, chunk_records = load_chunk_records(manifest_path)
    official_rows = load_official_rows(official_csv_path)
    strict_index, relaxed_index = build_official_indexes(official_rows)
    overrides = override_lookup(load_csv_rows(override_path))

    stats: dict[str, int] = defaultdict(int)
    unresolved: list[tuple[Any, Any]] = []
    updated_chunks: list[tuple[Path, list[dict[str, Any]]]] = []

    for chunk_path, rows in chunk_records:
        updated_rows: list[dict[str, Any]] = []
        for record in rows:
            if record.get("municipality") != "大阪市":
                updated_rows.append(record)
                continue
            if current_area_label(record) == "大阪市（区未取得）":
                stats["before_unknown"] += 1
            enriched, source = enrich_record(record, strict_index, relaxed_index, overrides)
            stats[source] += 1
            if current_area_label(enriched) == "大阪市（区未取得）":
                unresolved.append((enriched.get("office_no"), enriched.get("office_name")))
            updated_rows.append(enriched)
        updated_chunks.append((chunk_path, updated_rows))

    after_unknown = 0
    total_osaka_city = 0
    for _, rows in updated_chunks:
        for row in rows:
            if row.get("municipality") != "大阪市":
                continue
            total_osaka_city += 1
            if current_area_label(row) == "大阪市（区未取得）":
                after_unknown += 1

    if unresolved:
        raise SystemExit(
            json.dumps(
                {
                    "message": "ward resolution incomplete",
                    "unresolved": unresolved,
                },
                ensure_ascii=False,
                indent=2,
            )
        )

    for chunk_path, rows in updated_chunks:
        write_json(chunk_path, rows, pretty=False)

    manifest.setdefault("meta", {})
    manifest["meta"]["osaka_city_ward_generated_at"] = datetime.now().astimezone().isoformat(timespec="seconds")
    manifest["meta"]["osaka_city_ward_coverage_count"] = total_osaka_city - after_unknown
    manifest["meta"]["osaka_city_ward_unknown_count"] = after_unknown
    write_json(manifest_path, manifest, pretty=True)

    print(
        json.dumps(
            {
                "official_csv_cache": str(official_csv_path),
                "before_unknown": stats["before_unknown"],
                "after_unknown": after_unknown,
                "matched_by_wam_address": stats["wam_address"],
                "matched_by_manual_override": stats["manual_override"],
                "matched_by_official_name_strict": stats["official_name_strict"],
                "matched_by_official_name_relaxed": stats["official_name_relaxed"],
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
