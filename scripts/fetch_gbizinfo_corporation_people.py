#!/usr/bin/env python3
from __future__ import annotations

import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from html import unescape
import json
import re
import time
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen


USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/136.0.0.0 Safari/537.36"
)
GBIZINFO_URL_TEMPLATE = "https://info.gbiz.go.jp/hojin/ichiran?hojinBango={corporation_number}"
TITLE_RE = re.compile(r"<title>\s*(.*?)\s*\|\s*(\d{13})\s*\|\s*Gビズインフォ", re.S)
FIELD_BLOCK_RE = re.compile(
    r'<p class="fw-bold[^"]*">\s*(?P<label>[^<]+?)\s*</p>\s*'
    r'<p class="[^"]*col-md-10[^"]*">\s*(?P<value>.*?)</p>',
    re.S,
)
TAG_RE = re.compile(r"<[^>]+>")
REPRESENTATIVE_NOTE_RE = re.compile(r"\s*[（(][^()（）]*(職場情報総合サイト|GEPS)[^()（）]*[)）]\s*$")
ROLE_PREFIXES = [
    "代表取締役社長",
    "代表取締役会長",
    "代表取締役",
    "取締役社長",
    "代表社員",
    "代表理事",
    "理事長",
    "会長",
    "社長",
    "代表",
    "理事",
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Gビズインフォの公開ページから法人番号ごとの代表者名を取得する。"
    )
    parser.add_argument(
        "--input",
        default="data/exports/r6kouchinjissekib/normalized/shuro_b_dashboard.json",
        help="法人番号を含む JSON 入力",
    )
    parser.add_argument(
        "--output",
        default="data/exports/gbizinfo/osaka_shuro_b/gbizinfo_corporation_people.json",
        help="代表者情報の JSON 出力",
    )
    parser.add_argument(
        "--by-number-dir",
        default="data/exports/gbizinfo/osaka_shuro_b/by_number",
        help="法人番号別キャッシュディレクトリ",
    )
    parser.add_argument(
        "--max-workers",
        type=int,
        default=4,
        help="同時取得数",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=0,
        help="取得法人数の上限。0 は無制限",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="既存キャッシュがあっても再取得する",
    )
    return parser.parse_args()


def resolve_path(value: str) -> Path:
    path = Path(value)
    if path.is_absolute():
        return path
    return Path(__file__).resolve().parents[1] / path


def load_records(path: Path) -> list[dict[str, Any]]:
    payload = json.loads(path.read_text())
    if isinstance(payload, dict):
        records = payload.get("records", [])
    else:
        records = payload
    if not isinstance(records, list):
        raise ValueError(f"records not found in {path}")
    return [row for row in records if isinstance(row, dict)]


def normalize_whitespace(value: str) -> str:
    return re.sub(r"\s+", " ", value.replace("\u3000", " ")).strip()


def clean_representative_note(value: str | None) -> str:
    return REPRESENTATIVE_NOTE_RE.sub("", normalize_whitespace(value or ""))


def normalize_person_name(value: str | None) -> str | None:
    if not value:
        return None
    text = clean_representative_note(value).lower()
    text = text.replace("　", " ")
    text = re.sub(r"[・･·]", "", text)
    text = re.sub(r"\s+", "", text)
    return text or None


def split_representative(raw_value: str | None) -> tuple[str | None, str | None]:
    if not raw_value:
        return None, None
    value = clean_representative_note(raw_value)
    for prefix in ROLE_PREFIXES:
        if value.startswith(prefix):
            person = value[len(prefix) :].strip(" :：　・･")
            while person:
                person = re.sub(r"^[・･\s]+", "", person)
                nested_prefix = next((item for item in ROLE_PREFIXES if person.startswith(item)), None)
                if not nested_prefix:
                    break
                person = person[len(nested_prefix) :].strip(" :：　・･")
            return prefix, person or None
    return None, value


def strip_tags(value: str) -> str:
    return normalize_whitespace(unescape(TAG_RE.sub(" ", value)))


def parse_html(corporation_number: str, html_text: str) -> dict[str, Any]:
    title_match = TITLE_RE.search(html_text)
    field_map: dict[str, str] = {}
    for match in FIELD_BLOCK_RE.finditer(html_text):
        label = strip_tags(match.group("label"))
        value = strip_tags(match.group("value"))
        field_map[label] = value

    representative_raw = field_map.get("代表者名") or None
    representative_role, representative_name = split_representative(representative_raw)
    corporation_name = field_map.get("法人名") or (strip_tags(title_match.group(1)) if title_match else None)
    if corporation_name == corporation_number:
        corporation_name = None

    return {
        "corporation_number": corporation_number,
        "corporation_name": corporation_name,
        "representative_raw": representative_raw,
        "representative_role": representative_role,
        "representative_name": representative_name,
        "representative_name_normalized": normalize_person_name(representative_name),
        "source_url": GBIZINFO_URL_TEMPLATE.format(corporation_number=corporation_number),
        "fetch_status": "ok",
        "fetch_error": "",
    }


def fetch_corporation(corporation_number: str) -> dict[str, Any]:
    url = GBIZINFO_URL_TEMPLATE.format(corporation_number=corporation_number)
    request = Request(url, headers={"User-Agent": USER_AGENT})
    try:
        with urlopen(request, timeout=20) as response:
            body = response.read().decode("utf-8", errors="replace")
    except HTTPError as exc:
        return {
            "corporation_number": corporation_number,
            "corporation_name": None,
            "representative_raw": None,
            "representative_role": None,
            "representative_name": None,
            "representative_name_normalized": None,
            "source_url": url,
            "fetch_status": "http_error",
            "fetch_error": f"{exc.code} {exc.reason}",
        }
    except URLError as exc:
        return {
            "corporation_number": corporation_number,
            "corporation_name": None,
            "representative_raw": None,
            "representative_role": None,
            "representative_name": None,
            "representative_name_normalized": None,
            "source_url": url,
            "fetch_status": "network_error",
            "fetch_error": str(exc.reason),
        }

    result = parse_html(corporation_number, body)
    result["fetched_at"] = datetime.now().astimezone().isoformat(timespec="seconds")
    return result


def write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n")


def main() -> None:
    args = parse_args()
    input_path = resolve_path(args.input)
    output_path = resolve_path(args.output)
    by_number_dir = resolve_path(args.by_number_dir)

    records = load_records(input_path)
    corporation_rows: dict[str, dict[str, Any]] = {}
    for record in records:
        corporation_number = str(record.get("corporation_number") or "").strip()
        if not corporation_number:
            continue
        row = corporation_rows.setdefault(
            corporation_number,
            {
                "corporation_number": corporation_number,
                "corporation_names": set(),
                "office_count": 0,
                "office_names": [],
            },
        )
        if record.get("corporation_name"):
            row["corporation_names"].add(str(record["corporation_name"]).strip())
        if record.get("office_name"):
            row["office_names"].append(str(record["office_name"]).strip())
        row["office_count"] += 1

    corporation_numbers = sorted(corporation_rows)
    if args.limit > 0:
        corporation_numbers = corporation_numbers[: args.limit]

    by_number_dir.mkdir(parents=True, exist_ok=True)
    results: dict[str, dict[str, Any]] = {}
    to_fetch: list[str] = []

    for corporation_number in corporation_numbers:
        cache_path = by_number_dir / f"{corporation_number}.json"
        if cache_path.exists() and not args.force:
            results[corporation_number] = json.loads(cache_path.read_text())
        else:
            to_fetch.append(corporation_number)

    if to_fetch:
        with ThreadPoolExecutor(max_workers=max(1, args.max_workers)) as executor:
            future_map = {
                executor.submit(fetch_corporation, corporation_number): corporation_number
                for corporation_number in to_fetch
            }
            for index, future in enumerate(as_completed(future_map), start=1):
                corporation_number = future_map[future]
                result = future.result()
                results[corporation_number] = result
                write_json(by_number_dir / f"{corporation_number}.json", result)
                if index % 50 == 0 or index == len(to_fetch):
                    print(f"fetched {index}/{len(to_fetch)}")
                time.sleep(0.05)

    output_rows: list[dict[str, Any]] = []
    for corporation_number in corporation_numbers:
        result = dict(results[corporation_number])
        summary = corporation_rows[corporation_number]
        result["source_corporation_names"] = sorted(summary["corporation_names"])
        result["source_office_count"] = summary["office_count"]
        result["source_office_names"] = sorted(set(summary["office_names"]))[:30]
        output_rows.append(result)

    coverage_count = sum(1 for row in output_rows if row.get("representative_name"))
    summary_payload = {
        "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "input": str(input_path),
        "corporation_count": len(output_rows),
        "representative_name_count": coverage_count,
        "records": output_rows,
    }
    write_json(output_path, summary_payload)
    print(
        json.dumps(
            {
                "corporation_count": len(output_rows),
                "representative_name_count": coverage_count,
                "output": str(output_path),
            },
            ensure_ascii=False,
        )
    )


if __name__ == "__main__":
    main()
