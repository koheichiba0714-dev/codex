#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
from datetime import datetime
import html
import json
from pathlib import Path
import re
import time
from typing import Any
from urllib.parse import urlparse
import unicodedata

import requests


SEARCH_URL = "https://html.duckduckgo.com/html/"
USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36"
)
OUTPUT_FIELDS = [
    "office_no",
    "municipality",
    "corporation_name",
    "office_name",
    "wam_office_url",
    "homepage_url",
    "homepage_source",
    "homepage_confidence",
    "instagram_url",
    "instagram_source",
    "instagram_confidence",
    "search_query",
    "search_checked_at",
]

DIRECTORY_DOMAINS = {
    "www.wam.go.jp",
    "wam.go.jp",
    "snabi.jp",
    "www.snabi.jp",
    "shohgaisha.com",
    "www.shohgaisha.com",
    "litalico-c.jp",
    "www.litalico-c.jp",
}
SOCIAL_DOMAINS = {
    "instagram.com",
    "www.instagram.com",
    "m.instagram.com",
    "facebook.com",
    "www.facebook.com",
    "x.com",
    "www.x.com",
    "twitter.com",
    "www.twitter.com",
    "tiktok.com",
    "www.tiktok.com",
    "youtube.com",
    "www.youtube.com",
}
INVALID_INSTAGRAM_SEGMENTS = {"p", "reel", "reels", "stories", "explore", "tv"}
LEGAL_SUFFIXES = [
    "株式会社",
    "合同会社",
    "一般社団法人",
    "一般財団法人",
    "社会福祉法人",
    "医療法人",
    "特定非営利活動法人",
    "npo法人",
]
RESULT_RE = re.compile(
    r'<a rel="nofollow" class="result__a" href="(?P<url>.*?)">(?P<title>.*?)</a>.*?'
    r'(?:<a class="result__snippet" href=".*?">(?P<snippet>.*?)</a>)?',
    re.S,
)
SEARCH_DISABLED = False


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="大阪市の就労継続支援B型について、ホームページとInstagramを検索補完する。"
    )
    parser.add_argument(
        "--dashboard-input",
        default="apps/r6kouchin-dashboard/data/dashboard-data.json",
        help="dashboard manifest json path",
    )
    parser.add_argument(
        "--output",
        default="data/inputs/web_links/osakashi_shuro_b_office_links.csv",
        help="output csv path",
    )
    parser.add_argument(
        "--municipality",
        default="大阪市",
        help="target municipality label",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=0,
        help="limit records for testing",
    )
    parser.add_argument(
        "--sleep-seconds",
        type=float,
        default=0.55,
        help="sleep between search requests",
    )
    parser.add_argument(
        "--refresh",
        action="store_true",
        help="ignore existing output cache and rebuild",
    )
    return parser.parse_args()


def repo_root() -> Path:
    return Path(__file__).resolve().parents[1]


def resolve_path(path_str: str) -> Path:
    path = Path(path_str).expanduser()
    if path.is_absolute():
        return path.resolve()
    return (repo_root() / path).resolve()


def load_dashboard_records(manifest_path: Path) -> list[dict[str, Any]]:
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    app_root = manifest_path.parent.parent
    records: list[dict[str, Any]] = []
    for rel_path in manifest.get("data_files", {}).get("record_chunks", []):
        records.extend(json.loads((app_root / rel_path).read_text(encoding="utf-8")))
    return records


def load_existing_rows(path: Path) -> dict[str, dict[str, str]]:
    if not path.exists():
        return {}
    with path.open(newline="", encoding="utf-8") as file_obj:
        return {
            str(row.get("office_no") or "").strip(): row
            for row in csv.DictReader(file_obj)
            if str(row.get("office_no") or "").strip()
        }


def write_rows(path: Path, rows: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as file_obj:
        writer = csv.DictWriter(file_obj, fieldnames=OUTPUT_FIELDS)
        writer.writeheader()
        for row in rows:
            writer.writerow({field: row.get(field, "") for field in OUTPUT_FIELDS})


def normalize_text(value: str | None) -> str:
    text = unicodedata.normalize("NFKC", value or "").lower()
    text = text.replace("　", " ")
    text = re.sub(r"\s+", "", text)
    text = re.sub(r"[!-/:-@[-`{-~、。・･「」『』（）【】［］\[\]<>〈〉《》]", "", text)
    return text


def strip_legal_suffix(value: str | None) -> str:
    text = value or ""
    for suffix in LEGAL_SUFFIXES:
        text = text.replace(suffix, "")
    return text.strip()


def clean_html_fragment(value: str | None) -> str:
    if not value:
        return ""
    return re.sub(r"<.*?>", "", html.unescape(value)).strip()


def canonicalize_url(value: str | None) -> str | None:
    if not value:
        return None
    candidate = html.unescape(str(value)).strip()
    if not candidate:
        return None
    if candidate.startswith("//"):
        candidate = f"https:{candidate}"
    if not re.match(r"^https?://", candidate, flags=re.I):
        return None
    try:
        parsed = urlparse(candidate)
    except ValueError:
        return None
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        return None
    normalized_path = parsed.path.rstrip("/") or ""
    return f"{parsed.scheme}://{parsed.netloc.lower()}{normalized_path}"


def hostname(value: str | None) -> str:
    if not value:
        return ""
    try:
        return urlparse(value).netloc.lower()
    except ValueError:
        return ""


def is_instagram_url(value: str | None) -> bool:
    return "instagram.com" in hostname(value)


def normalize_instagram_profile_url(value: str | None) -> str | None:
    url = canonicalize_url(value)
    if not is_instagram_url(url):
        return None
    try:
        parsed = urlparse(url)
    except ValueError:
        return None
    segments = [segment for segment in parsed.path.split("/") if segment]
    if not segments or segments[0].lower() in INVALID_INSTAGRAM_SEGMENTS:
        return None
    return f"https://www.instagram.com/{segments[0]}"


def is_directory_url(value: str | None) -> bool:
    host = hostname(value)
    return host in DIRECTORY_DOMAINS


def is_homepage_candidate(value: str | None) -> bool:
    host = hostname(value)
    if not host:
        return False
    if host in SOCIAL_DOMAINS or host in DIRECTORY_DOMAINS:
        return False
    return True


def build_query(record: dict[str, Any]) -> str:
    office_name = record.get("office_name") or ""
    municipality = record.get("municipality") or ""
    parts = [
        f'"{office_name}"' if office_name else "",
        f'"{municipality}"' if municipality else "",
        strip_legal_suffix(record.get("corporation_name") or ""),
        "就労継続支援B型",
        "Instagram",
    ]
    return " ".join(part for part in parts if part).strip()


def search_duckduckgo(session: requests.Session, query: str) -> list[dict[str, str]]:
    global SEARCH_DISABLED
    if SEARCH_DISABLED:
        return []
    response = session.post(
        SEARCH_URL,
        data={"q": query},
        timeout=20,
        headers={"User-Agent": USER_AGENT},
    )
    if response.status_code == 403:
        SEARCH_DISABLED = True
        return []
    response.raise_for_status()
    results: list[dict[str, str]] = []
    for match in RESULT_RE.finditer(response.text):
        url = canonicalize_url(match.group("url"))
        if not url:
            continue
        results.append(
            {
                "url": url,
                "title": clean_html_fragment(match.group("title")),
                "snippet": clean_html_fragment(match.group("snippet")),
            }
        )
        if len(results) >= 8:
            break
    return results


def office_markers(record: dict[str, Any]) -> dict[str, str]:
    office = normalize_text(record.get("office_name"))
    corporation = normalize_text(strip_legal_suffix(record.get("corporation_name")))
    municipality = normalize_text(record.get("municipality"))
    return {
        "office": office,
        "corporation": corporation,
        "municipality": municipality,
    }


def score_homepage(record: dict[str, Any], result: dict[str, str]) -> tuple[int, str]:
    url = result["url"]
    if not is_homepage_candidate(url):
        return (-999, "social_or_directory")
    combined = normalize_text(" ".join([result["title"], result["snippet"], url]))
    markers = office_markers(record)
    score = 0
    reason_parts: list[str] = []
    if markers["office"] and markers["office"] in combined:
        score += 6
        reason_parts.append("office")
    if markers["corporation"] and markers["corporation"] in combined:
        score += 3
        reason_parts.append("corp")
    if markers["municipality"] and markers["municipality"] in combined:
        score += 1
        reason_parts.append("area")
    if "就労継続支援" in combined or "b型" in combined or "障がい" in combined:
        score += 1
        reason_parts.append("service")
    if "instagram" in combined:
        score -= 2
    return (score, "+".join(reason_parts) or "weak")


def score_instagram(record: dict[str, Any], result: dict[str, str]) -> tuple[int, str]:
    url = normalize_instagram_profile_url(result["url"])
    if not url:
        return (-999, "not_instagram")
    combined = normalize_text(" ".join([result["title"], result["snippet"], url]))
    markers = office_markers(record)
    score = 0
    reason_parts: list[str] = []
    if markers["office"] and markers["office"] in combined:
        score += 6
        reason_parts.append("office")
    if markers["corporation"] and markers["corporation"] in combined:
        score += 3
        reason_parts.append("corp")
    if markers["municipality"] and markers["municipality"] in combined:
        score += 1
        reason_parts.append("area")
    if "公式" in combined or "official" in combined:
        score += 1
        reason_parts.append("official")
    return (score, "+".join(reason_parts) or "weak")


def confidence_from_score(score: int) -> str:
    if score >= 8:
        return "high"
    if score >= 5:
        return "medium"
    return "low"


def pick_best_result(
    record: dict[str, Any],
    results: list[dict[str, str]],
    scorer,
    minimum_score: int,
    url_transformer=lambda value: value,
) -> tuple[str | None, str | None, str | None]:
    best_url: str | None = None
    best_score = -999
    best_reason: str | None = None
    for result in results:
        score, reason = scorer(record, result)
        if score > best_score:
            best_score = score
            best_url = url_transformer(result["url"])
            best_reason = reason
    if best_score < minimum_score:
        return (None, None, None)
    return (best_url, "ddg_search", confidence_from_score(best_score))


def build_row(
    record: dict[str, Any],
    session: requests.Session,
    sleep_seconds: float,
    cached_row: dict[str, str] | None,
    refresh: bool,
) -> dict[str, Any]:
    office_no = str(record.get("office_no") or "").strip()
    wam_office_url = canonicalize_url(record.get("wam_office_url"))

    if cached_row and not refresh:
        return {field: cached_row.get(field, "") for field in OUTPUT_FIELDS}

    homepage_url = None
    homepage_source = None
    homepage_confidence = None
    instagram_url = None
    instagram_source = None
    instagram_confidence = None

    if is_instagram_url(wam_office_url):
        normalized_instagram = normalize_instagram_profile_url(wam_office_url)
        if normalized_instagram:
            instagram_url = normalized_instagram
            instagram_source = "wam_url"
            instagram_confidence = "high"
    elif is_homepage_candidate(wam_office_url):
        homepage_url = wam_office_url
        homepage_source = "wam_url"
        homepage_confidence = "high"

    search_query = ""
    if not homepage_url or not instagram_url:
        search_query = build_query(record)
        results: list[dict[str, str]] = []
        searched = False
        if not SEARCH_DISABLED:
            searched = True
            try:
                results = search_duckduckgo(session, search_query)
            except requests.RequestException:
                results = []
        if not homepage_url:
            homepage_url, homepage_source, homepage_confidence = pick_best_result(
                record,
                results,
                score_homepage,
                6,
            )
        if not instagram_url:
            instagram_url, instagram_source, instagram_confidence = pick_best_result(
                record,
                results,
                score_instagram,
                5,
                url_transformer=normalize_instagram_profile_url,
            )
        if searched and not SEARCH_DISABLED:
            time.sleep(sleep_seconds)

    return {
        "office_no": office_no,
        "municipality": record.get("municipality") or "",
        "corporation_name": record.get("corporation_name") or "",
        "office_name": record.get("office_name") or "",
        "wam_office_url": wam_office_url or "",
        "homepage_url": homepage_url or "",
        "homepage_source": homepage_source or "",
        "homepage_confidence": homepage_confidence or "",
        "instagram_url": instagram_url or "",
        "instagram_source": instagram_source or "",
        "instagram_confidence": instagram_confidence or "",
        "search_query": search_query,
        "search_checked_at": datetime.now().isoformat(timespec="seconds"),
    }


def main() -> None:
    args = parse_args()
    dashboard_input = resolve_path(args.dashboard_input)
    output_path = resolve_path(args.output)

    records = [
        record
        for record in load_dashboard_records(dashboard_input)
        if record.get("municipality") == args.municipality
    ]
    records.sort(key=lambda row: (str(row.get("office_no") or ""), str(row.get("office_name") or "")))
    if args.limit > 0:
        records = records[: args.limit]

    cached = load_existing_rows(output_path)
    session = requests.Session()
    session.headers.update({"User-Agent": USER_AGENT})

    rows = [
        build_row(
            record,
            session=session,
            sleep_seconds=args.sleep_seconds,
            cached_row=cached.get(str(record.get("office_no") or "").strip()),
            refresh=args.refresh,
        )
        for record in records
    ]
    write_rows(output_path, rows)

    homepage_count = sum(1 for row in rows if row.get("homepage_url"))
    instagram_count = sum(1 for row in rows if row.get("instagram_url"))
    print(
        json.dumps(
            {
                "record_count": len(rows),
                "homepage_count": homepage_count,
                "instagram_count": instagram_count,
                "output": str(output_path),
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
