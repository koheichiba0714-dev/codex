#!/usr/bin/env python3
from __future__ import annotations

import argparse
from concurrent.futures import ThreadPoolExecutor
import csv
from datetime import datetime
import html
import json
from pathlib import Path
import re
import time
from typing import Any
from urllib.parse import urljoin, urlparse
import unicodedata
import xml.etree.ElementTree as ET

from pypdf import PdfReader
import requests


SEARCH_URL = "https://html.duckduckgo.com/html/"
BING_SEARCH_URL = "https://www.bing.com/search"
OSAKA_CITY_PDF_URL = "https://www.city.osaka.lg.jp/fukushi/cmsfiles/contents/0000603/603679/b_r6_shuurouninnzuuchousa.pdf"
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
PROFILE_BRIDGE_DOMAINS = {
    "linktr.ee",
    "www.linktr.ee",
    "lit.link",
    "www.lit.link",
    "lit.link",
    "bio.site",
    "www.bio.site",
    "carrd.co",
    "www.carrd.co",
    "instabio.cc",
    "www.instabio.cc",
    "taplink.cc",
    "www.taplink.cc",
}
NON_HTML_SUFFIXES = (
    ".css",
    ".js",
    ".json",
    ".xml",
    ".rss",
    ".txt",
    ".pdf",
    ".jpg",
    ".jpeg",
    ".png",
    ".gif",
    ".svg",
    ".webp",
    ".ico",
    ".mp4",
    ".mov",
    ".webm",
    ".mp3",
    ".wav",
    ".woff",
    ".woff2",
    ".ttf",
    ".otf",
    ".zip",
)
INTERNAL_PAGE_KEYWORDS = (
    "/about",
    "/aboutus",
    "/company",
    "/facility",
    "/office",
    "/access",
    "/contact",
    "/profile",
    "/service",
    "/news",
    "/blog",
    "/staff",
    "/guide",
    "/support",
    "/works",
    "/business",
    "/jigyosyo",
    "/center",
    "/shop",
    "/tenpo",
    "/day",
    "/information",
    "/sns",
)
SITEMAP_HINT_PATHS = ("/robots.txt", "/sitemap.xml", "/sitemap_index.xml")
SITEMAP_FETCH_LIMIT = 2
SITEMAP_PAGE_LIMIT = 6
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
HREF_RE = re.compile(r'href=["\'](?P<href>[^"\']+)["\']', re.I)
PDF_URL_RE = re.compile(r"(?:https?://|https//|http//|hxxps?://|hyyps?://|www\.)[^\s　]+", re.I)
PDF_ROW_RE = re.compile(r"^(?P<office_number>\d{10})\s+(?P<office_name>.+?)\s+大阪市", re.S)
SEARCH_DISABLED = False
INSTAGRAM_SEARCH_DISABLED = False


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
        "--osaka-city-pdf-cache",
        default="data/inputs/web_links/osaka_city_b_r6_services.pdf",
        help="cached official Osaka City PDF path",
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
        default=0.2,
        help="sleep between search requests",
    )
    parser.add_argument(
        "--workers",
        type=int,
        default=8,
        help="parallel workers for homepage / instagram fetch",
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


def build_session() -> requests.Session:
    session = requests.Session()
    session.headers.update({"User-Agent": USER_AGENT})
    return session


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
    candidate = candidate.rstrip(".,、。)>］】")
    candidate = candidate.replace("hxxps://", "https://").replace("hxxp://", "http://")
    candidate = candidate.replace("hyyps://", "https://").replace("hyyP://", "https://")
    if re.match(r"^https?//", candidate, flags=re.I):
        candidate = re.sub(r"^(https?)//", r"\1://", candidate, flags=re.I)
    if candidate.startswith("//"):
        candidate = f"https:{candidate}"
    if candidate.startswith("www."):
        candidate = f"https://{candidate}"
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


def decode_embedded_urls(value: str) -> str:
    decoded = html.unescape(value or "")
    replacements = (
        ("\\/", "/"),
        ("\\u002f", "/"),
        ("\\u002F", "/"),
        ("\\x2f", "/"),
        ("\\x2F", "/"),
        ("\\u003a", ":"),
        ("\\u003A", ":"),
        ("\\x3a", ":"),
        ("\\x3A", ":"),
        ("\\u0026", "&"),
    )
    for before, after in replacements:
        decoded = decoded.replace(before, after)
    return decoded


def is_bridge_profile_url(value: str | None) -> bool:
    return hostname(value) in PROFILE_BRIDGE_DOMAINS


def is_probably_html_page(value: str | None) -> bool:
    if not value:
        return False
    try:
        parsed = urlparse(value)
    except ValueError:
        return False
    path = (parsed.path or "").lower()
    if path.endswith(NON_HTML_SUFFIXES):
        return False
    if path.startswith("/wp-content/") or path.startswith("/wp-includes/"):
        return False
    if any(fragment in path for fragment in ("/feed", "/comments/", "/wp-json", "/xmlrpc.php")):
        return False
    return True


def rank_internal_link(url: str) -> tuple[int, int, str]:
    lower = url.lower()
    keyword_score = sum(1 for keyword in INTERNAL_PAGE_KEYWORDS if keyword in lower)
    return (-keyword_score, len(lower), lower)


def parse_sitemap_locations(xml_text: str) -> list[str]:
    return [clean_html_fragment(match) for match in re.findall(r"<loc>(.*?)</loc>", xml_text, re.I | re.S)]


def discover_sitemap_page_urls(session: requests.Session, homepage_url: str) -> list[str]:
    try:
        parsed_homepage = urlparse(homepage_url)
    except ValueError:
        return []
    if not parsed_homepage.scheme or not parsed_homepage.netloc:
        return []
    origin = f"{parsed_homepage.scheme}://{parsed_homepage.netloc}"
    sitemap_urls: list[str] = []
    checked_hints: set[str] = set()
    for hint_path in SITEMAP_HINT_PATHS:
        hint_url = urljoin(origin, hint_path)
        if hint_url in checked_hints:
            continue
        checked_hints.add(hint_url)
        try:
            response = session.get(hint_url, timeout=4, headers={"User-Agent": USER_AGENT})
            response.raise_for_status()
        except requests.RequestException:
            continue
        text = response.text
        if hint_path == "/robots.txt":
            for match in re.findall(r"(?im)^sitemap:\s*(\S+)", text):
                normalized = canonicalize_url(match)
                if normalized:
                    sitemap_urls.append(normalized)
        else:
            sitemap_urls.append(response.url)

    sitemap_urls = list(dict.fromkeys(sitemap_urls))[:SITEMAP_FETCH_LIMIT]
    page_urls: list[str] = []
    checked_sitemaps: set[str] = set()
    for sitemap_url in sitemap_urls:
        if sitemap_url in checked_sitemaps:
            continue
        checked_sitemaps.add(sitemap_url)
        try:
            response = session.get(sitemap_url, timeout=4, headers={"User-Agent": USER_AGENT})
            response.raise_for_status()
        except requests.RequestException:
            continue
        locations = parse_sitemap_locations(response.text)
        child_sitemaps: list[str] = []
        for location in locations:
            normalized = canonicalize_url(location)
            if not normalized:
                continue
            if normalized.endswith(".xml"):
                child_sitemaps.append(normalized)
                continue
            if hostname(normalized) != parsed_homepage.netloc.lower() or not is_probably_html_page(normalized):
                continue
            page_urls.append(normalized)
        for child_sitemap in child_sitemaps[:1]:
            if child_sitemap in checked_sitemaps:
                continue
            checked_sitemaps.add(child_sitemap)
            try:
                child_response = session.get(child_sitemap, timeout=4, headers={"User-Agent": USER_AGENT})
                child_response.raise_for_status()
            except requests.RequestException:
                continue
            for location in parse_sitemap_locations(child_response.text):
                normalized = canonicalize_url(location)
                if (
                    normalized
                    and hostname(normalized) == parsed_homepage.netloc.lower()
                    and not normalized.endswith(".xml")
                    and is_probably_html_page(normalized)
                ):
                    page_urls.append(normalized)
        if page_urls:
            break
    ordered = sorted(dict.fromkeys(page_urls), key=rank_internal_link)
    return ordered[:SITEMAP_PAGE_LIMIT]


def fetch_html_page(session: requests.Session, url: str, timeout: float) -> tuple[str, str] | None:
    try:
        response = session.get(url, timeout=timeout, headers={"User-Agent": USER_AGENT})
        response.raise_for_status()
    except requests.RequestException:
        return None
    content_type = (response.headers.get("content-type") or "").lower()
    if content_type and all(token not in content_type for token in ("text/html", "application/xhtml+xml", "text/plain")):
        return None
    return response.url, response.text


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


def build_instagram_query(record: dict[str, Any]) -> str:
    office_name = record.get("office_name") or ""
    municipality = record.get("municipality") or ""
    parts = [
        f'"{office_name}"' if office_name else "",
        f'"{municipality}"' if municipality else "",
        strip_legal_suffix(record.get("corporation_name") or ""),
        "site:instagram.com",
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


def search_bing_rss(session: requests.Session, query: str) -> list[dict[str, str]]:
    response = session.get(
        BING_SEARCH_URL,
        params={"q": query, "format": "rss"},
        timeout=20,
        headers={"User-Agent": USER_AGENT},
    )
    response.raise_for_status()
    root = ET.fromstring(response.text)
    results: list[dict[str, str]] = []
    for item in root.findall("./channel/item")[:8]:
        url = canonicalize_url(item.findtext("link"))
        if not url:
            continue
        results.append(
            {
                "url": url,
                "title": clean_html_fragment(item.findtext("title")),
                "snippet": clean_html_fragment(item.findtext("description")),
            }
        )
    return results


def ensure_osaka_city_pdf(cache_path: Path) -> Path:
    if cache_path.exists():
        return cache_path
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    response = requests.get(OSAKA_CITY_PDF_URL, timeout=30, headers={"User-Agent": USER_AGENT})
    response.raise_for_status()
    cache_path.write_bytes(response.content)
    return cache_path


def parse_osaka_city_pdf_homepages(pdf_path: Path) -> list[dict[str, str]]:
    reader = PdfReader(str(pdf_path))
    rows: list[dict[str, str]] = []
    current = ""
    for page in reader.pages:
        for raw_line in (page.extract_text() or "").splitlines():
            line = raw_line.strip()
            if not line:
                continue
            if "大阪市 就労継続支援B型事業所でのサービス提供内容について" in line:
                continue
            if "事業所ホームページ" in line and "事業所所在地" in line:
                continue
            if re.fullmatch(r"\d+/\d+", line):
                continue
            if re.match(r"^\d{10}\s+", line):
                if current:
                    row = parse_pdf_row(current)
                    if row:
                        rows.append(row)
                current = line
            elif current:
                current = f"{current} {line}"
        if current:
            row = parse_pdf_row(current)
            if row:
                rows.append(row)
            current = ""
    return rows


def parse_pdf_row(text: str) -> dict[str, str] | None:
    normalized = re.sub(r"\s+", " ", text).strip()
    match = PDF_ROW_RE.match(normalized)
    if not match:
        return None
    office_number = match.group("office_number")
    office_name = match.group("office_name").strip()
    url_match = PDF_URL_RE.search(normalized)
    homepage_url = canonicalize_url(url_match.group(0)) if url_match else None
    return {
        "office_number": office_number,
        "office_name": office_name,
        "homepage_url": homepage_url or "",
    }


def pdf_homepage_lookups(rows: list[dict[str, str]]) -> tuple[dict[str, str], dict[str, str]]:
    by_number = {
        row["office_number"]: row["homepage_url"]
        for row in rows
        if row.get("office_number") and row.get("homepage_url")
    }
    by_name: dict[str, str] = {}
    name_counts: dict[str, int] = {}
    for row in rows:
        name_key = normalize_text(row.get("office_name"))
        if not name_key or not row.get("homepage_url"):
            continue
        name_counts[name_key] = name_counts.get(name_key, 0) + 1
        by_name[name_key] = row["homepage_url"]
    by_name = {name: url for name, url in by_name.items() if name_counts.get(name) == 1}
    return by_number, by_name


def homepage_from_pdf(
    record: dict[str, Any],
    by_number: dict[str, str],
    by_name: dict[str, str],
) -> str | None:
    office_number = str(record.get("wam_office_number") or "").strip()
    if office_number and office_number in by_number:
        return by_number[office_number]
    return by_name.get(normalize_text(record.get("office_name")))


def extract_instagram_candidates(base_url: str, html_text: str) -> list[str]:
    decoded_html = decode_embedded_urls(html_text)
    candidates: list[str] = []
    for match in re.finditer(r"https?://(?:www\.)?instagram\.com/[^\"'<>/\s]+(?:/[^\"'<>/\s]+)?", decoded_html, re.I):
        normalized = normalize_instagram_profile_url(match.group(0))
        if normalized:
            candidates.append(normalized)
    for href_match in HREF_RE.finditer(decoded_html):
        href = href_match.group("href")
        absolute = canonicalize_url(urljoin(base_url, href))
        normalized = normalize_instagram_profile_url(absolute)
        if normalized:
            candidates.append(normalized)
    return list(dict.fromkeys(candidates))


def extract_internal_links(base_url: str, html_text: str) -> list[str]:
    base_host = hostname(base_url)
    decoded_html = decode_embedded_urls(html_text)
    links: list[str] = []
    fallback_links: list[str] = []
    for href_match in HREF_RE.finditer(decoded_html):
        href = href_match.group("href")
        absolute = canonicalize_url(urljoin(base_url, href))
        if (
            not absolute
            or hostname(absolute) != base_host
            or absolute == canonicalize_url(base_url)
            or not is_probably_html_page(absolute)
        ):
            continue
        lower = absolute.lower()
        if any(keyword in lower for keyword in INTERNAL_PAGE_KEYWORDS):
            links.append(absolute)
        else:
            fallback_links.append(absolute)
    ordered = sorted(dict.fromkeys(links), key=rank_internal_link)
    if len(ordered) < 8:
        ordered.extend(
            link for link in sorted(dict.fromkeys(fallback_links), key=rank_internal_link) if link not in ordered
        )
    return ordered[:10]


def extract_bridge_profile_links(base_url: str, html_text: str) -> list[str]:
    decoded_html = decode_embedded_urls(html_text)
    links: list[str] = []
    for href_match in HREF_RE.finditer(decoded_html):
        href = href_match.group("href")
        absolute = canonicalize_url(urljoin(base_url, href))
        if absolute and is_bridge_profile_url(absolute):
            links.append(absolute)
    return list(dict.fromkeys(links))[:3]


def instagram_from_homepage(session: requests.Session, homepage_url: str) -> str | None:
    root_page = fetch_html_page(session, homepage_url, timeout=8)
    if not root_page:
        return None
    queue: list[str] = []
    bridge_queue: list[str] = []
    visited_pages: set[str] = set()
    visited_bridges: set[str] = set()

    def inspect_page(page_url: str, html_text: str) -> str | None:
        direct_links = extract_instagram_candidates(page_url, html_text)
        if direct_links:
            return direct_links[0]
        for bridge_url in extract_bridge_profile_links(page_url, html_text):
            if bridge_url not in visited_bridges and bridge_url not in bridge_queue:
                bridge_queue.append(bridge_url)
        for internal_url in extract_internal_links(page_url, html_text):
            if internal_url not in visited_pages and internal_url not in queue:
                queue.append(internal_url)
        return None

    direct = inspect_page(*root_page)
    if direct:
        return direct

    while bridge_queue:
        bridge_url = bridge_queue.pop(0)
        if bridge_url in visited_bridges:
            continue
        visited_bridges.add(bridge_url)
        bridge_page = fetch_html_page(session, bridge_url, timeout=6)
        if not bridge_page:
            continue
        bridge_links = extract_instagram_candidates(*bridge_page)
        if bridge_links:
            return bridge_links[0]

    internal_checks = 0
    while queue and internal_checks < 8:
        internal_url = queue.pop(0)
        if internal_url in visited_pages:
            continue
        visited_pages.add(internal_url)
        internal_page = fetch_html_page(session, internal_url, timeout=6)
        if not internal_page:
            continue
        internal_checks += 1
        direct = inspect_page(*internal_page)
        if direct:
            return direct
        while bridge_queue:
            bridge_url = bridge_queue.pop(0)
            if bridge_url in visited_bridges:
                continue
            visited_bridges.add(bridge_url)
            bridge_page = fetch_html_page(session, bridge_url, timeout=6)
            if not bridge_page:
                continue
            bridge_links = extract_instagram_candidates(*bridge_page)
            if bridge_links:
                return bridge_links[0]
    sitemap_queue = [
        url
        for url in discover_sitemap_page_urls(session, homepage_url)
        if url not in visited_pages and url not in queue
    ]
    sitemap_checks = 0
    while sitemap_queue and sitemap_checks < 4:
        sitemap_url = sitemap_queue.pop(0)
        if sitemap_url in visited_pages:
            continue
        visited_pages.add(sitemap_url)
        sitemap_page = fetch_html_page(session, sitemap_url, timeout=5)
        if not sitemap_page:
            continue
        sitemap_checks += 1
        direct = inspect_page(*sitemap_page)
        if direct:
            return direct
        while bridge_queue:
            bridge_url = bridge_queue.pop(0)
            if bridge_url in visited_bridges:
                continue
            visited_bridges.add(bridge_url)
            bridge_page = fetch_html_page(session, bridge_url, timeout=6)
            if not bridge_page:
                continue
            bridge_links = extract_instagram_candidates(*bridge_page)
            if bridge_links:
                return bridge_links[0]
    return None


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
    pdf_homepages_by_number: dict[str, str],
    pdf_homepages_by_name: dict[str, str],
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

    pdf_homepage_url = homepage_from_pdf(record, pdf_homepages_by_number, pdf_homepages_by_name)
    if pdf_homepage_url and is_instagram_url(pdf_homepage_url):
        normalized_instagram = normalize_instagram_profile_url(pdf_homepage_url)
        if normalized_instagram:
            instagram_url = normalized_instagram
            instagram_source = "osaka_city_pdf_r6"
            instagram_confidence = "high"
    elif pdf_homepage_url and is_homepage_candidate(pdf_homepage_url):
        homepage_url = pdf_homepage_url
        homepage_source = "osaka_city_pdf_r6"
        homepage_confidence = "high"
    elif is_instagram_url(wam_office_url):
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
    if homepage_url and not instagram_url:
        instagram_url = instagram_from_homepage(session, homepage_url)
        if instagram_url:
            instagram_source = "homepage_crawl"
            instagram_confidence = "high"

    if not homepage_url or not instagram_url:
        searched = False
        if not homepage_url:
            search_query = build_query(record)
            results: list[dict[str, str]] = []
            if not SEARCH_DISABLED:
                searched = True
                try:
                    results = search_duckduckgo(session, search_query)
                except requests.RequestException:
                    results = []
            homepage_url, homepage_source, homepage_confidence = pick_best_result(
                record,
                results,
                score_homepage,
                6,
            )
            if homepage_url and not instagram_url:
                instagram_url = instagram_from_homepage(session, homepage_url)
                if instagram_url:
                    instagram_source = "homepage_crawl"
                    instagram_confidence = "high"
        if not instagram_url:
            search_query = build_instagram_query(record)
            insta_results: list[dict[str, str]] = []
            try:
                insta_results = search_bing_rss(session, search_query)
            except (requests.RequestException, ET.ParseError):
                insta_results = []
            instagram_url, instagram_source, instagram_confidence = pick_best_result(
                record,
                insta_results,
                score_instagram,
                5,
                url_transformer=normalize_instagram_profile_url,
            )
            if instagram_source == "ddg_search":
                instagram_source = "bing_rss_search"
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
    pdf_cache_path = resolve_path(args.osaka_city_pdf_cache)

    records = [
        record
        for record in load_dashboard_records(dashboard_input)
        if record.get("municipality") == args.municipality
    ]
    records.sort(key=lambda row: (str(row.get("office_no") or ""), str(row.get("office_name") or "")))
    if args.limit > 0:
        records = records[: args.limit]

    cached = load_existing_rows(output_path)
    pdf_rows = parse_osaka_city_pdf_homepages(ensure_osaka_city_pdf(pdf_cache_path))
    pdf_homepages_by_number, pdf_homepages_by_name = pdf_homepage_lookups(pdf_rows)

    def process_record(record: dict[str, Any]) -> dict[str, Any]:
        session = build_session()
        try:
            return build_row(
                record,
                session=session,
                sleep_seconds=args.sleep_seconds,
                cached_row=cached.get(str(record.get("office_no") or "").strip()),
                refresh=args.refresh,
                pdf_homepages_by_number=pdf_homepages_by_number,
                pdf_homepages_by_name=pdf_homepages_by_name,
            )
        finally:
            session.close()

    if args.workers <= 1:
        rows = [process_record(record) for record in records]
    else:
        with ThreadPoolExecutor(max_workers=args.workers) as executor:
            rows = list(executor.map(process_record, records))

    rows.sort(key=lambda row: (str(row.get("office_no") or ""), str(row.get("office_name") or "")))
    write_rows(output_path, rows)

    homepage_count = sum(1 for row in rows if row.get("homepage_url"))
    instagram_count = sum(1 for row in rows if row.get("instagram_url"))
    print(
        json.dumps(
            {
                "record_count": len(rows),
                "pdf_homepage_rows": len(pdf_rows),
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
