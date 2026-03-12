#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
from collections import Counter
from datetime import datetime
from pathlib import Path
import re
from statistics import median
from typing import Any
import unicodedata
import zipfile

import requests
from pypdf import PdfReader

from extract_r6kouchinjissekib import (
    NORMALIZED_FIELD_ORDER,
    enrich_records_with_analytics,
    write_dict_csv,
    write_json,
)


PDF_URL = (
    "https://www.pref.hokkaido.lg.jp/fs/1/2/9/0/9/0/1/8/_/"
    "%E5%B0%B1%E5%8A%B4%E7%B6%99%E7%B6%9A%E6%94%AF%E6%8F%B4B%E5%9E%8B%20"
    "%E5%B7%A5%E8%B3%83%E5%AE%9F%E7%B8%BE%E4%B8%80%E8%A6%A7.pdf"
)

WAGE_FORMS = [
    "月給＋日給＋時給",
    "日給＋時給",
    "月給＋日給",
    "月給＋時給",
    "時間給",
    "日給",
    "月給",
    "時給",
]
MUNICIPALITY_NORMALIZATION = {
    "中標津": "中標津町",
    "小樽": "小樽市",
    "岩見沢": "岩見沢市",
    "留萌": "留萌市",
    "釧路": "釧路市",
}

HEADER_LINES = {
    "就労継続支援B型",
    "○令和６年度工賃実績一覧表（実績0円除く）",
    "工賃支払総額 利用者延人数 年間開所日数 １日平均",
    "利用者数 年間開所月数 工賃平均額",
    "主な作業内容１ 主な作業内容２ 主な作業内容３施設種別 賃金（工賃）",
    "形態",
    "月額",
    "定員所在地",
    "（市町村）地区 事業所名",
}
LOCATION_SUFFIXES = ("市", "町", "村", "区", "郡", "振興局")
DISTRICT_TOKENS = [
    "札幌市",
    "函館市",
    "旭川市",
    "石狩",
    "空知",
    "後志",
    "胆振",
    "日高",
    "渡島",
    "檜山",
    "上川",
    "留萌",
    "宗谷",
    "オホーツク",
    "十勝",
    "釧路",
    "根室",
]

RECORD_PREFIX_RE = re.compile(r"^就労継続支援B型\s+\d{2}\s*")
RECORD_BODY_RE = re.compile(
    r"^(?P<office_name>.+?)\s+"
    r"(?P<capacity>\d+(?:名)?)\s+"
    r"(?P<wage_payment_total_yen>[\d,]+)\s+"
    r"(?P<user_days_total>[\d,]+)\s+"
    r"(?P<annual_open_days>\d+)\s+"
    r"(?P<average_daily_users>\d+(?:\.\d+)?)\s+"
    r"(?P<annual_open_months>\d+)\s+"
    r"(?P<average_wage_yen>[\d,]+(?:\.\d+)?)\s*"
    r"(?P<tail>.*)$"
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="北海道の就労継続支援B型 工賃実績PDFをダッシュボード用JSONへ変換する。"
    )
    parser.add_argument(
        "--input-pdf",
        default="data/inputs/hokkaido/hokkaido_r6_shuro_b_wages.pdf",
        help="source PDF path; downloaded automatically when missing",
    )
    parser.add_argument(
        "--export-root",
        default="data/exports/hokkaido_shuro_b",
        help="directory for generated exports, relative to repo root if not absolute",
    )
    parser.add_argument(
        "--download-url",
        default=PDF_URL,
        help="official PDF URL",
    )
    parser.add_argument(
        "--wam-open-data-zip",
        default="data/inputs/wam/sfkopendata_202503_46.zip",
        help="WAM open data zip path used to infer Hokkaido municipality names",
    )
    return parser.parse_args()


def repo_root() -> Path:
    return Path(__file__).resolve().parents[1]


def resolve_path(path_str: str) -> Path:
    path = Path(path_str).expanduser()
    if path.is_absolute():
        return path.resolve()
    return (repo_root() / path).resolve()


def ensure_pdf(path: Path, url: str) -> None:
    if path.exists():
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    response = requests.get(url, timeout=60)
    response.raise_for_status()
    path.write_bytes(response.content)


def parse_int(value: str) -> int:
    return int(value.replace(",", "").replace("名", ""))


def parse_float(value: str) -> float:
    return float(value.replace(",", ""))


def is_header_line(line: str) -> bool:
    return line in HEADER_LINES


def extract_record_lines(pdf_path: Path) -> list[str]:
    reader = PdfReader(str(pdf_path))
    raw_records: list[str] = []
    current: str | None = None

    for page in reader.pages:
        text = page.extract_text() or ""
        for original_line in text.splitlines():
            line = original_line.strip()
            if not line or is_header_line(line):
                continue
            if RECORD_PREFIX_RE.match(line):
                if current:
                    raw_records.append(current)
                current = line
                continue
            if current:
                current = f"{current} {line}"

    if current:
        raw_records.append(current)
    return raw_records


def normalize_tail_spacing(value: str) -> str:
    text = re.sub(r"\s*＋\s*", "＋", value.strip())
    return re.sub(r"\s+", " ", text).strip()


def extract_hokkaido_municipality(value: str | None) -> str | None:
    text = (value or "").replace("北海道", "").strip()
    text = re.sub(r"^[^\s]+郡", "", text)
    match = re.match(r"^([^\d\s]+?[市町村])", text)
    return match.group(1) if match else None


def load_known_municipalities(zip_path: Path) -> list[str]:
    municipalities: set[str] = set()
    with zipfile.ZipFile(zip_path) as zip_file:
        member = next((name for name in zip_file.namelist() if name.endswith(".csv")), None)
        if member is None:
            return []
        with zip_file.open(member) as file_obj:
            rows = csv.DictReader((line.decode("utf-8-sig") for line in file_obj))
            for row in rows:
                if row.get("サービス種別") != "就労継続支援Ｂ型":
                    continue
                address_city = row.get("事業所住所（市区町村）") or ""
                if not address_city.startswith("北海道"):
                    continue
                municipality = extract_hokkaido_municipality(address_city)
                if municipality:
                    municipalities.add(municipality)
    return sorted(municipalities, key=len, reverse=True)


def normalize_office_name(value: str | None) -> str:
    text = unicodedata.normalize("NFKC", value or "").lower().replace("　", " ")
    return re.sub(r"[\s\W_]+", "", text)


def load_hokkaido_open_data_rows(zip_path: Path) -> list[dict[str, str]]:
    rows_out: list[dict[str, str]] = []
    with zipfile.ZipFile(zip_path) as zip_file:
        member = next((name for name in zip_file.namelist() if name.endswith(".csv")), None)
        if member is None:
            return []
        with zip_file.open(member) as file_obj:
            rows = csv.DictReader((line.decode("utf-8-sig") for line in file_obj))
            for row in rows:
                if row.get("サービス種別") != "就労継続支援Ｂ型":
                    continue
                address_city = row.get("事業所住所（市区町村）") or ""
                if not address_city.startswith("北海道"):
                    continue
                rows_out.append(row)
    return rows_out


def enrich_records_from_open_data(
    records: list[dict[str, Any]],
    open_data_rows: list[dict[str, str]],
    known_municipalities: list[str],
) -> None:
    by_name: dict[str, list[dict[str, str]]] = {}
    for row in open_data_rows:
        key = normalize_office_name(row.get("事業所の名称"))
        if key:
            by_name.setdefault(key, []).append(row)

    known_set = set(known_municipalities)
    for record in records:
        matches = by_name.get(normalize_office_name(record.get("office_name")))
        if not matches or len(matches) != 1:
            normalized = MUNICIPALITY_NORMALIZATION.get(str(record.get("municipality")))
            if normalized:
                record["municipality"] = normalized
            continue
        match = matches[0]
        municipality = extract_hokkaido_municipality(match.get("事業所住所（市区町村）"))
        if municipality and (
            record.get("municipality") not in known_set
            or not str(record.get("municipality")).endswith(("市", "町", "村"))
        ):
            record["municipality"] = municipality
        record["corporation_name"] = match.get("法人の名称") or None
        record["corporation_number"] = match.get("法人番号") or None

    for record in records:
        normalized = MUNICIPALITY_NORMALIZATION.get(str(record.get("municipality")))
        if normalized:
            record["municipality"] = normalized


def split_municipality_and_body(
    remainder: str,
    known_municipalities: list[str],
) -> tuple[str, str]:
    remainder = re.sub(r"^[^\s]+郡", "", remainder)
    for municipality in known_municipalities:
        if remainder.startswith(municipality):
            return municipality, remainder[len(municipality) :].strip()
    parts = remainder.split(maxsplit=1)
    municipality = parts[0]
    body = parts[1].strip() if len(parts) > 1 else ""
    return municipality, body


def looks_like_location_token(value: str) -> bool:
    return any(value.endswith(suffix) for suffix in LOCATION_SUFFIXES)


def parse_record_line(
    line: str,
    index: int,
    known_municipalities: list[str],
) -> dict[str, Any]:
    prefix_match = re.match(r"^就労継続支援B型\s+(\d{2})\s*(.+)$", line)
    if not prefix_match:
        raise ValueError(f"record prefix parse failed: {line}")

    area_code = prefix_match.group(1)
    prefix_body = prefix_match.group(2).strip()
    prefix_parts = prefix_body.split(maxsplit=1)
    if (
        len(prefix_parts) >= 2
        and not prefix_parts[0] in known_municipalities
        and any(prefix_parts[0].startswith(token) for token in DISTRICT_TOKENS)
    ):
        token = next(token for token in DISTRICT_TOKENS if prefix_parts[0].startswith(token))
        prefix_parts = [token, prefix_body[len(token) :].strip()]
    if len(prefix_parts) < 2:
        raise ValueError(f"location parse failed: {line}")
    first_token, rest_after_first = prefix_parts[0], prefix_parts[1].strip()

    if first_token in known_municipalities:
        municipality = first_token
        rest_parts = rest_after_first.split(maxsplit=1)
        second_token = rest_parts[0]
        if looks_like_location_token(second_token):
            district = second_token
            remainder = rest_parts[1].strip() if len(rest_parts) > 1 else ""
        else:
            district = municipality
            remainder = rest_after_first
    else:
        district = first_token
        municipality, remainder = split_municipality_and_body(rest_after_first, known_municipalities)

    body_match = RECORD_BODY_RE.match(remainder)
    if not body_match:
        for anchor in [
            "就労継続支援",
            "就労支援",
            "多機能型",
            "ワーク",
            "工房",
            "作業所",
            "センター",
            "サポート",
            "障がい者",
        ]:
            anchor_index = remainder.find(anchor)
            if anchor_index > 0:
                body_match = RECORD_BODY_RE.match(remainder[anchor_index:].strip())
                if body_match:
                    break
    if not body_match:
        raise ValueError(f"record body parse failed: {line}")

    tail = normalize_tail_spacing(body_match.group("tail"))
    wage_form = None
    activities: list[str] = [item for item in tail.split(" ") if item]
    for candidate in WAGE_FORMS:
        if tail.startswith(candidate):
            wage_form = candidate
            tail_rest = tail[len(candidate) :].strip()
            activities = [item for item in tail_rest.split(" ") if item]
            break

    return {
        "source_row": index,
        "prefecture": "北海道",
        "office_no": index,
        "municipality": municipality,
        "corporation_type_code": None,
        "corporation_type_label": None,
        "corporation_number": None,
        "corporation_name": None,
        "office_name": body_match.group("office_name").strip(),
        "capacity": parse_int(body_match.group("capacity")),
        "wage_payment_total_yen": parse_int(body_match.group("wage_payment_total_yen")),
        "user_days_total": parse_int(body_match.group("user_days_total")),
        "annual_open_days": parse_int(body_match.group("annual_open_days")),
        "average_daily_users": parse_float(body_match.group("average_daily_users")),
        "average_daily_users_error": None,
        "annual_open_months": parse_int(body_match.group("annual_open_months")),
        "average_wage_yen": parse_float(body_match.group("average_wage_yen")),
        "average_wage_error": None,
        "is_new_office": False,
        "remarks": None,
        "response_status": "answered",
        "noufuku_active": None,
        "noufuku_new": None,
        "noufuku_income_ratio_decimal": None,
        "noufuku_income_ratio_pct": None,
        "suifuku_active": None,
        "suifuku_new": None,
        "suifuku_income_ratio_decimal": None,
        "suifuku_income_ratio_pct": None,
        "rinfuku_active": None,
        "rinfuku_new": None,
        "rinfuku_income_ratio_decimal": None,
        "rinfuku_income_ratio_pct": None,
        "home_use_active": None,
        "home_use_user_ratio_decimal": None,
        "home_use_user_ratio_pct": None,
        "area_code": area_code,
        "district": district,
        "wage_form_label": wage_form,
        "primary_work_items": activities,
    }


def build_payload(
    pdf_path: Path,
    records: list[dict[str, Any]],
    parse_issue_count: int,
) -> dict[str, Any]:
    analytics = enrich_records_with_analytics(records)
    answered_records = [record for record in records if record["response_status"] != "unanswered"]
    answered_wages = [
        float(record["average_wage_yen"])
        for record in answered_records
        if record["average_wage_yen"] is not None
    ]
    response_breakdown = Counter(record["response_status"] for record in records)
    outlier_breakdown = Counter(record["wage_outlier_flag"] or "none" for record in records)
    municipalities = sorted(
        {
            record["municipality"]
            for record in records
            if isinstance(record.get("municipality"), str) and record["municipality"]
        }
    )
    missing_counts = {
        field_name: sum(1 for record in records if record.get(field_name) is None)
        for field_name in NORMALIZED_FIELD_ORDER
        if field_name not in {"source_row", "response_status"}
    }
    service_counts = {
        "noufuku_active": sum(1 for record in records if record.get("noufuku_active") is True),
        "suifuku_active": sum(1 for record in records if record.get("suifuku_active") is True),
        "rinfuku_active": sum(1 for record in records if record.get("rinfuku_active") is True),
        "home_use_active": sum(1 for record in records if record.get("home_use_active") is True),
    }

    issues = [
        {
            "sheet": "北海道 就労継続支援B型 工賃実績一覧PDF",
            "kind": "source_pdf",
            "detail": "北海道の公式PDFを行単位で抽出し、ダッシュボード互換JSONへ正規化した。",
        }
    ]
    if parse_issue_count:
        issues.append(
            {
                "sheet": "北海道 就労継続支援B型 工賃実績一覧PDF",
                "kind": "pdf_parse_warnings",
                "count": parse_issue_count,
            }
        )

    notes = [
        {
            "source_row": 0,
            "note_text": "法人名、法人種別、在宅利用、農福連携などPDF非掲載項目は欠損として扱う。",
        }
    ]

    return {
        "meta": {
            "source_pdf": pdf_path.name,
            "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
            "record_count": len(records),
        },
        "summary": {
            "record_count": len(records),
            "note_row_count": len(notes),
            "response_breakdown": dict(response_breakdown),
            "answered_average_wage_yen_mean": round(sum(answered_wages) / len(answered_wages), 3)
            if answered_wages
            else None,
            "answered_average_wage_yen_median": round(median(answered_wages), 3)
            if answered_wages
            else None,
            "municipality_count": len(municipalities),
            "new_office_count": 0,
            "formula_error_rows": 0,
            "high_outlier_count": outlier_breakdown.get("high", 0),
            "low_outlier_count": outlier_breakdown.get("low", 0),
            "duplicate_office_no_count": 0,
            "service_counts": service_counts,
            "corporation_type_breakdown": {"unknown": len(records)},
            "missing_counts": missing_counts,
            "overall_wage_stats": analytics["overall_wage_stats"],
            "overall_utilization_stats": analytics["overall_utilization_stats"],
        },
        "issues": issues,
        "notes": notes,
        "lookups": {
            "corporation_type_lookup": [],
        },
        "analytics": analytics,
        "records": records,
    }


def main() -> None:
    args = parse_args()
    pdf_path = resolve_path(args.input_pdf)
    export_root = resolve_path(args.export_root)
    wam_zip_path = resolve_path(args.wam_open_data_zip)
    normalized_dir = export_root / "normalized"
    analytics_dir = export_root / "analytics"
    normalized_dir.mkdir(parents=True, exist_ok=True)
    analytics_dir.mkdir(parents=True, exist_ok=True)

    ensure_pdf(pdf_path, args.download_url)
    known_municipalities = load_known_municipalities(wam_zip_path)
    open_data_rows = load_hokkaido_open_data_rows(wam_zip_path)
    raw_lines = extract_record_lines(pdf_path)

    records: list[dict[str, Any]] = []
    parse_issues: list[dict[str, Any]] = []
    for index, line in enumerate(raw_lines, start=1):
        try:
            records.append(parse_record_line(line, index, known_municipalities))
        except Exception as error:
            parse_issues.append({"line_index": index, "raw_line": line, "error": str(error)})

    enrich_records_from_open_data(records, open_data_rows, known_municipalities)
    payload = build_payload(pdf_path, records, len(parse_issues))

    write_dict_csv(normalized_dir / "shuro_b_records.csv", list(records[0].keys()) if records else [], records)
    write_json(normalized_dir / "shuro_b_records.json", records)
    write_json(normalized_dir / "shuro_b_dashboard.json", payload)
    write_json(normalized_dir / "parse_issues.json", parse_issues)
    write_dict_csv(
        analytics_dir / "municipality_summary.csv",
        list(payload["analytics"]["municipality_summary"][0].keys()) if payload["analytics"]["municipality_summary"] else [],
        payload["analytics"]["municipality_summary"],
    )
    print(
        {
            "record_count": len(records),
            "parse_issue_count": len(parse_issues),
            "output": str((normalized_dir / "shuro_b_dashboard.json").relative_to(repo_root())),
        }
    )


if __name__ == "__main__":
    main()
