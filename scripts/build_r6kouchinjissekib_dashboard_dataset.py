#!/usr/bin/env python3
from __future__ import annotations

import argparse
from collections import Counter, defaultdict
from copy import deepcopy
from datetime import datetime
import csv
import difflib
import json
import math
from pathlib import Path
from statistics import median
from typing import Any
import unicodedata


WAM_EXTRA_FIELD_ORDER = [
    "wam_match_status",
    "wam_match_confidence",
    "wam_match_method",
    "wam_match_score",
    "wam_office_number",
    "wam_fetch_status",
    "wam_corporation_name",
    "wam_corporation_number",
    "wam_designated_agency",
    "wam_office_address_city",
    "wam_office_address_line",
    "wam_office_phone",
    "wam_office_url",
    "wam_office_capacity",
    "wam_capacity_gap",
    "wam_weekday_office_hours",
    "wam_saturday_office_hours",
    "wam_sunday_office_hours",
    "wam_holiday_office_hours",
    "wam_weekday_service_hours",
    "wam_saturday_service_hours",
    "wam_sunday_service_hours",
    "wam_holiday_service_hours",
    "wam_closed_days",
    "wam_transport_available",
    "wam_meal_support_addon",
    "wam_regional_collaboration_addon",
    "wam_medical_partner",
    "wam_primary_activity_type",
    "wam_primary_activity_detail",
    "wam_average_wage_monthly_yen",
    "wam_average_wage_hourly_yen",
    "wam_average_wage_gap_yen",
    "wam_annual_sales_yen",
    "wam_annual_costs_yen",
    "wam_annual_wage_payment_yen",
    "wam_full_time_weekly_hours",
    "wam_welfare_staff_fte_total",
    "wam_actual_user_count",
    "wam_users_per_staff_monthly",
    "wam_manager_multi_post",
    "wam_manager_qualified",
    "wam_manager_qualification_name",
    "wam_health_check_performed",
    "wam_training_plan_available",
    "wam_training_implemented",
    "wam_decision_support_training",
    "wam_abuse_prevention_training",
    "wam_service_manager_actual",
    "wam_service_manager_fte",
    "wam_employment_support_actual",
    "wam_employment_support_fte",
    "wam_vocational_instructor_actual",
    "wam_vocational_instructor_fte",
    "wam_life_support_actual",
    "wam_life_support_fte",
    "wam_key_staff_actual_total",
    "wam_key_staff_fte_total",
    "wam_key_staff_fte_per_capacity",
    "wam_key_staff_fte_per_average_daily_user",
    "wam_welfare_staff_fte_per_capacity",
    "wam_welfare_staff_fte_per_average_daily_user",
    "wam_care_worker_qualification_total",
    "wam_social_worker_qualification_total",
    "wam_mental_health_social_worker_qualification_total",
    "wam_psychologist_qualification_total",
    "wam_initial_training_qualification_total",
    "wam_practical_training_qualification_total",
    "wam_service_addon_count",
    "wam_hr_practice_count",
    "wam_support_category_total_users",
    "wam_support_category_none_users",
    "wam_support_category_1_users",
    "wam_support_category_2_users",
    "wam_support_category_3_users",
    "wam_support_category_4_users",
    "wam_support_category_5_users",
    "wam_support_category_6_users",
    "wam_medical_care_user_total",
    "wam_staffing_outlier_flag",
    "wam_staffing_outlier_severity",
    "wam_staffing_efficiency_quadrant",
]

WAM_DIRECTORY_FIELD_ORDER = [
    "office_number",
    "office_name",
    "corporation_name",
    "corporation_number",
    "designated_agency",
    "office_address_city",
    "office_address_line",
    "office_phone",
    "office_url",
    "office_capacity",
    "transport_available",
    "meal_support_addon",
    "regional_collaboration_addon",
    "primary_activity_type",
    "primary_activity_detail",
    "average_wage_monthly_yen",
    "average_wage_hourly_yen",
    "welfare_staff_fte_total",
    "key_staff_fte_total",
    "key_staff_fte_per_capacity",
    "service_addon_count",
    "hr_practice_count",
    "fetch_status",
    "fetch_error",
    "fetched_at",
]

APP_RECORD_FIELD_ORDER = [
    "office_no",
    "municipality",
    "corporation_type_label",
    "corporation_name",
    "office_name",
    "response_status",
    "remarks",
    "capacity",
    "average_daily_users",
    "average_wage_yen",
    "average_wage_error",
    "daily_user_capacity_ratio",
    "wage_ratio_to_overall_mean",
    "wage_ratio_to_municipality_mean",
    "wage_ratio_to_corporation_type_mean",
    "wage_ratio_to_capacity_band_mean",
    "wage_z_score",
    "wage_outlier_flag",
    "wage_outlier_severity",
    "wage_band_label",
    "capacity_band_label",
    "market_position_quadrant",
    "is_new_office",
    "home_use_active",
    "home_use_user_ratio_decimal",
    "home_use_user_ratio_pct",
    "noufuku_active",
    "wam_match_status",
    "wam_match_confidence",
    "wam_fetch_status",
    "wam_office_number",
    "wam_office_url",
    "wam_office_phone",
    "wam_office_address_city",
    "wam_office_address_line",
    "wam_office_capacity",
    "wam_primary_activity_type",
    "wam_primary_activity_detail",
    "wam_transport_available",
    "wam_meal_support_addon",
    "wam_regional_collaboration_addon",
    "wam_manager_multi_post",
    "wam_welfare_staff_fte_total",
    "wam_key_staff_fte_per_capacity",
    "wam_staffing_efficiency_quadrant",
    "wam_staffing_outlier_flag",
    "wam_staffing_outlier_severity",
    "wam_average_wage_monthly_yen",
    "wam_average_wage_gap_yen",
    "wam_service_manager_fte",
]

APP_RECORD_CHUNK_SIZE = 250


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Excel 工賃データと WAM 大阪市B型詳細を統合し、ダッシュボード用 JSON を生成する。"
    )
    parser.add_argument(
        "--dashboard-input",
        default="data/exports/r6kouchinjissekib/normalized/shuro_b_dashboard.json",
        help="base dashboard json from extract script",
    )
    parser.add_argument(
        "--wam-details-input",
        default="data/exports/wam/osakashi_shuro_b/details/osakashi_shuro_b_details.json",
        help="WAM details json path",
    )
    parser.add_argument(
        "--wam-match-override-input",
        default="data/inputs/wam/osakashi_shuro_b_match_overrides.csv",
        help="manual workbook office_no to WAM office_number overrides",
    )
    parser.add_argument(
        "--integrated-output",
        default="data/exports/r6kouchinjissekib/integrated/shuro_b_dashboard_integrated.json",
        help="integrated dashboard json output path",
    )
    parser.add_argument(
        "--records-output-json",
        default="data/exports/r6kouchinjissekib/integrated/shuro_b_osakashi_wam_enriched_records.json",
        help="integrated record json output path",
    )
    parser.add_argument(
        "--records-output-csv",
        default="data/exports/r6kouchinjissekib/integrated/shuro_b_osakashi_wam_enriched_records.csv",
        help="integrated record csv output path",
    )
    parser.add_argument(
        "--match-report-output",
        default="data/exports/r6kouchinjissekib/integrated/shuro_b_osakashi_wam_match_report.csv",
        help="match report csv output path",
    )
    parser.add_argument(
        "--app-output",
        default="apps/r6kouchin-dashboard/data/dashboard-data.json",
        help="dashboard app data output path",
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


def load_csv_rows(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []
    with path.open(newline="", encoding="utf-8") as file_obj:
        return list(csv.DictReader(file_obj))


def write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def write_dict_csv(path: Path, fieldnames: list[str], rows: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as file_obj:
        writer = csv.DictWriter(file_obj, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def chunked_rows(rows: list[dict[str, Any]], chunk_size: int) -> list[list[dict[str, Any]]]:
    return [rows[index : index + chunk_size] for index in range(0, len(rows), chunk_size)]


def slim_records_for_app(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [{field: record.get(field) for field in APP_RECORD_FIELD_ORDER} for record in records]


def normalize_name(value: str | None, relaxed: bool = False) -> str:
    text = unicodedata.normalize("NFKC", value or "").lower().replace("　", " ")
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
        "就労継続支援b型",
        "就労継続支援b",
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


def similarity(left: str, right: str) -> float:
    if not left or not right:
        return 0.0
    return difflib.SequenceMatcher(None, left, right).ratio()


def common_suffix_length(left: str, right: str) -> int:
    if not left or not right:
        return 0
    length = 0
    for left_char, right_char in zip(reversed(left), reversed(right)):
        if left_char != right_char:
            break
        length += 1
    return length


def is_numeric(value: Any) -> bool:
    return isinstance(value, (int, float)) and not math.isnan(value)


def round_nullable(value: float | None, digits: int = 3) -> float | None:
    if value is None:
        return None
    return round(value, digits)


def percentile_interpolated(sorted_values: list[float], percentile: float) -> float | None:
    if not sorted_values:
        return None
    if len(sorted_values) == 1:
        return float(sorted_values[0])
    position = (len(sorted_values) - 1) * (percentile / 100)
    lower_index = math.floor(position)
    upper_index = math.ceil(position)
    lower_value = sorted_values[lower_index]
    upper_value = sorted_values[upper_index]
    if lower_index == upper_index:
        return float(lower_value)
    fraction = position - lower_index
    return float(lower_value + (upper_value - lower_value) * fraction)


def compute_numeric_stats(values: list[float]) -> dict[str, float | int | None]:
    sorted_values = sorted(float(value) for value in values)
    if not sorted_values:
        return {
            "count": 0,
            "mean": None,
            "median": None,
            "q1": None,
            "q3": None,
            "p90": None,
            "iqr": None,
            "lower_fence": None,
            "upper_fence": None,
            "extreme_lower_fence": None,
            "extreme_upper_fence": None,
        }
    mean_value = sum(sorted_values) / len(sorted_values)
    q1 = percentile_interpolated(sorted_values, 25)
    q3 = percentile_interpolated(sorted_values, 75)
    iqr = q3 - q1 if q1 is not None and q3 is not None else None
    return {
        "count": len(sorted_values),
        "mean": round_nullable(mean_value),
        "median": round_nullable(float(median(sorted_values))),
        "q1": round_nullable(q1),
        "q3": round_nullable(q3),
        "p90": round_nullable(percentile_interpolated(sorted_values, 90)),
        "iqr": round_nullable(iqr),
        "lower_fence": round_nullable(max((q1 or 0) - 1.5 * (iqr or 0), 0)) if q1 is not None and iqr is not None else None,
        "upper_fence": round_nullable((q3 or 0) + 1.5 * (iqr or 0)) if q3 is not None and iqr is not None else None,
        "extreme_lower_fence": round_nullable(max((q1 or 0) - 3 * (iqr or 0), 0)) if q1 is not None and iqr is not None else None,
        "extreme_upper_fence": round_nullable((q3 or 0) + 3 * (iqr or 0)) if q3 is not None and iqr is not None else None,
    }


def classify_outlier(value: float | None, stats: dict[str, float | int | None]) -> tuple[str | None, str | None]:
    if value is None:
        return None, None
    lower = stats.get("lower_fence")
    upper = stats.get("upper_fence")
    extreme_lower = stats.get("extreme_lower_fence")
    extreme_upper = stats.get("extreme_upper_fence")
    if extreme_upper is not None and value >= float(extreme_upper):
        return "high", "extreme"
    if upper is not None and value > float(upper):
        return "high", "moderate"
    if extreme_lower is not None and value <= float(extreme_lower):
        return "low", "extreme"
    if lower is not None and value < float(lower):
        return "low", "moderate"
    return None, None


def dedupe_wam_rows(wam_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    seen_office_numbers: set[str] = set()
    deduped: list[dict[str, Any]] = []
    for row in wam_rows:
        office_number = str(row.get("office_number") or "").strip()
        if office_number and office_number in seen_office_numbers:
            continue
        if office_number:
            seen_office_numbers.add(office_number)
        deduped.append(row)
    return deduped


def staffing_quadrant(
    wage_value: float | None,
    staffing_value: float | None,
    wage_threshold: float | None,
    staffing_threshold: float | None,
) -> str | None:
    if None in {wage_value, staffing_value, wage_threshold, staffing_threshold}:
        return None
    if wage_value >= wage_threshold and staffing_value >= staffing_threshold:
        return "高工賃 × 厚い人員"
    if wage_value >= wage_threshold and staffing_value < staffing_threshold:
        return "高工賃 × 少ない人員"
    if wage_value < wage_threshold and staffing_value >= staffing_threshold:
        return "低工賃 × 厚い人員"
    return "低工賃 × 少ない人員"


def build_match_proposals(
    records: list[dict[str, Any]],
    wam_rows: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    wam_by_corporation: dict[str, list[tuple[int, dict[str, Any]]]] = defaultdict(list)
    wam_by_corporation_name: dict[str, list[tuple[int, dict[str, Any]]]] = defaultdict(list)
    wam_strict_name_counts = Counter(normalize_name(row.get("office_name")) for row in wam_rows)
    wam_relaxed_name_counts = Counter(normalize_name(row.get("office_name"), relaxed=True) for row in wam_rows)

    for wam_index, wam_row in enumerate(wam_rows):
        corporation_number = str(wam_row.get("corporation_number") or "").strip()
        corporation_name = normalize_name(wam_row.get("corporation_name"))
        if corporation_number:
            wam_by_corporation[corporation_number].append((wam_index, wam_row))
        if corporation_name:
            wam_by_corporation_name[corporation_name].append((wam_index, wam_row))

    proposals: list[dict[str, Any]] = []
    for record_index, record in enumerate(records):
        if record.get("municipality") != "大阪市":
            continue

        strict_record_name = normalize_name(record.get("office_name"))
        relaxed_record_name = normalize_name(record.get("office_name"), relaxed=True)
        strict_record_corporation = normalize_name(record.get("corporation_name"))
        corporation_number = str(record.get("corporation_number") or "").strip()
        capacity = record.get("capacity")

        corporation_candidates = wam_by_corporation.get(corporation_number, [])
        corporation_name_candidates = (
            wam_by_corporation_name.get(strict_record_corporation, []) if strict_record_corporation else []
        )
        search_space_map: dict[int, dict[str, Any]] = {}
        for wam_index, wam_row in corporation_candidates:
            search_space_map[wam_index] = wam_row
        for wam_index, wam_row in corporation_name_candidates:
            search_space_map.setdefault(wam_index, wam_row)

        if search_space_map:
            search_space = sorted(search_space_map.items(), key=lambda item: item[0])
            if corporation_candidates and corporation_name_candidates:
                corporation_mode = "corp_number+corp_name"
            elif corporation_candidates:
                corporation_mode = "corp_number"
            else:
                corporation_mode = "corp_name"
        else:
            strict_unique_matches = [
                (wam_index, wam_row)
                for wam_index, wam_row in enumerate(wam_rows)
                if strict_record_name
                and strict_record_name == normalize_name(wam_row.get("office_name"))
                and wam_strict_name_counts[strict_record_name] == 1
            ]
            relaxed_unique_matches = [
                (wam_index, wam_row)
                for wam_index, wam_row in enumerate(wam_rows)
                if relaxed_record_name
                and relaxed_record_name == normalize_name(wam_row.get("office_name"), relaxed=True)
                and wam_relaxed_name_counts[relaxed_record_name] == 1
            ]
            search_space = strict_unique_matches or relaxed_unique_matches
            corporation_mode = "name_only"

        for wam_index, wam_row in search_space:
            strict_wam_name = normalize_name(wam_row.get("office_name"))
            relaxed_wam_name = normalize_name(wam_row.get("office_name"), relaxed=True)
            strict_similarity = similarity(strict_record_name, strict_wam_name)
            relaxed_similarity = similarity(relaxed_record_name, relaxed_wam_name)
            suffix_length = common_suffix_length(relaxed_record_name, relaxed_wam_name)
            strict_exact = strict_record_name and strict_record_name == strict_wam_name
            strict_substring = bool(
                strict_record_name and strict_wam_name and (
                    strict_record_name in strict_wam_name or strict_wam_name in strict_record_name
                )
            )
            relaxed_exact = relaxed_record_name and relaxed_record_name == relaxed_wam_name
            relaxed_substring = bool(
                relaxed_record_name and relaxed_wam_name and (
                    relaxed_record_name in relaxed_wam_name or relaxed_wam_name in relaxed_record_name
                )
            )
            corporation_exact = bool(
                corporation_number
                and corporation_number == str(wam_row.get("corporation_number") or "").strip()
            )
            corporation_name_exact = strict_record_corporation and (
                strict_record_corporation == normalize_name(wam_row.get("corporation_name"))
            )
            wam_capacity = wam_row.get("office_capacity")
            capacity_exact = (
                is_numeric(capacity)
                and is_numeric(wam_capacity)
                and int(float(capacity)) == int(float(wam_capacity))
            )
            capacity_close = (
                is_numeric(capacity)
                and is_numeric(wam_capacity)
                and abs(int(float(capacity)) - int(float(wam_capacity))) <= 2
            )

            confidence: str | None = None
            method = ""
            score = 0.0

            if corporation_exact and strict_exact:
                confidence = "high"
                method = "corp_exact_name_strict"
                score = 100 + strict_similarity + (0.15 if capacity_exact else 0)
            elif corporation_exact and (strict_substring or strict_similarity >= 0.92):
                confidence = "high"
                method = "corp_exact_name_strict"
                score = 98 + strict_similarity + (0.15 if capacity_exact else 0)
            elif corporation_exact and (relaxed_exact or relaxed_substring) and (
                capacity_exact or len(corporation_candidates) == 1
            ):
                confidence = "medium"
                method = "corp_exact_name_relaxed"
                score = 85 + relaxed_similarity + (0.15 if capacity_exact else 0)
            elif corporation_exact and strict_similarity >= 0.78 and relaxed_similarity >= 0.88:
                confidence = "medium"
                method = "corp_exact_name_fuzzy"
                score = 80 + strict_similarity + 0.5 * relaxed_similarity + (0.1 if capacity_close else 0)
            elif corporation_exact and suffix_length >= 2:
                confidence = "medium"
                method = "corp_exact_suffix_match"
                score = 74 + suffix_length + 0.5 * strict_similarity + (0.15 if capacity_exact else 0)
            elif corporation_name_exact and strict_exact:
                confidence = "high"
                method = "corp_name_exact_name_strict"
                score = 72 + strict_similarity + (0.1 if capacity_exact else 0)
            elif corporation_name_exact and relaxed_exact:
                confidence = "medium"
                method = "corp_name_exact_name_relaxed"
                score = 68 + relaxed_similarity + (0.1 if capacity_exact else 0)
            elif corporation_name_exact and suffix_length >= 2:
                confidence = "medium"
                method = "corp_name_exact_suffix_match"
                score = 64 + suffix_length + 0.5 * strict_similarity + (0.1 if capacity_exact else 0)
            elif strict_exact and wam_strict_name_counts[strict_record_name] == 1:
                confidence = "medium"
                method = "unique_name_strict"
                score = 70 + strict_similarity + (0.05 if corporation_name_exact else 0)
            elif relaxed_exact and wam_relaxed_name_counts[relaxed_record_name] == 1 and strict_similarity >= 0.72:
                confidence = "low"
                method = "unique_name_relaxed"
                score = 60 + strict_similarity + (0.1 if capacity_exact else 0)

            if confidence is None:
                continue

            proposals.append(
                {
                    "record_index": record_index,
                    "wam_index": wam_index,
                    "confidence": confidence,
                    "method": method,
                    "score": round(score, 4),
                    "strict_similarity": round(strict_similarity, 4),
                    "relaxed_similarity": round(relaxed_similarity, 4),
                    "suffix_length": suffix_length,
                    "capacity_exact": capacity_exact,
                    "capacity_close": capacity_close,
                    "corporation_mode": corporation_mode,
                }
            )
    return proposals


def build_override_proposals(
    records: list[dict[str, Any]],
    wam_rows: list[dict[str, Any]],
    override_rows: list[dict[str, str]],
) -> list[dict[str, Any]]:
    record_index_by_office_no = {
        str(record.get("office_no")): record_index
        for record_index, record in enumerate(records)
        if record.get("municipality") == "大阪市"
    }
    wam_index_by_office_number = {
        str(wam_row.get("office_number")): wam_index
        for wam_index, wam_row in enumerate(wam_rows)
        if wam_row.get("office_number")
    }
    proposals: list[dict[str, Any]] = []
    for override in override_rows:
        record_office_no = str(override.get("record_office_no") or "").strip()
        wam_office_number = str(override.get("wam_office_number") or "").strip()
        override_reason = str(override.get("reason") or "").strip() or None
        record_index = record_index_by_office_no.get(record_office_no)
        wam_index = wam_index_by_office_number.get(wam_office_number)
        if record_index is None or wam_index is None:
            continue
        proposals.append(
            {
                "record_index": record_index,
                "wam_index": wam_index,
                "confidence": "high",
                "method": "manual_override",
                "score": 1000.0,
                "strict_similarity": 1.0,
                "relaxed_similarity": 1.0,
                "suffix_length": 0,
                "capacity_exact": False,
                "capacity_close": False,
                "corporation_mode": "manual_override",
                "override_reason": override_reason,
            }
        )
    return proposals


def assign_matches(
    records: list[dict[str, Any]],
    wam_rows: list[dict[str, Any]],
    override_rows: list[dict[str, str]] | None = None,
) -> tuple[dict[int, dict[str, Any]], list[dict[str, Any]]]:
    proposals = build_match_proposals(records, wam_rows)
    proposals.extend(build_override_proposals(records, wam_rows, override_rows or []))
    confidence_priority = {"high": 3, "medium": 2, "low": 1}
    proposals_by_record: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for proposal in proposals:
        proposals_by_record[proposal["record_index"]].append(proposal)
    for proposal_list in proposals_by_record.values():
        proposal_list.sort(
            key=lambda proposal: (
                confidence_priority[proposal["confidence"]],
                proposal["score"],
                proposal["strict_similarity"],
                proposal["relaxed_similarity"],
            ),
            reverse=True,
        )

    proposals.sort(
        key=lambda proposal: (
            confidence_priority[proposal["confidence"]],
            proposal["score"],
            proposal["strict_similarity"],
            proposal["relaxed_similarity"],
        ),
        reverse=True,
    )
    assigned_records: set[int] = set()
    assigned_wam: set[int] = set()
    chosen: dict[int, dict[str, Any]] = {}
    match_rows: list[dict[str, Any]] = []
    for proposal in proposals:
        record_index = proposal["record_index"]
        wam_index = proposal["wam_index"]
        candidates = proposals_by_record[record_index]
        best_score = candidates[0]["score"]
        second_score = candidates[1]["score"] if len(candidates) > 1 else None
        ambiguous_tie = second_score is not None and abs(best_score - second_score) < 0.25
        if ambiguous_tie and proposal["score"] == best_score and proposal["method"] not in {
            "corp_exact_name_strict",
            "corp_name_exact_name_strict",
        }:
            continue
        if record_index in assigned_records or wam_index in assigned_wam:
            continue
        assigned_records.add(record_index)
        assigned_wam.add(wam_index)
        chosen[record_index] = proposal
        record = records[record_index]
        wam_row = wam_rows[wam_index]
        match_rows.append(
            {
                "record_index": record_index,
                "source_row": record.get("source_row"),
                "office_no": record.get("office_no"),
                "municipality": record.get("municipality"),
                "record_corporation_name": record.get("corporation_name"),
                "record_office_name": record.get("office_name"),
                "record_capacity": record.get("capacity"),
                "wam_office_number": wam_row.get("office_number"),
                "wam_corporation_name": wam_row.get("corporation_name"),
                "wam_office_name": wam_row.get("office_name"),
                "wam_capacity": wam_row.get("office_capacity"),
                "confidence": proposal["confidence"],
                "method": proposal["method"],
                "score": proposal["score"],
                "strict_similarity": proposal["strict_similarity"],
                "relaxed_similarity": proposal["relaxed_similarity"],
                "suffix_length": proposal["suffix_length"],
                "capacity_exact": proposal["capacity_exact"],
                "capacity_close": proposal["capacity_close"],
                "override_reason": proposal.get("override_reason"),
            }
        )
    return chosen, match_rows


def merge_wam_fields(record: dict[str, Any], wam_row: dict[str, Any] | None, match: dict[str, Any] | None) -> dict[str, Any]:
    merged = deepcopy(record)
    merged.update({field: None for field in WAM_EXTRA_FIELD_ORDER})

    if wam_row is None or match is None:
        merged["wam_match_status"] = "unmatched"
        return merged

    average_daily_users = record.get("average_daily_users")
    wam_key_staff_fte_total = wam_row.get("key_staff_fte_total")
    wam_welfare_staff_fte_total = wam_row.get("welfare_staff_fte_total")
    wam_capacity = wam_row.get("office_capacity")

    merged.update(
        {
            "wam_match_status": "matched",
            "wam_match_confidence": match["confidence"],
            "wam_match_method": match["method"],
            "wam_match_score": match["score"],
            "wam_office_number": wam_row.get("office_number"),
            "wam_fetch_status": wam_row.get("fetch_status"),
            "wam_corporation_name": wam_row.get("corporation_name"),
            "wam_corporation_number": wam_row.get("corporation_number"),
            "wam_designated_agency": wam_row.get("designated_agency"),
            "wam_office_address_city": wam_row.get("office_address_city"),
            "wam_office_address_line": wam_row.get("office_address_line"),
            "wam_office_phone": wam_row.get("office_phone"),
            "wam_office_url": wam_row.get("office_url"),
            "wam_office_capacity": wam_capacity,
            "wam_capacity_gap": round_nullable(float(wam_capacity) - float(record["capacity"]))
            if is_numeric(wam_capacity) and is_numeric(record.get("capacity"))
            else None,
            "wam_weekday_office_hours": wam_row.get("weekday_office_hours"),
            "wam_saturday_office_hours": wam_row.get("saturday_office_hours"),
            "wam_sunday_office_hours": wam_row.get("sunday_office_hours"),
            "wam_holiday_office_hours": wam_row.get("holiday_office_hours"),
            "wam_weekday_service_hours": wam_row.get("weekday_service_hours"),
            "wam_saturday_service_hours": wam_row.get("saturday_service_hours"),
            "wam_sunday_service_hours": wam_row.get("sunday_service_hours"),
            "wam_holiday_service_hours": wam_row.get("holiday_service_hours"),
            "wam_closed_days": wam_row.get("closed_days"),
            "wam_transport_available": wam_row.get("transport_available"),
            "wam_meal_support_addon": wam_row.get("meal_support_addon"),
            "wam_regional_collaboration_addon": wam_row.get("regional_collaboration_addon"),
            "wam_medical_partner": wam_row.get("medical_partner"),
            "wam_primary_activity_type": wam_row.get("primary_activity_type"),
            "wam_primary_activity_detail": wam_row.get("primary_activity_detail"),
            "wam_average_wage_monthly_yen": wam_row.get("average_wage_monthly_yen"),
            "wam_average_wage_hourly_yen": wam_row.get("average_wage_hourly_yen"),
            "wam_average_wage_gap_yen": round_nullable(
                float(wam_row["average_wage_monthly_yen"]) - float(record["average_wage_yen"])
            )
            if is_numeric(wam_row.get("average_wage_monthly_yen")) and is_numeric(record.get("average_wage_yen"))
            else None,
            "wam_annual_sales_yen": wam_row.get("annual_sales_yen"),
            "wam_annual_costs_yen": wam_row.get("annual_costs_yen"),
            "wam_annual_wage_payment_yen": wam_row.get("annual_wage_payment_yen"),
            "wam_full_time_weekly_hours": wam_row.get("full_time_weekly_hours"),
            "wam_welfare_staff_fte_total": wam_welfare_staff_fte_total,
            "wam_actual_user_count": wam_row.get("actual_user_count"),
            "wam_users_per_staff_monthly": wam_row.get("users_per_staff_monthly"),
            "wam_manager_multi_post": wam_row.get("manager_multi_post"),
            "wam_manager_qualified": wam_row.get("manager_qualified"),
            "wam_manager_qualification_name": wam_row.get("manager_qualification_name"),
            "wam_health_check_performed": wam_row.get("health_check_performed"),
            "wam_training_plan_available": wam_row.get("training_plan_available"),
            "wam_training_implemented": wam_row.get("training_implemented"),
            "wam_decision_support_training": wam_row.get("decision_support_training"),
            "wam_abuse_prevention_training": wam_row.get("abuse_prevention_training"),
            "wam_service_manager_actual": wam_row.get("service_manager_actual"),
            "wam_service_manager_fte": wam_row.get("service_manager_fte"),
            "wam_employment_support_actual": wam_row.get("employment_support_actual"),
            "wam_employment_support_fte": wam_row.get("employment_support_fte"),
            "wam_vocational_instructor_actual": wam_row.get("vocational_instructor_actual"),
            "wam_vocational_instructor_fte": wam_row.get("vocational_instructor_fte"),
            "wam_life_support_actual": wam_row.get("life_support_actual"),
            "wam_life_support_fte": wam_row.get("life_support_fte"),
            "wam_key_staff_actual_total": wam_row.get("key_staff_actual_total"),
            "wam_key_staff_fte_total": wam_key_staff_fte_total,
            "wam_key_staff_fte_per_capacity": wam_row.get("key_staff_fte_per_capacity"),
            "wam_key_staff_fte_per_average_daily_user": round_nullable(
                float(wam_key_staff_fte_total) / float(average_daily_users)
            )
            if is_numeric(wam_key_staff_fte_total) and is_numeric(average_daily_users) and float(average_daily_users) > 0
            else None,
            "wam_welfare_staff_fte_per_capacity": wam_row.get("welfare_staff_fte_per_capacity"),
            "wam_welfare_staff_fte_per_average_daily_user": round_nullable(
                float(wam_welfare_staff_fte_total) / float(average_daily_users)
            )
            if is_numeric(wam_welfare_staff_fte_total) and is_numeric(average_daily_users) and float(average_daily_users) > 0
            else None,
            "wam_care_worker_qualification_total": wam_row.get("care_worker_qualification_total"),
            "wam_social_worker_qualification_total": wam_row.get("social_worker_qualification_total"),
            "wam_mental_health_social_worker_qualification_total": wam_row.get("mental_health_social_worker_qualification_total"),
            "wam_psychologist_qualification_total": wam_row.get("psychologist_qualification_total"),
            "wam_initial_training_qualification_total": wam_row.get("initial_training_qualification_total"),
            "wam_practical_training_qualification_total": wam_row.get("practical_training_qualification_total"),
            "wam_service_addon_count": wam_row.get("service_addon_count"),
            "wam_hr_practice_count": wam_row.get("hr_practice_count"),
            "wam_support_category_total_users": wam_row.get("support_category_total_users"),
            "wam_support_category_none_users": wam_row.get("support_category_none_users"),
            "wam_support_category_1_users": wam_row.get("support_category_1_users"),
            "wam_support_category_2_users": wam_row.get("support_category_2_users"),
            "wam_support_category_3_users": wam_row.get("support_category_3_users"),
            "wam_support_category_4_users": wam_row.get("support_category_4_users"),
            "wam_support_category_5_users": wam_row.get("support_category_5_users"),
            "wam_support_category_6_users": wam_row.get("support_category_6_users"),
            "wam_medical_care_user_total": wam_row.get("medical_care_user_total"),
        }
    )
    return merged


def summarize_wam_roles(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    role_rows = [
        ("サービス管理責任者", "wam_service_manager_actual", "wam_service_manager_fte"),
        ("就労支援員", "wam_employment_support_actual", "wam_employment_support_fte"),
        ("職業指導員", "wam_vocational_instructor_actual", "wam_vocational_instructor_fte"),
        ("生活支援員", "wam_life_support_actual", "wam_life_support_fte"),
    ]
    summary: list[dict[str, Any]] = []
    for label, actual_key, fte_key in role_rows:
        actual_values = [float(record[actual_key]) for record in records if is_numeric(record.get(actual_key))]
        fte_values = [float(record[fte_key]) for record in records if is_numeric(record.get(fte_key))]
        summary.append(
            {
                "role_label": label,
                "office_count": len(fte_values),
                "average_actual": round_nullable(sum(actual_values) / len(actual_values)) if actual_values else None,
                "average_fte": round_nullable(sum(fte_values) / len(fte_values)) if fte_values else None,
                "nonzero_office_count": sum(1 for value in fte_values if value > 0),
            }
        )
    return summary


def summarize_feature_effects(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    feature_rows = [
        ("送迎あり", "wam_transport_available"),
        ("食事提供体制加算あり", "wam_meal_support_addon"),
        ("地域協働加算あり", "wam_regional_collaboration_addon"),
        ("管理者兼務あり", "wam_manager_multi_post"),
    ]
    summary: list[dict[str, Any]] = []
    for label, key in feature_rows:
        true_records = [record for record in records if record.get(key) is True]
        false_records = [record for record in records if record.get(key) is False]
        true_wages = [float(record["average_wage_yen"]) for record in true_records if is_numeric(record.get("average_wage_yen"))]
        false_wages = [float(record["average_wage_yen"]) for record in false_records if is_numeric(record.get("average_wage_yen"))]
        true_staffing = [
            float(record["wam_key_staff_fte_per_capacity"])
            for record in true_records
            if is_numeric(record.get("wam_key_staff_fte_per_capacity"))
        ]
        summary.append(
            {
                "feature_label": label,
                "true_count": len(true_records),
                "false_count": len(false_records),
                "true_average_wage_yen": round_nullable(sum(true_wages) / len(true_wages)) if true_wages else None,
                "false_average_wage_yen": round_nullable(sum(false_wages) / len(false_wages)) if false_wages else None,
                "true_average_key_staff_fte_per_capacity": round_nullable(
                    sum(true_staffing) / len(true_staffing)
                )
                if true_staffing
                else None,
            }
        )
    return summary


def build_watchlists(records: list[dict[str, Any]]) -> dict[str, list[dict[str, Any]]]:
    matched = [record for record in records if record.get("wam_match_status") == "matched"]
    lean_high = [
        record
        for record in matched
        if record.get("wam_staffing_efficiency_quadrant") == "高工賃 × 少ない人員"
    ]
    heavy_low = [
        record
        for record in matched
        if record.get("wam_staffing_efficiency_quadrant") == "低工賃 × 厚い人員"
    ]
    lean_high.sort(key=lambda row: (-(row.get("average_wage_yen") or 0), row.get("wam_key_staff_fte_per_capacity") or 999))
    heavy_low.sort(key=lambda row: ((row.get("average_wage_yen") or 999999), -(row.get("wam_key_staff_fte_per_capacity") or 0)))
    picked_fields = [
        "office_no",
        "municipality",
        "corporation_name",
        "office_name",
        "average_wage_yen",
        "daily_user_capacity_ratio",
        "wam_key_staff_fte_per_capacity",
        "wam_transport_available",
        "wam_primary_activity_type",
    ]
    return {
        "lean_high_wage": [{field: row.get(field) for field in picked_fields} for row in lean_high[:12]],
        "heavy_low_wage": [{field: row.get(field) for field in picked_fields} for row in heavy_low[:12]],
    }


def enrich_staffing_analytics(records: list[dict[str, Any]]) -> dict[str, Any]:
    matched_records = [
        record
        for record in records
        if record.get("wam_match_status") == "matched" and record.get("wam_fetch_status") == "ok"
    ]
    staffing_values = [
        float(record["wam_key_staff_fte_per_capacity"])
        for record in matched_records
        if is_numeric(record.get("wam_key_staff_fte_per_capacity"))
    ]
    wage_values = [
        float(record["average_wage_yen"])
        for record in matched_records
        if is_numeric(record.get("average_wage_yen"))
    ]
    staffing_stats = compute_numeric_stats(staffing_values)
    wage_stats = compute_numeric_stats(wage_values)
    staffing_median = staffing_stats["median"]
    wage_median = wage_stats["median"]

    for record in records:
        staffing_value = float(record["wam_key_staff_fte_per_capacity"]) if is_numeric(record.get("wam_key_staff_fte_per_capacity")) else None
        outlier_flag, outlier_severity = classify_outlier(staffing_value, staffing_stats)
        record["wam_staffing_outlier_flag"] = outlier_flag
        record["wam_staffing_outlier_severity"] = outlier_severity
        record["wam_staffing_efficiency_quadrant"] = staffing_quadrant(
            float(record["average_wage_yen"]) if is_numeric(record.get("average_wage_yen")) else None,
            staffing_value,
            float(wage_median) if is_numeric(wage_median) else None,
            float(staffing_median) if is_numeric(staffing_median) else None,
        )

    staffing_outliers = [
        {
            "office_no": record.get("office_no"),
            "municipality": record.get("municipality"),
            "corporation_name": record.get("corporation_name"),
            "office_name": record.get("office_name"),
            "average_wage_yen": record.get("average_wage_yen"),
            "wam_key_staff_fte_per_capacity": record.get("wam_key_staff_fte_per_capacity"),
            "wam_staffing_outlier_flag": record.get("wam_staffing_outlier_flag"),
            "wam_staffing_outlier_severity": record.get("wam_staffing_outlier_severity"),
            "wam_transport_available": record.get("wam_transport_available"),
            "wam_primary_activity_type": record.get("wam_primary_activity_type"),
        }
        for record in matched_records
        if record.get("wam_staffing_outlier_flag")
    ]
    staffing_outliers.sort(
        key=lambda row: (
            row["wam_staffing_outlier_flag"] != "high",
            -(row["wam_key_staff_fte_per_capacity"] or 0),
        )
    )

    quadrant_summary = Counter(
        record["wam_staffing_efficiency_quadrant"]
        for record in matched_records
        if record.get("wam_staffing_efficiency_quadrant")
    )
    return {
        "matched_record_count": len(matched_records),
        "staffing_stats": staffing_stats,
        "wage_stats": wage_stats,
        "role_summary": summarize_wam_roles(matched_records),
        "feature_summary": summarize_feature_effects(matched_records),
        "staffing_outliers": staffing_outliers[:24],
        "staffing_quadrant_summary": [
            {"label": label, "office_count": quadrant_summary.get(label, 0)}
            for label in ["高工賃 × 厚い人員", "高工賃 × 少ない人員", "低工賃 × 厚い人員", "低工賃 × 少ない人員"]
        ],
        "watchlists": build_watchlists(records),
    }


def main() -> None:
    args = parse_args()
    dashboard_input = resolve_path(args.dashboard_input)
    wam_input = resolve_path(args.wam_details_input)
    override_input = resolve_path(args.wam_match_override_input)
    integrated_output = resolve_path(args.integrated_output)
    records_output_json = resolve_path(args.records_output_json)
    records_output_csv = resolve_path(args.records_output_csv)
    match_report_output = resolve_path(args.match_report_output)
    app_output = resolve_path(args.app_output)
    app_records_dir = app_output.parent / "records"
    app_root_dir = app_output.parent.parent

    dashboard_payload = load_json(dashboard_input)
    wam_rows = dedupe_wam_rows(load_json(wam_input))
    override_rows = load_csv_rows(override_input)

    records = deepcopy(dashboard_payload.get("records", []))
    chosen_matches, match_rows = assign_matches(records, wam_rows, override_rows)

    merged_records: list[dict[str, Any]] = []
    matched_wam_indices = {proposal["wam_index"] for proposal in chosen_matches.values()}
    for index, record in enumerate(records):
        match = chosen_matches.get(index)
        wam_row = wam_rows[match["wam_index"]] if match else None
        merged_records.append(merge_wam_fields(record, wam_row, match))

    wam_analytics = enrich_staffing_analytics(merged_records)

    workbook_osaka_records = [record for record in merged_records if record.get("municipality") == "大阪市"]
    unmatched_workbook_osaka_records = [
        {
            "office_no": record.get("office_no"),
            "corporation_name": record.get("corporation_name"),
            "office_name": record.get("office_name"),
            "capacity": record.get("capacity"),
            "average_wage_yen": record.get("average_wage_yen"),
        }
        for record in workbook_osaka_records
        if record.get("wam_match_status") != "matched"
    ]
    unmatched_wam_rows = [
        {
            "office_number": wam_row.get("office_number"),
            "corporation_name": wam_row.get("corporation_name"),
            "office_name": wam_row.get("office_name"),
            "office_capacity": wam_row.get("office_capacity"),
            "fetch_status": wam_row.get("fetch_status"),
        }
        for wam_index, wam_row in enumerate(wam_rows)
        if wam_index not in matched_wam_indices
    ]

    match_summary = {
        "wam_directory_count": len(wam_rows),
        "workbook_osaka_record_count": len(workbook_osaka_records),
        "matched_record_count": sum(1 for record in workbook_osaka_records if record.get("wam_match_status") == "matched"),
        "unmatched_record_count": len(unmatched_workbook_osaka_records),
        "matched_wam_count": len(matched_wam_indices),
        "unmatched_wam_count": len(unmatched_wam_rows),
        "confidence_breakdown": dict(
            Counter(
                record["wam_match_confidence"]
                for record in workbook_osaka_records
                if record.get("wam_match_confidence")
            )
        ),
        "method_breakdown": dict(
            Counter(
                record["wam_match_method"]
                for record in workbook_osaka_records
                if record.get("wam_match_method")
            )
        ),
        "manual_override_count": sum(
            1 for record in workbook_osaka_records if record.get("wam_match_method") == "manual_override"
        ),
    }

    slim_wam_directory = [
        {field: wam_row.get(field) for field in WAM_DIRECTORY_FIELD_ORDER}
        for wam_row in wam_rows
    ]

    integrated_payload = deepcopy(dashboard_payload)
    generated_at = datetime.now().astimezone().isoformat(timespec="seconds")
    integrated_payload["meta"]["generated_at"] = generated_at
    integrated_payload["meta"]["wam_generated_at"] = generated_at
    integrated_payload["meta"]["wam_directory_count"] = len(wam_rows)
    integrated_payload["meta"]["wam_matched_record_count"] = match_summary["matched_record_count"]
    integrated_payload["summary"]["wam_match_summary"] = match_summary
    integrated_payload["summary"]["wam_staffing_summary"] = {
        "matched_record_count": wam_analytics["matched_record_count"],
        "staffing_stats": wam_analytics["staffing_stats"],
        "feature_summary": wam_analytics["feature_summary"],
    }
    integrated_payload["analytics"]["wam_match_summary"] = match_summary
    integrated_payload["analytics"]["wam_staffing"] = wam_analytics
    integrated_payload["analytics"]["wam_unmatched_workbook_osaka_records"] = unmatched_workbook_osaka_records[:120]
    integrated_payload["analytics"]["wam_unmatched_directory"] = unmatched_wam_rows[:120]
    integrated_payload["issues"] = list(integrated_payload.get("issues", []))
    integrated_payload["issues"].append(
        {
            "sheet": "WAM 大阪市B型統合",
            "kind": "wam_match_coverage",
            "detail": f"大阪市レコード {match_summary['matched_record_count']} / {match_summary['workbook_osaka_record_count']} 件を WAM 詳細へ一致させた。",
        }
    )
    integrated_payload["records"] = merged_records
    integrated_payload["wam_directory"] = slim_wam_directory

    app_payload = deepcopy(integrated_payload)
    app_payload.pop("wam_directory", None)
    app_payload.pop("lookups", None)
    app_payload["analytics"] = {
        "overall_wage_stats": integrated_payload["analytics"].get("overall_wage_stats"),
        "overall_utilization_stats": integrated_payload["analytics"].get("overall_utilization_stats"),
        "wam_match_summary": integrated_payload["analytics"].get("wam_match_summary"),
    }
    app_records = slim_records_for_app(merged_records)
    app_payload["records"] = []
    app_payload["meta"].pop("source_workbook", None)
    app_payload["meta"]["app_record_field_count"] = len(APP_RECORD_FIELD_ORDER)
    app_payload["meta"]["app_record_chunk_size"] = APP_RECORD_CHUNK_SIZE

    app_records_dir.mkdir(parents=True, exist_ok=True)
    for stale_file in app_records_dir.glob("dashboard-records-*.json"):
        stale_file.unlink()

    record_chunk_files: list[str] = []
    for chunk_index, chunk_rows in enumerate(chunked_rows(app_records, APP_RECORD_CHUNK_SIZE), start=1):
        chunk_path = app_records_dir / f"dashboard-records-{chunk_index:03d}.json"
        write_json(chunk_path, chunk_rows)
        record_chunk_files.append(str(chunk_path.relative_to(app_root_dir)))

    app_payload["data_files"] = {
        "record_chunks": record_chunk_files,
        "record_count": len(app_records),
    }

    write_json(integrated_output, integrated_payload)
    write_json(records_output_json, merged_records)
    all_fieldnames = list(merged_records[0].keys()) if merged_records else []
    write_dict_csv(records_output_csv, all_fieldnames, merged_records)
    write_dict_csv(match_report_output, list(match_rows[0].keys()) if match_rows else [], match_rows)
    write_json(app_output, app_payload)

    summary = {
        "generated_at": generated_at,
        "dashboard_input": str(dashboard_input),
        "wam_input": str(wam_input),
        "matched_record_count": match_summary["matched_record_count"],
        "unmatched_record_count": match_summary["unmatched_record_count"],
        "matched_wam_count": match_summary["matched_wam_count"],
        "unmatched_wam_count": match_summary["unmatched_wam_count"],
        "files": {
            "integrated_dashboard_json": str(integrated_output.relative_to(repo_root())),
            "integrated_records_json": str(records_output_json.relative_to(repo_root())),
            "integrated_records_csv": str(records_output_csv.relative_to(repo_root())),
            "match_report_csv": str(match_report_output.relative_to(repo_root())),
            "app_dashboard_json": str(app_output.relative_to(repo_root())),
            "app_record_chunk_count": len(record_chunk_files),
        },
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
