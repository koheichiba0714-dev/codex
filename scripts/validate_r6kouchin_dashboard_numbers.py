#!/usr/bin/env python3
from __future__ import annotations

import argparse
from collections import Counter
import json
import math
from pathlib import Path
import sys
from typing import Any


SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from build_r6kouchinjissekib_dashboard_dataset import (  # type: ignore
    classify_outlier as classify_staffing_outlier,
    compute_numeric_stats as compute_staffing_stats,
    is_numeric,
    staffing_quadrant,
)
from extract_r6kouchinjissekib import (  # type: ignore
    capacity_band_label,
    classify_outlier as classify_wage_outlier,
    compute_numeric_stats as compute_wage_stats,
    market_position_quadrant,
    wage_band_label,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Dashboard 数値整合性チェック")
    parser.add_argument(
        "--app-manifest",
        default="apps/r6kouchin-dashboard/data/dashboard-data.json",
        help="app manifest json path",
    )
    parser.add_argument(
        "--integrated-dashboard",
        default="data/exports/r6kouchinjissekib/integrated/shuro_b_dashboard_integrated.json",
        help="integrated dashboard json path",
    )
    parser.add_argument(
        "--output-json",
        default="artifacts/analysis_20260310_r6kouchin_dashboard_numeric_validation.json",
        help="validation result json path",
    )
    return parser.parse_args()


def repo_root() -> Path:
    return SCRIPT_DIR.parent


def resolve_repo_path(path_str: str) -> Path:
    path = Path(path_str).expanduser()
    if path.is_absolute():
      return path.resolve()
    return (repo_root() / path).resolve()


def load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def read_app_records(manifest_path: Path) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    manifest = load_json(manifest_path)
    chunk_paths = manifest.get("data_files", {}).get("record_chunks", [])
    records: list[dict[str, Any]] = []
    for relative_path in chunk_paths:
        chunk_path = (manifest_path.parent.parent / relative_path).resolve()
        records.extend(load_json(chunk_path))
    return manifest, records


def scoped_integrated_records(records: list[dict[str, Any]], scope_label: str | None) -> list[dict[str, Any]]:
    if not scope_label:
        return records
    return [record for record in records if record.get("municipality") == scope_label]


def is_close(left: float | None, right: float | None, tolerance: float = 0.001) -> bool:
    if left is None and right is None:
        return True
    if left is None or right is None:
        return False
    return math.isclose(float(left), float(right), rel_tol=0.0, abs_tol=tolerance)


def office_key(record: dict[str, Any]) -> str:
    return str(record.get("office_no"))


def index_by_office(records: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    return {office_key(record): record for record in records}


def summarize_check(name: str, passed: bool, detail: str, samples: list[dict[str, Any]] | None = None) -> dict[str, Any]:
    return {
        "name": name,
        "passed": passed,
        "detail": detail,
        "sample_mismatches": samples or [],
    }


def duplicate_count(records: list[dict[str, Any]]) -> int:
    counts = Counter(office_key(record) for record in records if record.get("office_no") is not None)
    return sum(count - 1 for count in counts.values() if count > 1)


def response_breakdown(records: list[dict[str, Any]]) -> dict[str, int]:
    return dict(Counter(str(record.get("response_status")) for record in records if record.get("response_status")))


def numeric_records(records: list[dict[str, Any]], key: str) -> list[float]:
    return [float(record[key]) for record in records if is_numeric(record.get(key))]


def main() -> None:
    args = parse_args()
    manifest_path = resolve_repo_path(args.app_manifest)
    integrated_path = resolve_repo_path(args.integrated_dashboard)
    output_path = resolve_repo_path(args.output_json)

    manifest, app_records = read_app_records(manifest_path)
    integrated = load_json(integrated_path)
    integrated_records: list[dict[str, Any]] = integrated["records"]
    app_scope = manifest.get("meta", {}).get("scope")
    scoped_integrated = scoped_integrated_records(integrated_records, app_scope)
    integrated_by_office = index_by_office(integrated_records)
    scoped_integrated_by_office = index_by_office(scoped_integrated)
    app_by_office = index_by_office(app_records)

    checks: list[dict[str, Any]] = []

    manifest_count = manifest["meta"]["record_count"]
    chunk_count = manifest["data_files"]["record_count"]
    app_total_records = manifest["meta"].get("total_records", manifest_count)
    integrated_meta_count = integrated["meta"]["record_count"]
    integrated_record_count = len(integrated_records)
    checks.append(
        summarize_check(
            "record_count_consistency",
            manifest_count == chunk_count == app_total_records == len(app_records),
            f"app_meta={manifest_count}, chunks={chunk_count}, app_total_records={app_total_records}, app_records={len(app_records)}, integrated_meta={integrated_meta_count}, integrated_records={integrated_record_count}",
        )
    )

    checks.append(
        summarize_check(
            "office_key_coverage",
            set(app_by_office) == set(scoped_integrated_by_office),
            f"app_unique={len(app_by_office)}, scoped_integrated_unique={len(scoped_integrated_by_office)}, scope={app_scope}",
        )
    )

    app_summary = manifest.get("summary", {})
    app_analytics = manifest.get("analytics", {})
    app_response = response_breakdown(app_records)
    app_summary_count_ok = (
        app_summary.get("record_count") == len(app_records)
        and app_summary.get("municipality_count") == len({record.get("municipality") for record in app_records if record.get("municipality")})
    )
    checks.append(
        summarize_check(
            "app_summary_counts",
            app_summary_count_ok,
            f"summary_record_count={app_summary.get('record_count')}, app_records={len(app_records)}, summary_municipality_count={app_summary.get('municipality_count')}",
        )
    )
    checks.append(
        summarize_check(
            "app_summary_response_breakdown",
            app_summary.get("response_breakdown", {}) == app_response,
            f"summary={app_summary.get('response_breakdown', {})}, actual={app_response}",
        )
    )

    app_wage_stats = compute_wage_stats(numeric_records(app_records, "average_wage_yen"))
    app_summary_wage_stats = app_summary.get("overall_wage_stats", {})
    app_analytics_wage_stats = app_analytics.get("overall_wage_stats", {})
    app_wage_stats_ok = all(
        is_close(app_wage_stats.get(key), app_summary_wage_stats.get(key))
        and is_close(app_wage_stats.get(key), app_analytics_wage_stats.get(key))
        for key in ["mean", "median", "q1", "q3", "p90", "lower_fence", "upper_fence", "extreme_upper_fence"]
    ) and (
        app_wage_stats.get("count") == app_summary_wage_stats.get("count") == app_analytics_wage_stats.get("count")
    )
    checks.append(
        summarize_check(
            "app_overall_wage_stats",
            app_wage_stats_ok,
            f"computed_mean={app_wage_stats.get('mean')}, summary_mean={app_summary_wage_stats.get('mean')}, analytics_mean={app_analytics_wage_stats.get('mean')}",
        )
    )

    app_utilization_stats = compute_wage_stats(numeric_records(app_records, "daily_user_capacity_ratio"))
    app_summary_utilization = app_summary.get("overall_utilization_stats", {})
    app_analytics_utilization = app_analytics.get("overall_utilization_stats", {})
    app_utilization_ok = all(
        is_close(app_utilization_stats.get(key), app_summary_utilization.get(key))
        and is_close(app_utilization_stats.get(key), app_analytics_utilization.get(key))
        for key in ["mean", "median", "q1", "q3", "p90"]
    ) and (
        app_utilization_stats.get("count")
        == app_summary_utilization.get("count")
        == app_analytics_utilization.get("count")
    )
    checks.append(
        summarize_check(
            "app_overall_utilization_stats",
            app_utilization_ok,
            f"computed_mean={app_utilization_stats.get('mean')}, summary_mean={app_summary_utilization.get('mean')}, analytics_mean={app_analytics_utilization.get('mean')}",
        )
    )

    expected_duplicates = int(integrated["summary"].get("duplicate_office_no_count", 0))
    actual_duplicates = duplicate_count(integrated_records)
    checks.append(
        summarize_check(
            "duplicate_office_numbers",
            expected_duplicates == actual_duplicates,
            f"summary={expected_duplicates}, actual={actual_duplicates}",
        )
    )

    expected_response = integrated["summary"].get("response_breakdown", {})
    actual_response = response_breakdown(integrated_records)
    checks.append(
        summarize_check(
            "response_breakdown",
            expected_response == actual_response,
            f"summary={expected_response}, actual={actual_response}",
        )
    )

    wage_stats = compute_wage_stats(numeric_records(integrated_records, "average_wage_yen"))
    expected_wage_stats = integrated["analytics"]["overall_wage_stats"]
    wage_stats_ok = all(
        is_close(wage_stats.get(key), expected_wage_stats.get(key))
        for key in ["mean", "median", "q1", "q3", "p90", "lower_fence", "upper_fence", "extreme_upper_fence"]
    ) and wage_stats.get("count") == expected_wage_stats.get("count")
    checks.append(
        summarize_check(
            "overall_wage_stats",
            wage_stats_ok,
            f"computed_mean={wage_stats.get('mean')}, summary_mean={expected_wage_stats.get('mean')}, computed_median={wage_stats.get('median')}, summary_median={expected_wage_stats.get('median')}",
        )
    )

    utilization_stats = compute_wage_stats(numeric_records(integrated_records, "daily_user_capacity_ratio"))
    expected_utilization = integrated["analytics"]["overall_utilization_stats"]
    utilization_ok = all(
        is_close(utilization_stats.get(key), expected_utilization.get(key))
        for key in ["mean", "median", "q1", "q3", "p90"]
    ) and utilization_stats.get("count") == expected_utilization.get("count")
    checks.append(
        summarize_check(
            "overall_utilization_stats",
            utilization_ok,
            f"computed_mean={utilization_stats.get('mean')}, summary_mean={expected_utilization.get('mean')}, computed_median={utilization_stats.get('median')}, summary_median={expected_utilization.get('median')}",
        )
    )

    ratio_samples: list[dict[str, Any]] = []
    daily_ratio_mismatch = 0
    for record in integrated_records:
        capacity = record.get("capacity")
        average_daily_users = record.get("average_daily_users")
        stored_ratio = record.get("daily_user_capacity_ratio")
        if not (is_numeric(capacity) and is_numeric(average_daily_users) and float(capacity) > 0):
            continue
        expected_ratio = round(float(average_daily_users) / float(capacity), 3)
        if not is_close(expected_ratio, stored_ratio):
            daily_ratio_mismatch += 1
            if len(ratio_samples) < 5:
                ratio_samples.append(
                    {
                        "office_no": record.get("office_no"),
                        "expected": expected_ratio,
                        "stored": stored_ratio,
                    }
                )
    checks.append(
        summarize_check(
            "daily_user_capacity_ratio",
            daily_ratio_mismatch == 0,
            f"mismatches={daily_ratio_mismatch}",
            ratio_samples,
        )
    )

    overall_ratio_mismatch = 0
    overall_ratio_samples: list[dict[str, Any]] = []
    overall_mean = expected_wage_stats.get("mean")
    for record in integrated_records:
        wage = record.get("average_wage_yen")
        ratio = record.get("wage_ratio_to_overall_mean")
        if not (is_numeric(wage) and is_numeric(overall_mean)):
            continue
        expected_ratio = round(float(wage) / float(overall_mean), 3)
        if not is_close(expected_ratio, ratio):
            overall_ratio_mismatch += 1
            if len(overall_ratio_samples) < 5:
                overall_ratio_samples.append(
                    {
                        "office_no": record.get("office_no"),
                        "expected": expected_ratio,
                        "stored": ratio,
                    }
                )
    checks.append(
        summarize_check(
            "wage_ratio_to_overall_mean",
            overall_ratio_mismatch == 0,
            f"mismatches={overall_ratio_mismatch}",
            overall_ratio_samples,
        )
    )

    grouped_ratio_checks = [
        ("wage_ratio_to_municipality_mean", "municipality_average_wage_yen"),
        ("wage_ratio_to_corporation_type_mean", "corporation_type_average_wage_yen"),
        ("wage_ratio_to_capacity_band_mean", "capacity_band_average_wage_yen"),
    ]
    for ratio_key, base_key in grouped_ratio_checks:
        mismatch_count = 0
        mismatch_samples: list[dict[str, Any]] = []
        for record in integrated_records:
            wage = record.get("average_wage_yen")
            base_value = record.get(base_key)
            stored_ratio = record.get(ratio_key)
            if not (is_numeric(wage) and is_numeric(base_value) and float(base_value) > 0):
                continue
            expected_ratio = round(float(wage) / float(base_value), 3)
            if not is_close(expected_ratio, stored_ratio):
                mismatch_count += 1
                if len(mismatch_samples) < 5:
                    mismatch_samples.append(
                        {
                            "office_no": record.get("office_no"),
                            "expected": expected_ratio,
                            "stored": stored_ratio,
                        }
                    )
        checks.append(
            summarize_check(
                ratio_key,
                mismatch_count == 0,
                f"mismatches={mismatch_count}",
                mismatch_samples,
            )
        )

    label_samples: list[dict[str, Any]] = []
    label_mismatch = 0
    for record in integrated_records:
        wage = record.get("average_wage_yen")
        capacity = record.get("capacity")
        expected_wage_band = wage_band_label(float(wage)) if is_numeric(wage) else None
        expected_capacity_band = capacity_band_label(float(capacity)) if is_numeric(capacity) else None
        if expected_wage_band != record.get("wage_band_label") or expected_capacity_band != record.get("capacity_band_label"):
            label_mismatch += 1
            if len(label_samples) < 5:
                label_samples.append(
                    {
                        "office_no": record.get("office_no"),
                        "expected_wage_band": expected_wage_band,
                        "stored_wage_band": record.get("wage_band_label"),
                        "expected_capacity_band": expected_capacity_band,
                        "stored_capacity_band": record.get("capacity_band_label"),
                    }
                )
    checks.append(
        summarize_check(
            "band_labels",
            label_mismatch == 0,
            f"mismatches={label_mismatch}",
            label_samples,
        )
    )

    quadrant_samples: list[dict[str, Any]] = []
    quadrant_mismatch = 0
    utilization_median = expected_utilization.get("median")
    for record in integrated_records:
        wage = record.get("average_wage_yen")
        utilization_ratio = record.get("daily_user_capacity_ratio")
        if not (is_numeric(wage) and is_numeric(utilization_ratio) and is_numeric(overall_mean) and is_numeric(utilization_median)):
            continue
        expected_quadrant = market_position_quadrant(
            float(wage),
            float(utilization_ratio),
            float(overall_mean),
            float(utilization_median),
        )
        if expected_quadrant != record.get("market_position_quadrant"):
            quadrant_mismatch += 1
            if len(quadrant_samples) < 5:
                quadrant_samples.append(
                    {
                        "office_no": record.get("office_no"),
                        "expected": expected_quadrant,
                        "stored": record.get("market_position_quadrant"),
                    }
                )
    checks.append(
        summarize_check(
            "market_position_quadrant",
            quadrant_mismatch == 0,
            f"mismatches={quadrant_mismatch}",
            quadrant_samples,
        )
    )

    wage_outlier_mismatch = 0
    wage_outlier_samples: list[dict[str, Any]] = []
    high_outlier_count = 0
    low_outlier_count = 0
    for record in integrated_records:
        wage = record.get("average_wage_yen")
        z_score = record.get("wage_z_score")
        if not is_numeric(wage):
            continue
        expected_flag, expected_severity = classify_wage_outlier(float(wage), float(z_score) if is_numeric(z_score) else None, expected_wage_stats)
        if expected_flag == "high":
            high_outlier_count += 1
        if expected_flag == "low":
            low_outlier_count += 1
        if expected_flag != record.get("wage_outlier_flag") or expected_severity != record.get("wage_outlier_severity"):
            wage_outlier_mismatch += 1
            if len(wage_outlier_samples) < 5:
                wage_outlier_samples.append(
                    {
                        "office_no": record.get("office_no"),
                        "expected_flag": expected_flag,
                        "stored_flag": record.get("wage_outlier_flag"),
                        "expected_severity": expected_severity,
                        "stored_severity": record.get("wage_outlier_severity"),
                    }
                )
    checks.append(
        summarize_check(
            "wage_outlier_flags",
            wage_outlier_mismatch == 0,
            f"mismatches={wage_outlier_mismatch}, high={high_outlier_count}, low={low_outlier_count}",
            wage_outlier_samples,
        )
    )
    checks.append(
        summarize_check(
            "summary_outlier_counts",
            high_outlier_count == integrated["summary"]["high_outlier_count"] and low_outlier_count == integrated["summary"]["low_outlier_count"],
            f"computed_high={high_outlier_count}, summary_high={integrated['summary']['high_outlier_count']}, computed_low={low_outlier_count}, summary_low={integrated['summary']['low_outlier_count']}",
        )
    )

    match_summary = integrated["summary"]["wam_match_summary"]
    matched_osaka = [record for record in integrated_records if record.get("municipality") == "大阪市" and record.get("wam_match_status") == "matched"]
    checks.append(
        summarize_check(
            "wam_match_summary",
            len(matched_osaka) == match_summary["matched_record_count"] == manifest["meta"]["wam_matched_record_count"],
            f"computed_matched_osaka={len(matched_osaka)}, summary_matched={match_summary['matched_record_count']}, manifest_meta={manifest['meta']['wam_matched_record_count']}",
        )
    )

    match_confidence_counts = dict(Counter(record.get("wam_match_confidence") for record in matched_osaka if record.get("wam_match_confidence")))
    checks.append(
        summarize_check(
            "wam_confidence_breakdown",
            match_confidence_counts == match_summary["confidence_breakdown"],
            f"computed={match_confidence_counts}, summary={match_summary['confidence_breakdown']}",
        )
    )

    gap_mismatch = 0
    gap_samples: list[dict[str, Any]] = []
    for record in integrated_records:
        monthly_wage = record.get("wam_average_wage_monthly_yen")
        workbook_wage = record.get("average_wage_yen")
        stored_gap = record.get("wam_average_wage_gap_yen")
        if not (is_numeric(monthly_wage) and is_numeric(workbook_wage)):
            continue
        expected_gap = round(float(monthly_wage) - float(workbook_wage), 3)
        if not is_close(expected_gap, stored_gap):
            gap_mismatch += 1
            if len(gap_samples) < 5:
                gap_samples.append(
                    {
                        "office_no": record.get("office_no"),
                        "expected": expected_gap,
                        "stored": stored_gap,
                    }
                )
    checks.append(
        summarize_check(
            "wam_average_wage_gap_yen",
            gap_mismatch == 0,
            f"mismatches={gap_mismatch}",
            gap_samples,
        )
    )

    staffing_ratio_mismatch = 0
    staffing_ratio_samples: list[dict[str, Any]] = []
    for record in integrated_records:
        total_fte = record.get("wam_key_staff_fte_total")
        wam_capacity = record.get("wam_office_capacity")
        stored_ratio = record.get("wam_key_staff_fte_per_capacity")
        if not (is_numeric(total_fte) and is_numeric(wam_capacity) and float(wam_capacity) > 0):
            continue
        expected_ratio = round(float(total_fte) / float(wam_capacity), 3)
        if not is_close(expected_ratio, stored_ratio):
            staffing_ratio_mismatch += 1
            if len(staffing_ratio_samples) < 5:
                staffing_ratio_samples.append(
                    {
                        "office_no": record.get("office_no"),
                        "expected": expected_ratio,
                        "stored": stored_ratio,
                    }
                )
    checks.append(
        summarize_check(
            "wam_key_staff_fte_per_capacity",
            staffing_ratio_mismatch == 0,
            f"mismatches={staffing_ratio_mismatch}",
            staffing_ratio_samples,
        )
    )

    matched_records = [
        record
        for record in integrated_records
        if record.get("wam_match_status") == "matched" and record.get("wam_fetch_status") == "ok"
    ]
    staffing_values = numeric_records(matched_records, "wam_key_staff_fte_per_capacity")
    staffing_stats = compute_staffing_stats(staffing_values)
    staffing_median = staffing_stats["median"]
    wage_median = compute_staffing_stats(numeric_records(matched_records, "average_wage_yen"))["median"]

    staffing_outlier_mismatch = 0
    staffing_quadrant_mismatch = 0
    staffing_samples: list[dict[str, Any]] = []
    quadrant_samples_staffing: list[dict[str, Any]] = []
    for record in matched_records:
        staffing_value = record.get("wam_key_staff_fte_per_capacity")
        wage_value = record.get("average_wage_yen")
        if is_numeric(staffing_value):
            expected_flag, expected_severity = classify_staffing_outlier(float(staffing_value), staffing_stats)
            if expected_flag != record.get("wam_staffing_outlier_flag") or expected_severity != record.get("wam_staffing_outlier_severity"):
                staffing_outlier_mismatch += 1
                if len(staffing_samples) < 5:
                    staffing_samples.append(
                        {
                            "office_no": record.get("office_no"),
                            "expected_flag": expected_flag,
                            "stored_flag": record.get("wam_staffing_outlier_flag"),
                            "expected_severity": expected_severity,
                            "stored_severity": record.get("wam_staffing_outlier_severity"),
                        }
                    )
        if is_numeric(staffing_value) and is_numeric(wage_value) and is_numeric(staffing_median) and is_numeric(wage_median):
            expected_quadrant = staffing_quadrant(float(wage_value), float(staffing_value), float(wage_median), float(staffing_median))
            if expected_quadrant != record.get("wam_staffing_efficiency_quadrant"):
                staffing_quadrant_mismatch += 1
                if len(quadrant_samples_staffing) < 5:
                    quadrant_samples_staffing.append(
                        {
                            "office_no": record.get("office_no"),
                            "expected": expected_quadrant,
                            "stored": record.get("wam_staffing_efficiency_quadrant"),
                        }
                    )
    checks.append(
        summarize_check(
            "wam_staffing_outlier_flags",
            staffing_outlier_mismatch == 0,
            f"mismatches={staffing_outlier_mismatch}",
            staffing_samples,
        )
    )
    checks.append(
        summarize_check(
            "wam_staffing_efficiency_quadrant",
            staffing_quadrant_mismatch == 0,
            f"mismatches={staffing_quadrant_mismatch}",
            quadrant_samples_staffing,
        )
    )

    app_field_subset_mismatch = 0
    app_subset_samples: list[dict[str, Any]] = []
    for office_no, app_record in app_by_office.items():
        integrated_record = scoped_integrated_by_office.get(office_no)
        if integrated_record is None:
            app_field_subset_mismatch += 1
            continue
        for key, value in app_record.items():
            if key not in integrated_record:
                continue
            if integrated_record.get(key) != value:
                app_field_subset_mismatch += 1
                if len(app_subset_samples) < 5:
                    app_subset_samples.append(
                        {
                            "office_no": office_no,
                            "field": key,
                            "app_value": value,
                            "integrated_value": integrated_record.get(key),
                        }
                    )
                break
    checks.append(
        summarize_check(
            "app_record_subset_integrity",
            app_field_subset_mismatch == 0,
            f"mismatched_records={app_field_subset_mismatch}",
            app_subset_samples,
        )
    )

    passed = sum(1 for check in checks if check["passed"])
    failed_checks = [check["name"] for check in checks if not check["passed"]]
    payload = {
        "generated_at": __import__("datetime").datetime.now().astimezone().isoformat(timespec="seconds"),
        "app_manifest": str(manifest_path),
        "integrated_dashboard": str(integrated_path),
        "summary": {
            "check_count": len(checks),
            "passed_count": passed,
            "failed_count": len(checks) - passed,
            "failed_checks": failed_checks,
        },
        "checks": checks,
    }

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(payload["summary"], ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
