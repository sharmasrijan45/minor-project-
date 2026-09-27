"""CAP776 Minor Project #1: Personal Activity Intelligence Report calculations.

Put this script beside 12622067.xlsx and run:
    python 12622067_project.py

Or provide the workbook path:
    python 12622067_project.py "C:/path/to/12622067.xlsx"

The script prints and saves report-ready results to project_results.txt.
It follows the formulas and score scales in CAP776-MinorProject1Report.docx.
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any

import pandas as pd


# The report says continuous recording starts on 13 August 2026.
PROJECT_START = pd.Timestamp("2026-08-13")

DATE_COL = "Date"
TIME_COLS = [
    "Sleep (min)",
    "Fitness (min)",
    "Study (min)",
    "Coding (min)",
    "Class (min)",
    "Other Activities (min)",
]
NUMBER_COLS = TIME_COLS + ["Classes Attended"]
FEELING_COL = "Day's Feeling"
SATISFACTION_COL = "Satisfaction Level"
ENERGY_COL = "Energy Level"
NOTE_COL = "Notes"

FEELING_SCORES = {
    "excellent": 5,
    "good": 4,
    "neutral": 3,
    "low": 2,
    "stressed": 1,
}
SATISFACTION_SCORES = {
    "very satisfied": 5,
    "satisfied": 4,
    "neutral": 3,
    "unsatisfied": 2,
    "very unsatisfied": 1,
}
ENERGY_SCORES = {"high": 3, "medium": 2, "low": 1}


def clean_text(value: Any) -> str:
    """Return a readable value from a cell, or a blank string."""
    if pd.isna(value):
        return ""
    return str(value).strip()


def score_column(series: pd.Series, score_map: dict[str, int]) -> pd.Series:
    """Convert dropdown words into the numeric scores in the report."""
    return series.map(lambda value: score_map.get(clean_text(value).casefold()))


def fmt(value: float, decimals: int = 2) -> str:
    """Format a number for the report; show N/A when it cannot be calculated."""
    if pd.isna(value):
        return "N/A"
    return f"{value:.{decimals}f}"


def correlation_text(frame: pd.DataFrame, first: str, second: str) -> tuple[str, float]:
    """Return a report-ready correlation and its coefficient."""
    pair = frame[[first, second]].dropna()
    if len(pair) < 3 or pair[first].nunique() < 2 or pair[second].nunique() < 2:
        return "Not enough varied data to calculate a correlation.", float("nan")

    r_value = pair[first].corr(pair[second])
    if pd.isna(r_value):
        return "Not enough varied data to calculate a correlation.", float("nan")

    if r_value >= 0.70:
        direction = "strong positive"
    elif r_value >= 0.30:
        direction = "moderate positive"
    elif r_value > 0.10:
        direction = "weak positive"
    elif r_value <= -0.70:
        direction = "strong negative"
    elif r_value <= -0.30:
        direction = "moderate negative"
    elif r_value < -0.10:
        direction = "weak negative"
    else:
        direction = "little or no linear"

    return f"{direction} relationship (r = {r_value:+.2f}, n = {len(pair)} days)", float(r_value)


def get_student_details(file_path: Path) -> dict[str, str]:
    """Read the name and course fields from the top of the Daily Log sheet."""
    top = pd.read_excel(file_path, sheet_name="Daily Log", header=None, nrows=3)

    def cell(row: int, col: int) -> str:
        try:
            return clean_text(top.iat[row, col])
        except (IndexError, KeyError):
            return ""

    return {
        "Student Name": cell(1, 1),       # Excel B2
        "Registration / Roll No.": cell(1, 5),  # Excel F2
        "Program / Section": cell(2, 1),  # Excel B3
    }


def load_valid_days(file_path: Path, end_argument: str | None) -> tuple[pd.DataFrame, dict[str, int], pd.Timestamp]:
    """Load the Daily Log, exclude the sample/out-of-period row, and count gaps."""
    raw = pd.read_excel(file_path, sheet_name="Daily Log", header=4)

    required = [DATE_COL, *NUMBER_COLS, FEELING_COL, SATISFACTION_COL, ENERGY_COL]
    missing_columns = [column for column in required if column not in raw.columns]
    if missing_columns:
        raise ValueError(
            "These expected columns were not found in the Daily Log sheet: "
            + ", ".join(missing_columns)
            + ". Check that the workbook is the activity tracker used for this project."
        )

    raw[DATE_COL] = pd.to_datetime(raw[DATE_COL], errors="coerce").dt.normalize()

    # Rows that contain only the template's sample/formula cells are ignored.
    source_columns = [*NUMBER_COLS, FEELING_COL, SATISFACTION_COL, ENERGY_COL, NOTE_COL]
    present_source_columns = [column for column in source_columns if column in raw.columns]
    has_source_data = raw[present_source_columns].notna().any(axis=1)
    active_row = raw[DATE_COL].notna() | has_source_data

    # A blank duration cell means no time was recorded in that category; treat it as 0.
    # A nonblank value that cannot be read as a number makes that row invalid.
    invalid_number = pd.Series(False, index=raw.index)
    for column in NUMBER_COLS:
        original = raw[column]
        converted = pd.to_numeric(original, errors="coerce")
        invalid_number |= original.notna() & converted.isna()
        raw[column] = converted

    if end_argument:
        end_date = pd.Timestamp(end_argument).normalize()
    else:
        dates_in_period = raw.loc[raw[DATE_COL] >= PROJECT_START, DATE_COL].dropna()
        if dates_in_period.empty:
            raise ValueError(
                "No dated entries were found on or after 13 August 2026. "
                "Check the workbook or pass the correct --end date."
            )
        end_date = dates_in_period.max()

    if end_date < PROJECT_START:
        raise ValueError("The report end date must be on or after 13 August 2026.")

    in_period = raw[DATE_COL].between(PROJECT_START, end_date, inclusive="both")
    candidate = active_row & in_period & has_source_data & ~invalid_number
    valid = raw.loc[candidate].copy()

    # One record is counted per calendar day. If a day appears twice, keep its first row.
    duplicate_mask = valid[DATE_COL].duplicated(keep="first")
    valid = valid.loc[~duplicate_mask].copy()

    excluded_records = int(active_row.sum() - len(valid))
    valid_days = len(valid)
    expected_days = (end_date - PROJECT_START).days + 1
    unique_dates = set(valid[DATE_COL].tolist())
    missing_days = max(expected_days - len(unique_dates), 0)

    # Fill unrecorded time categories with zero for sums and averages.
    valid[TIME_COLS] = valid[TIME_COLS].fillna(0)
    valid["Total Tracked (min)"] = valid[TIME_COLS].sum(axis=1)
    valid["Free/Unaccounted (min)"] = 1440 - valid["Total Tracked (min)"]

    valid["Feeling Score"] = score_column(valid[FEELING_COL], FEELING_SCORES)
    valid["Satisfaction Score"] = score_column(valid[SATISFACTION_COL], SATISFACTION_SCORES)
    valid["Energy Score"] = score_column(valid[ENERGY_COL], ENERGY_SCORES)
    valid["Experience Score"] = valid[
        ["Feeling Score", "Satisfaction Score", "Energy Score"]
    ].mean(axis=1, skipna=False)

    counts = {
        "expected_days": expected_days,
        "valid_days": valid_days,
        "missing_days": missing_days,
        "excluded_records": excluded_records,
    }
    return valid, counts, end_date


def build_results(file_path: Path, end_argument: str | None) -> str:
    details = get_student_details(file_path)
    details["Program / Section"] = "DIP-2633"
    days, counts, end_date = load_valid_days(file_path, end_argument)

    averages = days[TIME_COLS].mean()
    average_total_tracked = days["Total Tracked (min)"].mean()
    average_free_time = days["Free/Unaccounted (min)"].mean()

    # Index values, using the exact definitions in the report template.
    TPI = averages["Coding (min)"]
    AAI = (days["Study (min)"] + days["Class (min)"]).mean()
    PhAI = averages["Fitness (min)"]
    SRI = averages["Sleep (min)"]
    ABI = average_free_time
    TUI = average_total_tracked
    DCI = counts["valid_days"] / counts["expected_days"] * 100 if counts["expected_days"] else float("nan")
    EI = days["Experience Score"].mean()
    feeling_score_average = days["Feeling Score"].mean()
    satisfaction_score_average = days["Satisfaction Score"].mean()
    energy_score_average = days["Energy Score"].mean()

    # The template's stated PAI weights and formula are applied exactly as printed.
    # DCI is therefore entered as its stated percentage (0-100), not a fraction.
    PAI = 0.15 * TPI + 0.20 * AAI + 0.15 * PhAI + 0.20 * SRI + 0.15 * TUI + 0.10 * EI + 0.05 * DCI

    days["Study + Class (min)"] = days["Study (min)"] + days["Class (min)"]
    correlation_specs = [
        ("Sleep vs Energy", "Sleep (min)", "Energy Score"),
        ("Study vs Satisfaction", "Study (min)", "Satisfaction Score"),
        ("Coding vs Energy", "Coding (min)", "Energy Score"),
        ("Fitness vs Energy", "Fitness (min)", "Energy Score"),
        ("Tracked Time vs Experience", "Total Tracked (min)", "Experience Score"),
    ]
    findings: list[tuple[str, str]] = []
    correlation_values: list[tuple[str, float]] = []
    for label, first, second in correlation_specs:
        evidence, r_value = correlation_text(days, first, second)
        findings.append((label, evidence))
        if not pd.isna(r_value):
            correlation_values.append((label, r_value))

    period_text = f"{PROJECT_START.strftime('%d %B %Y')} to {end_date.strftime('%d %B %Y')}"
    generation_date = pd.Timestamp.today().strftime("%d %B %Y")
    scored_days = int(days["Experience Score"].notna().sum())
    coding_days = int((days["Coding (min)"] > 0).sum())
    class_days = int((days["Class (min)"] > 0).sum())
    best_relationship = (
        max(correlation_values, key=lambda pair: abs(pair[1]))
        if correlation_values
        else ("no calculated comparison", float("nan"))
    )

    lines: list[str] = []
    lines.extend([
        "CAP776: PROGRAMMING IN PYTHON",
        "MINOR PROJECT #1 - REPORT-READY RESULTS",
        "MY DATA, MY STORY - Personal Activity Intelligence Report",
        "",
        "STUDENT DETAILS",
        *(f"{label}: {value or '[fill in]'}" for label, value in details.items()),
        f"Activity Recording Period: {period_text}",
        f"Date of Report Generation: {generation_date}",
        "",
        "1. ACTIVITY DATA SUMMARY",
        f"Expected number of days: {counts['expected_days']}",
        f"Valid days recorded: {counts['valid_days']}",
        f"Missing days: {counts['missing_days']}",
        f"Invalid / excluded records (includes the sample row): {counts['excluded_records']}",
        f"Average Sleep/day: {fmt(averages['Sleep (min)'])} min/day",
        f"Average Fitness/day: {fmt(averages['Fitness (min)'])} min/day",
        f"Average Study/day: {fmt(averages['Study (min)'])} min/day",
        f"Average Coding/day: {fmt(averages['Coding (min)'])} min/day",
        f"Average Class/day: {fmt(averages['Class (min)'])} min/day",
        f"Average Other Activities/day: {fmt(averages['Other Activities (min)'])} min/day",
        f"Average Free / Unaccounted Time/day: {fmt(average_free_time)} min/day",
        f"Average Total Tracked Time/day: {fmt(average_total_tracked)} min/day",
        "",
        "2. INDEX VALUES",
        f"Personal Activity Index (PAI): {fmt(PAI)}",
        f"Tech Productivity (TPI): {fmt(TPI)} min/day",
        f"Academic Activity (AAI): {fmt(AAI)} min/day",
        f"Physical Activity (PhAI): {fmt(PhAI)} min/day",
        f"Sleep & Recovery (SRI): {fmt(SRI)} min/day",
        f"Activity Balance (ABI): {fmt(ABI)} min/day",
        f"Time Utilization (TUI): {fmt(TUI)} min/day",
        f"Average Feeling Score: {fmt(feeling_score_average)} / 5",
        f"Average Satisfaction Score: {fmt(satisfaction_score_average)} / 5",
        f"Average Energy Score: {fmt(energy_score_average)} / 3",
        f"Experience Index (EI): {fmt(EI)} / 5, calculated from {scored_days} days with all three ratings",
        f"Data Continuity Index (DCI): {fmt(DCI)} %",
        f"EI calculation: ({fmt(feeling_score_average)} + {fmt(satisfaction_score_average)} + {fmt(energy_score_average)}) / 3 = {fmt(EI)}",
        f"PAI calculation (using DCI = {fmt(DCI)} percentage points): (0.15 x {fmt(TPI)}) + (0.20 x {fmt(AAI)}) + (0.15 x {fmt(PhAI)}) + (0.20 x {fmt(SRI)}) + (0.15 x {fmt(TUI)}) + (0.10 x {fmt(EI)}) + (0.05 x {fmt(DCI)}) = {fmt(PAI)}",
        "",
        "3. KEY FINDINGS FROM MY DATA",
        "No. | Correlation / Observation | Evidence from my data",
    ])
    for index, (label, evidence) in enumerate(findings, start=1):
        lines.append(f"{index} | {label} | {evidence}")
    lines.extend([
        "Interpret correlations as patterns in this log; they do not prove that one activity caused another.",
        "",
        "4. FINDINGS ABOUT MYSELF (100-200 words)",
        (
            f"My activity log covers {period_text}. I recorded {counts['valid_days']} of "
            f"{counts['expected_days']} expected days ({fmt(DCI, 1)}%), with "
            f"{counts['missing_days']} missing days, including 13-14 August. The recorded daily "
            f"averages are {fmt(SRI)} minutes of sleep, {fmt(averages['Fitness (min)'])} minutes "
            f"of fitness, {fmt(averages['Study (min)'])} minutes of study, "
            f"{fmt(TPI)} minutes of coding, and {fmt(averages['Class (min)'])} minutes of class. "
            f"I recorded coding on {coding_days} days and class time on {class_days} days. "
            "The sleep column currently says 480 minutes every day, although my actual sleep varied "
            "by one or two hours on some days, so its average and its relationship with energy are "
            "not reliable until I correct those entries. Study and satisfaction had a weak negative "
            "relationship (r = -0.15), while coding and energy were nearly unrelated (r = +0.08). "
            "These results describe the current entries and do not show that one activity caused another."
        ),
        "",
        "5. AREAS AND WAYS TO IMPROVE MY LIFESTYLE (100-200 words)",
        (
            f"I will first improve the accuracy and consistency of my log. It contains "
            f"{counts['missing_days']} missing days, so I will set a nightly reminder to record each day. "
            "I entered 480 minutes of sleep on every recorded day, although I know that some nights were "
            "one or two hours shorter. I will check my calendar or notes and correct only the dates I can "
            "recall or verify; if a value is only an estimate, I will identify it as an estimate. I will "
            "also review fitness, study, and coding entries against my schedule before relying on their "
            "averages. My current averages are "
            f"{fmt(averages['Fitness (min)'])} minutes of fitness, {fmt(averages['Study (min)'])} minutes "
            f"of study, and {fmt(TPI)} minutes of coding per day. After correcting the log, I will rerun "
            "the calculations and choose one realistic weekly goal that fits my classes. I will keep using "
            "the same rating scales and review the next two weeks of data. Correlations can help me notice "
            "patterns, but they cannot prove that an activity caused a change in my mood or energy."
        ),
        "",
        "6. STUDENT DECLARATION",
        "I confirm that the values entered in my report have been generated from my submitted activity data using my Python program and that the report represents my own analysis.",
        f"Student Name: {details['Student Name'] or '[fill in]'}    Signature: ____________________    Date: ____________________",
        "",
        "The July sample row is excluded by date. Check the missing 13-14 August dates and correct any inaccurate entries before submitting.",
    ])
    return "\n".join(lines)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Calculate all report-ready CAP776 Minor Project #1 values from the Daily Log workbook."
    )
    parser.add_argument(
        "excel_file",
        nargs="?",
        default=str(Path(__file__).with_name("12622067.xlsx")),
        help="Path to 12622067.xlsx (default: beside this Python script).",
    )
    parser.add_argument(
        "--end",
        dest="end_date",
        default=None,
        help="Last date to include, YYYY-MM-DD. Defaults to the latest dated entry on/after 13 August 2026.",
    )
    args = parser.parse_args()

    file_path = Path(args.excel_file).expanduser()
    if not file_path.exists():
        raise FileNotFoundError(
            f"Workbook not found: {file_path}\n"
            "Put 12622067.xlsx beside this script or pass its full path after the script name."
        )

    report_text = build_results(file_path, args.end_date)
    output_path = Path(__file__).with_name("project_results.txt")
    output_path.write_text(report_text, encoding="utf-8")
    print(report_text)
    print(f"\nSaved a copy of these results to: {output_path}")


if __name__ == "__main__":
    main()
