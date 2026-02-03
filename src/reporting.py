"""Ticket reporting automation utilities."""
from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path

import numpy as np
import pandas as pd


DATE_FORMAT = "%m/%d/%Y %I:%M:%S %p"


@dataclass(frozen=True)
class ShutdownWindow:
    start: str
    end: str


DEFAULT_SHUTDOWNS = (
    ShutdownWindow(start="2025-08-05", end="2025-08-22"),
    ShutdownWindow(start="2025-12-24", end="2026-01-02"),
)


def _build_shutdown_days(shutdowns: tuple[ShutdownWindow, ...]) -> np.ndarray:
    shutdown_ranges = [
        pd.date_range(start=window.start, end=window.end, freq="B")
        for window in shutdowns
    ]
    if not shutdown_ranges:
        return np.array([], dtype="datetime64[D]")
    return np.concatenate([rng.values for rng in shutdown_ranges]).astype("datetime64[D]")


def _parse_datetime(series: pd.Series) -> pd.Series:
    parsed = pd.to_datetime(series, format=DATE_FORMAT, errors="coerce")
    if parsed.isna().any():
        parsed = parsed.fillna(pd.to_datetime(series, errors="coerce"))
    return parsed


def load_ticket_data(file_path: Path) -> pd.DataFrame:
    data = pd.read_excel(file_path, engine="openpyxl")
    data = data.dropna(axis=1, how="all")
    data["Created"] = _parse_datetime(data["Created"])
    data["Closed"] = _parse_datetime(data["Closed"])
    return data


def calculate_resolution_times(
    tickets: pd.DataFrame,
    report_year: int,
    shutdowns: tuple[ShutdownWindow, ...] = DEFAULT_SHUTDOWNS,
) -> pd.DataFrame:
    shutdown_days = _build_shutdown_days(shutdowns)
    mask = (tickets["Closed"].dt.year == report_year) & (tickets["Closed"].notna())
    calc_df = tickets.loc[mask].copy()

    calc_df["Resolution_time_raw"] = (
        (calc_df["Closed"] - calc_df["Created"]).dt.total_seconds() / 3600
    )
    calc_df["Total_Days"] = (calc_df["Closed"] - calc_df["Created"]).dt.days

    start_arr = calc_df["Created"].values.astype("datetime64[D]")
    end_arr = calc_df["Closed"].values.astype("datetime64[D]")

    working_days = np.busday_count(start_arr, end_arr, holidays=shutdown_days)
    total_days = calc_df["Closed"].sub(calc_df["Created"]).dt.days
    off_days = total_days - working_days

    calc_df["Resolution_time_real"] = (
        calc_df["Resolution_time_raw"] - (off_days * 24)
    ).clip(lower=0)

    calc_df.attrs["shutdown_days"] = shutdown_days
    return calc_df


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Compute resolution times for ticket data.")
    parser.add_argument(
        "file_path",
        type=Path,
        help="Path to the exported Excel file (e.g., data/ExportData.xlsx).",
    )
    parser.add_argument(
        "--year",
        type=int,
        default=2025,
        help="Report year for closed tickets (default: 2025).",
    )
    return parser


def main() -> None:
    parser = _build_parser()
    args = parser.parse_args()

    tickets = load_ticket_data(args.file_path)
    resolved = calculate_resolution_times(tickets, args.year)
    shutdown_days = resolved.attrs.get("shutdown_days", np.array([]))

    preview_cols = [
        col
        for col in ["Ticket ID", "Location", "Resolution_time_raw", "Resolution_time_real"]
        if col in resolved.columns
    ]
    if preview_cols:
        print(resolved[preview_cols].head())
    else:
        print(resolved.head())

    avg_real = resolved["Resolution_time_real"].mean()
    print(f"\nAverage Global Resolution (Business Hours): {avg_real:.2f} hours")
    print(f"Factory Shutdowns accounted for: {len(shutdown_days)} working days removed.")


if __name__ == "__main__":
    main()
