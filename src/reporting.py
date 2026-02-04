"""Ticket reporting automation utilities."""
from __future__ import annotations

import argparse
import base64
from io import BytesIO
from dataclasses import dataclass
from pathlib import Path

import matplotlib.pyplot as plt
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
    parser.add_argument(
        "--output-report",
        type=Path,
        default=Path("report.html"),
        help="Output path for the one-page HTML report (default: report.html).",
    )
    return parser


def _encode_chart() -> str:
    buffer = BytesIO()
    plt.tight_layout()
    plt.savefig(buffer, format="png", dpi=144, bbox_inches="tight")
    plt.close()
    buffer.seek(0)
    return base64.b64encode(buffer.read()).decode("utf-8")


def _format_hours(value: float | int) -> str:
    return f"{value:.1f} h"


def _select_submitter_column(tickets: pd.DataFrame) -> str | None:
    for candidate in (
        "Submitted By",
        "Submitter",
        "Requester",
        "Opened By",
        "Created By",
        "User",
    ):
        if candidate in tickets.columns:
            return candidate
    return None


def _select_location_column(tickets: pd.DataFrame) -> str | None:
    for candidate in ("Location", "Site", "Office", "Plant"):
        if candidate in tickets.columns:
            return candidate
    return None


def _select_escalation_column(tickets: pd.DataFrame) -> str | None:
    for candidate in ("Escalated", "Escalation", "Escalation Status", "Escalation Level"):
        if candidate in tickets.columns:
            return candidate
    return None


def _select_reaction_escalation_column(tickets: pd.DataFrame) -> str | None:
    for candidate in (
        "Reaction Time Escalated Trigger",
        "Reaction Escalated Trigger",
        "Reaction Time Escalated",
        "Reaction Escalated",
        "Reaction Escalation",
    ):
        if candidate in tickets.columns:
            return candidate
    return None


def _select_solution_escalation_column(tickets: pd.DataFrame) -> str | None:
    for candidate in (
        "Solution Time Escalated Trigger",
        "Solution Escalated Trigger",
        "Solution Time Escalated",
        "Solution Escalated",
        "Solution Escalation",
    ):
        if candidate in tickets.columns:
            return candidate
    return None


def _build_top_bar_chart(tickets: pd.DataFrame, column: str, title: str) -> tuple[str, str]:
    counts = tickets[column].fillna("Unknown").value_counts().head(10)
    palette = ["#A7C7E7", "#F7C8E0", "#B4F8C8", "#FDE2A7", "#CDB4DB", "#B8E0D2"]
    plt.figure(figsize=(7.2, 4.6))
    plt.barh(counts.index.astype(str), counts.values, color=palette[0])
    plt.gca().invert_yaxis()
    plt.title(title)
    plt.xlabel("Tickets")
    plt.ylabel(column)
    return (title, _encode_chart())


def _build_report_charts(resolved: pd.DataFrame) -> list[tuple[str, str]]:
    charts: list[tuple[str, str]] = []
    palette = ["#A7C7E7", "#F7C8E0", "#B4F8C8", "#FDE2A7", "#CDB4DB", "#B8E0D2"]

    if "Created" in resolved.columns:
        monthly_total = (
            resolved["Created"]
            .dt.to_period("M")
            .value_counts()
            .sort_index()
        )
        reaction_column = _select_reaction_escalation_column(resolved)
        solution_column = _select_solution_escalation_column(resolved)

        reaction_series = None
        solution_series = None
        if reaction_column:
            reaction_mask = resolved[reaction_column].astype(str).str.contains(
                "1|true|yes", case=False, na=False
            )
            reaction_series = (
                resolved.loc[reaction_mask, "Created"]
                .dt.to_period("M")
                .value_counts()
                .sort_index()
            ).reindex(monthly_total.index, fill_value=0)

        if solution_column:
            solution_mask = resolved[solution_column].astype(str).str.contains(
                "1|true|yes", case=False, na=False
            )
            solution_series = (
                resolved.loc[solution_mask, "Created"]
                .dt.to_period("M")
                .value_counts()
                .sort_index()
            ).reindex(monthly_total.index, fill_value=0)

        plt.figure(figsize=(8.4, 4.8))
        plt.plot(
            monthly_total.index.astype(str),
            monthly_total.values,
            marker="o",
            color=palette[0],
            linewidth=2.5,
            label="Total Tickets",
        )
        if reaction_series is not None:
            plt.plot(
                monthly_total.index.astype(str),
                reaction_series.values,
                marker="o",
                color=palette[1],
                linewidth=2.5,
                label="Reaction Time Escalated",
            )
        if solution_series is not None:
            plt.plot(
                monthly_total.index.astype(str),
                solution_series.values,
                marker="o",
                color=palette[2],
                linewidth=2.5,
                label="Solution Time Escalated",
            )
        plt.title("Ticket Volume vs Escalations (Monthly)")
        plt.xlabel("Month")
        plt.ylabel("Ticket Count")
        plt.xticks(rotation=35, ha="right")
        plt.legend(frameon=False)
        charts.append(("Ticket Volume vs Escalations (Monthly)", _encode_chart()))

    return charts


def build_html_report(resolved: pd.DataFrame, report_year: int, output_path: Path) -> None:
    median_hours = resolved["Resolution_time_real"].median()
    pct_90 = resolved["Resolution_time_real"].quantile(0.9)
    pct_10 = resolved["Resolution_time_real"].quantile(0.1)
    tickets_count = len(resolved)
    median_class = "kpi-good" if median_hours < 6 else ""

    charts = _build_report_charts(resolved)
    title = f"Report technical support Bourdon {report_year}"
    submitter_column = _select_submitter_column(resolved)
    location_column = _select_location_column(resolved)

    chart_blocks = ""
    if charts:
        chart_blocks = "\n".join(
            f"""
            <div class=\"chart\">
                <h3>{label}</h3>
                <img src=\"data:image/png;base64,{image}\" alt=\"{label}\" />
            </div>
            """
            for label, image in charts
        )
    else:
        chart_blocks = "<p class=\"placeholder\">No charts available for the current data.</p>"

    table_blocks = ""
    if submitter_column:
        submitter_title = "Top 10 Tickets Submitted by User"
        submitter_chart = _build_top_bar_chart(resolved, submitter_column, submitter_title)
        table_blocks += f"""
            <div class=\"chart\">
                <h3>{submitter_chart[0]}</h3>
                <img src=\"data:image/png;base64,{submitter_chart[1]}\" alt=\"{submitter_chart[0]}\" />
            </div>
        """
    if location_column:
        location_title = "Top 10 Tickets Submitted by Location"
        location_chart = _build_top_bar_chart(resolved, location_column, location_title)
        table_blocks += f"""
            <div class=\"chart\">
                <h3>{location_chart[0]}</h3>
                <img src=\"data:image/png;base64,{location_chart[1]}\" alt=\"{location_chart[0]}\" />
            </div>
        """
    if not table_blocks:
        table_blocks = "<p class=\"placeholder\">No ticket breakdown charts available.</p>"

    html = f"""
    <!DOCTYPE html>
    <html lang=\"en\">
    <head>
        <meta charset=\"UTF-8\" />
        <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\" />
        <title>{title}</title>
        <link href=\"https://fonts.googleapis.com/css2?family=Roboto+Condensed:wght@300;400;600;700&display=swap\" rel=\"stylesheet\">
        <style>
            :root {{
                --bg: #f7f6fb;
                --card: #ffffff;
                --text: #2b2b2b;
                --muted: #6f6f6f;
                --accent: #a7c7e7;
                --success: #43a047;
            }}
            body {{
                margin: 0;
                font-family: "Roboto Condensed", "Segoe UI", system-ui, sans-serif;
                color: var(--text);
                background: var(--bg);
            }}
            .page {{
                max-width: 1400px;
                margin: 0 auto;
                padding: 28px 36px 44px;
            }}
            header {{
                display: flex;
                justify-content: space-between;
                align-items: center;
                margin-bottom: 24px;
            }}
            header h1 {{
                margin: 0;
                font-size: 28px;
            }}
            header span {{
                color: var(--muted);
                font-size: 14px;
            }}
            .kpi-grid {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
                gap: 16px;
                margin-bottom: 28px;
            }}
            .kpi-card {{
                background: var(--card);
                padding: 18px 20px;
                border-radius: 16px;
                box-shadow: 0 8px 18px rgba(35, 35, 35, 0.08);
            }}
            .kpi-card h2 {{
                margin: 0 0 6px;
                font-size: 14px;
                color: var(--muted);
                font-weight: 600;
                letter-spacing: 0.04em;
                text-transform: uppercase;
            }}
            .kpi-card p {{
                margin: 0;
                font-size: 26px;
                font-weight: 700;
                color: var(--text);
            }}
            .kpi-card p.kpi-good {{
                color: var(--success);
            }}
            .charts {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(360px, 1fr));
                gap: 18px;
            }}
            .chart {{
                background: var(--card);
                padding: 16px;
                border-radius: 16px;
                box-shadow: 0 8px 18px rgba(35, 35, 35, 0.08);
            }}
            .chart h3 {{
                margin: 0 0 10px;
                font-size: 16px;
            }}
            .chart img {{
                width: 100%;
                border-radius: 12px;
                background: #ffffff;
            }}
            .tables {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(360px, 1fr));
                gap: 18px;
                margin-top: 22px;
            }}
            .placeholder {{
                background: var(--card);
                padding: 20px;
                border-radius: 12px;
                color: var(--muted);
            }}
        </style>
    </head>
    <body>
        <div class=\"page\">
            <header>
                <div>
                    <h1>{title}</h1>
                    <span>One-page KPI summary</span>
                </div>
            </header>
            <section class=\"kpi-grid\">
                <div class=\"kpi-card\">
                    <h2>Median Answering Time</h2>
                    <p class=\"{median_class}\">{_format_hours(median_hours)}</p>
                </div>
                <div class=\"kpi-card\">
                    <h2>Tickets Treated</h2>
                    <p>{tickets_count}</p>
                </div>
                <div class=\"kpi-card\">
                    <h2>90th Percentile</h2>
                    <p>{_format_hours(pct_90)}</p>
                </div>
                <div class=\"kpi-card\">
                    <h2>10th Percentile</h2>
                    <p>{_format_hours(pct_10)}</p>
                </div>
            </section>
            <section class=\"charts\">
                {chart_blocks}
            </section>
            <section class=\"tables\">
                {table_blocks}
            </section>
        </div>
    </body>
    </html>
    """
    output_path.write_text(html, encoding="utf-8")


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

    build_html_report(resolved, args.year, args.output_report)
    print(f"Report written to {args.output_report.resolve()}")


if __name__ == "__main__":
    main()
