from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path

import matplotlib.dates as mdates
import matplotlib.pyplot as plt

from .chart_common import configure_logging, init_fonts, load_result_sheet, parse_date, run_single_chart_script


REQUIRED_KEY = "最近四个正式学期"


def render_one(xlsx_path: str | Path, out_path: str | Path) -> None:
    wb, ws = load_result_sheet(xlsx_path)

    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    idx = {str(h).strip(): i + 1 for i, h in enumerate(headers) if h is not None}

    required = ["学期名称", "开始日期", "结束日期"]
    for k in required:
        if k not in idx:
            raise RuntimeError(f"missing column: {k}; headers={headers}")

    terms: list[tuple[str, datetime, datetime]] = []
    for r in range(2, ws.max_row + 1):
        name_v = ws.cell(r, idx["学期名称"]).value
        start_v = ws.cell(r, idx["开始日期"]).value
        end_v = ws.cell(r, idx["结束日期"]).value
        if name_v is None or start_v is None or end_v is None:
            continue

        start = datetime.strptime(parse_date(start_v), "%Y-%m-%d")
        end = datetime.strptime(parse_date(end_v), "%Y-%m-%d")
        if end < start:
            raise RuntimeError(f"end date earlier than start date: {name_v}")
        terms.append((str(name_v).strip(), start, end))

    if not terms:
        raise RuntimeError("no rows parsed")

    terms.sort(key=lambda item: item[1])
    labels = [item[0] for item in terms]
    starts = [mdates.date2num(item[1]) for item in terms]
    durations = [(item[2] - item[1]).days + 1 for item in terms]
    y_pos = list(range(len(terms)))

    init_fonts()
    fig_h = max(4.5, len(terms) * 1.1 + 1.5)
    fig, ax = plt.subplots(figsize=(14, fig_h))

    colors = ["#2563EB", "#10B981", "#F59E0B", "#7C3AED"]
    bars = ax.barh(
        y_pos,
        durations,
        left=starts,
        height=0.58,
        color=[colors[i % len(colors)] for i in range(len(terms))],
    )

    for i, bar in enumerate(bars):
        duration = durations[i]
        start = terms[i][1]
        end = terms[i][2]
        center_x = starts[i] + duration / 2
        ax.text(
            center_x,
            y_pos[i],
            f"{duration}天",
            ha="center",
            va="center",
            color="white",
            fontsize=9,
            fontweight="bold",
        )
        ax.text(
            starts[i] + duration + 5,
            y_pos[i],
            f"{start:%Y-%m-%d} - {end:%Y-%m-%d}",
            va="center",
            ha="left",
            fontsize=9,
            color="#374151",
        )

    ax.set_yticks(y_pos)
    ax.set_yticklabels(labels)
    ax.invert_yaxis()
    ax.set_title(f"最近正式学期时间轴（{len(terms)}个学期）")
    ax.set_xlabel("日期")
    ax.grid(True, axis="x", alpha=0.25)
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=1))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m"))
    fig.autofmt_xdate(rotation=30, ha="right")
    fig.subplots_adjust(left=0.28, right=0.88)
    fig.tight_layout()
    fig.savefig(out_path, dpi=180)
    plt.close(fig)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Render one chart from matched Excel files.")
    parser.add_argument("--data-dir", required=True)
    parser.add_argument("--output-dir", required=True)
    args = parser.parse_args(argv)

    configure_logging(Path(__file__).stem)

    try:
        result = run_single_chart_script(
            data_dir=args.data_dir,
            output_dir=args.output_dir,
            required_key=REQUIRED_KEY,
            render_one=render_one,
        )
    except Exception as e:
        import logging

        logging.exception("script failed: %s", e)
        return 1

    if result.failures:
        import logging

        logging.error("script completed with failures: %d", len(result.failures))
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
