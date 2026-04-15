from __future__ import annotations

import argparse
from collections import defaultdict
from pathlib import Path

import matplotlib.pyplot as plt

from .chart_common import configure_logging, init_fonts, load_result_sheet, run_single_chart_script


REQUIRED_KEY = "考试创建数"


def render_one(xlsx_path: str | Path, out_path: str | Path) -> None:
    wb, ws = load_result_sheet(xlsx_path)

    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    idx = {str(h).strip(): i + 1 for i, h in enumerate(headers) if h is not None}

    required = ["月", "考试创建数", "考场状态"]
    for k in required:
        if k not in idx:
            raise RuntimeError(f"missing column: {k}; headers={headers}")

    by_month: dict[str, dict[str, int]] = defaultdict(lambda: defaultdict(int))
    status_total: dict[str, int] = defaultdict(int)

    for r in range(2, ws.max_row + 1):
        month_v = ws.cell(r, idx["月"]).value
        count_v = ws.cell(r, idx["考试创建数"]).value
        status_v = ws.cell(r, idx["考场状态"]).value
        if month_v is None:
            continue

        month = str(month_v).strip()
        status = str(status_v).strip() if status_v is not None and str(status_v).strip() != "" else "未知"
        try:
            count = int(count_v) if count_v is not None else 0
        except Exception:
            count = 0

        by_month[month][status] += count
        status_total[status] += count

    months = sorted(by_month.keys())
    if not months:
        raise RuntimeError("no rows parsed")

    statuses = [name for name, _ in sorted(status_total.items(), key=lambda item: (-item[1], item[0]))]
    x = list(range(len(months)))
    series: dict[str, list[int]] = {status: [] for status in statuses}
    totals: list[int] = []

    for month in months:
        total = 0
        month_data = by_month[month]
        for status in statuses:
            value = int(month_data.get(status, 0))
            series[status].append(value)
            total += value
        totals.append(total)

    init_fonts()
    fig, ax = plt.subplots(figsize=(12, 6))

    palette = ["#2563EB", "#10B981", "#F59E0B", "#7C3AED", "#06B6D4", "#EF4444"]
    bottom = [0] * len(months)
    for i, status in enumerate(statuses):
        values = series[status]
        if max(values) == 0:
            continue
        ax.bar(
            x,
            values,
            bottom=bottom,
            width=0.7,
            label=status,
            color=palette[i % len(palette)],
        )
        bottom = [b + v for b, v in zip(bottom, values)]

    ax.plot(x, totals, label="合计", linewidth=2.2, linestyle="--", color="#111827", marker="o")
    ax.set_xticks(x)
    ax.set_xticklabels(months, rotation=30, ha="right")
    ax.set_title(f"考试创建数（月度，按考场状态）\n数据截至 {months[-1]}")
    ax.set_xlabel("月份")
    ax.set_ylabel("考试创建数")
    ax.grid(True, axis="y", alpha=0.25)
    ax.legend(frameon=False, ncol=3, loc="upper left")

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
