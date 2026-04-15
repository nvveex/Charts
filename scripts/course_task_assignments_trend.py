from __future__ import annotations

import argparse
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import matplotlib.pyplot as plt

from .chart_common import configure_logging, init_fonts, load_result_sheet, parse_date, run_single_chart_script


REQUIRED_KEY = "每周课程任务布置数（近"


def render_one(xlsx_path: str | Path, out_path: str | Path) -> None:
    wb, ws = load_result_sheet(xlsx_path)

    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    idx = {str(h).strip(): i + 1 for i, h in enumerate(headers) if h is not None}

    required = ["周起始日", "type", "任务布置数"]
    for k in required:
        if k not in idx:
            raise RuntimeError(f"missing column: {k}; headers={headers}")

    by_week: dict[str, dict[str, int]] = {}
    type_total: dict[str, int] = defaultdict(int)

    for r in range(2, ws.max_row + 1):
        week_v = ws.cell(r, idx["周起始日"]).value
        type_v = ws.cell(r, idx["type"]).value
        cnt_v = ws.cell(r, idx["任务布置数"]).value
        if week_v is None or type_v is None:
            continue

        week = parse_date(week_v)
        t = str(type_v).strip()
        try:
            cnt = int(cnt_v) if cnt_v is not None else 0
        except Exception:
            cnt = 0

        if week not in by_week:
            by_week[week] = {}
        by_week[week][t] = by_week[week].get(t, 0) + cnt
        type_total[t] += cnt

    if not by_week:
        raise RuntimeError("no rows parsed")

    types_sorted = sorted(type_total.items(), key=lambda kv: kv[1], reverse=True)
    top_types = [t for t, _ in types_sorted[:5]]
    if not top_types:
        raise RuntimeError("no types parsed")

    dates = sorted(by_week.keys())
    x = [datetime.strptime(d, "%Y-%m-%d") for d in dates]

    series: dict[str, list[int]] = {t: [] for t in top_types}
    other: list[int] = []
    total: list[int] = []

    for d in dates:
        m = by_week[d]
        s_total = sum(int(v) for v in m.values())
        total.append(s_total)

        s_other = 0
        for t, v in m.items():
            if t in series:
                series[t].append(int(v))
            else:
                s_other += int(v)

        # Ensure equal length for plotting
        for t in top_types:
            if len(series[t]) < len(total):
                series[t].append(0)

        other.append(s_other)

    init_fonts()
    plt.figure(figsize=(14, 6))
    palette = ["#2563EB", "#10B981", "#F59E0B", "#8B5CF6", "#06B6D4"]

    for i, t in enumerate(top_types):
        y = series[t]
        if max(y) == 0:
            continue
        plt.plot(x, y, label=t, linewidth=2, color=palette[i % len(palette)])

    if max(other) > 0:
        plt.plot(x, other, label="其他", linewidth=2, color="#6B7280")

    plt.plot(x, total, label="合计", linewidth=2.5, linestyle="--", color="#111827")
    plt.title("每周课程任务布置数趋势（近14个月）\n数据截至 2026-03-13")
    plt.xlabel("周起始日")
    plt.ylabel("任务布置数")
    plt.grid(True, axis="y", alpha=0.25)
    plt.legend(ncol=4, frameon=False)
    plt.tight_layout()
    plt.savefig(out_path, dpi=180)
    plt.close()


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

