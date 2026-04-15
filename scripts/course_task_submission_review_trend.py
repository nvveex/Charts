from __future__ import annotations

import argparse
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import matplotlib.pyplot as plt

from .chart_common import configure_logging, init_fonts, load_result_sheet, parse_date, place_legend_outside, run_single_chart_script


REQUIRED_KEY = "每周课程任务提交与批阅数（近"


def render_one(xlsx_path: str | Path, out_path: str | Path) -> None:
    wb, ws = load_result_sheet(xlsx_path)

    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    idx = {str(h).strip(): i + 1 for i, h in enumerate(headers) if h is not None}

    required = ["周起始日", "状态", "任务数"]
    for k in required:
        if k not in idx:
            raise RuntimeError(f"missing column: {k}; headers={headers}")

    by_week: dict[str, dict[str, int]] = {}
    statuses: set[str] = set()

    for r in range(2, ws.max_row + 1):
        week_v = ws.cell(r, idx["周起始日"]).value
        status_v = ws.cell(r, idx["状态"]).value
        cnt_v = ws.cell(r, idx["任务数"]).value
        if week_v is None or status_v is None:
            continue

        week = parse_date(week_v)
        st = str(status_v).strip()
        statuses.add(st)
        try:
            cnt = int(cnt_v) if cnt_v is not None else 0
        except Exception:
            cnt = 0

        if week not in by_week:
            by_week[week] = {}
        by_week[week][st] = by_week[week].get(st, 0) + cnt

    if not by_week:
        raise RuntimeError("no rows parsed")

    order = ["已批阅", "已提交", "未提交"]
    for st in sorted(statuses):
        if st not in order:
            order.append(st)

    dates = sorted(by_week.keys())
    x = [datetime.strptime(d, "%Y-%m-%d") for d in dates]

    series: dict[str, list[int]] = {st: [] for st in order}
    total: list[int] = []

    for d in dates:
        m = by_week[d]
        total.append(sum(int(v) for v in m.values()))
        for st in order:
            series[st].append(int(m.get(st, 0)))

    init_fonts()
    plt.figure(figsize=(14, 6))

    colors = {"已批阅": "#10B981", "已提交": "#2563EB", "未提交": "#F59E0B"}
    for st in order:
        y = series[st]
        if max(y) == 0:
            continue
        plt.plot(x, y, label=st, linewidth=2, color=colors.get(st))

    plt.plot(x, total, label="合计", linewidth=2.5, linestyle="--", color="#111827")
    plt.title("每周课程任务提交与批阅趋势（近14个月，阈值：5）\n数据截至 2026-03-13")
    plt.xlabel("周起始日")
    plt.ylabel("任务数")
    plt.grid(True, axis="y", alpha=0.25)
    place_legend_outside(plt.gca(), ncol=4)
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
