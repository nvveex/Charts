from __future__ import annotations

import argparse
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import matplotlib.pyplot as plt

from .chart_common import configure_logging, init_fonts, load_result_sheet, parse_date, place_legend_outside, run_single_chart_script


REQUIRED_KEY = "每周学生"


def simplify_leave_type(tp: str) -> str:
    # Keep the original business grouping logic.
    if "代请假" in tp:
        return "导师代请假 (Tutor Proxy)"
    if "批量" in tp:
        return "批量请假 (Batch)"
    if "学生请假" in tp:
        return "学生自请假 (Student Self)"
    return "其他 (Other)"


def render_one(xlsx_path: str | Path, out_path: str | Path) -> None:
    wb, ws = load_result_sheet(xlsx_path)

    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    idx = {str(h).strip(): i + 1 for i, h in enumerate(headers) if h is not None}

    required_cols = ["周起", "类型", "数量"]
    for k in required_cols:
        if k not in idx:
            raise RuntimeError(f"missing column: {k}; headers={headers}")

    by_week: dict[str, dict[str, int]] = defaultdict(lambda: defaultdict(int))
    types: set[str] = set()

    for r in range(2, ws.max_row + 1):
        d = ws.cell(r, idx["周起"]).value
        t_v = ws.cell(r, idx["类型"]).value
        cnt_v = ws.cell(r, idx["数量"]).value
        if d is None or t_v is None:
            continue

        ds = parse_date(d)
        tp = simplify_leave_type(str(t_v).strip())

        try:
            c = int(cnt_v) if cnt_v is not None else 0
        except Exception:
            c = 0

        types.add(tp)
        by_week[ds][tp] += c

    dates = sorted(by_week.keys())
    if not dates:
        raise RuntimeError("no rows parsed")

    x = [datetime.strptime(d, "%Y-%m-%d") for d in dates]
    type_list = sorted(types)
    if not type_list:
        raise RuntimeError("no leave types parsed")

    series: dict[str, list[int]] = {t: [] for t in type_list}
    total: list[int] = []

    for d in dates:
        m = by_week[d]
        s_total = 0
        for t in type_list:
            val = int(m.get(t, 0))
            series[t].append(val)
            s_total += val
        total.append(s_total)

    init_fonts()
    plt.figure(figsize=(14, 6))
    colors = ["#2563EB", "#F59E0B", "#10B981", "#8B5CF6"]

    bottom = [0] * len(dates)
    for i, t in enumerate(type_list):
        y = series[t]
        if max(y) == 0:
            continue
        plt.bar(x, y, bottom=bottom, label=t, color=colors[i % len(colors)], width=5)
        bottom = [b + v for b, v in zip(bottom, y)]

    plt.plot(x, total, label="合计 (Total)", linewidth=2, linestyle="--", color="#111827")
    plt.title("每周学生请假数趋势（按申请类型，近一年）\n数据截至 2026-03-13")
    plt.xlabel("周起始日")
    plt.ylabel("请假数量")
    plt.grid(True, axis="y", alpha=0.25)
    place_legend_outside(plt.gca())
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
