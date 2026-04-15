from __future__ import annotations

import argparse
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import matplotlib.pyplot as plt

from .chart_common import configure_logging, init_fonts, load_result_sheet, parse_date, run_single_chart_script


REQUIRED_KEY = "每周课表查询次数（近一年）"


def render_one(xlsx_path: str | Path, out_path: str | Path) -> None:
    wb, ws = load_result_sheet(xlsx_path)

    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    idx = {str(h).strip(): i + 1 for i, h in enumerate(headers) if h is not None}
    required = ["周开始(周一)", "类型", "访问次数"]
    for k in required:
        if k not in idx:
            raise RuntimeError(f"missing column: {k}; headers={headers}")

    by_week: dict[str, dict[str, int]] = defaultdict(lambda: defaultdict(int))
    type_total: dict[str, int] = defaultdict(int)

    for r in range(2, ws.max_row + 1):
        d_v = ws.cell(r, idx["周开始(周一)"]).value
        t_v = ws.cell(r, idx["类型"]).value
        c_v = ws.cell(r, idx["访问次数"]).value
        if d_v is None or t_v is None:
            continue
        d = parse_date(d_v)
        t = str(t_v).strip()
        try:
            c = int(c_v) if c_v is not None else 0
        except Exception:
            c = 0
        by_week[d][t] += c
        type_total[t] += c

    weeks = sorted(by_week.keys())
    if not weeks:
        raise RuntimeError("no rows parsed")

    top_types = [t for t, _ in sorted(type_total.items(), key=lambda kv: kv[1], reverse=True)[:5]]
    x = [datetime.strptime(d, "%Y-%m-%d") for d in weeks]
    total = [sum(by_week[d].values()) for d in weeks]

    series: dict[str, list[int]] = {t: [] for t in top_types}
    for d in weeks:
        m = by_week[d]
        for t in top_types:
            series[t].append(int(m.get(t, 0)))

    init_fonts()
    plt.figure(figsize=(14, 6))
    palette = ["#10B981", "#2563EB", "#F59E0B", "#8B5CF6", "#06B6D4"]

    for i, t in enumerate(top_types):
        y = series[t]
        if max(y) == 0:
            continue
        plt.plot(x, y, linewidth=2, label=t, color=palette[i % len(palette)])

    plt.plot(x, total, linewidth=2.5, linestyle="--", color="#111827", label="合计")
    plt.title("每周课表查询次数趋势（近一年）\n数据截至 2026-03-13")
    plt.xlabel("周起始日")
    plt.ylabel("访问次数")
    plt.grid(True, axis="y", alpha=0.25)
    plt.legend(frameon=False, ncol=3)
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

