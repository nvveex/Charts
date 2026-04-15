from __future__ import annotations

import argparse
import re
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import matplotlib.pyplot as plt

from .chart_common import configure_logging, init_fonts, load_result_sheet, run_single_chart_script


# Key is derived from file name rule: first '_' before date suffix.
REQUIRED_KEY = "每月教师考核分数录入数（近一年）"


def _parse_month_to_datetime(month_str: str) -> datetime:
    s = month_str.strip()
    if re.match(r"^\d{4}-\d{2}$", s):
        return datetime.strptime(s, "%Y-%m")
    if re.match(r"^\d{4}-\d{2}-\d{2}$", s):
        return datetime.strptime(s, "%Y-%m-%d")
    # Fallback: try ISO parsing.
    return datetime.fromisoformat(s)


def render_one(xlsx_path: str | Path, out_path: str | Path) -> None:
    wb, ws = load_result_sheet(xlsx_path)

    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    idx = {str(h).strip(): i + 1 for i, h in enumerate(headers) if h is not None}

    required = ["月份", "评教名称", "教师评价打分人次"]
    for k in required:
        if k not in idx:
            raise RuntimeError(f"missing column: {k}; headers={headers}")

    month_total: dict[str, int] = defaultdict(int)
    month_names: dict[str, set[str]] = defaultdict(set)

    for r in range(2, ws.max_row + 1):
        m_v = ws.cell(r, idx["月份"]).value
        n_v = ws.cell(r, idx["评教名称"]).value
        c_v = ws.cell(r, idx["教师评价打分人次"]).value
        if m_v is None:
            continue

        m = str(m_v).strip()
        try:
            c = int(c_v) if c_v is not None else 0
        except Exception:
            c = 0

        month_total[m] += c
        if n_v is not None:
            n = str(n_v).strip()
            if n:
                month_names[m].add(n)

    months = sorted(month_total.keys())
    if not months:
        raise RuntimeError("no rows parsed")

    x = [_parse_month_to_datetime(m) for m in months]
    y = [int(month_total[m]) for m in months]
    y2 = [len(month_names.get(m, set())) for m in months]

    init_fonts()
    fig, ax = plt.subplots(figsize=(14, 6))

    ax.bar(x, y, width=20, color="#F59E0B", label="教师评价打分人次（合计）")
    last_date_str = x[-1].date().isoformat() if x else ""
    ax.set_title(f"每月教师考核分数录入趋势（近一年）\n数据截至 {last_date_str}")
    ax.set_xlabel("月份")
    ax.set_ylabel("录入数")
    ax.grid(True, axis="y", alpha=0.25)

    ax2 = ax.twinx()
    if max(y2) > 0:
        ax2.plot(x, y2, color="#2563EB", linewidth=2, label="评教类型数")
        ax2.set_ylabel("评教类型数")

    h1, l1 = ax.get_legend_handles_labels()
    h2, l2 = ax2.get_legend_handles_labels()
    ax.legend(h1 + h2, l1 + l2, frameon=False, ncol=2, loc="upper left")

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

