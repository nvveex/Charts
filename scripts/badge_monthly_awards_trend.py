from __future__ import annotations

import argparse
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import matplotlib.pyplot as plt

from .chart_common import configure_logging, init_fonts, load_result_sheet, parse_date, run_single_chart_script


REQUIRED_KEY = "每月小红花（认证）颁发数量（近"


def render_one(xlsx_path: str | Path, out_path: str | Path) -> None:
    wb, ws = load_result_sheet(xlsx_path)

    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    idx = {str(h).strip(): i + 1 for i, h in enumerate(headers) if h is not None}

    required = ["月", "认证颁发数量"]
    for k in required:
        if k not in idx:
            raise RuntimeError(f"missing column: {k}; headers={headers}")

    month_total: dict[str, int] = defaultdict(int)
    month_name_count: dict[str, set[str]] = defaultdict(set)
    name_col = idx.get("name")

    for r in range(2, ws.max_row + 1):
        month_v = ws.cell(r, idx["月"]).value
        cnt_v = ws.cell(r, idx["认证颁发数量"]).value
        if month_v is None:
            continue
        month = str(month_v).strip()
        try:
            cnt = int(cnt_v) if cnt_v is not None else 0
        except Exception:
            cnt = 0
        month_total[month] += cnt

        if name_col is not None:
            name_v = ws.cell(r, name_col).value
            if name_v is not None and str(name_v).strip() != "":
                month_name_count[month].add(str(name_v).strip())

    months = sorted(month_total.keys())
    if not months:
        raise RuntimeError("no rows parsed")

    x = [datetime.strptime(m + "-01", "%Y-%m-%d") for m in months]
    y = [int(month_total[m]) for m in months]
    y2 = [len(month_name_count.get(m, set())) for m in months]

    init_fonts()
    fig, ax = plt.subplots(figsize=(14, 6))
    ax.bar(x, y, width=20, color="#F59E0B", label="颁发数量（合计）")
    ax.set_title("每月小红花（认证）颁发数量（近13个月）\n数据截至 2026-03-13")
    ax.set_xlabel("月份")
    ax.set_ylabel("颁发数量")
    ax.grid(True, axis="y", alpha=0.25)

    ax2 = ax.twinx()
    if max(y2) > 0:
        ax2.plot(x, y2, color="#2563EB", linewidth=2, label="认证类型数")
        ax2.set_ylabel("认证类型数")

    handles, labels = ax.get_legend_handles_labels()
    handles2, labels2 = ax2.get_legend_handles_labels()
    ax.legend(handles + handles2, labels + labels2, frameon=False, ncol=2, loc="upper left")

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

