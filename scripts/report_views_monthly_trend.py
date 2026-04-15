from __future__ import annotations

import argparse
from pathlib import Path

import matplotlib.pyplot as plt

from .chart_common import configure_logging, init_fonts, load_result_sheet, run_single_chart_script


REQUIRED_KEY = "每月自定义报告单查看人次（近"


def render_one(xlsx_path: str | Path, out_path: str | Path) -> None:
    wb, ws = load_result_sheet(xlsx_path)

    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    idx = {str(h).strip(): i + 1 for i, h in enumerate(headers) if h is not None}

    required_cols = ["月份", "查看总人次"]
    for k in required_cols:
        if k not in idx:
            raise RuntimeError(f"missing column: {k}; headers={headers}")

    data: dict[str, int] = {}
    for r in range(2, ws.max_row + 1):
        m = ws.cell(r, idx["月份"]).value
        cnt = ws.cell(r, idx["查看总人次"]).value
        if m is None:
            continue
        mo = str(m).strip()
        try:
            c = int(cnt) if cnt is not None else 0
        except Exception:
            c = 0
        data[mo] = c

    months = sorted(data.keys())
    if not months:
        raise RuntimeError("no rows parsed")

    values = [data[m] for m in months]

    init_fonts()
    plt.figure(figsize=(10, 6))
    bars = plt.bar(months, values, color="#8B5CF6", width=0.6)

    for bar in bars:
        height = bar.get_height()
        plt.text(
            bar.get_x() + bar.get_width() / 2.0,
            height,
            f"{int(height)}",
            ha="center",
            va="bottom",
        )

    plt.title("每月自定义报告单查看人次（近5个月）\n数据截至 2026-03-13")
    plt.xlabel("月份")
    plt.ylabel("查看人次")
    plt.grid(True, axis="y", alpha=0.25)
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

