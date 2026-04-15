from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path

import matplotlib.pyplot as plt

from .chart_common import configure_logging, init_fonts, load_result_sheet, parse_date, place_legend_outside, run_single_chart_script


REQUIRED_KEY = "每周排课操作数（近一年）"


def render_one(xlsx_path: str | Path, out_path: str | Path) -> None:
    wb, ws = load_result_sheet(xlsx_path)

    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    idx = {str(h).strip(): i + 1 for i, h in enumerate(headers) if h is not None}
    required = ["week_start", "snapshot_count"]
    for k in required:
        if k not in idx:
            raise RuntimeError(f"missing column: {k}; headers={headers}")

    data: dict[str, int] = {}
    for r in range(2, ws.max_row + 1):
        d_v = ws.cell(r, idx["week_start"]).value
        c_v = ws.cell(r, idx["snapshot_count"]).value
        if d_v is None:
            continue
        d = parse_date(d_v)
        try:
            c = int(c_v) if c_v is not None else 0
        except Exception:
            c = 0
        data[d] = data.get(d, 0) + c

    if not data:
        raise RuntimeError("no rows parsed")

    dates = sorted(data.keys())
    x = [datetime.strptime(d, "%Y-%m-%d") for d in dates]
    y = [int(data[d]) for d in dates]

    init_fonts()
    plt.figure(figsize=(14, 6))
    plt.plot(x, y, linewidth=2.2, color="#2563EB", label="排课操作数")
    plt.title("每周排课操作数趋势（近一年）\n数据截至 2026-03-13")
    plt.xlabel("周起始日")
    plt.ylabel("操作数")
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
