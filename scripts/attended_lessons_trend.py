from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path

import matplotlib.pyplot as plt

from .chart_common import configure_logging, init_fonts, load_result_sheet, parse_date, place_legend_outside, run_single_chart_script


REQUIRED_KEY = "每周已考勤课节数（近"


def render_one(xlsx_path: str | Path, out_path: str | Path) -> None:
    wb, ws = load_result_sheet(xlsx_path)

    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    idx = {str(h).strip(): i + 1 for i, h in enumerate(headers) if h is not None}
    required = ["周起", "已考勤课节数"]
    for k in required:
        if k not in idx:
            raise RuntimeError(f"missing column: {k}; headers={headers}")

    data: dict[str, int] = {}
    for r in range(2, ws.max_row + 1):
        d = ws.cell(r, idx["周起"]).value
        cnt = ws.cell(r, idx["已考勤课节数"]).value
        if d is None:
            continue
        ds = parse_date(d)
        try:
            c = int(cnt) if cnt is not None else 0
        except Exception:
            c = 0
        data[ds] = data.get(ds, 0) + c

    dates = sorted(data.keys())
    if not dates:
        raise RuntimeError("no rows parsed")

    values = [data[d] for d in dates]
    x = [datetime.strptime(d, "%Y-%m-%d") for d in dates]

    init_fonts()
    plt.figure(figsize=(14, 6))
    plt.plot(x, values, label="已考勤课节数", linewidth=2, color="#10B981")
    plt.title("每周已考勤课节数趋势（近13个月）\n数据截至 2026-03-13")
    plt.xlabel("周起始日")
    plt.ylabel("课节数量")
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
