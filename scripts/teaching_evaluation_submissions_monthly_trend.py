from __future__ import annotations

import argparse
from pathlib import Path

import matplotlib.pyplot as plt

from .chart_common import configure_logging, init_fonts, load_result_sheet, run_single_chart_script


REQUIRED_KEY = "每月评教提交人次（近"


def render_one(xlsx_path: str | Path, out_path: str | Path) -> None:
    wb, ws = load_result_sheet(xlsx_path)

    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    idx = {str(h).strip(): i + 1 for i, h in enumerate(headers) if h is not None}

    month_col = "月"
    name_col = "评教名称" if "评教名称" in idx else "name" if "name" in idx else None
    count_col = "评教提交次数"

    required_cols = [month_col, count_col]
    if name_col is None:
        raise RuntimeError(f"missing column: 评教名称/name; headers={headers}")
    for k in required_cols:
        if k not in idx:
            raise RuntimeError(f"missing column: {k}; headers={headers}")

    data: dict[str, list[tuple[str, int]]] = {}
    for r in range(2, ws.max_row + 1):
        m = ws.cell(r, idx[month_col]).value
        name = ws.cell(r, idx[name_col]).value
        cnt = ws.cell(r, idx[count_col]).value
        if m is None:
            continue
        mo = str(m).strip()
        name_str = str(name).strip() if name is not None else ""
        try:
            c = int(cnt) if cnt is not None else 0
        except Exception:
            c = 0

        data.setdefault(mo, []).append((name_str, c))

    months_sorted = sorted(data.keys())
    if not months_sorted:
        raise RuntimeError("no rows parsed")

    display_months: list[str] = []
    names: list[str] = []
    values: list[int] = []
    for month in months_sorted:
        for name, cnt in data[month]:
            display_months.append(month)
            names.append(name)
            values.append(cnt)

    init_fonts()
    plt.figure(figsize=(12, 8))
    bars = plt.bar(display_months, values, color="#4F46E5", width=0.5)

    for bar, name in zip(bars, names):
        height = bar.get_height()
        plt.text(
            bar.get_x() + bar.get_width() / 2.0,
            height,
            f"{name}\n({int(height)})",
            ha="center",
            va="bottom",
            fontsize=9,
            fontweight="bold",
            wrap=True,
        )

    plt.title("每月评教提交人次（近 14 个月）", fontsize=14, pad=40)
    plt.xlabel("月份", fontsize=12)
    plt.ylabel("评教提交次数", fontsize=12)

    if values:
        plt.ylim(0, max(values) * 1.3)

    plt.xticks(fontsize=10)
    plt.grid(True, axis="y", alpha=0.25)
    plt.tight_layout()
    plt.savefig(out_path, dpi=180)
    plt.close()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Auto rich chart for one key.")
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
    return 0 if not result.failures else 1


if __name__ == "__main__":
    raise SystemExit(main())
