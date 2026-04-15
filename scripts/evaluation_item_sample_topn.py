from __future__ import annotations

import argparse
from collections import defaultdict
from pathlib import Path

import matplotlib.pyplot as plt

from .chart_common import configure_logging, init_fonts, load_result_sheet, run_single_chart_script


REQUIRED_KEY = "近半年评价项名称抽样"
TOP_N = 10


def render_one(xlsx_path: str | Path, out_path: str | Path) -> None:
    wb, ws = load_result_sheet(xlsx_path)

    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    idx = {str(h).strip(): i + 1 for i, h in enumerate(headers) if h is not None}

    required = ["评价项名称", "过评项创建数"]
    for k in required:
        if k not in idx:
            raise RuntimeError(f"missing column: {k}; headers={headers}")

    counts: dict[str, int] = defaultdict(int)
    for r in range(2, ws.max_row + 1):
        name_v = ws.cell(r, idx["评价项名称"]).value
        count_v = ws.cell(r, idx["过评项创建数"]).value
        if name_v is None:
            continue

        name = str(name_v).strip()
        if name == "":
            continue

        try:
            count = int(count_v) if count_v is not None else 0
        except Exception:
            count = 0
        counts[name] += count

    ranking = sorted(counts.items(), key=lambda item: (-item[1], item[0]))[:TOP_N]
    if not ranking:
        raise RuntimeError("no rows parsed")

    labels = [item[0] for item in ranking]
    values = [item[1] for item in ranking]
    y_pos = list(range(len(ranking)))

    init_fonts()
    fig_h = max(5, len(ranking) * 0.65 + 1.8)
    fig, ax = plt.subplots(figsize=(13, fig_h))
    bars = ax.barh(y_pos, values, color="#2563EB", height=0.58)

    x_max = max(values)
    ax.set_xlim(0, x_max * 1.15 if x_max > 0 else 1)

    for i, bar in enumerate(bars):
        width = bar.get_width()
        ax.text(width + max(x_max * 0.015, 0.5), y_pos[i], f"{int(width)}", va="center", ha="left")

    ax.set_yticks(y_pos)
    ax.set_yticklabels(labels)
    ax.invert_yaxis()
    ax.set_title(f"近半年评价项名称抽样 Top {len(ranking)}（按过评项创建数）")
    ax.set_xlabel("过评项创建数")
    ax.grid(True, axis="x", alpha=0.25)

    fig.subplots_adjust(left=0.28)
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
