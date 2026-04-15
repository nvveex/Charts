from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path

import matplotlib.pyplot as plt

from .chart_common import configure_logging, init_fonts, load_result_sheet, parse_date, run_single_chart_script


REQUIRED_KEY = "选课列表"


def render_one(xlsx_path: str | Path, out_path: str | Path) -> None:
    wb, ws = load_result_sheet(xlsx_path)

    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    idx = {str(h).strip(): i + 1 for i, h in enumerate(headers) if h is not None}

    required = ["名称", "学生数量", "已经选中的学生数量", "班级数量", "模式"]
    for k in required:
        if k not in idx:
            raise RuntimeError(f"missing column: {k}; headers={headers}")

    rows: list[tuple[str, int, int, int, str]] = []
    for r in range(2, ws.max_row + 1):
        name_v = ws.cell(r, idx["名称"]).value
        if name_v is None or str(name_v).strip() == "":
            continue
        name = str(name_v).strip()

        total_v = ws.cell(r, idx["学生数量"]).value
        selected_v = ws.cell(r, idx["已经选中的学生数量"]).value
        classes_v = ws.cell(r, idx["班级数量"]).value
        mode_v = ws.cell(r, idx["模式"]).value

        try:
            total = int(total_v) if total_v is not None else 0
        except Exception:
            total = 0
        try:
            selected = int(selected_v) if selected_v is not None else 0
        except Exception:
            selected = 0
        try:
            classes = int(classes_v) if classes_v is not None else 0
        except Exception:
            classes = 0

        mode = str(mode_v).strip() if mode_v is not None else ""
        rows.append((name, total, selected, classes, mode))

    if not rows:
        raise RuntimeError("no rows parsed")

    rows.sort(key=lambda x: x[1], reverse=True)

    labels: list[str] = []
    selected_vals: list[int] = []
    remaining_vals: list[int] = []

    for name, total, selected, classes, mode in rows:
        rate = (selected / total) if total else 0.0
        _ = rate  # computed for readability parity with original code
        title = f"{name}（{mode}，{classes}班）"
        labels.append(title)
        selected_vals.append(max(0, min(selected, total)))
        remaining_vals.append(max(0, total - max(0, min(selected, total))))

    init_fonts()
    fig_h = max(6, 0.45 * len(labels) + 2)
    plt.figure(figsize=(14, fig_h))

    y = list(range(len(labels)))
    plt.barh(y, selected_vals, color="#10B981", label="已选中")
    plt.barh(y, remaining_vals, left=selected_vals, color="#E5E7EB", label="未选中")

    for i, (name, total, selected, classes, mode) in enumerate(rows):
        rate = (selected / total) if total else 0.0
        x_text = total + max(3, total * 0.01)
        plt.text(
            x_text,
            i,
            f"{selected}/{total}（{rate:.1%}）",
            va="center",
            fontsize=10,
            color="#111827",
        )

    plt.yticks(y, labels)
    plt.gca().invert_yaxis()
    plt.title("选课活动覆盖与完成度（按学生数排序）\n数据截至 2026-03-13")
    plt.xlabel("学生数量")
    plt.grid(True, axis="x", alpha=0.2)
    plt.legend(frameon=False, ncol=2, loc="lower right")
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

