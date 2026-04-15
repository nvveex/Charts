from __future__ import annotations

import argparse
from pathlib import Path

from .chart_common import auto_render_rich_chart, configure_logging, run_single_chart_script


REQUIRED_KEY = "每月校历事件数（近一年）"


def render_one(xlsx_path: str | Path, out_path: str | Path) -> None:
    auto_render_rich_chart(
        xlsx_path=xlsx_path,
        out_path=out_path,
        title="每月校历事件数趋势（按类型）",
    )


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

