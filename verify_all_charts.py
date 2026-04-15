from __future__ import annotations

import json
import traceback
from datetime import datetime, timezone
from importlib import import_module
from pathlib import Path

from scripts.chart_common import ensure_dir, run_single_chart_script, scan_xlsx_by_key


ROOT = Path(__file__).resolve().parent
ORIGINALDATA_DIR = ROOT / "originaldata"
CHARTS_DIR = ROOT / "charts"

def discover_script_modules() -> list[str]:
    """
    Discover candidate chart scripts under scripts/.

    Only keeps modules that expose REQUIRED_KEY + render_one.
    """
    modules: list[str] = []
    scripts_dir = ROOT / "scripts"
    for p in sorted(scripts_dir.glob("*.py")):
        if p.name in {"__init__.py", "chart_common.py"}:
            continue
        if p.stem.startswith("_"):
            continue
        try:
            m = import_module(f"scripts.{p.stem}")
        except Exception:
            # If import fails, still record it later via exception path in main.
            modules.append(p.stem)
            continue
        if hasattr(m, "REQUIRED_KEY") and hasattr(m, "render_one"):
            modules.append(p.stem)
    return modules


def expected_png_paths(charts_dir: Path, key: str, match_count: int) -> list[Path]:
    if match_count <= 1:
        return [charts_dir / f"{key}.png"]
    return [charts_dir / f"{key}_{i}.png" for i in range(1, match_count + 1)]


def main() -> int:
    ensure_dir(CHARTS_DIR)
    ensure_dir(ORIGINALDATA_DIR)

    # Clear old chart outputs to avoid false positives.
    for p in CHARTS_DIR.glob("*.png"):
        try:
            p.unlink()
        except OSError:
            pass

    discovered_modules = discover_script_modules()

    report: dict[str, object] = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "python": str(Path(__import__("sys").executable)),
        "scripts": [],
    }

    failed = 0
    ok = 0
    skipped = 0

    for module_name in discovered_modules:
        case: dict[str, object] = {"module": module_name}
        try:
            m = import_module(f"scripts.{module_name}")
            required_key: str = getattr(m, "REQUIRED_KEY")
            render_one = getattr(m, "render_one")

            matches = scan_xlsx_by_key(ORIGINALDATA_DIR, required_key)
            match_count = len(matches)

            # If no input xlsx matches this REQUIRED_KEY, skip.
            if match_count == 0:
                skipped += 1
                case.update(
                    {
                        "required_key": required_key,
                        "matched_files": [],
                        "expected_png": [],
                        "missing_png": [],
                        "failures": [],
                        "status": "skipped",
                        "reason": "no input xlsx matched REQUIRED_KEY",
                    }
                )
            else:
                expected_paths = expected_png_paths(CHARTS_DIR, required_key, match_count)

                result = run_single_chart_script(
                    data_dir=ORIGINALDATA_DIR,
                    output_dir=CHARTS_DIR,
                    required_key=required_key,
                    render_one=render_one,
                )

                missing = [p.name for p in expected_paths if not p.exists()]
                status = "ok" if (not result.failures and not missing) else "failed"
                if status == "ok":
                    ok += 1
                else:
                    failed += 1

                case.update(
                    {
                        "required_key": required_key,
                        "matched_files": [p.name for p in matches],
                        "expected_png": [p.name for p in expected_paths],
                        "missing_png": missing,
                        "failures": [(p.name, msg) for p, msg in result.failures],
                        "status": status,
                    }
                )
        except Exception:
            failed += 1
            case.update(
                {
                    "status": "failed",
                    "traceback": traceback.format_exc(),
                }
            )

        report["scripts"].append(case)

    report_path = CHARTS_DIR / "verify_report.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    total_cnt = len(report["scripts"])
    print(f"Chart verification finished: ok={ok}, skipped={skipped}, failed={failed}, total={total_cnt}")
    print(f"Report: {report_path}")
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())

