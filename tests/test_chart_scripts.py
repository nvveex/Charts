from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
import unittest
from importlib import import_module
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
VENV_PYTHON = ROOT / ".chartvenv" / "bin" / "python"
REPO_ORIGINALDATA = ROOT / "originaldata"

sys.path.insert(0, str(ROOT))

from scripts.chart_common import key_from_filename  # noqa: E402


def discover_script_modules() -> list[str]:
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
            continue
        if hasattr(m, "REQUIRED_KEY") and hasattr(m, "main"):
            modules.append(p.stem)
    return modules


class TestKeyExtraction(unittest.TestCase):
    def test_key_from_filename_rules(self) -> None:
        self.assertEqual(
            key_from_filename("每周德育分数录入数（近_13_个月）_2026_03_13.xlsx"),
            "每周德育分数录入数（近",
        )
        self.assertEqual(
            key_from_filename("周活（近_14_个月）_2026_03_13 (1).xlsx"),
            "周活（近",
        )
        self.assertEqual(key_from_filename("考试分数录入数_2026_03_13.xlsx"), "考试分数录入数")
        self.assertEqual(key_from_filename("选课列表.xlsx"), "选课列表")


class TestScriptIntegration(unittest.TestCase):
    def test_scripts_generate_expected_outputs(self) -> None:
        self.assertTrue(VENV_PYTHON.exists(), f"missing python: {VENV_PYTHON}")

        modules = discover_script_modules()
        self.assertGreater(len(modules), 0, "no scripts discovered")

        # Build a map from key -> sample xlsx in repo originaldata
        key_to_sample: dict[str, Path] = {}
        for p in REPO_ORIGINALDATA.glob("*.xlsx"):
            key_to_sample.setdefault(key_from_filename(p), p)

        for module in modules:
            m = import_module(f"scripts.{module}")
            required_key = getattr(m, "REQUIRED_KEY")
            sample = key_to_sample.get(required_key)
            if sample is None:
                # No real input for this key in repo; skip this module.
                continue

            with self.subTest(module=module, key=required_key):
                with tempfile.TemporaryDirectory() as tmp:
                    data_dir = Path(tmp) / "originaldata"
                    out_dir = Path(tmp) / "charts"
                    data_dir.mkdir(parents=True, exist_ok=True)
                    out_dir.mkdir(parents=True, exist_ok=True)

                    # Simulate different date suffixes by copying the real file twice.
                    f1 = data_dir / f"{required_key}_2025_01_01.xlsx"
                    f2 = data_dir / f"{required_key}_2025_02_02.xlsx"
                    shutil.copyfile(sample, f1)
                    shutil.copyfile(sample, f2)

                    cmd = [
                        str(VENV_PYTHON),
                        "-m",
                        f"scripts.{module}",
                        "--data-dir",
                        str(data_dir),
                        "--output-dir",
                        str(out_dir),
                    ]
                    proc = subprocess.run(cmd, capture_output=True, text=True)
                    if proc.returncode != 0:
                        self.fail(
                            f"{module} failed.\nSTDOUT:\n{proc.stdout}\nSTDERR:\n{proc.stderr}\n"
                            f"cmd={' '.join(cmd)}"
                        )

                    expected_1 = out_dir / f"{required_key}_1.png"
                    expected_2 = out_dir / f"{required_key}_2.png"
                    self.assertTrue(expected_1.exists(), f"missing: {expected_1}")
                    self.assertTrue(expected_2.exists(), f"missing: {expected_2}")

                    pngs = list(out_dir.glob("*.png"))
                    self.assertEqual(len(pngs), 2, f"unexpected pngs: {[p.name for p in pngs]}")


if __name__ == "__main__":
    raise SystemExit(unittest.main())

