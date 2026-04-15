# Project Directory Structure

## 1. Directories

- `originaldata/`
  - 放置未经处理的原始 Excel 文件（`*.xlsx`）
  - 脚本会在此目录下扫描匹配的文件
- `scripts/`
  - 单图脚本（每个脚本只负责生成“一种图表”）
  - 统一依赖 `scripts/chart_common.py` 提供 key 匹配、输出命名、日期清洗等公共逻辑
- `charts/`
  - 脚本运行后生成的全部 PNG

## 2. Key 匹配与输出命名

### Key 规则（匹配 `originaldata/` 内 Excel）

- 对每个 Excel 文件名（不含扩展名）取“第一个 `_` 之前的字符串”作为 `key`
- 忽略 `_` 之后的日期后缀（以及其它可能变化的内容）
- 若文件名不包含 `_`，则 `key = 文件名（不含扩展名）`

### PNG 输出规则

- 每个脚本有一个固定 `REQUIRED_KEY`
- 若 `REQUIRED_KEY` 命中 `N` 个 Excel：
  - `N = 1` 时输出：`charts/<key>.png`
  - `N > 1` 时输出：`charts/<key>_1.png`、`charts/<key>_2.png` ... `charts/<key>_N.png`

## 3. 如何运行

### 3.1 验收（推荐：一键跑完所有图并产出检查报告）

```bash
./.chartvenv/bin/python verify_all_charts.py
```

运行完成后会生成：

- `charts/verify_report.json`

### 3.2 单元测试

由于本项目依赖 `openpyxl` 与 `matplotlib`，请使用项目虚拟环境运行：

```bash
./.chartvenv/bin/python -m unittest -v
```

## 4. 重要说明

- 14 个脚本均提供统一 CLI 参数：
  - `--data-dir`：Excel 所在目录（应为 `originaldata/`）
  - `--output-dir`：输出目录（应为 `charts/`）

