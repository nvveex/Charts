# Charts

基于 `originaldata/` 中的 Excel 数据，批量生成学校业务图表，并输出到 `charts/`。

项目当前包含：

- 29 份原始 Excel 数据
- 36 个图表脚本（已覆盖当前全部数据 key）
- 29 张已生成 PNG 图表
- 1 份验证报告 `charts/verify_report.json`

## 目录说明

- `originaldata/`
  - 原始 Excel 数据源
  - 文件匹配规则以“第一个下划线 `_` 之前的内容”为 key
- `scripts/`
  - 单图脚本目录
  - 每个脚本负责一种图表
  - 公共逻辑统一复用 `scripts/chart_common.py`
- `charts/`
  - 脚本运行后生成的 PNG 图表和验证报告
- `tests/`
  - 集成测试与 key 规则测试

## 脚本命名约定

脚本统一使用业务含义明确的英文名称，而不是随机哈希名，例如：

- `weekly_active_users_trend.py`
- `school_calendar_events_monthly_trend.py`
- `teaching_evaluation_submissions_monthly_trend.py`
- `formal_terms_timeline.py`
- `evaluation_item_sample_topn.py`

每个脚本都应暴露以下统一接口：

- `REQUIRED_KEY`
- `render_one(...)`
- `main(...)`

并支持统一 CLI 参数：

- `--data-dir`
- `--output-dir`

## 运行方式

项目依赖已安装在本地虚拟环境 `.chartvenv/` 中。

### 1. 一键验收并生成全部图表

```bash
./.chartvenv/bin/python verify_all_charts.py
```

运行完成后会生成：

- `charts/*.png`
- `charts/verify_report.json`

### 2. 运行单个脚本

示例：

```bash
./.chartvenv/bin/python -m scripts.teaching_evaluation_submissions_monthly_trend \
  --data-dir originaldata \
  --output-dir charts
```

若要生成“每月评教提交人次”图表，可直接运行：

```bash
./.chartvenv/bin/python -m scripts.teaching_evaluation_submissions_monthly_trend \
  --data-dir originaldata \
  --output-dir charts
```

### 3. 运行测试

```bash
./.chartvenv/bin/python -m unittest discover -s tests -v
```

## 关键规则

### Excel key 提取规则

- 文件名去掉扩展名后
- 若包含 `_`，取第一个 `_` 之前的字符串作为 key
- 若不包含 `_`，则整个文件名作为 key

例如：

- `每周德育分数录入数（近_13_个月）_2026_04_14.xlsx` -> `每周德育分数录入数（近`
- `考试创建数_2026_04_14.xlsx` -> `考试创建数`

### PNG 输出规则

- 一个 key 命中 1 个 Excel 时，输出 `charts/<key>.png`
- 一个 key 命中多个 Excel 时，输出 `charts/<key>_1.png`、`charts/<key>_2.png` ...

## 当前已覆盖的新增业务图表

本轮已补齐并验证以下缺失图表：

- `每周过评分录入数（近`
- `每月评教提交人次（近`
- `考试创建数`
- `最近四个正式学期`
- `近半年评价项名称抽样`
- `每月报表查看人次（近`
- `每月报表发布人次（近一年）`
- `每月教师考核分数录入数（近一年）`

对应脚本分别为：

- `evaluation_score_entries_trend.py`
- `teaching_evaluation_submissions_monthly_trend.py`
- `exam_creation_trend.py`
- `formal_terms_timeline.py`
- `evaluation_item_sample_topn.py`
- `report_views_monthly_trend.py`
- `report_publishers_monthly_trend.py`
- `exam_score_entries_trend.py`

## 验证结果

最近一次全量验收结果：

- `ok=29`
- `skipped=7`
- `failed=0`
- `total=36`

说明当前 `originaldata/` 中已有数据的 key，均可以成功生成对应图表。

## 中文说明补充

- `teaching_evaluation_submissions_monthly_trend.py` 用于生成“每月评教提交人次”图表。
- 脚本会优先读取 `月`、`评教名称`、`评教提交次数` 这几列。
- 若 Excel 第二列使用的是 `name` 而不是 `评教名称`，脚本也会兼容处理。
- 图中会在每根柱子上方直接标出评教名称和对应提交次数。
