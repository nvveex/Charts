# Charts

基于 `originaldata/` 中的 Excel 数据，批量生成学校业务图表，并输出到 `charts/`。


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

## 运行方式

项目依赖已安装在本地虚拟环境 `.chartvenv/` 中。

### 1. 一键验收并生成全部图表

```bash
./.chartvenv/bin/python verify_all_charts.py
```

运行完成后会生成：

- `charts/*.png`
- `charts/verify_report.json`

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

对应脚本分别为：

- `evaluation_score_entries_trend.py`
- `teaching_evaluation_submissions_monthly_trend.py`
- `exam_creation_trend.py`
- `formal_terms_timeline.py`
- `evaluation_item_sample_topn.py`

## 验证结果

最近一次全量验收结果：

- `ok=26`
- `skipped=10`
- `failed=0`
- `total=36`

说明当前 `originaldata/` 中已有数据的 key，均可以成功生成对应图表。
