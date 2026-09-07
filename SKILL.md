---
name: kb-report
description: "Use this skill to generate the final evaluation report (评测报告) from an already-evaluated knowledge base QA workbook: 统计摘要 (MOS/badcase rate/维度明细/分数段/badcase 清单) and Badcase归因分析 (覆盖率客观计算 + LLM 归因结论/分析/原文核验), in the reference report format. Trigger when the user asks to 生成报告/统计摘要/badcase归因/评测报告 from a *_evaluated_*.xlsx file, or wants the 参考表格式 report for en/zh/fr/ar/id/ru evaluations."
---
# KB QA Report (kb-report)

## Overview

把**已评测**的知识库问答工作簿（kb-qa-eval 输出或同构文件）转换成参考格式的最终评测报告：`统计摘要` + `Badcase归因分析`，支持单评测员报告与多评测员合并报告（N≥2）。

## When to Use

- 已有 `*_evaluated_*.xlsx`，需要统计摘要 + badcase 归因报告
- 多个评测员（模型或人工）各评一份后，需要合并报告
- 报告格式需对齐参考样表

**Do NOT use when:** 需要跑评测打分（用 kb-qa-eval）、筛选/抽样数据（用 kb-query-pipeline）。

## Entry

```bash
# 离线（覆盖率 + 模板归因）
python scripts/report.py --input <evaluated.xlsx> --language zh --no-llm

# 完整（LLM 归因；--doc-dirs 提供语料目录时启用原文核验）
python scripts/report.py --input <evaluated.xlsx> --language ar --model azure/gpt-5.4 --doc-dirs D1,D2

# 多评测员合并（--input 可重复，N≥2）
python scripts/report.py --input <m1.xlsx> --input <m2.xlsx> --language fr --model azure/gpt-5.4
```

`--language` 必传（zh/en/fr/ar/id/ru）。合并模式强校验：语言/表头/query 序/回答/行布局必须一致。完整参数、IO 契约、项目结构、运行时配置见 [README.md](README.md)。

## Rules

- 不修改输入文件内容；新报告写入 `--output`（默认 `<input>_report.xlsx`）
- 金标↔召回匹配为客观覆盖率计算（字符 bigram）；归因结论/分析/原文核验由 LLM 生成或模板兜底
- 覆盖率算法与归因口径见 [references/attribution_criteria.md](references/attribution_criteria.md)

## References

- 完整文档: [README.md](README.md)
- 归因口径: [references/attribution_criteria.md](references/attribution_criteria.md)
- 参考格式（单评测员）: `tests/fixtures/reference_layout.xlsx`
- 依赖 kb-qa-eval 的 venv（见 README#依赖）