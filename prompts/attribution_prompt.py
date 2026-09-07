"""Attribution prompt template (single shared Chinese template).

The 归因口径 knowhow is loaded from references/attribution_criteria.md at
runtime so criteria edits never require code changes.
"""

ATTRIBUTION_PROMPT = """你是知识库问答评测的 badcase 归因分析专家。请对以下低分条目进行归因分析。

## 归因口径（判定方法）

{criteria}

## 输入信息

- 用户问题: {query}
- 语料源文件: {filename}
- 金标答案段落（标注）:
{annotation}

- 实际召回段落（搜索结果）:
{retrieval}

- 金标↔召回匹配（客观覆盖率计算）: {coverage_text}
- 大模型回答:
{answer}

- 加权均分: {wa}
- 各维度评分与理由:
{dim_reasons}
- 原文核验材料（原文回溯用；材料已按金标覆盖度降序排列，文件头附客观覆盖率，超预算文件按金标相关段落截取）:
{original_text}

## 任务

1. 依据归因口径判定【归因结论】（A/B/C 分类，允许组合如 "A主+B次"）
2. 撰写【归因分析】（中文，150-250 字，主因→次因，引用金标/召回/回答三方具体证据）
3. 若提供了原文核验材料，撰写【原文核验】（比对金标句与召回段在原文中的实际位置，给出 ✅ 已核实 / ⚠️ 部分核实 / ⚠️ 标注问题 三态结论；未提供材料则该字段为空字符串）

## 输出格式（严格 JSON）

{{
  "归因结论": "...",
  "归因分析": "...",
  "原文核验": "..."
}}
"""
