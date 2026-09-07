# kb-report |  | 知识库问答评测报告生成工具

把**已评测**的知识库问答工作簿（kb-qa-eval 输出或同构文件）转换成参考格式的最终评测报告：
统计摘要 + Badcase 归因分析，支持单评测员报告与多评测员合并报告。

> 配套工具：评测打分用 [kb-qa-eval](internal-gitlab/eval_project)；
> 数据筛选/抽样用 [kb-query-pipeline](internal-gitlab/kb_query_pipeline)。

## 功能

- **统计摘要** sheet：评测总数 / 核心指标（MOS、Badcase 率与数、准出 ✅❌）/ 5 维度评分明细 / 分数段分布 / Badcase 明细
- **Badcase 归因分析** sheet（12 列）：序号 / Query / 语料源文件 / 金标段落 / 召回段落 / 金标↔召回匹配 / 回答 / WA / 归因结论 / 归因分析 / 原文核验 / Reason 摘录
  - 金标↔召回匹配 = **客观覆盖率计算**（字符 bigram），不靠 LLM 主观判断
  - 归因结论 / 分析 / 原文核验由 LLM 生成，无模型时用模板兜底
- **多评测员合并报告**（N≥2 个输入）：并排分数列 + 按 query 合并均分重算统计 + 合并 WA badcase 归因

## 何时用

- 已有 `*_evaluated_*.xlsx`，需要统计摘要 + badcase 归因报告
- 多个评测员（模型或人工）各评一份后，需要合并报告
- 报告格式需对齐参考样表

**不要用于**：跑评测打分（用 kb-qa-eval）、筛选/抽样数据（用 kb-query-pipeline）。

## 输入 → 输出

**输入**：已评测 xlsx（评测结果 sheet 含 5 维度分数/理由列 + 源数据列）；合并模式需 N≥2 个同源同序文件。

**输出**：
- 单输入：`<input>_report.xlsx` = 原工作簿 + `统计摘要` + `Badcase归因分析`
- 合并模式（默认 `<第一个输入>_merged_report.xlsx`）：
  - `Evaluation`：两行表头，公共列 + 每个评测员的 `分数_k / 理由_k / 加权均分_k / 综合评价_k`（N 可扩展）
  - `低分复核`：各评测员复核行纵向拼接
  - `统计摘要`：按合并均分（每 query 各评测员分数取均值）重算
  - `Badcase归因分析`：仅合并均分 < 阈值的条目，WA 列列出全部评测员值

## 用法

```bash
# 离线（覆盖率 + 模板归因，不调 LLM）
python scripts/report.py --input <evaluated.xlsx> --language zh --no-llm

# 完整（LLM 归因；--doc-dirs 提供语料目录时启用原文核验）
python scripts/report.py --input <evaluated.xlsx> --language ar \
  --model azure/gpt-5.4 --doc-dirs D1,D2

# 多评测员合并报告（--input 可重复，N≥2）
python scripts/report.py --input <evaluated_m1.xlsx> --input <evaluated_m2.xlsx> \
  --input <evaluated_m3.xlsx> --language fr --model azure/gpt-5.4 --doc-dirs D1
```

- `--language` 必传：`zh / en / fr / ar / id / ru`（决定报告标题前缀与统计摘要语言）
- 合并模式强校验：语言、表头、query 序、回答内容、原始行布局必须一致，不一致拒绝合并
- 合并宽表 WA 列显示按维度分重算的加权均分（与统计摘要同口径），不信任输入文件自带的 WA 列
- 完整参数：`python scripts/report.py --help`

## 运行时配置

API 配置同 kb-qa-eval，通过环境变量：

| 变量 | 说明 |
|------|------|
| `KB_QA_API_KEY` | LLM API Key（用 `--no-llm` 时不需要） |
| `KB_QA_API_BASE_URL` | API 基址（可选） |

## 项目结构

```
scripts/
  report.py          # 主入口：单/合并报告生成
  report_writer.py   # Excel 写入与样式（参考格式）
  attribution.py     # LLM 归因调用 + 模板兜底 + 429 重试
  doc_reader.py      # 原文读取：多文件金标覆盖率排序注入、stem 匹配、xlsx 支持
prompts/
  attribution_prompt.py   # 归因 prompt（运行时加载）
tests/
  test_report_merge.py      # 合并模式用例
  test_report_multifile.py  # 多文件原文核验用例
  test_doc_reader.py        # 原文读取用例
references/
  attribution_criteria.md   # 归因口径 knowhow
```

## 依赖

共享 kb-qa-eval 的 venv（`scripts/eval_stats` / `scripts/api_client` / `scripts/excel_style`）：
`D:\Users\admin\.agents\skills\kb-qa-eval\venv`

## 测试

```powershell
& D:\Users\admin\.agents\skills\kb-qa-eval\venv\Scripts\python.exe -m unittest discover -s tests
```

## 规则

- 不修改输入文件内容；新报告写入 `--output`（默认 `<input>_report.xlsx`）
- 覆盖率算法与归因口径见 `references/attribution_criteria.md`
