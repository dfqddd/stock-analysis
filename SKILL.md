---
name: stock-research
description: 个股深度调研技能。融合券商深度报告/业绩快评/年报精读/研报摘要/可比公司分析/行业研究六大体例，输入股票代码或名称（可选附年报/业绩公告/研报PDF），自动搜索公开数据生成结构化DOCX调研报告。覆盖核心叙事（公司在炒什么概念/什么逻辑）、商业模式与合作伙伴、行业全景与竞争格局、业绩与单季趋势、3年盈利预测与催化路径、多维估值矩阵、卖方观点分歧、催化剂时间线、情绪面分析。适用于"调研一下XX股票""分析一下XX公司""XX的投资逻辑是什么""XX估值合不合理""XX未来几年怎么看""帮我读这份年报/业绩公告/研报"等场景。
version: 2.1.0
---

# 个股深度调研技能（券商体例整合版）

## 七大核心目标（必须全部覆盖）
1. **公司在炒什么概念、什么逻辑** → 核心叙事章节
2. **行业地位如何** → 行业全景 + 竞争格局
3. **业绩如何** → 业绩快评（单季趋势+超预期判断+分项归因）
4. **未来几年预期** → 3年业绩展望 + 催化路径
5. **估值是否合理** → 多维估值矩阵 + 历史分位 + 三档情景隐含股价
6. **有没有跟什么公司合作** → 商业模式与产业链（上下游+大客户+合作披露）
7. **前景/卖方共识** → 卖方观点矩阵 + 催化剂时间线

## 执行流程

| Step | 动作 | 说明 |
|------|------|------|
| 1 | 确认标的与材料 | 确认股票代码/全称/简称，询问是否提供年报/业绩公告/研报PDF |
| 2 | 数据采集 | 有PDF → 材料精读分支；无PDF → 纯搜索分支（并行8路query） |
| 3 | 行业全景搜索 | 市场规模/产业链/竞争格局/政策环境 |
| 4 | 估值对标搜索 | 2-4家可比公司多维估值 + 历史分位 + 券商目标价 |
| 5 | 情绪与催化搜索 | 近期异动/资金流/板块联动/催化时间线 |
| 6 | 数据核实 | 交叉验证，统一口径 |
| 7 | 生成DOCX报告 | 13章结构，使用docx skill |
| 8 | 输出与展示 | 保存outputs/ + present_files + 3-5句核心结论 |

## 场景路由

| 精确场景 | 加载的参考文件 |
|---------|--------------|
| **标准调研**（无PDF，纯搜索） | [search-queries](references/search-queries.md), [report-structure](references/report-structure.md), [data-sources](references/data-sources.md) |
| **材料精读**（用户提供年报/业绩公告/研报PDF） | [material-analysis](references/material-analysis.md), [search-queries](references/search-queries.md), [report-structure](references/report-structure.md) |
| **估值专项**（用户只问估值/贵不贵/合理不合理） | [valuation-method](references/valuation-method.md), [search-queries](references/valuation-method.md) Step 4 |
| **批量调研**（多只股票并行） | [search-queries](references/search-queries.md), [report-structure](references/report-structure.md)，每只股票并行Task |
| **报告生成**（数据已齐，进入DOCX阶段） | [report-structure](references/report-structure.md), [quality-checklist](references/quality-checklist.md) |
| **质量检查**（报告生成后的验证环节） | [quality-checklist](references/quality-checklist.md) |

### 条件追加

| 条件 | 追加参考 |
|-----|---------|
| 半导体/AI等高成长板块 | [valuation-method](references/valuation-method.md)（板块PE环境章节） |
| 用户提供业绩公告/快报 | [material-analysis](references/material-analysis.md) A2业绩快评 |
| 用户提供多份研报PDF | [material-analysis](references/material-analysis.md) A3研报对比 |
| 需要生成DOCX | 加载docx skill |

## 硬红线（始终生效）

1. **不做买卖建议**：客观呈现数据，看多看空并列，让读者自己判断
2. **标注数据来源和口径**：PE必须标注TTM/静态/动态，同行对比统一口径
3. **不混用口径**：中财网中位PE与亿牛网TTM不可混用，混合口径即失真
4. **不用预测利润算PE**：只用已发布的实际利润
5. **情绪面与基本面分开**：情绪面分析独立成章，不与基本面混淆
6. **核心叙事用观察者视角**：禁用"我们认为"
7. **免责声明**：报告末尾必须加"以上内容仅作信息整理与分析，不构成任何投资建议"

## 输出格式

- 文件：`{公司简称}_深度调研_{YYYYMMDD}.docx`，保存到 outputs/
- 口头总结（3-5句）：核心叙事 → 估值判断 → 最大看点+最大风险 → 跟踪建议
