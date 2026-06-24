---
name: ab-test-analysis
description: 策略需求迭代上线后的AB实验数据分析。当用户说"做AB分析"、"实验数据分析"、"分析实验效果"、"跑一下AB数据"、"看下实验显著性"、"对比实验组"时触发。支持2-5组实验对比，自动计算交互率、完读率、交互深度及p值显著性检验。即使用户只是提到"分析数据"并且上下文涉及实验组/对照组/absid，也应触发此skill。
---

# AB实验数据分析

对多轮交互类产品（对话式AI、推荐系统等）的线上AB实验进行标准化分析。

## 触发时机

用户要分析策略迭代上线后的实验效果，涉及实验组vs对照组对比。

## 使用流程

### Step 1: 收集信息

执行前必须和用户确认以下信息：

1. **数据文件路径** — Excel文件，包含"在线"和"离线"两个sheet（或用户指定的sheet名）
2. **实验组absid** — 列出所有组的sid，明确哪个是对照组（第一个为对照）
3. **实验组label** — 每组的中文标签（如"线上对照"、"v4"、"v4+新prompt"）
4. **在线入口类型**（可选）— request_type中属于在线的值，默认 series/pin_dao/dongtai

确认模板：
```
请确认实验配置：
- 数据文件：xxx.xlsx
- 对照组：{sid} ({label})
- 实验组1：{sid} ({label})
- 实验组2：{sid} ({label})
- ...（支持2-5组）
- 在线入口：series, pin_dao, dongtai
```

### Step 2: 运行分析脚本（第一阶段：纯数据指标分析）

```bash
python3 <skill-path>/scripts/ab_analysis.py \
  --file "数据文件.xlsx" \
  --groups "对照sid,实验1sid,实验2sid,..." \
  --labels "对照label,实验1label,实验2label,..."
```

脚本会输出：
- 终端打印关键分析结果
- 同目录生成 `AB分析结果_<文件名>.xlsx`，包含多个sheet：
  - `中间统计_在线` / `中间统计_离线`: 各组各rank的query数、session数、完读数、完读率
  - `AB分析_在线` / `AB分析_离线`: 核心指标对比 + p值显著性检验
  - `交互深度_在线` / `交互深度_离线`: 平均深度、中位数 + Mann-Whitney U检验

运行完后将第一阶段Excel路径告知用户，再进入第二阶段。

### Step 3: 运行打标+分维度分析（第二阶段）

```bash
python3 <skill-path>/scripts/ab_label_analysis.py \
  --file "数据文件.csv" \
  --groups "对照sid,实验1sid,实验2sid,..." \
  --labels "对照label,实验1label,实验2label,..." \
  --exp-id 225457
```

- `--file` 支持原始CSV（含 `ab实验sid` 列）或已生成的 `入口分组_*.xlsx` 中间文件
- `--exp-id` 为CSV时必填，用于从复合 `ab实验sid`（如 `225457_2-230583_3`）中提取实验组

第二阶段为5个阶段，每阶段产出/更新Excel，中断可断点续传：

1. **阶段1 预处理**：提取实验组、按入口分组（保留全量数据，含time_diff=0），标记 `待打标=time_diff!=0`，输出中间文件 `入口分组_<文件名>.xlsx`（sheet `在线`/`离线`）
2. **阶段2 基础AB分析**：基于全量数据做无标签AB分析（在线：一轮→二轮转化率/一轮完读率/二轮完读率；离线：二轮→三轮转化率/二轮完读率/三轮完读率）
3. **阶段3 LLM打标**：对 `待打标=True` 的query用 deepseek-v4-flash + glm-5.1 并发打标（query类型：客观/主观，18类汽车分类）。**先对unique query去重打标再map回**（约省50%），checkpoint断点续传，每500条落盘
4. **阶段4 标签维度AB分析**：标签取query自身（不传播）。在线锚定rank=1，离线锚定rank=2，按 `query类型` 和 `汽车分类` 拆组做AB分析
5. **阶段5 最终输出**：汇总所有sheet到 `AB分析结果_标签维度_<文件名>.xlsx`，清理checkpoint

输出Excel包含sheets：
- `AB分析_在线` / `AB分析_离线`: 无标签整体AB分析
- `按query类型_在线` / `按query类型_离线`: 按客观/主观拆组
- `按汽车分类_在线` / `按汽车分类_离线`: 按18个汽车子分类拆组
- `标注明细_在线` / `标注明细_离线`: 打标结果明细

**效率优化**：query去重（约省52%）+ 并发8 + batch50 + 多模型轮询，约提速5-6倍。

**容错**：checkpoint 仅在全流程结束后清理；打标过程每500条更新Excel快照，中途被kill不丢数据，重跑自动从checkpoint恢复。

### Step 4: 解读结果

根据输出进行分析解读：
- 显著性标记：*** p<0.001, ** p<0.01, * p<0.05, ns 不显著
- 关注显著性结果的业务含义
- 给出结论性建议

## 必需字段

数据表需包含以下列：
- `实验组`: absid标识
- `request_type`: 入口类型
- `rank`: 交互轮次（1=一轮，2=二轮...）
- `recommend_prompt`: 完读标识（非空且非`""`即完读）
- `session_id`: 会话标识
- `format_date`: 日期

## 分析指标说明

**在线入口(series/pin_dao/dongtai)，锚定 rank=1:**
- 一轮→二轮转化率 = rank2 query数 / rank1 query数
- 一轮完读率 = rank1中recommend_prompt非空数 / rank1总数
- 二轮完读率 = rank2中recommend_prompt非空数 / rank2总数

**离线入口(其他所有request_type)，锚定 rank=2:**
- 二轮→三轮转化率 = rank3 query数 / rank2 query数
- 二轮完读率 = rank2中完读数 / rank2总数
- 三轮完读率 = rank3中完读数 / rank3总数

> 离线从rank=2起算：离线入口首轮(rank=1)绝大多数 time_diff=0（曝光未真正交互），无有效打标，故从二轮看才有意义。

**关键口径**：打标只针对 time_diff!=0 的query，但AB指标的分母必须用**全量数据**（含time_diff=0）；否则离线rank=1样本骤减，转化率会算出>1的异常值。
