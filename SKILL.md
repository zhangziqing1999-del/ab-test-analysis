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
  --file "数据文件.xlsx" \
  --groups "对照sid,实验1sid,实验2sid,..." \
  --labels "对照label,实验1label,实验2label,..."
```

第二阶段会：
1. 对timediff非0的query调用DSV4Flash进行LLM打标（query类型：客观/主观，汽车视角分类）
2. 按"query类型"和"汽车分类"分别拆组做AB分析（交互率、完读率、深度）
3. 输出 `AB分析结果_标签维度_<文件名>.xlsx`，包含：
   - `按类型_在线` / `按类型_离线`: 按客观/主观拆组分析
   - `按汽车分类_在线` / `按汽车分类_离线`: 按18个汽车子分类拆组分析
   - `标注明细_在线` / `标注明细_离线`: 打标结果明细

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

**在线入口(series/pin_dao/dongtai):**
- 二轮交互率 = rank2 query数 / rank1 query数
- 一轮完读率 = rank1中recommend_prompt非空数 / rank1总数
- 二轮完读率 = rank2中recommend_prompt非空数 / rank2总数

**离线入口(其他所有request_type):**
- 二轮交互率、二→三轮交互率
- 二轮完读率、三轮完读率

**交互深度:**
- 在线：人均交互轮次 = 总query / 唯一session数
- 离线：有二轮用户的人均交互轮次
