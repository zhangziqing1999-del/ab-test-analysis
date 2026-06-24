#!/usr/bin/env python3
"""
AB实验打标+分维度分析脚本（第二阶段）
用法: python3 ab_label_analysis.py --file <csv/excel> --groups <sid1,sid2,...> --labels <l1,l2,...>
前置：需先运行 ab_analysis.py 获得第一阶段结果。
输出: 同目录下生成 AB分析结果_标签维度_<filename>.xlsx
"""
import argparse, os, time, json
import pandas as pd
import numpy as np
from scipy import stats
import requests

API_URL = "https://oneapi-comate.baidu-int.com/v1/chat/completions"
API_KEY = "sk-qeNvUukW0VTSMp600eC1De92D2154c05AcAe8aD48a45CaC4"
MODEL = "deepseek-v4-flash"
BATCH_SIZE = 30

CATEGORIES = "整体评价类、性价比类、值得买类、用养成本类、口碑类、价格类、场景适配类、保值类、空间类、完整配置类、续航/充电类、驾驶与操控类、智能配置类、舒适配置类、安全配置类、外观类、车系推荐类、其他"

SYSTEM_PROMPT = f"""你是汽车领域query分类器。对每条query严格输出一行JSON，不要任何解释：
{{"type":"客观"或"主观", "category":"子分类"}}

类型规则：
- 客观：有确定答案的事实性问题（参数、价格、配置等）
- 主观：需要个人判断、评价、推荐的问题

可选子分类（只能从以下选一个）：
{CATEGORIES}
如果query与汽车无关或无法归类，category填"其他"。"""


def proportion_z_test(c1, n1, c2, n2):
    if n1 == 0 or n2 == 0: return np.nan
    p_pool = (c1 + c2) / (n1 + n2)
    if p_pool == 0 or p_pool == 1: return np.nan
    se = np.sqrt(p_pool * (1 - p_pool) * (1/n1 + 1/n2))
    z = (c1/n1 - c2/n2) / se
    return 2 * (1 - stats.norm.cdf(abs(z)))


def sig(p):
    if pd.isna(p): return '-'
    if p < 0.001: return '***'
    if p < 0.01: return '**'
    if p < 0.05: return '*'
    return 'ns'


def is_read(s):
    return s.notna() & (s != '') & (s != '""')


def classify_batch(queries):
    """调用LLM对一批query打标"""
    content = "\n".join([f"{i+1}. {q}" for i, q in enumerate(queries)])
    user_msg = f"对以下{len(queries)}条query逐条分类，每条输出一行JSON，共{len(queries)}行：\n{content}"
    headers = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}

    for _ in range(3):
        try:
            resp = requests.post(API_URL, headers=headers, json={
                "model": MODEL,
                "messages": [{"role": "system", "content": SYSTEM_PROMPT},
                             {"role": "user", "content": user_msg}],
                "max_tokens": len(queries) * 80, "temperature": 0
            }, timeout=120)
            if resp.status_code == 429:
                time.sleep(3); continue
            if resp.status_code != 200:
                time.sleep(1); continue
            text = resp.json()['choices'][0]['message'].get('content', '')
            results = []
            for line in text.strip().split('\n'):
                line = line.strip()
                if not line: continue
                if line[0].isdigit():
                    idx = line.find('{')
                    if idx >= 0: line = line[idx:]
                try:
                    results.append(json.loads(line))
                except:
                    results.append(None)
            while len(results) < len(queries): results.append(None)
            return results[:len(queries)]
        except:
            time.sleep(2)
    return [None] * len(queries)


def label_dataframe(df):
    """对df中timediff!=0的query进行打标，返回带标签的df"""
    mask = df['time_diff'] != 0
    queries = df.loc[mask, 'query'].fillna('').tolist()
    total = len(queries)
    if total == 0:
        df['query类型'] = None
        df['汽车分类'] = None
        return df

    print(f"  打标中: {total} queries, batch={BATCH_SIZE}")
    all_types = [None] * total
    all_cats = [None] * total

    for i in range(0, total, BATCH_SIZE):
        batch = queries[i:i+BATCH_SIZE]
        results = classify_batch(batch)
        for j, r in enumerate(results):
            if r and isinstance(r, dict):
                all_types[i+j] = r.get('type', '')
                all_cats[i+j] = r.get('category', '')
        if (i // BATCH_SIZE + 1) % 20 == 0:
            print(f"    进度: {min(i+BATCH_SIZE, total)}/{total}")
        time.sleep(0.3)

    df.loc[mask, 'query类型'] = all_types
    df.loc[mask, '汽车分类'] = all_cats
    print(f"  打标完成，成功率: {sum(1 for t in all_types if t)*100//total}%")
    return df


def ab_by_dimension(df, control, exp_groups, labels, dim_col, dim_name):
    """按某个维度分组做AB分析，返回DataFrame"""
    rows = []
    for dim_val in sorted(df[dim_col].dropna().unique()):
        sub = df[df[dim_col] == dim_val]
        ctrl = sub[sub['实验组'] == control]
        nc1 = len(ctrl[ctrl['rank'] == 1])
        nc2 = len(ctrl[ctrl['rank'] == 2])
        nc1r = int(is_read(ctrl[ctrl['rank'] == 1]['recommend_prompt']).sum())
        nc2r = int(is_read(ctrl[ctrl['rank'] == 2]['recommend_prompt']).sum())

        for g in exp_groups:
            gdf = sub[sub['实验组'] == g]
            n1 = len(gdf[gdf['rank'] == 1])
            n2 = len(gdf[gdf['rank'] == 2])
            n1r = int(is_read(gdf[gdf['rank'] == 1]['recommend_prompt']).sum())
            n2r = int(is_read(gdf[gdf['rank'] == 2]['recommend_prompt']).sum())

            p_interact = proportion_z_test(n2, n1, nc2, nc1)
            p_read1 = proportion_z_test(n1r, n1, nc1r, nc1)

            # 交互深度
            ctrl_d = ctrl.groupby('session_id').size()
            exp_d = gdf.groupby('session_id').size()
            if len(exp_d) > 0 and len(ctrl_d) > 0:
                _, p_depth = stats.mannwhitneyu(exp_d, ctrl_d, alternative='two-sided')
            else:
                p_depth = np.nan

            rows.append({
                dim_name: dim_val, '实验组': labels[g],
                'query数_实验': n1, 'query数_对照': nc1,
                '二轮交互率': f"{n2/n1:.4f}" if n1 else '-',
                '对照_二轮交互率': f"{nc2/nc1:.4f}" if nc1 else '-',
                '交互率_p值': round(p_interact, 4) if not pd.isna(p_interact) else '-',
                '交互率_显著性': sig(p_interact),
                '一轮完读率': f"{n1r/n1:.4f}" if n1 else '-',
                '对照_一轮完读率': f"{nc1r/nc1:.4f}" if nc1 else '-',
                '完读率_p值': round(p_read1, 4) if not pd.isna(p_read1) else '-',
                '完读率_显著性': sig(p_read1),
                '平均深度_实验': round(exp_d.mean(), 3) if len(exp_d) > 0 else '-',
                '平均深度_对照': round(ctrl_d.mean(), 3) if len(ctrl_d) > 0 else '-',
                '深度_p值': round(p_depth, 4) if not pd.isna(p_depth) else '-',
                '深度_显著性': sig(p_depth),
            })
    return pd.DataFrame(rows)


def main():
    parser = argparse.ArgumentParser(description='AB实验打标+分维度分析（第二阶段）')
    parser.add_argument('--file', required=True)
    parser.add_argument('--groups', required=True)
    parser.add_argument('--labels', required=True)
    parser.add_argument('--online-sheet', default='在线')
    parser.add_argument('--offline-sheet', default='离线')
    args = parser.parse_args()

    groups = args.groups.split(',')
    labels_list = args.labels.split(',')
    if len(groups) != len(labels_list):
        print("ERROR: groups数量必须和labels数量一致"); return

    control = groups[0]
    exp_groups = groups[1:]
    labels = dict(zip(groups, labels_list))

    # 读取数据
    file_path = args.file
    if file_path.endswith('.csv'):
        df_all = pd.read_csv(file_path)
        df_online = df_all[df_all['request_type'].isin(['series', 'pin_dao', 'dongtai'])]
        df_offline = df_all[~df_all['request_type'].isin(['series', 'pin_dao', 'dongtai'])]
    else:
        df_online = pd.read_excel(file_path, sheet_name=args.online_sheet, engine='openpyxl')
        df_offline = pd.read_excel(file_path, sheet_name=args.offline_sheet, engine='openpyxl')

    df_online = df_online[df_online['实验组'].isin(groups)].copy()
    df_offline = df_offline[df_offline['实验组'].isin(groups)].copy()
    print(f"读取: {file_path}")
    print(f"在线: {len(df_online)}, 离线: {len(df_offline)}")

    # 打标
    print("\n[阶段1] 在线数据打标")
    df_online = label_dataframe(df_online)
    print("\n[阶段2] 离线数据打标")
    df_offline = label_dataframe(df_offline)

    # 分维度分析（仅timediff!=0的数据）
    td_online = df_online[df_online['time_diff'] != 0]
    td_offline = df_offline[df_offline['time_diff'] != 0]

    sheets = {}
    print("\n[阶段3] 分维度AB分析")

    if len(td_online) > 0:
        sheets['按类型_在线'] = ab_by_dimension(td_online, control, exp_groups, labels, 'query类型', 'query类型')
        sheets['按汽车分类_在线'] = ab_by_dimension(td_online, control, exp_groups, labels, '汽车分类', '汽车分类')

    if len(td_offline) > 0:
        sheets['按类型_离线'] = ab_by_dimension(td_offline, control, exp_groups, labels, 'query类型', 'query类型')
        sheets['按汽车分类_离线'] = ab_by_dimension(td_offline, control, exp_groups, labels, '汽车分类', '汽车分类')

    # 标注明细
    sheets['标注明细_在线'] = td_online[['format_date', 'session_id', 'query', 'rank', 'time_diff', '实验组', 'query类型', '汽车分类']].head(5000)
    sheets['标注明细_离线'] = td_offline[['format_date', 'session_id', 'query', 'rank', 'time_diff', '实验组', 'query类型', '汽车分类']].head(5000)

    # 输出
    base = os.path.splitext(os.path.basename(file_path))[0]
    out_dir = os.path.dirname(os.path.abspath(file_path))
    out_path = os.path.join(out_dir, f"AB分析结果_标签维度_{base}.xlsx")

    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        for name, d in sheets.items():
            d.to_excel(writer, sheet_name=name, index=False)

    print(f"\n结果已保存: {out_path}")
    print(f"sheets: {list(sheets.keys())}")

    # 打印关键结果
    for key in ['按类型_在线', '按类型_离线', '按汽车分类_在线', '按汽车分类_离线']:
        if key in sheets:
            print(f"\n{'='*50}\n{key}\n{'='*50}")
            print(sheets[key].to_string(index=False))


if __name__ == '__main__':
    main()
