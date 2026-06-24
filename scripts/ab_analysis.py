#!/usr/bin/env python3
"""
AB实验数据分析脚本
用法: python3 ab_analysis.py --file <excel/csv> --groups <sid1,sid2,...> --labels <l1,l2,...> [--online-sheet ...] [--offline-sheet ...]
第一个group为对照组，支持2-5组。
输出: 同目录下生成 AB分析结果_<filename>.xlsx，包含中间统计表和最终分析表。
"""
import argparse, os
import pandas as pd
import numpy as np
from scipy import stats


def proportion_z_test(c1, n1, c2, n2):
    if n1 == 0 or n2 == 0:
        return np.nan
    p_pool = (c1 + c2) / (n1 + n2)
    if p_pool == 0 or p_pool == 1:
        return np.nan
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


def build_intermediate_stats(df, groups, labels, entry_type):
    """构建中间统计表：各组各rank的query数、session数、完读数"""
    rows = []
    for g in groups:
        gdf = df[df['实验组'] == g]
        for rank in sorted(gdf['rank'].unique()):
            rdf = gdf[gdf['rank'] == rank]
            rows.append({
                '入口': entry_type, '实验组': labels[g], 'rank': int(rank),
                'query数': len(rdf),
                'session数': rdf['session_id'].nunique(),
                '完读数': int(is_read(rdf['recommend_prompt']).sum()),
                '完读率': is_read(rdf['recommend_prompt']).mean() if len(rdf) > 0 else 0
            })
    return pd.DataFrame(rows)


def build_online_result(df, control, exp_groups, labels):
    """在线AB分析结果表"""
    ctrl = df[df['实验组'] == control]
    nc1, nc2 = len(ctrl[ctrl['rank'] == 1]), len(ctrl[ctrl['rank'] == 2])
    nc1r = int(is_read(ctrl[ctrl['rank'] == 1]['recommend_prompt']).sum())
    nc2r = int(is_read(ctrl[ctrl['rank'] == 2]['recommend_prompt']).sum())

    rows = []
    for g in exp_groups:
        gdf = df[df['实验组'] == g]
        n1, n2 = len(gdf[gdf['rank'] == 1]), len(gdf[gdf['rank'] == 2])
        n1r = int(is_read(gdf[gdf['rank'] == 1]['recommend_prompt']).sum())
        n2r = int(is_read(gdf[gdf['rank'] == 2]['recommend_prompt']).sum())

        p_interact = proportion_z_test(n2, n1, nc2, nc1)
        p_read1 = proportion_z_test(n1r, n1, nc1r, nc1)
        p_read2 = proportion_z_test(n2r, n2, nc2r, nc2)

        rows.append({
            '实验组': labels[g], '对照组': labels[control],
            '二轮交互率': f"{n2/n1:.4f}" if n1 else '-',
            '对照_二轮交互率': f"{nc2/nc1:.4f}" if nc1 else '-',
            '二轮交互率_p值': round(p_interact, 4) if not pd.isna(p_interact) else '-',
            '二轮交互率_显著性': sig(p_interact),
            '一轮完读率': f"{n1r/n1:.4f}" if n1 else '-',
            '对照_一轮完读率': f"{nc1r/nc1:.4f}" if nc1 else '-',
            '一轮完读率_p值': round(p_read1, 4) if not pd.isna(p_read1) else '-',
            '一轮完读率_显著性': sig(p_read1),
            '二轮完读率': f"{n2r/n2:.4f}" if n2 else '-',
            '对照_二轮完读率': f"{nc2r/nc2:.4f}" if nc2 else '-',
            '二轮完读率_p值': round(p_read2, 4) if not pd.isna(p_read2) else '-',
            '二轮完读率_显著性': sig(p_read2),
        })
    return pd.DataFrame(rows)


def build_offline_result(df, control, exp_groups, labels):
    """离线AB分析结果表"""
    ctrl = df[df['实验组'] == control]
    nc1, nc2, nc3 = len(ctrl[ctrl['rank'] == 1]), len(ctrl[ctrl['rank'] == 2]), len(ctrl[ctrl['rank'] == 3])
    nc2r = int(is_read(ctrl[ctrl['rank'] == 2]['recommend_prompt']).sum())
    nc3r = int(is_read(ctrl[ctrl['rank'] == 3]['recommend_prompt']).sum())

    rows = []
    for g in exp_groups:
        gdf = df[df['实验组'] == g]
        n1, n2, n3 = len(gdf[gdf['rank'] == 1]), len(gdf[gdf['rank'] == 2]), len(gdf[gdf['rank'] == 3])
        n2r = int(is_read(gdf[gdf['rank'] == 2]['recommend_prompt']).sum())
        n3r = int(is_read(gdf[gdf['rank'] == 3]['recommend_prompt']).sum())

        p_i2 = proportion_z_test(n2, n1, nc2, nc1)
        p_i3 = proportion_z_test(n3, n2, nc3, nc2)
        p_r2 = proportion_z_test(n2r, n2, nc2r, nc2)
        p_r3 = proportion_z_test(n3r, n3, nc3r, nc3)

        rows.append({
            '实验组': labels[g], '对照组': labels[control],
            '二轮交互率': f"{n2/n1:.4f}" if n1 else '-',
            '对照_二轮交互率': f"{nc2/nc1:.4f}" if nc1 else '-',
            '二轮交互率_p值': round(p_i2, 4) if not pd.isna(p_i2) else '-',
            '二轮交互率_显著性': sig(p_i2),
            '三轮交互率': f"{n3/n2:.4f}" if n2 else '-',
            '对照_三轮交互率': f"{nc3/nc2:.4f}" if nc2 else '-',
            '三轮交互率_p值': round(p_i3, 4) if not pd.isna(p_i3) else '-',
            '三轮交互率_显著性': sig(p_i3),
            '二轮完读率': f"{n2r/n2:.4f}" if n2 else '-',
            '对照_二轮完读率': f"{nc2r/nc2:.4f}" if nc2 else '-',
            '二轮完读率_p值': round(p_r2, 4) if not pd.isna(p_r2) else '-',
            '二轮完读率_显著性': sig(p_r2),
            '三轮完读率': f"{n3r/n3:.4f}" if n3 else '-',
            '对照_三轮完读率': f"{nc3r/nc3:.4f}" if nc3 else '-',
            '三轮完读率_p值': round(p_r3, 4) if not pd.isna(p_r3) else '-',
            '三轮完读率_显著性': sig(p_r3),
        })
    return pd.DataFrame(rows)


def build_depth_result(df, control, exp_groups, labels, require_r2=False):
    """交互深度分析结果"""
    ctrl = df[df['实验组'] == control]
    if require_r2:
        cs = ctrl[ctrl['rank'] >= 2]['session_id'].unique()
        ctrl = ctrl[ctrl['session_id'].isin(cs)]
    ctrl_d = ctrl.groupby('session_id').size()

    rows = []
    for g in exp_groups:
        gdf = df[df['实验组'] == g]
        if require_r2:
            s2 = gdf[gdf['rank'] >= 2]['session_id'].unique()
            gdf = gdf[gdf['session_id'].isin(s2)]
        exp_d = gdf.groupby('session_id').size()
        if len(exp_d) > 0 and len(ctrl_d) > 0:
            _, p = stats.mannwhitneyu(exp_d, ctrl_d, alternative='two-sided')
        else:
            p = np.nan
        rows.append({
            '实验组': labels[g], '实验组_平均深度': round(exp_d.mean(), 3),
            '对照组_平均深度': round(ctrl_d.mean(), 3),
            '实验组_中位数': round(exp_d.median(), 1),
            '对照组_中位数': round(ctrl_d.median(), 1),
            '实验组_session数': len(exp_d),
            'p值': round(p, 4) if not pd.isna(p) else '-',
            '显著性': sig(p),
        })
    return pd.DataFrame(rows)


def main():
    parser = argparse.ArgumentParser(description='AB实验数据分析（支持2-5组）')
    parser.add_argument('--file', required=True, help='数据文件路径(Excel或CSV)')
    parser.add_argument('--groups', required=True, help='实验组sid,逗号分隔,第一个为对照组')
    parser.add_argument('--labels', required=True, help='各组标签,逗号分隔')
    parser.add_argument('--online-sheet', default='在线', help='在线数据sheet名')
    parser.add_argument('--offline-sheet', default='离线', help='离线数据sheet名')
    args = parser.parse_args()

    groups = args.groups.split(',')
    labels_list = args.labels.split(',')
    if len(groups) != len(labels_list):
        print("ERROR: groups数量必须和labels数量一致"); return
    if len(groups) < 2 or len(groups) > 5:
        print("ERROR: 支持2-5组（第一个为对照组）"); return

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

    df_online = df_online[df_online['实验组'].isin(groups)]
    df_offline = df_offline[df_offline['实验组'].isin(groups)]
    print(f"读取: {file_path}")
    print(f"在线: {len(df_online)} rows, 离线: {len(df_offline)} rows")
    print(f"对照: {control} ({labels[control]}), 实验组: {[f'{g}({labels[g]})' for g in exp_groups]}")

    # 构建所有表
    sheets = {}

    # 中间统计表
    stats_online = build_intermediate_stats(df_online, groups, labels, '在线')
    stats_offline = build_intermediate_stats(df_offline, groups, labels, '离线')
    sheets['中间统计_在线'] = stats_online
    sheets['中间统计_离线'] = stats_offline

    # 最终分析表
    sheets['AB分析_在线'] = build_online_result(df_online, control, exp_groups, labels)
    sheets['AB分析_离线'] = build_offline_result(df_offline, control, exp_groups, labels)
    sheets['交互深度_在线'] = build_depth_result(df_online, control, exp_groups, labels, False)
    sheets['交互深度_离线'] = build_depth_result(df_offline, control, exp_groups, labels, True)

    # 输出Excel
    base = os.path.splitext(os.path.basename(file_path))[0]
    out_dir = os.path.dirname(os.path.abspath(file_path))
    out_path = os.path.join(out_dir, f"AB分析结果_{base}.xlsx")

    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"\n结果已保存: {out_path}")
    print(f"包含sheets: {list(sheets.keys())}")

    # 同时打印关键结果到stdout
    print("\n" + "=" * 60)
    print("在线AB分析")
    print("=" * 60)
    print(sheets['AB分析_在线'].to_string(index=False))
    print("\n" + "=" * 60)
    print("离线AB分析")
    print("=" * 60)
    print(sheets['AB分析_离线'].to_string(index=False))
    print("\n" + "=" * 60)
    print("交互深度")
    print("=" * 60)
    print("在线:")
    print(sheets['交互深度_在线'].to_string(index=False))
    print("离线(有二轮用户):")
    print(sheets['交互深度_离线'].to_string(index=False))


if __name__ == '__main__':
    main()
