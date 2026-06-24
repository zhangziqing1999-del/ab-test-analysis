#!/usr/bin/env python3
"""
AB实验打标+分维度分析脚本
用法: python3 ab_label_analysis.py --file <csv/excel> --groups <sid1,sid2,...> --labels <l1,l2,...>
流程: 预处理→基础AB分析→LLM打标→标签维度AB分析→最终输出
输出: 同目录下生成 AB分析结果_标签维度_<filename>.xlsx
"""
import argparse, os, time, json, hashlib
from concurrent.futures import ThreadPoolExecutor, as_completed
import pandas as pd
import numpy as np
from scipy import stats
import requests

API_URL = "https://oneapi-comate.baidu-int.com/v1/chat/completions"
API_KEY = "sk-qeNvUukW0VTSMp600eC1De92D2154c05AcAe8aD48a45CaC4"
MODELS = ["deepseek-v4-flash", "glm-5.1"]
BATCH_SIZE = 50
CONCURRENCY = 8
SAVE_INTERVAL = 500  # 每500条落盘一次Excel

CATEGORIES = "整体评价类、性价比类、值得买类、用养成本类、口碑类、价格类、场景适配类、保值类、空间类、完整配置类、续航/充电类、驾驶与操控类、智能配置类、舒适配置类、安全配置类、外观类、车系推荐类、其他"

SYSTEM_PROMPT = f"""你是汽车领域query分类器。对每条query严格输出一行JSON，不要任何解释：
{{"type":"客观"或"主观", "category":"子分类"}}

类型规则：
- 客观：有确定答案的事实性问题（参数、价格、配置等）
- 主观：需要个人判断、评价、推荐的问题

可选子分类（只能从以下选一个）：
{CATEGORIES}
如果query与汽车无关或无法归类，category填"其他"。"""


# ===== 工具函数 =====

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


def classify_batch(queries, model=None):
    """调用LLM对一批query打标"""
    if model is None:
        model = MODELS[0]
    content = "\n".join([f"{i+1}. {q}" for i, q in enumerate(queries)])
    user_msg = f"对以下{len(queries)}条query逐条分类，每条输出一行JSON，共{len(queries)}行：\n{content}"
    headers = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}

    for _ in range(3):
        try:
            resp = requests.post(API_URL, headers=headers, json={
                "model": model,
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


def _checkpoint_path(file_path, suffix):
    h = hashlib.md5(os.path.abspath(file_path).encode()).hexdigest()[:8]
    return os.path.join(os.path.dirname(os.path.abspath(file_path)), f".label_ckpt_{h}_{suffix}.json")


def _load_checkpoint(ckpt_path):
    if os.path.exists(ckpt_path):
        with open(ckpt_path, 'r') as f:
            return json.load(f)
    return None


def _save_checkpoint(ckpt_path, data):
    with open(ckpt_path, 'w') as f:
        json.dump(data, f, ensure_ascii=False)


# ===== 阶段1: 预处理 + 入口分组 =====

def stage1_preprocess(file_path, groups, exp_id):
    """读取CSV，提取实验组，按入口分组（保留全量），标记待打标，输出Excel中间文件"""
    out_dir = os.path.dirname(os.path.abspath(file_path))
    base = os.path.splitext(os.path.basename(file_path))[0]
    mid_path = os.path.join(out_dir, f"入口分组_{base}.xlsx")

    # 如果输入本身是入口分组Excel，直接读取
    if file_path.endswith('.xlsx'):
        try:
            xls = pd.ExcelFile(file_path, engine='openpyxl')
            if '在线' in xls.sheet_names and '离线' in xls.sheet_names:
                df_online = pd.read_excel(file_path, sheet_name='在线', engine='openpyxl')
                df_offline = pd.read_excel(file_path, sheet_name='离线', engine='openpyxl')
                print(f"[阶段1] 检测到已有入口分组文件，跳过预处理")
                print(f"  在线: {len(df_online)}, 离线: {len(df_offline)}")
                return df_online, df_offline, file_path
        except:
            pass

    # 如果中间文件已存在，直接读取
    if os.path.exists(mid_path):
        try:
            xls = pd.ExcelFile(mid_path, engine='openpyxl')
            if '在线' in xls.sheet_names and '离线' in xls.sheet_names:
                df_online = pd.read_excel(mid_path, sheet_name='在线', engine='openpyxl')
                df_offline = pd.read_excel(mid_path, sheet_name='离线', engine='openpyxl')
                if len(df_online) > 0 or len(df_offline) > 0:
                    print(f"[阶段1] 从已有中间文件恢复: {mid_path}")
                    print(f"  在线: {len(df_online)}, 离线: {len(df_offline)}")
                    return df_online, df_offline, mid_path
        except:
            pass

    print(f"[阶段1] 预处理: {file_path}")
    df_all = pd.read_csv(file_path)
    print(f"  原始数据: {len(df_all)} 行")

    # 提取实验组
    if '实验组' not in df_all.columns and 'ab实验sid' in df_all.columns:
        if exp_id:
            def extract_group(sid, eid):
                if pd.isna(sid): return None
                for part in str(sid).split('-'):
                    if part.startswith(f"{eid}_"):
                        return part.split('_')[1]
                return None
            df_all['实验组'] = df_all['ab实验sid'].apply(lambda x: extract_group(x, exp_id))

    # 统一转str，避免类型不匹配
    df_all['实验组'] = df_all['实验组'].astype(str)
    groups_str = [str(g) for g in groups]
    df_all = df_all[df_all['实验组'].isin(groups_str)].copy()
    print(f"  筛选实验组{groups_str}: {len(df_all)} 行")

    # 保留全量数据，标记待打标（time_diff!=0 才打标）
    df_all['待打标'] = df_all['time_diff'] != 0
    print(f"  全量: {len(df_all)} 行, 其中待打标(time_diff!=0): {df_all['待打标'].sum()} 行")

    # 按入口分组（保留全量）
    online_set = ['series', 'pin_dao', 'dongtai']
    df_online = df_all[df_all['request_type'].isin(online_set)].copy()
    df_offline = df_all[~df_all['request_type'].isin(online_set)].copy()
    print(f"  在线全量: {len(df_online)} (待打标{df_online['待打标'].sum()}), 离线全量: {len(df_offline)} (待打标{df_offline['待打标'].sum()})")

    # 落盘Excel
    with pd.ExcelWriter(mid_path, engine='openpyxl') as writer:
        df_online.to_excel(writer, sheet_name='在线', index=False)
        df_offline.to_excel(writer, sheet_name='离线', index=False)
    print(f"  中间文件已保存: {mid_path}")

    return df_online, df_offline, mid_path


# ===== 阶段2: 基础AB分析 =====

def compute_metrics(df, control, exp_groups, labels, mode='online'):
    """计算AB指标，mode='online'或'offline'决定指标集"""
    rows = []
    for g in exp_groups:
        ctrl = df[df['实验组'] == control]
        gdf = df[df['实验组'] == g]

        # rank各层数量
        nc1 = len(ctrl[ctrl['rank'] == 1])
        nc2 = len(ctrl[ctrl['rank'] == 2])
        n1 = len(gdf[gdf['rank'] == 1])
        n2 = len(gdf[gdf['rank'] == 2])

        # 完读
        nc1r = int(is_read(ctrl[ctrl['rank'] == 1]['recommend_prompt']).sum())
        nc2r = int(is_read(ctrl[ctrl['rank'] == 2]['recommend_prompt']).sum())
        n1r = int(is_read(gdf[gdf['rank'] == 1]['recommend_prompt']).sum())
        n2r = int(is_read(gdf[gdf['rank'] == 2]['recommend_prompt']).sum())

        # 二轮交互率
        p_interact2 = proportion_z_test(n2, n1, nc2, nc1)

        row = {
            '实验组': labels[g],
            'query数_实验_rank1': n1, 'query数_对照_rank1': nc1,
            '二轮交互率_实验': f"{n2/n1:.4f}" if n1 else '-',
            '二轮交互率_对照': f"{nc2/nc1:.4f}" if nc1 else '-',
            '二轮交互率_p值': round(p_interact2, 4) if not pd.isna(p_interact2) else '-',
            '二轮交互率_显著性': sig(p_interact2),
        }

        if mode == 'online':
            # 一轮完读率
            p_read1 = proportion_z_test(n1r, n1, nc1r, nc1)
            row.update({
                '一轮完读率_实验': f"{n1r/n1:.4f}" if n1 else '-',
                '一轮完读率_对照': f"{nc1r/nc1:.4f}" if nc1 else '-',
                '一轮完读率_p值': round(p_read1, 4) if not pd.isna(p_read1) else '-',
                '一轮完读率_显著性': sig(p_read1),
            })
            # 二轮完读率
            p_read2 = proportion_z_test(n2r, n2, nc2r, nc2)
            row.update({
                '二轮完读率_实验': f"{n2r/n2:.4f}" if n2 else '-',
                '二轮完读率_对照': f"{nc2r/nc2:.4f}" if nc2 else '-',
                '二轮完读率_p值': round(p_read2, 4) if not pd.isna(p_read2) else '-',
                '二轮完读率_显著性': sig(p_read2),
            })
        else:  # offline
            # 二轮→三轮交互率
            nc3 = len(ctrl[ctrl['rank'] == 3])
            n3 = len(gdf[gdf['rank'] == 3])
            p_interact3 = proportion_z_test(n3, n2, nc3, nc2)
            row.update({
                '二轮→三轮交互率_实验': f"{n3/n2:.4f}" if n2 else '-',
                '二轮→三轮交互率_对照': f"{nc3/nc2:.4f}" if nc2 else '-',
                '二轮→三轮交互率_p值': round(p_interact3, 4) if not pd.isna(p_interact3) else '-',
                '二轮→三轮交互率_显著性': sig(p_interact3),
            })
            # 二轮完读率
            p_read2 = proportion_z_test(n2r, n2, nc2r, nc2)
            row.update({
                '二轮完读率_实验': f"{n2r/n2:.4f}" if n2 else '-',
                '二轮完读率_对照': f"{nc2r/nc2:.4f}" if nc2 else '-',
                '二轮完读率_p值': round(p_read2, 4) if not pd.isna(p_read2) else '-',
                '二轮完读率_显著性': sig(p_read2),
            })
            # 三轮完读率
            nc3r = int(is_read(ctrl[ctrl['rank'] == 3]['recommend_prompt']).sum())
            n3r = int(is_read(gdf[gdf['rank'] == 3]['recommend_prompt']).sum())
            p_read3 = proportion_z_test(n3r, n3, nc3r, nc3)
            row.update({
                '三轮完读率_实验': f"{n3r/n3:.4f}" if n3 else '-',
                '三轮完读率_对照': f"{nc3r/nc3:.4f}" if nc3 else '-',
                '三轮完读率_p值': round(p_read3, 4) if not pd.isna(p_read3) else '-',
                '三轮完读率_显著性': sig(p_read3),
            })

        rows.append(row)
    return pd.DataFrame(rows)


def stage2_basic_ab(df_online, df_offline, control, exp_groups, labels, out_path):
    """基础AB分析（无标签），输出到Excel"""
    print(f"\n[阶段2] 基础AB分析")
    sheets = {}

    if len(df_online) > 0:
        sheets['AB分析_在线'] = compute_metrics(df_online, control, exp_groups, labels, mode='online')
        print(f"  在线AB分析完成")
    if len(df_offline) > 0:
        sheets['AB分析_离线'] = compute_metrics(df_offline, control, exp_groups, labels, mode='offline')
        print(f"  离线AB分析完成")

    # 写入Excel
    _append_sheets(out_path, sheets)
    print(f"  结果已追加到: {out_path}")
    return sheets


# ===== 阶段3: LLM打标 =====

def label_dataframe(df, ckpt_suffix, file_path, out_path=None):
    """对df中待打标=True的query进行打标。先对unique query打标再map回，支持断点续传+并发多模型。"""
    # 初始化标签列
    if 'query类型' not in df.columns:
        df['query类型'] = ''
    if '汽车分类' not in df.columns:
        df['汽车分类'] = ''

    # 只对待打标的行打标
    if '待打标' in df.columns:
        mask = df['待打标'].astype(bool)
    else:
        mask = df['time_diff'] != 0

    if mask.sum() == 0:
        print(f"  无待打标数据")
        return df

    # 去重：只对unique query打标
    all_queries = df.loc[mask, 'query'].fillna('').tolist()
    unique_queries = list(dict.fromkeys(all_queries))  # 保序去重
    total = len(unique_queries)
    print(f"  待打标 {len(all_queries)} 条 -> 去重后 {total} 条 unique ({100*(1-total/len(all_queries)):.0f}% 减少)")

    total_batches = (total + BATCH_SIZE - 1) // BATCH_SIZE
    ckpt_path = _checkpoint_path(file_path, ckpt_suffix) if file_path else ''

    # checkpoint存储 unique query 维度的结果
    all_types = [''] * total
    all_cats = [''] * total
    start_batch = 0
    if ckpt_path:
        ckpt = _load_checkpoint(ckpt_path)
        if ckpt and ckpt.get('total') == total:
            all_types = ckpt['types']
            all_cats = ckpt['cats']
            start_batch = ckpt['done_batches']
            print(f"  恢复checkpoint: 已完成 {start_batch}/{total_batches} batches")

    if start_batch < total_batches:
        print(f"  打标中: {total} unique queries, {total_batches} batches, 并发={CONCURRENCY}, 模型={MODELS}")
        t0 = time.time()
        success = sum(1 for t in all_types if t)
        labeled_since_save = 0

        batch_indices = list(range(start_batch * BATCH_SIZE, total, BATCH_SIZE))
        i = 0
        while i < len(batch_indices):
            chunk = batch_indices[i:i+CONCURRENCY]
            futures = {}
            with ThreadPoolExecutor(max_workers=CONCURRENCY) as executor:
                for ci, offset in enumerate(chunk):
                    batch = unique_queries[offset:offset+BATCH_SIZE]
                    model = MODELS[ci % len(MODELS)]
                    futures[executor.submit(classify_batch, batch, model)] = offset
                for future in as_completed(futures):
                    offset = futures[future]
                    results = future.result()
                    for j, r in enumerate(results):
                        if r and isinstance(r, dict):
                            all_types[offset+j] = r.get('type', '')
                            all_cats[offset+j] = r.get('category', '')
                            success += 1

            i += CONCURRENCY
            done_batches = start_batch + i
            done_queries = min(done_batches * BATCH_SIZE, total)
            labeled_since_save += len(chunk) * BATCH_SIZE

            # 保存checkpoint
            if ckpt_path:
                _save_checkpoint(ckpt_path, {
                    'total': total, 'done_batches': done_batches,
                    'types': all_types, 'cats': all_cats
                })

            # 定期落盘Excel（map回后保存）
            if out_path and labeled_since_save >= SAVE_INTERVAL:
                _map_labels_back(df, mask, unique_queries, all_types, all_cats)
                sheet_name = f'标注明细_{ckpt_suffix}'
                labeled = df[mask & (df['query类型'] != '')]
                _append_sheets(out_path, {sheet_name: labeled.head(5000)})
                labeled_since_save = 0

            # 进度报告
            elapsed = time.time() - t0
            speed = (i * BATCH_SIZE) / elapsed if elapsed > 0 else 0
            remaining = total - done_queries
            eta = remaining / speed if speed > 0 else 0
            print(f"    [{done_batches}/{total_batches}] {done_queries}/{total} ({100*done_queries//total}%) | "
                  f"耗时{elapsed:.0f}s | 剩余{eta:.0f}s | 成功率{100*success//max(done_queries,1)}%")
        print(f"  打标完成，成功率: {success*100//total}%")
    else:
        print(f"  打标已完成（从checkpoint恢复）")

    # map回全部待打标行
    _map_labels_back(df, mask, unique_queries, all_types, all_cats)
    # 注意：不在这里删除checkpoint，等全流程结束后统一清理
    return df


def _map_labels_back(df, mask, unique_queries, all_types, all_cats):
    """将unique query的标签结果map回df的待打标行"""
    type_map = dict(zip(unique_queries, all_types))
    cat_map = dict(zip(unique_queries, all_cats))
    q = df.loc[mask, 'query'].fillna('')
    df.loc[mask, 'query类型'] = q.map(type_map).fillna('')
    df.loc[mask, '汽车分类'] = q.map(cat_map).fillna('')


def stage3_labeling(df_online, df_offline, mid_path, out_path):
    """LLM打标阶段（只对待打标=True的行）"""
    print(f"\n[阶段3] LLM打标")

    if len(df_online) > 0:
        print(f"  在线打标")
        df_online = label_dataframe(df_online, 'online', mid_path, out_path)
    if len(df_offline) > 0:
        print(f"  离线打标")
        df_offline = label_dataframe(df_offline, 'offline', mid_path, out_path)

    # 打标完成后更新中间文件
    with pd.ExcelWriter(mid_path, engine='openpyxl') as writer:
        df_online.to_excel(writer, sheet_name='在线', index=False)
        df_offline.to_excel(writer, sheet_name='离线', index=False)
    print(f"  标签已写回中间文件: {mid_path}")

    return df_online, df_offline


# ===== 阶段4: 基于标签的AB分析 =====

def _session_rank_presence(df):
    """返回 {session_id: set(ranks)}，以及 {(session_id,rank): row_index} 便于查完读"""
    sess_ranks = df.groupby('session_id')['rank'].apply(set).to_dict()
    return sess_ranks


def compute_label_metrics(df, control, exp_groups, labels, dim_col, anchor_rank):
    """按标签锚定在某个rank做AB分析。
    anchor_rank=1(在线): 看 二轮渗透率(rank2/rank1)、一轮完读率、二轮完读率
    anchor_rank=2(离线): 看 三轮渗透率(rank3/rank2)、二轮完读率、三轮完读率
    标签取该query自身的标签（不传播），分母为该标签锚定rank的query数。
    """
    next_rank = anchor_rank + 1
    rows = []
    # 该维度所有标签值
    labeled_anchor = df[(df['rank'] == anchor_rank) &
                        df[dim_col].notna() &
                        (df[dim_col].astype(str).str.strip() != '') &
                        (df[dim_col].astype(str) != 'nan')]
    dim_vals = sorted(labeled_anchor[dim_col].astype(str).unique())

    for dim_val in dim_vals:
        # 对照组锚定
        def anchored(group):
            gd = df[df['实验组'] == group]
            a = gd[(gd['rank'] == anchor_rank) & (gd[dim_col].astype(str) == dim_val)]
            sids = set(a['session_id'])
            # 下一轮：同session且rank==next_rank
            nxt = gd[(gd['rank'] == next_rank) & (gd['session_id'].isin(sids))]
            return a, nxt
        ca, cn = anchored(control)
        nc_anchor = len(ca)
        nc_next = len(cn)
        nc_anchor_r = int(is_read(ca['recommend_prompt']).sum())
        nc_next_r = int(is_read(cn['recommend_prompt']).sum())

        for g in exp_groups:
            ga, gn = anchored(g)
            n_anchor = len(ga)
            n_next = len(gn)
            n_anchor_r = int(is_read(ga['recommend_prompt']).sum())
            n_next_r = int(is_read(gn['recommend_prompt']).sum())

            # 渗透率 = next_rank数 / anchor_rank数
            p_pen = proportion_z_test(n_next, n_anchor, nc_next, nc_anchor)
            # anchor轮完读率
            p_ar = proportion_z_test(n_anchor_r, n_anchor, nc_anchor_r, nc_anchor)
            # next轮完读率
            p_nr = proportion_z_test(n_next_r, n_next, nc_next_r, nc_next)

            pen_name = f"{_cn(anchor_rank)}→{_cn(next_rank)}转化率"
            ar_name = f"{_cn(anchor_rank)}完读率"
            nr_name = f"{_cn(next_rank)}完读率"

            rows.append({
                dim_col: dim_val, '实验组': labels[g],
                f'query数_实验_rank{anchor_rank}': n_anchor,
                f'query数_对照_rank{anchor_rank}': nc_anchor,
                f'{pen_name}_实验': f"{n_next/n_anchor:.4f}" if n_anchor else '-',
                f'{pen_name}_对照': f"{nc_next/nc_anchor:.4f}" if nc_anchor else '-',
                f'{pen_name}_p值': round(p_pen, 4) if not pd.isna(p_pen) else '-',
                f'{pen_name}_显著性': sig(p_pen),
                f'{ar_name}_实验': f"{n_anchor_r/n_anchor:.4f}" if n_anchor else '-',
                f'{ar_name}_对照': f"{nc_anchor_r/nc_anchor:.4f}" if nc_anchor else '-',
                f'{ar_name}_p值': round(p_ar, 4) if not pd.isna(p_ar) else '-',
                f'{ar_name}_显著性': sig(p_ar),
                f'{nr_name}_实验': f"{n_next_r/n_next:.4f}" if n_next else '-',
                f'{nr_name}_对照': f"{nc_next_r/nc_next:.4f}" if nc_next else '-',
                f'{nr_name}_p值': round(p_nr, 4) if not pd.isna(p_nr) else '-',
                f'{nr_name}_显著性': sig(p_nr),
            })
    return pd.DataFrame(rows)


def _cn(rank):
    """rank数字转中文轮次"""
    m = {1: '一轮', 2: '二轮', 3: '三轮', 4: '四轮', 5: '五轮'}
    return m.get(rank, f'{rank}轮')


def stage4_label_ab(df_online, df_offline, control, exp_groups, labels, out_path):
    """基于标签维度的AB分析。在线锚定rank=1，离线锚定rank=2。标签取query自身，不传播。"""
    print(f"\n[阶段4] 基于标签的AB分析")
    sheets = {}

    if len(df_online) > 0 and 'query类型' in df_online.columns:
        sheets['按query类型_在线'] = compute_label_metrics(df_online, control, exp_groups, labels, 'query类型', anchor_rank=1)
        sheets['按汽车分类_在线'] = compute_label_metrics(df_online, control, exp_groups, labels, '汽车分类', anchor_rank=1)
        print(f"  在线标签分析完成（锚定rank=1）")

    if len(df_offline) > 0 and 'query类型' in df_offline.columns:
        sheets['按query类型_离线'] = compute_label_metrics(df_offline, control, exp_groups, labels, 'query类型', anchor_rank=2)
        sheets['按汽车分类_离线'] = compute_label_metrics(df_offline, control, exp_groups, labels, '汽车分类', anchor_rank=2)
        print(f"  离线标签分析完成（锚定rank=2）")

    _append_sheets(out_path, sheets)
    print(f"  结果已追加到: {out_path}")
    return sheets


# ===== 阶段5: 最终输出 =====

def stage5_finalize(df_online, df_offline, mid_path, out_path, all_sheets):
    """最终输出：标注明细 + 清理checkpoint"""
    print(f"\n[阶段5] 最终输出")

    # 添加标注明细sheets（只展示已打标的行，全量展示）
    detail_cols = ['format_date', 'session_id', 'query', 'rank', 'time_diff', 'request_type', '实验组', 'query类型', '汽车分类']

    def labeled_rows(d):
        cols = [c for c in detail_cols if c in d.columns]
        m = d['query类型'].notna() & (d['query类型'].astype(str).isin(['主观', '客观', '其他']))
        return d.loc[m, cols]

    if len(df_online) > 0:
        all_sheets['标注明细_在线'] = labeled_rows(df_online)
    if len(df_offline) > 0:
        all_sheets['标注明细_离线'] = labeled_rows(df_offline)

    # 写入最终Excel
    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        for name, d in all_sheets.items():
            if isinstance(d, pd.DataFrame) and len(d) > 0:
                d.to_excel(writer, sheet_name=name[:31], index=False)

    print(f"  最终结果: {out_path}")
    print(f"  sheets: {list(all_sheets.keys())}")

    # 清理checkpoint
    for suffix in ['online', 'offline']:
        ckpt = _checkpoint_path(mid_path, suffix)
        if os.path.exists(ckpt):
            os.remove(ckpt)
            print(f"  清理checkpoint: {ckpt}")

    # 打印关键结论
    for key in ['AB分析_在线', 'AB分析_离线', '按query类型_在线', '按query类型_离线', '按汽车分类_在线', '按汽车分类_离线']:
        if key in all_sheets and len(all_sheets[key]) > 0:
            print(f"\n{'='*50}\n{key}\n{'='*50}")
            print(all_sheets[key].to_string(index=False))


# ===== 辅助: 追加sheets到Excel =====

def _append_sheets(path, sheets):
    """将sheets追加/覆盖到Excel文件"""
    if not sheets:
        return
    if os.path.exists(path):
        from openpyxl import load_workbook
        book = load_workbook(path)
        with pd.ExcelWriter(path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            for name, d in sheets.items():
                if isinstance(d, pd.DataFrame):
                    d.to_excel(writer, sheet_name=name[:31], index=False)
    else:
        with pd.ExcelWriter(path, engine='openpyxl') as writer:
            for name, d in sheets.items():
                if isinstance(d, pd.DataFrame):
                    d.to_excel(writer, sheet_name=name[:31], index=False)


# ===== 主函数 =====

def main():
    parser = argparse.ArgumentParser(description='AB实验打标+分维度分析')
    parser.add_argument('--file', required=True)
    parser.add_argument('--groups', required=True)
    parser.add_argument('--labels', required=True)
    parser.add_argument('--exp-id', default=None, help='实验ID（如225457），CSV时从ab实验sid中提取')
    args = parser.parse_args()

    groups = [str(g) for g in args.groups.split(',')]
    labels_list = args.labels.split(',')
    if len(groups) != len(labels_list):
        print("ERROR: groups数量必须和labels数量一致"); return

    control = groups[0]
    exp_groups = groups[1:]
    labels = dict(zip(groups, labels_list))

    file_path = args.file
    out_dir = os.path.dirname(os.path.abspath(file_path))
    base = os.path.splitext(os.path.basename(file_path))[0]
    out_path = os.path.join(out_dir, f"AB分析结果_标签维度_{base}.xlsx")

    # 阶段1
    df_online, df_offline, mid_path = stage1_preprocess(file_path, groups, args.exp_id)

    # 阶段2
    sheets_basic = stage2_basic_ab(df_online, df_offline, control, exp_groups, labels, out_path)

    # 阶段3
    df_online, df_offline = stage3_labeling(df_online, df_offline, mid_path, out_path)

    # 阶段4
    sheets_label = stage4_label_ab(df_online, df_offline, control, exp_groups, labels, out_path)

    # 阶段5
    all_sheets = {}
    all_sheets.update(sheets_basic)
    all_sheets.update(sheets_label)
    stage5_finalize(df_online, df_offline, mid_path, out_path, all_sheets)


if __name__ == '__main__':
    main()
