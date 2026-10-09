# -*- coding: utf-8 -*-
"""
序时账抽样（voucher-sampling skill v2 执行器）
=============================================
确定性部分全部在本脚本内完成：列识别、凭证聚合、重复行标记、借贷平衡检查、
分层、样本量分配、随机抽样、全量风险筛查、输出工作簿、打印结构化摘要。

按 SKILL.md 的固定规则实现，参数以本脚本为唯一权威来源。
AI 不得在脚本之外自行计算笔数、金额、抽样率或覆盖率。

依赖：pandas + openpyxl（不依赖 sklearn/scipy）

用法：
  python sample_vouchers.py --input "序时账.xlsx" --output "序时账抽样清单.xlsx"
  python sample_vouchers.py --input "序时账.xlsx" --output "out.xlsx" --target 30
  python sample_vouchers.py --input "序时账.xlsx" --output "out.xlsx" --seed 20250401 --sheet 序时账
"""
from __future__ import annotations

import argparse
import hashlib
import json
import math
import re
import sys
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# 强制 UTF-8 输出：避免在非 UTF-8 控制台/管道中出现中文乱码（含 argparse --help）
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding='utf-8')
    except Exception:
        pass

# ============================================================
# 固定规则常量（SKILL.md §二 / §三）
# ============================================================
BINS = [0.0, 1_000.0, 5_000.0, 20_000.0, 100_000.0, 500_000.0, 10_000_000.0, math.inf]
LAYERS = ['A', 'B', 'C', 'D', 'E', 'F', 'G']
LAYER_LABEL = {
    'A': 'A ≤1千元', 'B': 'B 1千-5千元', 'C': 'C 5千-2万元', 'D': 'D 2万-10万元',
    'E': 'E 10万-50万元', 'F': 'F 50万-1000万元', 'G': 'G >1000万元',
    'Z': 'Z 零金额',
}
# 默认配置（未给 --target 时）
DEFAULT_PLAN = {
    'G': ('all', None), 'F': ('toprandom', (40, 20)), 'E': ('toprandom', (25, 15)),
    'D': ('toprandom', (15, 15)), 'C': ('random', 20), 'B': ('random', 12),
    'A': ('random', 8), 'Z': ('all', None),
}
# Z 层（零金额凭证）抽样率上限：零金额凭证不产生金额覆盖，限额抽样即可
Z_LAYER_MAX = 3
RISK_DEFS = [
    ('整数金额', '|金额| ≥ 1万 且为整千'),
    ('冲销调整', '摘要含 冲销/冲回/调整/更正/红字'),
    ('借款备用金往来', '摘要含 借款/备用金/往来/暂付/暂收/借支'),
    ('负数红字凭证', '凭证金额 < 0'),
    ('含重复分录', '同凭证 科目编码+借方金额+贷方金额+部门 完全相同（排除分录数≥50）'),
    ('分录数异常', '凭证分录行数 ≥ 50'),
    ('摘要缺失', '摘要为空或去空格后 ≤2 字'),
    ('借贷不平', '|借方合计-贷方合计| > 容差'),
]
FLAG_COLS = ['风险-' + n for n, _ in RISK_DEFS]

ALIAS = {
    '凭证编号': ['凭证编号', '凭证号', '凭证字号', '记账凭证号', '凭证号数'],
    '制单日期': ['制单日期', '日期', '记账日期', '业务日期', '制证日期'],
    '摘要': ['摘要', '业务摘要', '摘要说明', '摘要内容'],
    '科目编码': ['科目编码', '科目代码', '科目号', '会计科目编码'],
    '科目名称': ['科目名称', '科目', '会计科目'],
    '部门名称': ['部门名称', '部门'],
    '借方金额': ['借方金额', '借方', '借方发生额', '借方本币金额'],
    '贷方金额': ['贷方金额', '贷方', '贷方发生额', '贷方本币金额'],
    '制单人': ['制单人', '制单', '编制人'],
    '审核人': ['审核人', '审核'],
}

HDR_FILL = PatternFill('solid', fgColor='305496')
HDR_FONT = Font(bold=True, color='FFFFFF', size=10)
TITLE_FONT = Font(bold=True, size=14)
THIN = Side(style='thin', color='B0B0B0')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
RED = PatternFill('solid', fgColor='FFC7CE')


def log(msg: str = '') -> None:
    """输出摘要。强制 UTF-8，避免在非 UTF-8 控制台或管道中出现乱码。"""
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
    print(msg, flush=True)


def norm(s) -> str:
    return re.sub(r'[\s\u3000]+', '', str(s)).lower()


# ============================================================
# 1. 读取与列识别
# ============================================================
def find_header_and_sheet(path: Path, sheet: str | None):
    """返回 (sheet名, header行号(0基), DataFrame)。逐个 sheet 尝试，不靠猜。"""
    xls = pd.ExcelFile(path)
    candidates = [sheet] if sheet else list(xls.sheet_names)
    tried = []
    for name in candidates:
        if name not in xls.sheet_names:
            raise SystemExit(f'[错误] 工作表不存在：{name}；可选：{xls.sheet_names}')
        for hdr in range(0, 8):
            df = xls.parse(name, header=hdr)
            if df.dropna(how='all').empty:
                continue
            cols = [norm(c) for c in df.columns]
            hit = sum(1 for f, al in ALIAS.items()
                      if any(norm(a) in cols for a in al))
            if hit >= 3:
                return name, hdr, df, tried, hit
        tried.append(f'{name}(无有效表头)')
    raise SystemExit(f'[错误] 未能在任何工作表定位表头（需至少识别 3 个字段）。尝试过：{tried}')


def map_columns(df: pd.DataFrame) -> dict:
    """按别名映射字段，返回 {字段: 实际列名}。"""
    cols = list(df.columns)
    nc = {c: norm(c) for c in cols}
    mapping = {}
    for field, aliases in ALIAS.items():
        for a in aliases:
            na = norm(a)
            exact = [c for c in cols if nc[c] == na]
            if exact:
                mapping[field] = exact[0]
                break
            part = [c for c in cols if na in nc[c]]
            if part:
                mapping[field] = part[0]
                break
    return mapping


def to_num(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.replace(r'[,\s\u3000￥¥]', '', regex=True)
    s = s.replace({'': np.nan, 'nan': np.nan, 'None': np.nan, '-': np.nan})
    return pd.to_numeric(s, errors='coerce')


# ============================================================
# 2. 凭证聚合与质量检查
# ============================================================
def build_vouchers(df: pd.DataFrame, mapping: dict, tol: float):
    q = {}
    d = pd.DataFrame()
    d['凭证编号'] = df[mapping['凭证编号']].astype(str).str.strip()
    d['制单日期'] = pd.to_datetime(df[mapping['制单日期']], errors='coerce') if '制单日期' in mapping else pd.NaT
    d['摘要'] = df[mapping['摘要']].fillna('').astype(str).str.strip() if '摘要' in mapping else ''
    d['科目编码'] = df[mapping['科目编码']].fillna('').astype(str).str.strip() if '科目编码' in mapping else ''
    d['科目名称'] = df[mapping['科目名称']].fillna('').astype(str).str.strip() if '科目名称' in mapping else ''
    d['部门名称'] = df[mapping['部门名称']].fillna('').astype(str).str.strip() if '部门名称' in mapping else ''
    d['制单人'] = df[mapping['制单人']].fillna('').astype(str).str.strip() if '制单人' in mapping else ''
    d['审核人'] = df[mapping['审核人']].fillna('').astype(str).str.strip() if '审核人' in mapping else ''
    d['借方金额'] = to_num(df[mapping['借方金额']]) if '借方金额' in mapping else 0.0
    d['贷方金额'] = to_num(df[mapping['贷方金额']]) if '贷方金额' in mapping else 0.0

    # 只有"非空但解析失败"才算无法解析（拆分列中一列为空属正常）
    raw_dr = df[mapping['借方金额']] if '借方金额' in mapping else pd.Series([pd.NA] * len(df))
    raw_cr = df[mapping['贷方金额']] if '贷方金额' in mapping else pd.Series([pd.NA] * len(df))
    bad = 0
    for raw, parsed in ((raw_dr, d['借方金额']), (raw_cr, d['贷方金额'])):
        rs = raw.astype('string').str.strip()
        nonempty = rs.notna() & (rs != '') & (~rs.isin(['nan', 'None', '-']))
        bad += int((nonempty & parsed.isna()).sum())
    q['无法解析金额单元格数'] = bad
    d['借方金额'] = d['借方金额'].fillna(0.0)
    d['贷方金额'] = d['贷方金额'].fillna(0.0)

    d = d[d['凭证编号'].str.len() > 0].copy()
    q['读入分录行数'] = int(len(d))

    # 完全重复行：全部保留，仅标记（不静默删除，避免吞掉金额）
    # 判据不含摘要：摘要措辞可能逐行微调，但"同凭证+同科目+同金额+同部门"重复仍是重复
    dup_key = ['凭证编号', '科目编码', '借方金额', '贷方金额', '部门名称']
    all_dup = d.duplicated(subset=dup_key, keep=False)
    extra_dup = d.duplicated(subset=dup_key, keep='first')
    q['完全重复行数'] = int(all_dup.sum())
    q['完全重复行涉及凭证数'] = int(d.loc[all_dup, '凭证编号'].nunique())
    q['重复行多计借方'] = float(d.loc[extra_dup, '借方金额'].sum())
    q['重复行多计贷方'] = float(d.loc[extra_dup, '贷方金额'].sum())
    d['完全重复行'] = all_dup

    g = d.groupby('凭证编号', sort=True).agg(
        制单日期=('制单日期', 'min'),
        制单人=('制单人', 'first'),
        审核人=('审核人', 'first'),
        行数=('摘要', 'size'),
        借方合计=('借方金额', 'sum'),
        贷方合计=('贷方金额', 'sum'),
        科目数=('科目编码', lambda s: int(s[s != ''].nunique())),
        含空摘要=('摘要', lambda s: bool((s.str.strip().str.len() <= 2).any())),
        重复行数=('完全重复行', 'sum'),
    )
    g['摘要'] = d.groupby('凭证编号')['摘要'].apply(
        lambda s: '\n'.join(dict.fromkeys([x for x in s if x.strip()])))

    g['借方合计'] = g['借方合计'].round(2)
    g['贷方合计'] = g['贷方合计'].round(2)
    g['凭证金额'] = g[['借方合计', '贷方合计']].max(axis=1).round(2)
    g['绝对金额'] = g['凭证金额'].abs().round(2)
    g['差额'] = (g['借方合计'] - g['贷方合计']).round(2)
    g['不平'] = g['差额'].abs() > tol
    q['凭证张数'] = int(len(g))
    q['不平凭证张数'] = int(g['不平'].sum())
    q['不平差额合计'] = float(g.loc[g['不平'], '差额'].abs().sum())
    q['借方合计'] = float(g['借方合计'].sum())
    q['贷方合计'] = float(g['贷方合计'].sum())
    q['凭证金额合计'] = float(g['绝对金额'].sum())
    q['日期最小'] = None if g['制单日期'].isna().all() else str(g['制单日期'].min().date())
    q['日期最大'] = None if g['制单日期'].isna().all() else str(g['制单日期'].max().date())
    if q['日期最小']:
        per = g['制单日期'].dt.to_period('M').value_counts().sort_index()
        q['各月凭证数'] = {str(k): int(v) for k, v in per.items()}

    # 重复分录（凭证级）：存在完全重复行，且排除汇总类凭证（行数 ≥ 50）
    dup_keys = set(d.loc[d['完全重复行'], '凭证编号'])
    g['含重复分录'] = g.index.isin(dup_keys) & (g['行数'] < 50)
    g['分录数异常'] = g['行数'] >= 50
    g['负数红字凭证'] = g['凭证金额'] < 0
    g['摘要缺失'] = g['含空摘要']
    g['借贷不平'] = g['不平']
    g['整数金额'] = (g['绝对金额'] >= 10_000) & (np.isclose(g['绝对金额'] % 1_000, 0, atol=1e-6))
    smry = g['摘要'].astype(str)
    g['冲销调整'] = smry.str.contains('冲销|冲回|调整|更正|红字', regex=True)
    g['借款备用金往来'] = smry.str.contains('借款|备用金|往来|暂付|暂收|借支', regex=True)

    g['层'] = pd.cut(g['绝对金额'], bins=BINS, labels=LAYERS, right=True)
    g['层'] = g['层'].astype(object).where(g['绝对金额'] > 0, 'Z')
    g.loc[g['层'].isna(), '层'] = 'Z'
    return g, q


# ============================================================
# 3. 样本量分配
# ============================================================
def allocate(target: int, sizes: dict, amounts: dict, order: list) -> dict:
    """样本量分配（重写版，含不变量断言）。

    分配顺序：
      1) 每层最小 1 张（含 Z 层；额度不足时按金额从低到高让步）
      2) G 层（金额最高层）尽量全选，但不得挤占其他层的 1 张
      3) Z 层（零金额）上限 Z_LAYER_MAX
      4) 余额按 sqrt(层金额)×层张数 加权，一次最大余数法发完
    """
    alloc = {k: 0 for k in order}
    if target <= 0:
        return alloc
    cap = {k: sizes[k] for k in order}
    n_layers = sum(1 for k in order if cap[k] > 0)
    budget = min(target, sum(cap.values()))

    # ---- 1) G 层定档：为其他层各留 1 张后，G 尽量全选 ----
    others = [k for k in order if k != 'G' and cap[k] > 0]
    if 'G' in order and cap['G'] > 0:
        alloc['G'] = max(1, min(cap['G'], budget - len(others)))
    left = budget - sum(alloc.values())

    # ---- 2) 其余层最小 1 张；额度不足时优先满足高金额层 ----
    for k in order:
        if k == 'G' or cap[k] == 0:
            continue
        if left > 0:
            alloc[k] = 1
            left -= 1
    short = [k for k in order if k != 'G' and cap[k] > 0 and alloc[k] == 0]
    for k in short:                      # 从 G 让位（G 至少保留 1 张）
        if alloc['G'] > 1:
            alloc['G'] -= 1
            alloc[k] = 1
            left -= 1

    # ---- 3) Z 层上限 ----
    z_cap = min(cap.get('Z', 0), Z_LAYER_MAX)
    if 'Z' in order and alloc['Z'] > z_cap:
        left += alloc['Z'] - z_cap
        alloc['Z'] = z_cap

    # ---- 4) 余额按权重一次发完（最大余数法，永不超发） ----
    if left > 0:
        room = {}
        for k in order:
            limit = z_cap if k == 'Z' else cap[k]
            if limit - alloc[k] > 0:
                room[k] = limit - alloc[k]
        if room:
            w = {k: math.sqrt(max(amounts.get(k, 0.0), 0.0) + 1.0) * sizes.get(k, 0) for k in room}
            tot_w = sum(w.values())
            if tot_w > 0:
                quota = {k: left * w[k] / tot_w for k in room}
                gave = {k: min(int(math.floor(quota[k])), room[k]) for k in room}
                # 逐层发放，且总量不超过 left
                sent = 0
                for k in sorted(room, key=lambda x: -w[x]):
                    take = min(gave[k], left - sent)
                    alloc[k] += take
                    sent += take
                # 余数按小数部分降序补，补到 left 用完或所有层到顶
                rest = left - sent
                for k in sorted(room, key=lambda x: (-(quota[x] - gave[x]), -w[x])):
                    if rest <= 0:
                        break
                    limit = z_cap if k == 'Z' else cap[k]
                    if alloc[k] < limit:
                        alloc[k] += 1
                        rest -= 1
                # 仍有剩余（取整损耗）按权重顺序补满
                for k in sorted(room, key=lambda x: -w[x]):
                    if rest <= 0:
                        break
                    limit = z_cap if k == 'Z' else cap[k]
                    add = min(rest, limit - alloc[k])
                    alloc[k] += add
                    rest -= add
                left = rest

    # ---- 5) 单调化整理：高层抽样率不得低于低层（每层至少 1 张的硬约束下尽力而为） ----
    main = [k for k in order if k != 'Z' and cap[k] > 0]
    for i in range(len(main) - 2, -1, -1):
        lo, hi = main[i], main[i + 1]
        if cap[hi] == 0:
            continue
        rate_hi = alloc[hi] / cap[hi]
        cap_lo = math.floor(rate_hi * cap[lo] + 1e-9)
        # 不低于 1 张：每层最小覆盖优先于单调性
        cap_lo = max(cap_lo, 1 if cap[lo] > 0 else 0)
        if alloc[lo] > cap_lo:
            freed = alloc[lo] - cap_lo
            alloc[lo] = cap_lo
            for k in main[::-1]:
                if freed <= 0:
                    break
                add = min(freed, cap[k] - alloc[k])
                alloc[k] += add
                freed -= add

    # ---- 不变量断言：总量不超过预算、各层不超上限、每层不超总体 ----
    assert sum(alloc.values()) <= budget, (sum(alloc.values()), budget)
    for k in order:
        limit = z_cap if k == 'Z' else cap[k]
        assert 0 <= alloc[k] <= limit, (k, alloc[k], limit)
    return alloc


def default_alloc(sizes: dict, order: list) -> dict:
    alloc = {}
    for k in order:
        kind, param = DEFAULT_PLAN.get(k, ('random', 0))
        n = sizes.get(k, 0)
        if kind == 'all':
            alloc[k] = n
        elif kind == 'random':
            alloc[k] = min(param or 0, n)
        elif kind == 'toprandom':
            top, rnd = param
            alloc[k] = min((top or 0) + (rnd or 0), n)
        else:
            alloc[k] = 0
    if 'Z' in alloc:
        alloc['Z'] = min(alloc['Z'], Z_LAYER_MAX)
    return alloc


# ============================================================
# 4. 抽样
# ============================================================
def draw(sub: pd.DataFrame, layer: str, alloc_n: int, rng) -> tuple[list, str]:
    """返回 (选中的凭证编号列表, 抽样方式描述)"""
    n_total = len(sub)
    if alloc_n <= 0 or n_total == 0:
        return [], '未抽取'
    if alloc_n >= n_total:
        return sub.index.tolist(), '全选（100%检查）'
    kind, param = DEFAULT_PLAN.get(layer, ('random', alloc_n))
    if kind == 'toprandom' and param:
        top, _ = param
        top = max(1, round(alloc_n * top / (top + param[1])))
        deep = sub.head(top).index.tolist()
        rest = sub.iloc[top:].index.tolist()
        need = alloc_n - len(deep)
        rand = list(rng.choice(rest, size=need, replace=False)) if need > 0 and rest else []
        return deep + rand, f'金额降序前{len(deep)}张 + 随机{len(rand)}张'
    picked = list(rng.choice(sub.index.tolist(), size=alloc_n, replace=False))
    return picked, '随机抽样'


# ============================================================
# 5. 输出
# ============================================================
def write_sheet(ws, title, df, widths, money_cols=(), wrap_cols=(), notes=None):
    if title:
        ws['A1'] = title
        ws['A1'].font = TITLE_FONT
    r0 = 3
    if notes:
        for i, nt in enumerate(notes):
            ws.cell(row=2 + i, column=1, value=nt).font = Font(size=9, italic=True, color='808080')
        r0 = 2 + len(notes) + 1
    for j, c in enumerate(df.columns, 1):
        cell = ws.cell(row=r0, column=j, value=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = BORDER
    for i, (_, row) in enumerate(df.iterrows(), r0 + 1):
        for j, c in enumerate(df.columns, 1):
            v = row[c]
            if isinstance(v, (np.integer,)):
                v = int(v)
            elif isinstance(v, (np.floating,)):
                v = float(v)
            cell = ws.cell(row=i, column=j, value=v)
            cell.border = BORDER
            cell.font = Font(size=10)
            if c in money_cols:
                cell.number_format = '#,##0.00'
            if c in wrap_cols:
                cell.alignment = Alignment(wrap_text=True, vertical='top')
            else:
                cell.alignment = Alignment(vertical='top')
    for j, c in enumerate(df.columns, 1):
        ws.column_dimensions[get_column_letter(j)].width = widths.get(c, 12)
    ws.freeze_panes = ws.cell(row=r0 + 1, column=1)
    return r0


def main() -> int:
    ap = argparse.ArgumentParser(description='序时账抽样（voucher-sampling v2）')
    ap.add_argument('--input', '-i', required=True, help='序时账 Excel/CSV 路径')
    ap.add_argument('--output', '-o', required=True, help='输出工作簿路径（新建，不覆盖输入）')
    ap.add_argument('--target', '-t', type=int, default=None, help='目标样本量（凭证张数）；不给则用默认分层配置')
    ap.add_argument('--seed', type=int, default=20250401, help='随机种子（默认 20250401）')
    ap.add_argument('--tol', type=float, default=0.01, help='借贷平衡容差（默认 0.01）')
    ap.add_argument('--sheet', default=None, help='指定工作表名；默认自动识别')
    ap.add_argument('--json-out', default=None, help='把结构化摘要另存为 JSON')
    args = ap.parse_args()

    src = Path(args.input)
    if not src.exists():
        raise SystemExit(f'[错误] 输入文件不存在：{src}')
    out = Path(args.output)
    if out.resolve() == src.resolve():
        raise SystemExit('[错误] 输出路径不得与输入文件相同（本脚本不改动原始序时账）')
    if src.suffix.lower() == '.csv':
        df = pd.read_csv(src)
        sheet_name, hdr_row, tried, hit = None, 0, [], 0
    else:
        sheet_name, hdr_row, df, tried, hit = find_header_and_sheet(src, args.sheet)

    mapping = map_columns(df)
    missing = [f for f in ['凭证编号'] if f not in mapping]
    if missing:
        raise SystemExit(f'[错误] 未识别到必需字段 {missing}；实际列：{list(df.columns)}')
    if '借方金额' not in mapping and '贷方金额' not in mapping:
        raise SystemExit(f'[错误] 未识别到借方/贷方金额列；实际列：{list(df.columns)}')
    if '凭证编号' not in mapping and '制单日期' not in mapping:
        raise SystemExit('[错误] 凭证编号与制单日期均缺失，无法标识抽样单元')

    sha = hashlib.sha256(src.read_bytes()).hexdigest()
    g, q = build_vouchers(df, mapping, args.tol)

    order = [L for L in LAYERS if (g['层'] == L).sum() > 0] + (['Z'] if (g['层'] == 'Z').sum() > 0 else [])
    sizes = {L: int((g['层'] == L).sum()) for L in order}
    amounts = {L: float(g.loc[g['层'] == L, '绝对金额'].sum()) for L in order}
    if args.target is not None:
        alloc = allocate(args.target, sizes, amounts, order)
    else:
        alloc = default_alloc(sizes, order)

    rng = np.random.default_rng(args.seed)
    picked, how = {}, {}
    for L in order:
        sub = g[g['层'] == L].sort_values(['绝对金额', '凭证编号'], ascending=[False, True])
        ids, desc = draw(sub, L, alloc[L], rng)
        picked[L] = ids
        how[L] = desc

    all_ids = [i for L in order for i in picked[L]]
    sel = g.loc[all_ids].copy().reset_index()
    sel['层'] = pd.Categorical(sel['层'], categories=order, ordered=True)
    sel['层序'] = sel['层'].map({L: i for i, L in enumerate(order)})
    sel['抽样方式'] = sel['层'].map(how)
    sel['层内全选'] = sel['层'].map(lambda L: int(alloc[L] >= sizes[L]))
    sel = sel.sort_values(['层序', '绝对金额'], ascending=[True, False]).reset_index(drop=True)
    sel['序号'] = range(1, len(sel) + 1)
    sel['层标签'] = sel['层'].map(LAYER_LABEL)
    sel['抽样说明'] = sel.apply(
        lambda r: f"{r['层标签']}分层；{r['抽样方式']}", axis=1)

    flag_src = {
        '整数金额': sel['整数金额'], '冲销调整': sel['冲销调整'],
        '借款备用金往来': sel['借款备用金往来'], '负数红字凭证': sel['负数红字凭证'],
        '含重复分录': sel['含重复分录'], '分录数异常': sel['分录数异常'],
        '摘要缺失': sel['摘要缺失'], '借贷不平': sel['借贷不平'],
    }
    for name, mask in flag_src.items():
        sel['风险-' + name] = np.where(mask, '√', '')
    sel['风险标记数'] = (sel[FLAG_COLS] == '√').sum(axis=1)
    sel['风险标记'] = sel[FLAG_COLS].apply(
        lambda r: '、'.join(c.replace('风险-', '') for c in FLAG_COLS if r[c] == '√'), axis=1)

    # ---- 分层方案表 ----
    plan_rows = []
    for L in order:
        sub = g[g['层'] == L]
        n, amt = len(sub), float(sub['绝对金额'].sum())
        rate = (alloc[L] / n) if n else 0.0
        plan_rows.append({
            '层': LAYER_LABEL[L], '总体凭证数': n, '样本数': alloc[L],
            '抽样率': f'{rate * 100:.1f}%', '层金额合计(元)': amt,
            '层金额估算(元)': round(amt * rate, 2), '抽样方法': how[L],
        })
    plan_df = pd.DataFrame(plan_rows)
    plan_df.loc[len(plan_df)] = [
        '合计', q['凭证张数'], int(plan_df['样本数'].sum()),
        f"{plan_df['样本数'].sum() / q['凭证张数'] * 100:.1f}%",
        q['凭证金额合计'], round(plan_df['层金额估算(元)'].sum(), 2), '—']

    # ---- 全量风险分布 ----
    risk_rows = []
    for name, crit in RISK_DEFS:
        mask = g[name] if name in g.columns else pd.Series(False, index=g.index)
        risk_rows.append({'风险标记': name, '判据': crit,
                          '全量涉及张数': int(mask.sum()),
                          '全量涉及金额(元)': float(g.loc[mask, '绝对金额'].sum()),
                          '样本涉及张数': int(sel['风险-' + name].eq('√').sum())})
    risk_df = pd.DataFrame(risk_rows)

    # ---- 重点核查排序 ----
    key = sel.sort_values(['风险标记数', '绝对金额'], ascending=[False, False]).head(60).copy()
    key['核查优先级'] = ['高' if n >= 3 else ('中' if n >= 1 else '常规') for n in key['风险标记数']]
    key['核查重点建议'] = key.apply(lambda r: '、'.join(filter(None, [
        '检查整千大额资金流向与商业实质' if r['风险-整数金额'] == '√' else '',
        '核对冲销/调整依据及原凭证' if r['风险-冲销调整'] == '√' else '',
        '核实往来款性质、是否关联方及期后收回' if r['风险-借款备用金往来'] == '√' else '',
        '确认红字/负数分录的业务原因' if r['风险-负数红字凭证'] == '√' else '',
        '排查同一凭证重复入账（人工复核）' if r['风险-含重复分录'] == '√' else '',
        '关注复杂分录借贷对应关系' if r['风险-分录数异常'] == '√' else '',
        '补充摘要或核对业务内容' if r['风险-摘要缺失'] == '√' else '',
        '查明借贷不平原因' if r['风险-借贷不平'] == '√' else '',
    ])) or '按常规程序检查附件与审批', axis=1)

    actual_cov = sel['绝对金额'].sum() / q['凭证金额合计'] if q['凭证金额合计'] else 0
    est_cov = plan_df.iloc[-1]['层金额估算(元)'] / q['凭证金额合计'] if q['凭证金额合计'] else 0

    # 抽样率单调性检查（如实报告，不粉饰）
    main_layers = [L for L in order if L != 'Z' and sizes[L] > 0]
    rate_of = {L: (alloc[L] / sizes[L] if sizes[L] else 0.0) for L in order}
    inversions = []
    for i in range(len(main_layers) - 1):
        lo, hi = main_layers[i], main_layers[i + 1]
        if rate_of[hi] + 1e-12 < rate_of[lo]:
            inversions.append(f"{LAYER_LABEL[lo]}({rate_of[lo] * 100:.1f}%) > {LAYER_LABEL[hi]}({rate_of[hi] * 100:.1f}%)")
    mono_text = '抽样率随层金额单调不降' if not inversions else (
        '存在率倒挂（因"每层至少 1 张"的最小覆盖约束，仅出现在最低层）：' + '；'.join(inversions))

    # ============ 写工作簿 ============
    ordered = (['序号', '层标签', '凭证编号', '制单日期', '摘要', '科目数', '行数', '重复行数',
                '借方合计', '贷方合计', '凭证金额', '制单人', '审核人',
                '抽样方式', '抽样说明', '风险标记数', '风险标记'] + FLAG_COLS)
    out_df = sel[ordered].rename(columns={'层标签': '层'}).copy()
    out_df['制单日期'] = pd.to_datetime(out_df['制单日期']).dt.strftime('%Y-%m-%d').fillna('')

    wb = Workbook()
    ws1 = wb.active
    ws1.title = '抽样清单'
    w = {'序号': 6, '层': 16, '凭证编号': 22, '制单日期': 12, '摘要': 52, '科目数': 8, '行数': 7,
         '重复行数': 9, '借方合计': 14, '贷方合计': 14, '凭证金额': 14, '制单人': 10, '审核人': 10,
         '抽样方式': 30, '抽样说明': 32, '风险标记数': 10, '风险标记': 26}
    for c in FLAG_COLS:
        w[c] = 11
    notes1 = [
        f"抽样对象：{src.name}｜工作表：{sheet_name or 'CSV'}｜表头行：{hdr_row + 1}",
        f"抽样总体：{q['凭证张数']:,} 张凭证 / {q['读入分录行数']:,} 行分录"
        + f"｜实际期间：{q['日期最小']} ~ {q['日期最大']}",
        f"数据质量：完全重复行 {q['完全重复行数']} 行（涉及 {q['完全重复行涉及凭证数']} 张凭证，"
        f"多计借方 {q['重复行多计借方']:,.2f} 元）——已保留并标记，未做删除；借贷不平凭证 {q['不平凭证张数']} 张",
        f"抽样方法：分层判断抽样（非统计抽样），随机种子 {args.seed}，输入文件 SHA256 {sha[:16]}",
        '风险标记为辅助提示非审计结论：重复分录按"同凭证内科目编码+借贷金额+部门完全相同"判定并排除分录数≥50的汇总凭证；整数金额与往来款命中多为正常业务，需人工复核',
    ]
    hdr_row_out = write_sheet(ws1, f'序时账抽样清单（样本 {len(sel)} 张凭证）', out_df, w,
                              money_cols=['借方合计', '贷方合计', '凭证金额'],
                              wrap_cols=['摘要', '抽样方式', '抽样说明', '风险标记'],
                              notes=notes1)
    for i in range(len(out_df)):
        if sel.iloc[i]['风险标记数'] > 0:
            ws1.cell(row=hdr_row_out + 1 + i, column=ordered.index('风险标记') + 1).fill = RED
    ws1.row_dimensions[hdr_row_out].height = 30

    ws2 = wb.create_sheet('抽样方案与说明')
    write_sheet(ws2, '分层抽样方案与全量风险分布', plan_df,
                {'层': 18, '总体凭证数': 11, '样本数': 9, '抽样率': 9,
                 '层金额合计(元)': 18, '层金额估算(元)': 18, '抽样方法': 26},
                money_cols=['层金额合计(元)', '层金额估算(元)'])
    rt = 3 + 1 + len(plan_df) + 1
    ws2.cell(row=rt, column=1, value='全量风险分布（对全部凭证筛查，非仅样本）').font = Font(bold=True, size=12)
    rb = rt + 1
    for j, c in enumerate(risk_df.columns, 1):
        cell = ws2.cell(row=rb, column=j, value=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.border = BORDER
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for i, (_, row) in enumerate(risk_df.iterrows(), rb + 1):
        for j, c in enumerate(risk_df.columns, 1):
            v = row[c]
            if isinstance(v, (np.integer,)):
                v = int(v)
            elif isinstance(v, (np.floating,)):
                v = float(v)
            cell = ws2.cell(row=i, column=j, value=v)
            cell.border = BORDER
            cell.font = Font(size=10)
            cell.alignment = Alignment(wrap_text=(c == '判据'), vertical='top')
            if c == '全量涉及金额(元)':
                cell.number_format = '#,##0.00'
    for j, c in enumerate(risk_df.columns, 1):
        ws2.column_dimensions[get_column_letter(j)].width = {'风险标记': 16, '判据': 46,
                                                             '全量涉及张数': 12, '全量涉及金额(元)': 18,
                                                             '样本涉及张数': 12}.get(c, 12)
    ws2.freeze_panes = 'A4'

    ws3 = wb.create_sheet('抽样说明')
    method_rows = [
        ('抽样对象', f'{src.name}（工作表：{sheet_name or "CSV"}，表头第 {hdr_row + 1} 行）'),
        ('实际数据期间', f"{q['日期最小']} 至 {q['日期最大']}"
         + (f"（按月分布：" + '、'.join(f'{k}:{v}张' for k, v in q.get('各月凭证数', {}).items()) + '）'
            if q.get('各月凭证数') else '')),
        ('抽样总体', f"{q['凭证张数']:,} 张记账凭证 / {q['读入分录行数']:,} 行会计分录（原始数据未做任何删除）"),
        ('数据质量发现', f"完全重复行 {q['完全重复行数']} 行，涉及 {q['完全重复行涉及凭证数']} 张凭证，"
         f"重复行多计借方 {q['重复行多计借方']:,.2f} 元、多计贷方 {q['重复行多计贷方']:,.2f} 元"
         f"（判据：同凭证内 科目编码+借方金额+贷方金额+部门名称 完全相同，不含摘要）；"
         f"无法解析金额单元格 {q['无法解析金额单元格数']} 个"),
        ('总体发生额', f"借方 {q['借方合计']:,.2f} 元，贷方 {q['贷方合计']:,.2f} 元"
         + (f"；借贷不平凭证 {q['不平凭证张数']} 张，差额合计 {q['不平差额合计']:,.2f} 元"
            if q['不平凭证张数'] else '；借贷平衡')),
        ('抽样单元', '记账凭证（以"凭证编号"唯一标识，非会计分录行）'),
        ('抽样方法', '分层判断抽样（非统计抽样）：按凭证金额绝对值分层，大额层全选或定向抽取，其余层内随机'),
        ('分层依据', '凭证金额 = max(借方合计, 贷方合计) 的绝对值；阈值 1千/5千/2万/10万/50万/1000万元，零金额单列 Z 层'),
        ('分层与分配', '；'.join(
            f"{r['层']}：总体{r['总体凭证数']}张/样本{r['样本数']}张({r['抽样率']})" for _, r in plan_df[:-1].iterrows())),
        ('抽样率单调性', mono_text),
        ('随机种子', f'{args.seed}（层内随机结果可复现）'),
        ('输入文件哈希', f'SHA256 {sha}'),
        ('样本规模', f"{len(sel)} 张凭证，占总体张数 {len(sel) / q['凭证张数'] * 100:.1f}%"),
        ('覆盖率（实际）', f"样本凭证金额 {sel['绝对金额'].sum():,.2f} 元，占总体金额 {actual_cov * 100:.1f}%（已抽取凭证实际合计）"),
        ('覆盖率（估算）', f"按分层抽样估计 {plan_df.iloc[-1]['层金额估算(元)']:,.2f} 元，占总体金额 {est_cov * 100:.1f}%（仅用于方案说明，不得用于推断总体错报）"),
        ('借贷平衡检查', f"容差 {args.tol}；不平凭证 {q['不平凭证张数']} 张（不剔除，仍在抽样总体内）"),
        ('风险标记', '；'.join(f"{r['风险标记']}：全量{r['全量涉及张数']}张/样本{r['样本涉及张数']}张"
                            for _, r in risk_df.iterrows())),
        ('执行方式', '抽取样本后应核对：原始凭证与记账凭证一致、审批签字齐全、附件完整、账务处理与业务实质相符；对风险标记逐项落实'),
        ('已知误报边界', '含重复分录在 20–49 行多科目凭证上仍可能误报；整数金额与借款备用金往来命中多为集团内资金调拨、备用金借还等正常业务，属提示非结论'),
        ('局限性', '本方法为非统计判断抽样，样本结果不能用于推断总体错报金额；未读取原始凭证与外部单据，无法验证业务真实性；未执行替代程序，未评价未抽取项目'),
    ]
    ws3['A1'] = '抽样说明'
    ws3['A1'].font = TITLE_FONT
    rr = 3
    for k, v in method_rows:
        ws3.cell(row=rr, column=1, value=k).font = Font(bold=True, size=10)
        ws3.cell(row=rr, column=1).alignment = Alignment(vertical='top')
        c = ws3.cell(row=rr, column=2, value=v)
        c.alignment = Alignment(wrap_text=True, vertical='top')
        c.font = Font(size=10)
        ws3.row_dimensions[rr].height = max(28, 14 * (1 + len(str(v)) // 70))
        rr += 1
    ws3.column_dimensions['A'].width = 18
    ws3.column_dimensions['B'].width = 100

    ws4 = wb.create_sheet('重点核查排序')
    key_out = key[['序号', '核查优先级', '层标签', '凭证编号', '制单日期', '摘要', '凭证金额',
                   '风险标记数', '风险标记', '核查重点建议']].rename(columns={'层标签': '层'}).copy()
    key_out['制单日期'] = pd.to_datetime(key_out['制单日期']).dt.strftime('%Y-%m-%d').fillna('')
    r04 = write_sheet(ws4, f'重点核查排序（按风险标记数、金额降序，前 {len(key_out)} 张）', key_out,
                      {'序号': 6, '核查优先级': 10, '层': 16, '凭证编号': 22, '制单日期': 12,
                       '摘要': 50, '凭证金额': 15, '风险标记数': 10, '风险标记': 26, '核查重点建议': 46},
                      money_cols=['凭证金额'], wrap_cols=['摘要', '风险标记', '核查重点建议'],
                      notes=['优先级：风险标记数 ≥3 为"高"，1~2 为"中"，0 为"常规"'])
    for i in range(len(key_out)):
        if key_out.iloc[i]['核查优先级'] == '高':
            ws4.cell(row=r04 + 1 + i, column=2).fill = RED

    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)

    # ============ 结构化摘要 ============
    summary = {
        'input': str(src), 'input_sha256': sha, 'sheet': sheet_name, 'header_row': hdr_row + 1,
        'column_mapping': mapping, 'seed': args.seed, 'tol': args.tol, 'target': args.target,
        'quality': q,
        'layers': plan_rows,
        'sample': {
            'count': int(len(sel)),
            'share_of_vouchers_pct': round(len(sel) / q['凭证张数'] * 100, 2),
            'amount': float(sel['绝对金额'].sum()),
            'actual_coverage_pct': round(actual_cov * 100, 2),
            'estimated_coverage_pct': round(est_cov * 100, 2),
        },
        'rate_monotonic': mono_text,
        'all_risk_distribution': risk_rows,
        'output': str(out),
    }
    if args.json_out:
        Path(args.json_out).write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding='utf-8')

    log('=' * 78)
    log('序时账抽样 —— 结构化摘要（抽样说明中的数字必须取自此处）')
    log('=' * 78)
    log(f"输入文件      : {src.name}")
    log(f"选用工作表    : {sheet_name or 'CSV'}（表头第 {hdr_row + 1} 行）")
    log(f"字段映射      : {mapping}")
    log(f"输入 SHA256   : {sha[:16]}…")
    log(f"随机种子      : {args.seed}")
    log('-' * 78)
    log(f"抽样总体      : {q['凭证张数']:,} 张凭证 / {q['读入分录行数']:,} 行分录（原始数据未删除任何行）")
    log(f"完全重复行    : {q['完全重复行数']} 行，涉及 {q['完全重复行涉及凭证数']} 张凭证，"
        f"多计借方 {q['重复行多计借方']:,.2f} 元 / 多计贷方 {q['重复行多计贷方']:,.2f} 元（已保留并标记）")
    log(f"无法解析金额  : {q['无法解析金额单元格数']} 个单元格（非空但无法解析，按 0 处理归 Z 层）")
    log(f"实际数据期间  : {q['日期最小']} ~ {q['日期最大']}")
    if q.get('各月凭证数'):
        log(f"各月凭证数    : {q['各月凭证数']}")
    log(f"总体发生额    : 借 {q['借方合计']:,.2f} / 贷 {q['贷方合计']:,.2f}")
    log(f"借贷不平凭证  : {q['不平凭证张数']} 张（差额合计 {q['不平差额合计']:,.2f}）")
    log(f"凭证金额合计  : {q['凭证金额合计']:,.2f}")
    log('-' * 78)
    log(f"{'层':<18}{'总体':>8}{'样本':>8}{'抽样率':>9}{'层金额合计':>20}{'层金额估算':>20}  抽样方法")
    for r in plan_rows:
        log(f"{r['层']:<18}{r['总体凭证数']:>8}{r['样本数']:>8}{r['抽样率']:>9}"
            f"{r['层金额合计(元)']:>20,.2f}{r['层金额估算(元)']:>20,.2f}  {r['抽样方法']}")
    log(f"{'合计':<18}{q['凭证张数']:>8}{len(sel):>8}{len(sel) / q['凭证张数'] * 100:>8.1f}%"
        f"{q['凭证金额合计']:>20,.2f}{plan_df.iloc[-1]['层金额估算(元)']:>20,.2f}")
    log('-' * 78)
    log(f"样本规模      : {len(sel)} 张凭证（占总体张数 {len(sel) / q['凭证张数'] * 100:.1f}%）")
    log(f"覆盖率(实际)  : {sel['绝对金额'].sum():,.2f} 元 = {actual_cov * 100:.1f}%")
    log(f"覆盖率(估算)  : {plan_df.iloc[-1]['层金额估算(元)']:,.2f} 元 = {est_cov * 100:.1f}%")
    log(f"单调性检查    : {mono_text}")
    log('-' * 78)
    log('全量风险分布（对全部凭证筛查，非仅样本）：')
    log(f"{'风险标记':<16}{'全量张数':>9}{'全量金额':>18}{'样本张数':>9}  判据")
    for r in risk_rows:
        log(f"{r['风险标记']:<16}{r['全量涉及张数']:>9}{r['全量涉及金额(元)']:>18,.2f}"
            f"{r['样本涉及张数']:>9}  {r['判据']}")
    log('-' * 78)
    log(f"输出文件      : {out}")
    log('=' * 78)
    return 0


if __name__ == '__main__':
    sys.exit(main())
