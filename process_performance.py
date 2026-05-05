# -*- coding: utf-8 -*-
"""
绩效自动化核算脚本 v3.0
功能：读取04月份绩效.xlsx各日工作表，结合建立.xlsx规则库，统一重算得分，输出汇总表
"""
import sys, io, re, openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from collections import defaultdict

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# ============================================================
# 第一步：读取规则库
# ============================================================
def load_rules(rules_path='绩效规则库.xlsx'):
    wb = openpyxl.load_workbook(rules_path, data_only=True)
    ws = wb['绩效规则']
    rules = {}
    headers = []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=c).value
        headers.append(str(v).strip() if v else '')
    for r in range(2, ws.max_row + 1):
        code = ws.cell(row=r, column=1).value
        if code is None or str(code).strip() == '':
            continue
        code = str(code).strip()
        row_data = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            row_data[headers[c - 1]] = str(v).strip() if v else ''
        rules[code] = row_data
    return rules

# ============================================================
# 第二步：数据解析
# ============================================================
PROCESS_LIST = ['配料', '过筛', '混合', '预混', '制粒', '压片', '包衣', '内包', '外包']

PROCESS_COLUMNS = {
    '配料': ('配料基础分', '配料大清分', '配料定员（人）'),
    '过筛': ('过筛分', '', ''),
    '混合': ('混合基础分', '混合大清分', '混合定员'),
    '预混': ('预混基础分', '预混大清分', '预混定员'),
    '制粒': ('制粒基础分', '制粒大清', '制粒定员（人）'),
    '压片': ('压片基础分', '压片大清分', '压片定员（人）'),
    '包衣': ('包衣基础分', '包衣大清分', '包衣定员（人）'),
    '内包': ('内包基础分', '内包大清分', '内包定员（人）'),
    '外包': ('外包基础分', '外包大清分', '外包定员（人）'),
}

def _fuzzy_get(rule, keywords):
    """模糊匹配规则行中的列名：包含所有keywords的列即匹配"""
    if not keywords:
        return ''
    for k, v in rule.items():
        if isinstance(k, str) and all(kw in k for kw in keywords.split()):
            return v
    return ''

# 工序的关键词映射（含缩写/别名）
PROCESS_KEYWORDS = {
    '配料': ['配料'],
    '过筛': ['过筛'],
    '混合': ['混合'],
    '预混': ['预混'],
    '制粒': ['制粒'],
    '压片': ['压片', '65冲', '装机'],  # "65冲装机"是压片工序
    '包衣': ['包衣'],
    '内包': ['内包', '内'],
    '外包': ['外包', '外', '外包缺员', '外缺员'],  # "外缺员"指外包缺员
}


def identify_process(remark, reward, sub_cols):
    """识别该人员的主要工序（增强版）"""
    # 构建搜索文本
    text_parts = []
    if remark:
        text_parts.append(remark)
    if reward:
        text_parts.append(reward)
    text_parts.extend(sub_cols)
    all_text = ' '.join(text_parts)

    # 精确匹配：备注开头的工序名（最可靠）
    if remark:
        for p in PROCESS_LIST:
            if remark.startswith(p) or remark.startswith(f'{p}、') or remark.startswith(f'{p} '):
                return p
        # 检查备注中"："后的内容
        for sep in ['：', ':']:
            if sep in remark:
                after = remark.split(sep, 1)[1].strip()
                for p in PROCESS_LIST:
                    if after.startswith(p):
                        return p

    # 关键字匹配：按优先级检查
    for p in PROCESS_LIST:
        for kw in PROCESS_KEYWORDS[p]:
            if kw in all_text:
                return p

    # 特殊：子列中"外缺员" → 外包，"内缺员" → 内包
    for s in sub_cols:
        if '外缺员' in s or '外包' in s:
            return '外包'
        if '内缺员' in s or '内包' in s:
            return '内包'

    return '其他'


def parse_batch_prefixes(text):
    """从文本中提取批号前缀"""
    if not text:
        return []
    batches = re.findall(r'(\d{6,9})', text)
    results = set()
    for b in batches:
        if len(b) >= 2:
            results.add(b[:2])
        if len(b) >= 3:
            results.add(b[:3])
    return list(results)


def find_exclusive_code(remark, reward, rules, sub_cols):
    """提取专属号，优先匹配规则库"""
    all_prefixes = set()
    for txt in [remark or '', reward or '', ' '.join(sub_cols)]:
        for p in parse_batch_prefixes(txt):
            all_prefixes.add(p)

    # 先在规则库中匹配（长前缀优先）
    for prefix in sorted(all_prefixes, key=lambda x: (-len(x))):
        if prefix in rules:
            return prefix

    # 没找到匹配 → 取第一个2位前缀
    two_d = sorted([p for p in all_prefixes if len(p) == 2])
    if two_d:
        return two_d[0]

    three_d = sorted([p for p in all_prefixes if len(p) == 3])
    if three_d:
        return three_d[0]

    return ''


def parse_score_text(text):
    """解析评分文本"""
    if not text or text in ['', '—', '-']:
        return 0, 'simple'
    if '基础分=' in text or '实际生产箱数' in text:
        return text, 'formula'
    m = re.search(r'(\d+(?:\.\d+)?)\s*分/箱', text)
    if m:
        return float(m.group(1)), 'per_box'
    m = re.search(r'(\d+(?:\.\d+)?)\s*分', text)
    if m:
        return float(m.group(1)), 'simple'
    m = re.search(r'^(\d+(?:\.\d+)?)', text.strip())
    if m:
        return float(m.group(1)), 'simple'
    return 0, 'simple'


def extract_boxes(text):
    """提取箱数"""
    if not text:
        return 0
    total = 0
    for m in re.finditer(r'[（(](\d+(?:\.\d+)?)\s*箱[）)]', text):
        total += float(m.group(1))
    return total


def extract_special_items(text):
    """提取异常项"""
    if not text:
        return ''
    items = []
    for k in ['缺员', '质量分', '请假', '护理假', '休息', '返回原']:
        if k in text:
            items.append(k)
    return '、'.join(items)


def is_non_production(remark, reward):
    """判断是否非生产状态"""
    for kw in ['休息', '请假', '护理假', '领料', '返回原部门', '返回原车间', '返回原']:
        if kw in (remark or '') or kw in (reward or ''):
            return True
    # "整理记录"也是非生产
    if '整理记录' in (remark or ''):
        return True
    return False


def parse_sheet(ws, rules):
    """解析单日工作表 - 带上下文继承"""
    rows_data = []
    day_label = ws.title

    # 用于上下文继承：记录上一个有完整备注的人的工序和专属号
    last_context = {'工序': '', '专属号': '', '箱数': 0}

    for r in range(3, ws.max_row + 1):
        name = ws.cell(row=r, column=1).value
        if name is None or str(name).strip() == '':
            continue
        name = str(name).strip()
        if re.match(r'^[\s\-—=+*/\d.]+$', name):
            continue

        coeff_raw = ws.cell(row=r, column=2).value
        try:
            coeff = float(coeff_raw) if coeff_raw else 1.0
        except (ValueError, TypeError):
            coeff = 1.0

        remark = str(ws.cell(row=r, column=9).value or '').strip()
        reward = str(ws.cell(row=r, column=5).value or '').strip()

        sub_cols = []
        for c in range(10, min(ws.max_column + 1, 14)):
            v = ws.cell(row=r, column=c).value
            sub_cols.append(str(v).strip() if v else '')

        # 识别工序
        process = identify_process(remark, reward, sub_cols)

        # 如果工序是"其他"但上一次有上下文，继承
        if process == '其他' and last_context['工序']:
            # 检查奖罚说明中是否有"大清"或"缺员"等非工序信息
            # 如果有且备注为空，很可能继承上一个人的工序
            if not remark:
                process = last_context['工序']

        # 提取专属号
        exclusive_code = find_exclusive_code(remark, reward, rules, sub_cols)
        if not exclusive_code and last_context['专属号']:
            # 如果没找到专属号，尝试继承
            if not remark and not reward:
                exclusive_code = last_context['专属号']

        # 提取箱数
        box_count = extract_boxes(remark)
        if box_count == 0:
            box_count = extract_boxes(reward)
        if box_count == 0:
            box_count = last_context['箱数']

        special = extract_special_items(f"{remark} {reward}")
        non_prod = is_non_production(remark, reward)

        # 更新上下文（有备注的行才更新）
        if remark:
            last_context['工序'] = process if process != '其他' else last_context['工序']
            last_context['专属号'] = exclusive_code or last_context['专属号']
            last_context['箱数'] = box_count or last_context['箱数']

        rows_data.append({
            '日期': day_label, '姓名': name, '系数': coeff,
            '工序': process, '备注': remark, '奖罚说明': reward,
            '专属号': exclusive_code, '箱数': box_count,
            '特殊项': special, '非生产': non_prod, '子列': sub_cols,
        })

    return rows_data


# ============================================================
# 第三步：统一重算得分
# ============================================================
def evaluate_formula(formula_text, box_count):
    try:
        expr = formula_text.replace('基础分=', '').replace('实际生产箱数', str(box_count))
        return float(eval(expr, {'__builtins__': {}}, {}))
    except:
        return 0


def get_process_score(rules, exclusive_code, process, box_count=0):
    if exclusive_code not in rules or process not in PROCESS_COLUMNS:
        return 0, 0, 1, '无规则'

    base_col, daqing_col, ding_col = PROCESS_COLUMNS[process]
    rule = rules[exclusive_code]

    base_text = _fuzzy_get(rule, base_col)
    daqing_text = _fuzzy_get(rule, daqing_col) if daqing_col else ''
    ding_text = _fuzzy_get(rule, ding_col) if ding_col else ''

    base_val, base_type = parse_score_text(base_text)
    if base_type == 'formula':
        if isinstance(base_val, str):
            base_score = evaluate_formula(base_val, box_count)
        else:
            base_score = base_val
    elif base_type == 'per_box':
        base_score = base_val * box_count
    else:
        base_score = base_val

    daqing_val, _ = parse_score_text(daqing_text)
    daqing_score = daqing_val if daqing_val else 0

    # 定员
    dingyuan = 1
    if ding_text:
        m = re.search(r'单线(\d+)人', ding_text)
        if m:
            dingyuan = int(m.group(1))
        else:
            m = re.search(r'(\d+)人', ding_text)
            if m:
                dingyuan = int(m.group(1))

    return base_score, daqing_score, dingyuan, base_type


def calculate_scores(all_data, rules):
    """统一算法重算得分"""
    # 统计每天每工序的有效人数
    daily_process_count = defaultdict(lambda: defaultdict(int))
    for row in all_data:
        if not row['非生产'] and row['工序'] != '其他':
            daily_process_count[row['日期']][row['工序']] += 1

    for row in all_data:
        if row['非生产'] or row['工序'] == '其他':
            row.update({'基础分来源': '', '规则基础分': 0, '规则大清分': 0,
                        '重算基础分': 0, '重算大清分': 0, '重算得分': 0, '参与人数': 0})
            continue

        process = row['工序']
        code = row['专属号']
        box_cnt = row['箱数']
        coeff = row['系数']

        base_score, daqing_score, dingyuan, type_desc = get_process_score(
            rules, code, process, box_cnt)

        row['基础分来源'] = f'专属号{code}-{type_desc}'
        row['规则基础分'] = round(base_score, 2)
        row['规则大清分'] = daqing_score

        people_count = daily_process_count.get(row['日期'], {}).get(process, 1)
        row['参与人数'] = people_count

        recalc_base = (base_score / people_count * coeff) if people_count > 0 and base_score > 0 else 0

        has_daqing = '大清' in (row['备注'] or '') or '大清' in (row['奖罚说明'] or '')
        recalc_daqing = (daqing_score / people_count * coeff) if has_daqing and daqing_score > 0 and people_count > 0 else 0

        row.update({
            '重算基础分': round(recalc_base, 2),
            '重算大清分': round(recalc_daqing, 2),
            '重算得分': round(recalc_base + recalc_daqing, 2)
        })

    return all_data


# ============================================================
# 第四步：生成汇总表
# ============================================================
def generate_excel(all_data, output_path='4月最终核算汇总表.xlsx'):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '4月核算汇总'

    ws.merge_cells('A1:M1')
    t = ws['A1']
    t.value = '2026年4月绩效核算汇总表（统一算法重算）'
    t.font = Font(name='微软雅黑', size=14, bold=True)
    t.alignment = Alignment(horizontal='center', vertical='center')

    headers = ['日期', '姓名', '系数', '工序', '专属号', '箱数/来源',
               '规则基础分', '重算基础分', '重算大清分', '重算得分',
               '参与人数', '特殊项', '备注/奖罚说明']
    hf = Font(name='微软雅黑', size=10, bold=True, color='FFFFFF')
    hfill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    ha = Alignment(horizontal='center', vertical='center', wrap_text=True)
    bd = Border(left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin'))

    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=2, column=c, value=h)
        cell.font = hf
        cell.fill = hfill
        cell.alignment = ha
        cell.border = bd

    df = Font(name='微软雅黑', size=9)
    da = Alignment(horizontal='center', vertical='center', wrap_text=True)
    yellow = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

    all_data.sort(key=lambda x: (x['日期'], x['工序'], x['姓名']))
    ri = 3
    for row in all_data:
        bc = row.get('箱数', 0)
        source = f"{bc:.0f}箱" if bc > 0 else row.get('专属号', '')
        vals = [
            row['日期'], row['姓名'], row['系数'], row['工序'], row['专属号'],
            source, row.get('规则基础分', 0),
            row.get('重算基础分', 0), row.get('重算大清分', 0),
            row.get('重算得分', 0), row.get('参与人数', 0),
            row.get('特殊项', ''), (row.get('备注', '') or row.get('奖罚说明', ''))
        ]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=ri, column=c, value=v)
            cell.font = df
            cell.alignment = da
            cell.border = bd
            if c in [3, 7, 8, 9, 10] and isinstance(v, (int, float)):
                cell.number_format = '0.00'
        if row.get('特殊项', ''):
            ws.cell(row=ri, column=12).fill = yellow
        ri += 1

    for i, w in enumerate([8, 10, 7, 8, 8, 12, 10, 12, 12, 12, 8, 14, 30], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = 'A3'

    # ---- 个人汇总 ----
    ws2 = wb.create_sheet('个人汇总')
    for c, h in enumerate(['姓名', '出勤天数', '总得分', '平均日得分', '最高单日得分', '最低单日得分', '涉及工序'], 1):
        cell = ws2.cell(row=1, column=c, value=h)
        cell.font = hf
        cell.fill = hfill
        cell.alignment = ha
        cell.border = bd

    pd = defaultdict(lambda: {'days': set(), 'scores': [], 'procs': set()})
    for row in all_data:
        n = row['姓名']
        pd[n]['days'].add(row['日期'])
        pd[n]['scores'].append(row.get('重算得分', 0))
        pd[n]['procs'].add(row['工序'])

    ri = 2
    for name, info in sorted(pd.items(), key=lambda x: sum(x[1]['scores']), reverse=True):
        sc = info['scores']
        total = sum(sc)
        avg = total / len(sc) if sc else 0
        for c, v in enumerate([name, len(info['days']), round(total, 2), round(avg, 2),
                                round(max(sc), 2) if sc else 0,
                                round(min(sc), 2) if sc else 0,
                                '、'.join(sorted(info['procs']))], 1):
            cell = ws2.cell(row=ri, column=c, value=v)
            cell.font = df
            cell.alignment = da
            cell.border = bd
            if c in [3, 4, 5, 6] and isinstance(v, float):
                cell.number_format = '0.00'
        ri += 1

    for i, w in enumerate([10, 10, 12, 12, 12, 12, 20], 1):
        ws2.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws2.freeze_panes = 'A2'

    wb.save(output_path)
    return output_path


# ============================================================
# 主流程
# ============================================================
def main():
    print("=" * 60)
    print("绩效自动化核算脚本 v3.0")
    print("=" * 60)

    rules = load_rules('绩效规则库.xlsx')
    print(f"\n[1/5] 规则库: {len(rules)} 条产品规则")

    print("[2/5] 读取绩效数据...")
    perf_wb = openpyxl.load_workbook('04月份绩效.xlsx', data_only=True)
    all_data = []
    for sn in sorted(perf_wb.sheetnames):
        ws = perf_wb[sn]
        rows = parse_sheet(ws, rules)
        all_data.extend(rows)
    print(f"  共 {len(all_data)} 条记录 from {len(perf_wb.sheetnames)} 个工作日")

    # 专属号匹配统计
    print("[3/5] 专属号匹配分析...")
    code_stats = defaultdict(int)
    no_code = 0
    for row in all_data:
        if row['专属号']:
            code_stats[row['专属号']] += 1
        else:
            no_code += 1
    print(f"  匹配: {sum(code_stats.values())} 条, 未匹配: {no_code} 条")
    for code, cnt in sorted(code_stats.items(), key=lambda x: -x[1])[:8]:
        ok = '是' if code in rules else '否(不在规则库)'
        print(f"    {code}: {cnt}条 [规则匹配:{ok}]")

    print("[4/5] 重算得分...")
    all_data = calculate_scores(all_data, rules)
    prod = sum(1 for r in all_data if not r['非生产'] and r['工序'] != '其他')
    np = sum(1 for r in all_data if r['非生产'])
    other = sum(1 for r in all_data if not r['非生产'] and r['工序'] == '其他')
    scored = sum(1 for r in all_data if r.get('重算得分', 0) > 0)
    special = sum(1 for r in all_data if r.get('特殊项', ''))
    print(f"  生产人员: {prod} | 非生产: {np} | 未识别工序: {other}")
    print(f"  已计分: {scored} | 含异常项: {special}")

    print("[5/5] 生成汇总表...")
    output_path = '4月最终核算汇总表.xlsx'
    generate_excel(all_data, output_path)
    print(f"  -> {output_path}")
    print(f"  -> 包含: 4月核算汇总 / 个人汇总")
    print("=" * 60)

    # 预览
    print("\n预览 (前15条):")
    print(f"{'日期':<8} {'姓名':<8} {'工序':<6} {'专属号':<6} {'得分':<8} {'特殊项':<14}")
    print("-" * 55)
    for row in all_data[:15]:
        print(f"{row['日期']:<8} {row['姓名']:<8} {row['工序']:<6} "
              f"{str(row.get('专属号','')):<6} {row.get('重算得分', 0):<8.2f} "
              f"{row.get('特殊项', ''):<14}")


if __name__ == '__main__':
    main()
