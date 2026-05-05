# -*- coding: utf-8 -*-
"""
绩效自动化计算与飞书同步程序
功能：
  1. 从飞书API读取品种规则/系数/人员数据
  2. 读取04月份绩效.xlsx的生产数据
  3. 利用perf_engine重新计算绩效得分
  4. 将结果写回飞书多维表格
"""
import os, sys, json, time, datetime, openpyxl
from openpyxl.utils import get_column_letter
from collections import defaultdict
import re

from perf_engine import calc_perf, find_prefix
from feishu_client import (
    FeishuClient, TABLES, 
    build_rules_map, build_coeff_map, 
    build_person_coeff_map, build_person_list
)

# 绩效明细表字段映射
DETAIL_FIELDS = {
    "姓名": "姓名",
    "日期": "生产日期",  # 注意：perf_engine输出中是"日期"，而表字段是"生产日期"
    "工序": "工序",
    "批号": "批号",
    "批次量(箱)": "批次量(箱)",
    "基础分": "基础分",
    "大清分": "大清分",
    "缺员补偿分": "缺员补偿分",
    "小时缺员补偿分": "小时缺员补偿分",
    "质量补偿分": "质量补偿分",
    "系数得分": "系数得分",
    "不乘系数得分": "不乘系数得分",
    "个人总得分": "个人总得分",
    "备注": "备注",
    "过筛分": "过筛分",
    "装机分": "装机分",
    "粉碎分": "粉碎分",
    "转片子分": "转片子分",
    "更换模具分": "更换模具分",
    "模具确认分": "模具确认分",
    "捡药分": "捡药分",
    "工时": "工时",
    "批次批数": "批次批数",
    "缺员扣减分": "缺员扣减分",
}

def timestamp_from_date(date_str):
    """将日期字符串转为13位时间戳（飞书日期字段格式）"""
    try:
        # 先尝试直接解析工作表名称常见格式
        if '.' in date_str:
            date_parts = date_str.split('.')
            if len(date_parts) == 3:
                day, month, year = int(date_parts[0]), int(date_parts[1]), int(date_parts[2])
            else:
                # 可能是"月.日"或"日.月"格式
                if len(date_parts) == 2:
                    val1, val2 = int(date_parts[0]), int(date_parts[1])
                    # 判断格式：如果val1 >= 3（最小可行日数），认为是"月.日"(04.13 → 月=4,日=13)
                    if val1 < 3:
                        # val1可能是日，val2是月
                        day, month = val1, val2
                    elif val2 > 31:
                        # val1是月，val2不是日（可能是年份）
                        month, day = val1, val2
                    elif val1 <= 12 and val2 <= 31:
                        # 两种解析都有可能，但优先以月.日格式（val1是月，val2是日）
                        month, day = val1, val2
                    elif val1 <= 31 and val2 <= 12:
                        # 日.月格式
                        day, month = val1, val2
                    else:
                        # 都超过范围，使用当前日期
                        month, day = val1, val2
                    year = datetime.datetime.now().year
                else:
                    raise ValueError(f"不是标准日期格式: {date_str}")
        elif '-' in date_str:
            parts = date_str.split('-')
            if len(parts[0]) == 4:  # yyyy-mm-dd
                year, month, day = int(parts[0]), int(parts[1]), int(parts[2])
            else:  # dd-mm-yyyy 或 dd-mm
                if len(parts) == 3:
                    day, month, year = int(parts[0]), int(parts[1]), int(parts[2])
                elif len(parts) == 2:
                    day, month = int(parts[0]), int(parts[1])
                    year = datetime.datetime.now().year
                else:
                    raise ValueError("不是标准日期格式")
        else:
            # 尝试直接解析数字，格式为：yyyymmdd
            if len(date_str) == 8:
                year = int(date_str[0:4])
                month = int(date_str[4:6])
                day = int(date_str[6:8])
            else:
                # 工作表名可能只是简化的，如"1.4"或"02-4"或"413"(表示4月13日)
                if len(date_str) <= 3 and date_str.isdigit():
                    # 如413表示4月13日
                    if len(date_str) == 3:
                        month = int(date_str[0])
                        day = int(date_str[1:3])
                    # 如42表示4月2日
                    elif len(date_str) == 2:
                        month = int(date_str[0])
                        day = int(date_str[1])
                    else:
                        day = int(date_str)
                        month = datetime.datetime.now().month
                    year = datetime.datetime.now().year
                else:
                    print(f"警告: 无法解析的日期格式'{date_str}'，使用当前日期")
                    now = datetime.datetime.now()
                    day, month, year = now.day, now.month, now.year
        
        # 确保年份是4位数
        if year < 100:
            year += 2000
            
        dt = datetime.datetime(year, month, day, 0, 0, 0)
        return int(dt.timestamp() * 1000)
    except Exception as e:
        print(f"日期解析错误 '{date_str}': {e}")
        # 返回当前日期作为后备
        now = datetime.datetime.now()
        return int(now.timestamp() * 1000)


def parse_excel_sheets(excel_path, rules, person_coeff_map=None):
    """解析Excel工作表中的生产数据，准备计算"""
    if person_coeff_map is None:
        person_coeff_map = {}
        
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    all_form_data = []
    
    for sheet_name in sorted(wb.sheetnames):
        ws = wb[sheet_name]
        print(f"处理工作表: {sheet_name}")
        
        # 每个工作表生成一个或多个表单数据
        # 1. 按工序分组
        proc_groups = defaultdict(list)
        batch_details = {}  # 批号->(箱数,人员列表)
        date_str = sheet_name  # 工作表名为日期
        
        # 先扫描整个表，提取批号、箱数、工序信息
        for r in range(3, ws.max_row + 1):
            name = ws.cell(row=r, column=1).value
            if not name or str(name).strip() == '':
                continue
                
            # 人员系数
            coeff_raw = ws.cell(row=r, column=2).value
            try:
                coeff = float(coeff_raw) if coeff_raw is not None else 1.0
            except (ValueError, TypeError):
                coeff = 1.0
                
            remark = str(ws.cell(row=r, column=9).value or '').strip()
            reward_text = str(ws.cell(row=r, column=5).value or '').strip()
            
            # 提取工序、批号、箱数
            process = identify_process(remark, reward_text)
            if process != '其他':
                batch_no = extract_batch_no(remark) or ''
                box_count = extract_boxes(remark)
                if box_count == 0:
                    box_count = extract_boxes(reward_text)
                
                # 处理所有工序
                # 将批号加入批次字典, 记录箱数和人员
                if batch_no and process:
                    if batch_no not in batch_details:
                        batch_details[batch_no] = {'box': box_count, 'ban': 0, 'people': []}
                    
                    person_info = {
                        'name': str(name).strip(), 
                        'coeff': coeff,
                        'quality': 0.0,  # 初始质量分为0
                    }
                    batch_details[batch_no]['people'].append(person_info)
                    # 将人员加入工序分组
                    proc_groups[process].append(person_info)
        
        # 2. 为每个工序创建一个表单
        for process, persons in proc_groups.items():
            if not process or process == '其他':
                continue
                
            # 获取该工序涉及的批号
            proc_batches = []
            for batch_no, details in batch_details.items():
                # 检查是否该批次至少有一个该工序的人员
                if any(p['name'] in [x['name'] for x in persons] for p in details['people']):
                    proc_batches.append({
                        'batch_no': batch_no,
                        'box': details['box'],
                        'ban': details['ban']
                    })
            
            if not proc_batches:  # 跳过没有批号的工序
                continue
                
            # 标准工时8小时
            hours = 8.0
            
            # 创建表单数据
            form_data = {
                '工序': process,
                '日期': date_str,
                '工时': hours,
                'batches': proc_batches,
                'persons': persons,
                # 以下是可选项
                'hasDQ': False,  # 是否有大清分标志（依据奖罚说明识别）
                'hasGS': False,  # 是否过筛
                'hasZJ': False,  # 是否装机
                'ceshi': False,  # 是否试验品种
                'chong': '',  # 压片冲型
                'qyPersons': [],  # 缺员人员
                'zpBatches': 0,  # 转片批数
                'zpPersons': [],  # 转片人员
                'mgChangePersons': [],  # 模具更换人员
                'mgConfirmPersons': [],  # 模具确认人员
                'jyBaskets': 0,  # 捡药筐数
                'jyPersons': [],  # 捡药人员
                'fsScore': 0,  # 粉碎分
            }
            
            all_form_data.append(form_data)
    
    return all_form_data


# 工序识别，从process_performance.py简化
PROCESS_LIST = ['配料', '过筛', '混合', '预混', '制粒', '压片', '包衣', '内包', '外包']

PROCESS_KEYWORDS = {
    '配料': ['配料'],
    '过筛': ['过筛'],
    '混合': ['混合'],
    '预混': ['预混'],
    '制粒': ['制粒'],
    '压片': ['压片', '65冲', '装机'],
    '包衣': ['包衣'],
    '内包': ['内包', '内'],
    '外包': ['外包', '外', '外包缺员', '外缺员'],
}

def identify_process(remark, reward_text):
    """识别工序（从process_performance.py简化）"""
    if not remark and not reward_text:
        return '其他'
        
    all_text = f"{remark} {reward_text}"
    
    # 精确匹配：备注开头的工序名
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
    
    # 关键字匹配
    for p in PROCESS_LIST:
        for kw in PROCESS_KEYWORDS[p]:
            if kw in all_text:
                return p
                
    # 特殊：子列中"外缺员" → 外包，"内缺员" → 内包
    if '外缺员' in all_text or '外包' in all_text:
        return '外包'
    if '内缺员' in all_text or '内包' in all_text:
        return '内包'
        
    return '其他'


def extract_batch_no(text):
    """提取批号"""
    if not text:
        return None
    
    # 按优先级从高到低尝试匹配批号
    
    # 1. 先检查批号优先匹配工序名后跟的数字：如 "压片260426001"
    for proc in PROCESS_LIST:
        pattern = f"{proc}[\\s:：]*?(\\d{{8,9}})"
        m = re.search(pattern, text)
        if m:
            return m.group(1)
    
    # 2. 检查常见格式：类似 "工序260426001" 或 "工序 260425004（10箱）"
    m = re.search(r'(\d{8,9})', text)
    if m:
        return m.group(1)
    
    # 3. 匹配6-9位数字（若没有8-9位）
    m = re.search(r'(\d{6,9})', text)
    if m:
        return m.group(1)
    
    # 4. 最后尝试匹配任何看起来像批号的内容：2位专属号+其他数字
    for code in ['25', '26', '04', '11', '51', '06']:  # 常见专属号
        pattern = f"{code}\\d{{4,7}}"
        m = re.search(pattern, text)
        if m:
            return m.group(0)
    
    return None
    
    
def extract_boxes(text):
    """提取箱数"""
    if not text:
        return 0
    total = 0
    for m in re.finditer(r'[（(](\d+(?:\.\d+)?)\s*箱[）)]', text):
        total += float(m.group(1))
    return total


def create_performance_records(client, all_results, date_range=None):
    """转换计算结果为绩效明细表记录，分日期处理"""
    # 按日期分组
    records_by_date = defaultdict(list)
    
    for row in all_results:
        date_str = row['日期']
        
        # 如果有日期范围，检查该日期是否在范围内
        if date_range and date_str not in date_range:
            continue
            
        # 创建记录
        fields = {}
        for k, v in DETAIL_FIELDS.items():
            if k in row:
                # 日期需要特殊处理为时间戳
                if k == '日期':
                    fields[v] = timestamp_from_date(row[k])
                else:
                    fields[v] = row[k]
                    
        # 添加记录
        records_by_date[date_str].append({"fields": fields})
    
    # 批量写入每个日期的记录
    table_id = TABLES['个人绩效明细表']
    print(f"\n开始写入绩效明细表... 共 {sum(len(v) for v in records_by_date.values())} 条记录")
    
    for date, records in sorted(records_by_date.items()):
        print(f"  - {date}: {len(records)} 条记录")
        client.create_records(table_id, records)
        time.sleep(1)  # 避免API限流
        
    return sum(len(v) for v in records_by_date.values())


def update_performance_summary(client, all_results):
    """更新绩效考核表（月累计、内累计等）"""
    # 1. 获取现有人员记录
    persons = build_person_list(client)
    name_to_record = {p['name']: p for p in persons}
    
    # 2. 按人员汇总得分
    total_by_name = defaultdict(float)
    for row in all_results:
        name = row['姓名']
        score = float(row.get('个人总得分', 0))
        total_by_name[name] += score
        
    # 3. 准备更新记录
    records_to_update = []
    for name, total_score in total_by_name.items():
        if name in name_to_record:
            # 找到现有记录
            record = name_to_record[name]
            record_id = record['record_id']
            
            # 更新月累计（保留现有值）
            fields = {
                "月累计": round(float(record.get('月累计', 0)) + total_score, 2),
                "内累计": round(total_score, 2)
            }
            
            records_to_update.append({
                "record_id": record_id,
                "fields": fields
            })
    
    if not records_to_update:
        print("没有需要更新的记录")
        return 0
        
    # 批量更新
    table_id = TABLES['绩效考核表']
    print(f"\n更新绩效考核表... 共 {len(records_to_update)} 条记录")
    if records_to_update:  # 确保有数据才调用API
        client.update_records(table_id, records_to_update)
    return len(records_to_update)


def clear_performance_records(client, date_range=None):
    """清空指定日期范围内的绩效明细记录"""
    table_id = TABLES['个人绩效明细表']
    records = client.list_records(table_id)
    
    # 找出匹配日期范围的记录
    record_ids = []
    for r in records:
        fields = r.get('fields', {})
        date_ts = fields.get('生产日期')
        
        # 如果没有日期字段，跳过
        if not date_ts:
            continue
            
        # 转换时间戳为日期字符串
        dt = datetime.datetime.fromtimestamp(date_ts/1000)
        date_str = dt.strftime('%d.%m.%Y')
        
        # 检查是否在日期范围内
        if date_range is None or date_str in date_range:
            record_ids.append(r.get('record_id'))
    
    # 删除记录
    if record_ids:
        print(f"清理绩效明细表... 删除 {len(record_ids)} 条记录")
        client.delete_records(table_id, record_ids)
    else:
        print("没有需要清理的记录")
    
    return len(record_ids)


def main(excel_path='04月份绩效.xlsx', clear_existing=True):
    print("=" * 60)
    print("绩效自动化计算与飞书同步程序")
    print("=" * 60)
    
    # 检查 Excel 文件是否存在
    if not os.path.exists(excel_path):
        print(f"错误: Excel文件 '{excel_path}' 不存在!")
        return
    
    # 初始化客户端
    client = FeishuClient()
    
    # 第1步：从飞书获取规则和人员数据
    print("\n[1/5] 从飞书获取规则和人员数据...")
    rules = build_rules_map(client)
    coeff_map = build_coeff_map(client)
    person_coeff_map = build_person_coeff_map(client)
    persons = build_person_list(client)
    
    print(f"  规则: {len(rules)} 条")
    print(f"  系数: {len(coeff_map)} 条")
    print(f"  人员: {len(persons)} 人")
    
    # 第2步：解析Excel工作表
    print("\n[2/5] 解析Excel工作表...")
    form_data_list = parse_excel_sheets(excel_path, rules, person_coeff_map)
    print(f"  共解析 {len(form_data_list)} 个工序表单")
    
    # 第3步：使用perf_engine计算绩效
    print("\n[3/5] 计算绩效得分...")
    all_results = []
    for form_data in form_data_list:
        results = calc_perf(form_data, rules)
        all_results.extend(results)
    print(f"  共计算 {len(all_results)} 条得分记录")
    
    # 获取处理的唯一日期
    date_set = set(row['日期'] for row in all_results)
    print(f"  涉及日期: {', '.join(sorted(date_set))}")
    
    # 第4步：清空飞书现有记录
    if clear_existing:
        print("\n[4/5] 清理飞书现有记录...")
        deleted = clear_performance_records(client, date_set)
        print(f"  清理完成: 删除 {deleted} 条记录")
    else:
        print("\n[4/5] 跳过清理飞书现有记录")
    
    # 第5步：写入新记录
    print("\n[5/5] 写入飞书绩效记录...")
    # 写入明细表
    detail_count = create_performance_records(client, all_results, date_set)
    # 更新汇总表
    summary_count = update_performance_summary(client, all_results)
    
    print("\n完成!")
    print(f"- 写入绩效明细: {detail_count} 条")
    print(f"- 更新绩效汇总: {summary_count} 人")
    print("=" * 60)


if __name__ == "__main__":
    # 处理命令行参数
    excel_path = '04月份绩效.xlsx'
    clear_existing = True
    
    # 解析参数
    args = sys.argv[1:]
    for i, arg in enumerate(args):
        if arg.startswith('--'):
            # 处理选项参数
            if arg == '--no-clear':
                clear_existing = False
        else:
            # 第一个非选项参数视为Excel路径
            excel_path = arg
    
    main(excel_path, clear_existing)
