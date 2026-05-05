# -*- coding: utf-8 -*-
"""
飞书流水账驱动绩效计算程序
==============================
流程：
  组长通过飞书表单填报 → 写入"每日报工流水账"
  → 本脚本从流水账读取 → perf_engine 计算 → 写入绩效明细表 + 汇总表

用法：
  python app_feishu_flow.py              # 处理所有流水账记录
  python app_feishu_flow.py --no-clear   # 不清空旧绩效明细，追加写入
  python app_feishu_flow.py --dry-run    # 预览模式，不写入飞书
"""
import os, sys, json, time, datetime
from collections import defaultdict
from perf_engine import calc_perf, find_prefix
from feishu_client import FeishuClient, TABLES, build_rules_map, build_person_coeff_map, build_person_list

# ============================================================
# perf_engine 输出字段 → 飞书API中文字段名 映射
# perf_engine 返回的字典key如下：
# 日期, 姓名, 工序, 系数, 批号, 批量, 基础分, 生产得分, 奖罚说明,
# 大清, 过筛, 装机, 缺员补, 自动缺员补, 时缺补, 质量, 粉碎, 转片,
# 模具换, 模具确, 捡药, 扣, 系数分, 加减分, 总分, 备注, 显示备注
# ============================================================
DETAIL_FIELDS = {
    "姓名": "姓名",
    "日期": "生产日期",
    "工序": "工序",
    "批号": "批号",
    "批量": "批次量(箱)",
    "基础分": "基础分",
    "大清": "大清分",
    "缺员补": "缺员补偿分",
    "时缺补": "小时缺员补偿分",
    "质量": "质量补偿分",
    "系数分": "系数得分",
    "加减分": "不乘系数得分",
    "总分": "个人总得分",
    "备注": "备注",
    "过筛": "过筛分",
    "装机": "装机分",
    "粉碎": "粉碎分",
    "转片": "转片子分",
    "模具换": "更换模具分",
    "模具确": "模具确认分",
    "捡药": "捡药分",
    "扣": "缺员扣减分",
}

# ============================================================
# 工具函数
# ============================================================
def feishu_ts_to_date_str(ts_ms):
    """飞书时间戳(毫秒) → "日.月.年" 字符串"""
    if not ts_ms:
        return ""
    try:
        ts = int(ts_ms)
        # 统一转为秒级时间戳（飞书存的是毫秒）
        dt = datetime.datetime.fromtimestamp(ts // 1000)
        return dt.strftime("%d.%m.%Y")
    except (ValueError, OSError):
        return str(ts_ms)


def date_str_to_feishu_ts(date_str):
    """"日.月.年" 字符串 → 飞书时间戳(毫秒)"""
    try:
        parts = date_str.split(".")
        if len(parts) == 3:
            day, month, year = int(parts[0]), int(parts[1]), int(parts[2])
            if year < 100:
                year += 2000
            dt = datetime.datetime(year, month, day, 0, 0, 0)
            return int(dt.timestamp() * 1000)
    except (ValueError, OSError):
        pass
    return 0


def parse_comma_list(text):
    """解析逗号分隔的列表（英逗号或中逗号）"""
    if not text:
        return []
    text = str(text).strip()
    # 替换中文逗号
    text = text.replace("，", ",")
    parts = [p.strip() for p in text.split(",") if p.strip()]
    return parts


def parse_flag(text):
    """解析布尔标志字段"""
    if not text:
        return False
    text = str(text).strip()
    return text in ("是", "有", "true", "True", "1", "✓", "✔")


# ============================================================
# 核心：从流水账读取并转换为 form_data
# ============================================================
def read_flow_records(client: FeishuClient):
    """从飞书"每日报工流水账"读取所有记录"""
    table_id = TABLES.get("每日报工流水账")
    if not table_id:
        raise ValueError("未找到表: 每日报工流水账")
    
    records = client.list_records(table_id)
    print(f"  读取流水账: {len(records)} 条记录")
    
    parsed = []
    for r in records:
        f = r.get("fields", {})
        
        # 生产日期
        date_ts = f.get("生产日期")
        date_str = feishu_ts_to_date_str(date_ts)
        if not date_str:
            continue
        
        # 工序
        process = str(f.get("工序", "")).strip()
        if not process:
            continue
        
        # 批号
        batch_no = str(f.get("批号", "")).strip()
        
        # 批次量
        boxes = f.get("批次量(箱)")
        try:
            boxes = float(boxes) if boxes else 0
        except (ValueError, TypeError):
            boxes = 0
        
        # 生产批数
        batch_count = f.get("生产批数")
        try:
            batch_count = int(batch_count) if batch_count else 1
        except (ValueError, TypeError):
            batch_count = 1
        
        # 人员姓名和系数
        person_names = parse_comma_list(f.get("人员姓名"))
        person_coeffs_str = parse_comma_list(f.get("人员系数"))
        
        # 解析系数：与姓名一一对应
        person_coeffs = []
        for i in range(len(person_names)):
            if i < len(person_coeffs_str):
                try:
                    person_coeffs.append(float(person_coeffs_str[i]))
                except (ValueError, TypeError):
                    person_coeffs.append(1.0)
            else:
                person_coeffs.append(1.0)
        
        # 标志位
        has_dq = parse_flag(f.get("是否有大清分"))
        
        # 缺员信息
        qy_names = parse_comma_list(f.get("缺员人员"))
        qy_hours_str = parse_comma_list(f.get("缺员小时数"))
        qy_persons = []
        for i, name in enumerate(qy_names):
            hours = 0.0
            if i < len(qy_hours_str):
                try:
                    hours = float(qy_hours_str[i])
                except (ValueError, TypeError):
                    hours = 0.0
            qy_persons.append({"name": name, "hours": hours})
        
        # 缺员分、小时缺员分（手动补偿分）
        qy_total = f.get("缺员分")
        try:
            qy_total = float(qy_total) if qy_total else 0.0
        except (ValueError, TypeError):
            qy_total = 0.0
        
        xqy_total = f.get("小时缺员分")
        try:
            xqy_total = float(xqy_total) if xqy_total else 0.0
        except (ValueError, TypeError):
            xqy_total = 0.0
        
        # 质量分
        quality = f.get("质量分")
        try:
            quality = float(quality) if quality else 0.0
        except (ValueError, TypeError):
            quality = 0.0
        
        # 实际工时
        hours = f.get("实际工时")
        try:
            hours = float(hours) if hours else 8.0
        except (ValueError, TypeError):
            hours = 8.0
        
        # 备注
        remark = str(f.get("备注", "")).strip()
        
        parsed.append({
            "record_id": r.get("record_id", ""),
            "date_str": date_str,
            "date_ts": date_ts,
            "process": process,
            "batch_no": batch_no,
            "boxes": boxes,
            "batch_count": batch_count,
            "person_names": person_names,
            "person_coeffs": person_coeffs,
            "has_dq": has_dq,
            "qy_total": qy_total,
            "xqy_total": xqy_total,
            "quality": quality,
            "qy_persons": qy_persons,
            "hours": hours,
            "remark": remark,
        })
    
    return parsed


def group_flow_to_form_data(flow_records, person_coeff_map=None):
    """
    将流水账记录按 (日期, 工序) 分组，转换为 calc_perf 所需的 form_data 格式
    """
    if person_coeff_map is None:
        person_coeff_map = {}
    
    # 按 (日期, 工序) 分组
    groups = defaultdict(list)
    for rec in flow_records:
        key = (rec["date_str"], rec["process"])
        groups[key].append(rec)
    
    form_data_list = []
    
    for (date_str, process), recs in sorted(groups.items()):
        # 合并所有批次
        batches = []
        seen_batches = set()
        for rec in recs:
            if rec["batch_no"] and rec["batch_no"] not in seen_batches:
                seen_batches.add(rec["batch_no"])
                batches.append({
                    "batch_no": rec["batch_no"],
                    "box": int(rec["boxes"]),
                    "ban": 0,
                })
        
        if not batches:
            batches.append({"batch_no": "", "box": 0, "ban": 0})
        
        # 合并所有人员（去重）
        persons_dict = {}
        all_qy_persons = {}
        has_dq = False
        has_gs = False  # 过筛
        has_zj = False  # 装机
        has_pian = ""   # 冲型
        zp_batches = 0
        zp_persons_set = set()
        mg_change_set = set()
        mg_confirm_set = set()
        jy_baskets = 0
        jy_persons_set = set()
        fs_score = 0.0
        total_hours = 8.0
        
        for rec in recs:
            if rec["has_dq"]:
                has_dq = True
            if rec["hours"] > total_hours:
                total_hours = rec["hours"]
            
            for i, name in enumerate(rec["person_names"]):
                if name not in persons_dict:
                    coeff = rec["person_coeffs"][i] if i < len(rec["person_coeffs"]) else 1.0
                    if name in person_coeff_map:
                        coeff = person_coeff_map[name]
                    persons_dict[name] = {
                        "coeff": coeff,
                        "quality": float(rec.get("quality", 0) or 0),
                    }
            
            for qp in rec["qy_persons"]:
                name = qp["name"]
                if name in all_qy_persons:
                    all_qy_persons[name] += qp["hours"]
                else:
                    all_qy_persons[name] = qp["hours"]
        
        persons = [
            {"name": name, "coeff": info["coeff"], "quality": info["quality"]}
            for name, info in persons_dict.items()
        ]
        
        qy_persons_list = [{"name": k, "hours": v} for k, v in all_qy_persons.items()]
        
        form_data = {
            "工序": process,
            "日期": date_str,
            "工时": total_hours,
            "batches": batches,
            "persons": persons,
            "hasDQ": has_dq,
            "hasGS": has_gs,
            "hasZJ": has_zj,
            "ceshi": False,
            "chong": has_pian,
            "qyPersons": qy_persons_list,
            "zpBatches": zp_batches,
            "zpPersons": list(zp_persons_set),
            "mgChangePersons": list(mg_change_set),
            "mgConfirmPersons": list(mg_confirm_set),
            "jyBaskets": jy_baskets,
            "jyPersons": list(jy_persons_set),
            "fsScore": fs_score,
        }
        
        form_data_list.append(form_data)
    
    return form_data_list


# ============================================================
# 写入飞书
# ============================================================
# 飞书中类型为"数字"的字段
NUMBER_FEISHU_FIELDS = {
    "批次量(箱)",
    "批次批数",
    "工时",
}

def create_performance_records(client, all_results):
    """将计算结果批量写入飞书"个人绩效明细表" """
    table_id = TABLES["个人绩效明细表"]
    records_by_date = defaultdict(list)
    
    for row in all_results:
        date_str = row.get("日期", "")
        fields = {}
        for k, v in DETAIL_FIELDS.items():
            if k in row:
                if k == "日期":
                    fields[v] = date_str_to_feishu_ts(str(row[k]))
                elif v in NUMBER_FEISHU_FIELDS:
                    # 数字字段：确保传数字类型
                    val = row[k]
                    if isinstance(val, str):
                        # 从"5箱"等字符串中提取数字
                        import re
                        m = re.search(r'[\d.]+', val)
                        val = float(m.group()) if m else 0.0
                    elif isinstance(val, (int, float)):
                        val = float(val)
                    else:
                        val = 0.0
                    fields[v] = round(val, 2)
                else:
                    val = row[k]
                    if isinstance(val, float):
                        val = round(val, 2)
                    fields[v] = val
        records_by_date[date_str].append({"fields": fields})
    
    total_count = sum(len(v) for v in records_by_date.values())
    print(f"\n开始写入绩效明细表... 共 {total_count} 条记录")
    
    for date, records in sorted(records_by_date.items()):
        print(f"  - {date}: {len(records)} 条记录")
        client.create_records(table_id, records)
        time.sleep(0.5)
    
    return total_count


def update_performance_summary(client, all_results):
    """更新绩效考核汇总表 """
    persons = build_person_list(client)
    name_to_record = {p["name"]: p for p in persons}
    
    total_by_name = defaultdict(float)
    for row in all_results:
        name = row.get("姓名", "")
        score = float(row.get("总分", 0))
        total_by_name[name] += score
    
    records_to_update = []
    for name, total_score in total_by_name.items():
        if name in name_to_record:
            record = name_to_record[name]
            fields = {
                "月累计": round(float(record.get("月累计", 0)) + total_score, 2),
                "内累计": round(total_score, 2),
            }
            records_to_update.append({
                "record_id": record["record_id"],
                "fields": fields,
            })
    
    if not records_to_update:
        print("  没有需要更新的汇总记录")
        return 0
    
    table_id = TABLES["绩效考核表"]
    print(f"\n更新绩效考核表... 共 {len(records_to_update)} 条记录")
    client.update_records(table_id, records_to_update)
    return len(records_to_update)


def clear_performance_records(client, date_set):
    """清空指定日期的绩效明细"""
    table_id = TABLES["个人绩效明细表"]
    records = client.list_records(table_id)
    
    record_ids = []
    for r in records:
        fields = r.get("fields", {})
        date_ts = fields.get("生产日期")
        if not date_ts:
            continue
        date_str = feishu_ts_to_date_str(date_ts)
        if date_str in date_set:
            record_ids.append(r.get("record_id"))
    
    if record_ids:
        print(f"清理绩效明细表... 删除 {len(record_ids)} 条旧记录")
        client.delete_records(table_id, record_ids)
    
    return len(record_ids)


# ============================================================
# 主流程
# ============================================================
def main(clear_existing=True, dry_run=False):
    print("=" * 60)
    print("飞书流水账驱动绩效计算程序")
    print("=" * 60)
    
    client = FeishuClient()
    
    # [1/5] 从飞书获取规则和人员系数
    print("\n[1/5] 从飞书获取规则和人员数据...")
    rules = build_rules_map(client)
    person_coeff_map = build_person_coeff_map(client)
    persons = build_person_list(client)
    print(f"  品种规则: {len(rules)} 条")
    print(f"  人员系数: {len(person_coeff_map)} 人")
    print(f"  汇总表人员: {len(persons)} 人")
    
    # [2/5] 读取流水账
    print("\n[2/5] 读取流水账记录...")
    flow_records = read_flow_records(client)
    if not flow_records:
        print("  ⚠️ 流水账中没有记录！请先通过飞书表单填报数据。")
        return
    
    date_processes = set()
    for rec in flow_records:
        date_processes.add((rec["date_str"], rec["process"]))
    print(f"  有效记录: {len(flow_records)} 条")
    print(f"  涉及日期-工序组合: {len(date_processes)} 个")
    
    # [3/5] 转换为 form_data 并计算
    print("\n[3/5] 转换为计算格式并执行绩效计算...")
    form_data_list = group_flow_to_form_data(flow_records, person_coeff_map)
    print(f"  生成 {len(form_data_list)} 个计算任务")
    
    all_results = []
    for fd in form_data_list:
        results = calc_perf(fd, rules)
        all_results.extend(results)
    print(f"  计算结果: {len(all_results)} 条绩效明细")
    
    date_set = set(row["日期"] for row in all_results)
    print(f"  涉及日期: {', '.join(sorted(date_set))}")
    
    if dry_run:
        print("\n" + "=" * 60)
        print("🔍 预览模式 - 计算结果（不写入飞书）")
        print("=" * 60)
        for row in all_results[:20]:
            print(f"  {row.get('日期')} | {row.get('工序')} | {row.get('姓名')} | "
                  f"基础分={row.get('基础分')} | 大清={row.get('大清')} | "
                  f"系数分={row.get('系数分')} | 加减分={row.get('加减分')} | "
                  f"总分={row.get('总分')} | 备注={row.get('显示备注')}")
        if len(all_results) > 20:
            print(f"  ... 还有 {len(all_results) - 20} 条")
        return
    
    # [4/5] 清空旧记录
    if clear_existing:
        print(f"\n[4/5] 清理飞书旧记录...")
        deleted = clear_performance_records(client, date_set)
        print(f"  清理完成: 删除 {deleted} 条记录")
    else:
        print("\n[4/5] 跳过清理（--no-clear）")
    
    # [5/5] 写入
    print("\n[5/5] 写入飞书绩效记录...")
    detail_count = create_performance_records(client, all_results)
    summary_count = update_performance_summary(client, all_results)
    
    print("\n" + "=" * 60)
    print("完成！")
    print(f"  - 写入绩效明细: {detail_count} 条")
    print(f"  - 更新绩效汇总: {summary_count} 人")
    print("=" * 60)


if __name__ == "__main__":
    clear_existing = True
    dry_run = False
    
    args = sys.argv[1:]
    for arg in args:
        if arg == "--no-clear":
            clear_existing = False
        elif arg == "--dry-run":
            dry_run = True
        elif arg in ("-h", "--help"):
            print("用法:")
            print("  python app_feishu_flow.py              # 处理所有流水账记录")
            print("  python app_feishu_flow.py --no-clear   # 不清空旧绩效明细")
            print("  python app_feishu_flow.py --dry-run    # 预览模式，不写入")
            sys.exit(0)
    
    main(clear_existing, dry_run)