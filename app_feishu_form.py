# -*- coding: utf-8 -*-
"""
飞书报工 Web 表单
=================
组长手机浏览器打开 → 填报工单 → 一键写入飞书「每日报工流水账」
然后自动计算绩效 → 写入「个人绩效明细表」→ 更新「绩效考核表」汇总

部署到 Railway / Render，获得公网 URL，任何网络都能访问。
"""
import os
import datetime
import time
import re
import traceback
from collections import defaultdict
from flask import Flask, request, render_template_string, jsonify, redirect, url_for
from feishu_client import FeishuClient, TABLES, build_rules_map, build_person_coeff_map, build_person_list
from perf_engine import calc_perf

app = Flask(__name__)

# ============================================================
# 表单配置
# ============================================================
PROCESS_LIST = ['配料', '混合', '预混', '制粒', '压片', '包衣', '内包', '外包']
MODE_OPTIONS = [
    ('正常生产', '正常生产'),
    ('只报大清', '只报大清'),
]
CHONG_OPTIONS = ['', '65冲', '55冲', '47冲']

FORM_TEMPLATE = r'''
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
<title>飞书报工</title>
<style>
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background: #f0f2f5; padding: 12px; }
.container { max-width: 500px; margin: 0 auto; }
h1 { text-align: center; color: #1a1a2e; font-size: 20px; margin-bottom: 4px; }
.subtitle { text-align: center; color: #888; font-size: 12px; margin-bottom: 16px; }
.card { background: #fff; border-radius: 12px; padding: 16px; margin-bottom: 12px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }
.card-title { font-size: 15px; font-weight: 700; color: #333; margin-bottom: 12px; border-left: 3px solid #1890ff; padding-left: 8px; }
.row { display: flex; gap: 10px; margin-bottom: 10px; flex-wrap: wrap; }
.col { flex: 1; min-width: 120px; }
label { display: block; font-size: 13px; color: #555; margin-bottom: 4px; font-weight: 500; }
select, input[type="text"], input[type="number"], input[type="date"] {
    width: 100%; padding: 8px 10px; border: 1px solid #d9d9d9; border-radius: 6px;
    font-size: 14px; background: #fff; -webkit-appearance: none;
}
select:focus, input:focus { border-color: #1890ff; outline: none; box-shadow: 0 0 0 2px rgba(24,144,255,0.15); }
.person-row { display: flex; gap: 8px; margin-bottom: 8px; align-items: center; }
.person-row .seq { width: 24px; text-align: center; font-weight: 600; color: #1890ff; font-size: 14px; }
.person-row input { flex: 1; }
.person-row .coeff { max-width: 80px; }
.btn-row { display: flex; gap: 10px; justify-content: center; margin-top: 6px; }
.btn { padding: 6px 14px; border-radius: 6px; border: none; font-size: 13px; cursor: pointer; font-weight: 500; }
.btn-add { background: #e6f7ff; color: #1890ff; border: 1px solid #91d5ff; }
.btn-remove { background: #fff1f0; color: #ff4d4f; border: 1px solid #ffa39e; }
.submit-btn { width: 100%; padding: 14px; background: linear-gradient(135deg, #1890ff, #096dd9); color: #fff;
    border: none; border-radius: 10px; font-size: 17px; font-weight: 700; cursor: pointer; margin-top: 8px; }
.submit-btn:active { opacity: 0.85; }
.msg { padding: 10px 14px; border-radius: 8px; margin-bottom: 12px; font-size: 14px; text-align: center; }
.msg-success { background: #f6ffed; color: #52c41a; border: 1px solid #b7eb8f; }
.msg-error { background: #fff1f0; color: #ff4d4f; border: 1px solid #ffa39e; }
.check-row { display: flex; gap: 16px; flex-wrap: wrap; margin-bottom: 8px; }
.check-row label { display: flex; align-items: center; gap: 4px; font-size: 13px; color: #555; cursor: pointer; }
.check-row input[type="checkbox"] { width: 16px; height: 16px; }
.section-hidden { display: none; }
.section-visible { display: block; }
.info-text { font-size: 11px; color: #999; margin-top: 2px; }
</style>
</head>
<body>
<div class="container">
<h1>📋 飞书报工</h1>
<p class="subtitle">填报工单 → 自动写入飞书流水账 → 自动计算绩效</p>

{% if message %}
<div class="msg msg-{{ msg_type }}">{{ message }}</div>
{% endif %}

<form method="POST" id="mainForm">
<!-- 基础信息 -->
<div class="card">
<div class="card-title">📅 基础信息</div>
<div class="row">
<div class="col"><label>生产日期</label><input type="date" name="date" value="{{ today }}" required></div>
<div class="col"><label>工序</label>
<select name="process" id="process" onchange="toggleFields()" required>
<option value="">请选择</option>
{% for p in processes %}<option value="{{ p }}" {% if form_data.get('process') == p %}selected{% endif %}>{{ p }}</option>{% endfor %}
</select></div>
</div>
<div class="row">
<div class="col"><label>报工模式</label>
<select name="mode" required>
{% for v, label in modes %}<option value="{{ v }}" {% if form_data.get('mode') == v %}selected{% endif %}>{{ label }}</option>{% endfor %}
</select></div>
<div class="col"><label>实际工时 (h)</label><input type="number" name="hours" value="{{ form_data.get('hours', '8') }}" step="0.5" min="0"></div>
</div>
</div>

<!-- 批次信息 -->
<div class="card">
<div class="card-title">📦 批次信息</div>
<div class="row">
<div class="col"><label>批号</label><input type="text" name="batch_no" value="{{ form_data.get('batch_no', '') }}" placeholder="如 260426001"></div>
<div class="col"><label>批次量(箱)</label><input type="number" name="boxes" value="{{ form_data.get('boxes', '') }}" step="0.1" min="0"></div>
</div>
<div class="row">
<div class="col"><label>生产批数</label><input type="number" name="batch_count" value="{{ form_data.get('batch_count', '1') }}" min="1"></div>
<div class="col"><label>批次量(板数)</label><input type="number" name="ban_count" value="{{ form_data.get('ban_count', '0') }}" step="0.1" min="0"></div>
</div>
</div>

<!-- 人员信息 -->
<div class="card">
<div class="card-title">👥 人员信息 <span style="font-weight:400;font-size:12px;color:#999">(系数默认1.0)</span></div>
<div id="personContainer">
{% set pcount = form_data.get('person_count', '2')|int %}
{% if pcount < 1 %}{% set pcount = 2 %}{% endif %}
{% for i in range(pcount) %}
<div class="person-row" id="personRow{{ i }}">
<span class="seq">{{ i + 1 }}</span>
<input type="text" name="person_name_{{ i }}" placeholder="姓名" value="{{ form_data.get('person_name_' ~ i, '') }}">
<input type="number" class="coeff" name="person_coeff_{{ i }}" placeholder="系数" value="{{ form_data.get('person_coeff_' ~ i, '1.0') }}" step="0.01" min="0.5" max="2.0">
{% if i >= 2 %}<button type="button" class="btn btn-remove" onclick="removePerson({{ i }})">✕</button>{% endif %}
</div>
{% endfor %}
</div>
<input type="hidden" name="person_count" id="personCount" value="{{ form_data.get('person_count', '2') }}">
<div class="btn-row">
<button type="button" class="btn btn-add" onclick="addPerson()">+ 添加人员</button>
</div>
</div>

<!-- 加减分 -->
<div class="card">
<div class="card-title">⚙️ 加减分选项</div>

<!-- 大清 / 过筛 / 装机 -->
<div class="check-row">
<label><input type="checkbox" name="has_dq" value="1" {% if form_data.get('has_dq') == '1' %}checked{% endif %} onchange="toggleDQFields()" id="cbDQ"> 大清</label>
<label><input type="checkbox" name="has_gs" value="1" {% if form_data.get('has_gs') == '1' %}checked{% endif %}> 过筛</label>
<label><input type="checkbox" name="has_zj" value="1" {% if form_data.get('has_zj') == '1' %}checked{% endif %}> 干法制粒装机</label>
</div>

<!-- 缺员信息 -->
<div class="row">
<div class="col"><label>缺员人员</label><input type="text" name="qy_names" value="{{ form_data.get('qy_names', '') }}" placeholder="多人逗号分隔"></div>
<div class="col"><label>缺员小时数</label><input type="text" name="qy_hours" value="{{ form_data.get('qy_hours', '') }}" placeholder="逗号分隔，如 8,4"></div>
</div>

<!-- 缺员分 / 小时缺员分 / 质量分 -->
<div class="row">
<div class="col"><label>缺员分</label><input type="number" name="qy_total" value="{{ form_data.get('qy_total', '0') }}" step="0.1"></div>
<div class="col"><label>小时缺员分</label><input type="number" name="xqy_total" value="{{ form_data.get('xqy_total', '0') }}" step="0.1"></div>
<div class="col"><label>质量分</label><input type="number" name="quality" value="{{ form_data.get('quality', '0') }}" step="0.1"></div>
</div>

<!-- 压片冲型 -->
<div class="row" id="chongRow" style="display:{% if form_data.get('process') == '压片' %}flex{% else %}none{% endif %}">
<div class="col"><label>压片冲型</label>
<select name="chong">
{% for c in chong_opts %}<option value="{{ c }}" {% if form_data.get('chong') == c %}selected{% endif %}>{{ c or '不选' }}</option>{% endfor %}
</select></div>
</div>

<!-- 粉碎分 -->
<div class="row">
<div class="col"><label>粉碎分</label><input type="number" name="fs_score" value="{{ form_data.get('fs_score', '0') }}" step="0.1"></div>
</div>

<!-- 转片子 -->
<div class="row">
<div class="col"><label>转片子批数</label><input type="number" name="zp_batches" value="{{ form_data.get('zp_batches', '0') }}" min="0"></div>
<div class="col"><label>转片子人员</label><input type="text" name="zp_persons" value="{{ form_data.get('zp_persons', '') }}" placeholder="逗号分隔"></div>
</div>

<!-- 模具 -->
<div class="row">
<div class="col"><label>更换模具人员</label><input type="text" name="mg_change" value="{{ form_data.get('mg_change', '') }}" placeholder="逗号分隔"></div>
<div class="col"><label>模具确认人员</label><input type="text" name="mg_confirm" value="{{ form_data.get('mg_confirm', '') }}" placeholder="逗号分隔"></div>
</div>

<!-- 捡药 -->
<div class="row">
<div class="col"><label>捡药筐数</label><input type="number" name="jy_baskets" value="{{ form_data.get('jy_baskets', '0') }}" min="0"></div>
<div class="col"><label>捡药人员</label><input type="text" name="jy_persons" value="{{ form_data.get('jy_persons', '') }}" placeholder="逗号分隔"></div>
</div>
</div>

<!-- 备注 -->
<div class="card">
<div class="card-title">📝 备注</div>
<textarea name="remark" style="width:100%;height:60px;border:1px solid #d9d9d9;border-radius:6px;padding:8px;font-size:14px;" placeholder="填写备注信息...">{{ form_data.get('remark', '') }}</textarea>
</div>

<button type="submit" class="submit-btn">🚀 提交报工并自动计算绩效</button>
</form>
</div>

<script>
function addPerson() {
    var container = document.getElementById('personContainer');
    var count = parseInt(document.getElementById('personCount').value);
    var div = document.createElement('div');
    div.className = 'person-row';
    div.id = 'personRow' + count;
    div.innerHTML = '<span class="seq">' + (count + 1) + '</span>' +
        '<input type="text" name="person_name_' + count + '" placeholder="姓名">' +
        '<input type="number" class="coeff" name="person_coeff_' + count + '" placeholder="系数" value="1.0" step="0.01" min="0.5" max="2.0">' +
        (count >= 2 ? '<button type="button" class="btn btn-remove" onclick="removePerson(' + count + ')">✕</button>' : '');
    container.appendChild(div);
    document.getElementById('personCount').value = count + 1;
}

function removePerson(idx) {
    document.getElementById('personRow' + idx).remove();
    // compact remaining
    var container = document.getElementById('personContainer');
    var rows = container.querySelectorAll('.person-row');
    for (var i = 0; i < rows.length; i++) {
        rows[i].id = 'personRow' + i;
        rows[i].querySelector('.seq').textContent = i + 1;
        var inputs = rows[i].querySelectorAll('input');
        inputs[0].name = 'person_name_' + i;
        inputs[1].name = 'person_coeff_' + i;
        var btn = rows[i].querySelector('button');
        if (btn) btn.setAttribute('onclick', 'removePerson(' + i + ')');
        if (i < 2 && btn) btn.remove();
    }
    document.getElementById('personCount').value = rows.length;
}

function toggleFields() {
    var process = document.getElementById('process').value;
    document.getElementById('chongRow').style.display = (process === '压片') ? 'flex' : 'none';
}
</script>
</body>
</html>
'''


def parse_persons_from_form(form_data):
    """解析表单中的人员姓名和系数"""
    names = []
    coeffs = []
    person_count = int(form_data.get("person_count", 0) or 0)
    for i in range(max(person_count, 10)):
        name = form_data.get(f"person_name_{i}", "").strip()
        if not name:
            break
        coeff_str = form_data.get(f"person_coeff_{i}", "1.0").strip()
        try:
            coeff = float(coeff_str)
        except (ValueError, TypeError):
            coeff = 1.0
        names.append(name)
        coeffs.append(coeff)
    return names, coeffs


def write_to_feishu(form_data):
    """将表单数据写入飞书「每日报工流水账」"""
    client = FeishuClient()
    table_id = TABLES.get("每日报工流水账")
    if not table_id:
        raise ValueError("未找到表: 每日报工流水账")

    person_names, person_coeffs = parse_persons_from_form(form_data)

    # 生产日期：转成飞书时间戳
    date_str = form_data.get("date", "")
    try:
        dt = datetime.datetime.strptime(date_str, "%Y-%m-%d")
        date_ts = int(dt.timestamp() * 1000)
    except (ValueError, TypeError):
        date_ts = int(time.time() * 1000)

    # 数字字段
    def _float(key, default=0.0):
        try:
            val = form_data.get(key, "")
            if val is None or val == "":
                return default
            return float(val)
        except (ValueError, TypeError):
            return default

    def _int(key, default=0):
        try:
            val = form_data.get(key, "")
            if val is None or val == "":
                return default
            return int(float(val))
        except (ValueError, TypeError):
            return default

    mapped_fields = {
        "生产日期": date_ts,
        "工序": form_data.get("process", "").strip(),
        "报工模式": form_data.get("mode", "正常生产").strip(),
        "批号": form_data.get("batch_no", "").strip(),
        "批次量(箱)": _float("boxes"),
        "批次量(板数)": _float("ban_count"),
        "生产批数": _int("batch_count", 1),
        "实际工时": _float("hours", 8),
        "人员姓名": ",".join(person_names),
        "人员系数": ",".join(str(c) for c in person_coeffs),
        "是否有大清分": "是" if form_data.get("has_dq") == "1" else "否",
        "是否有过筛分": "是" if form_data.get("has_gs") == "1" else "否",
        "是否有干法制粒装机": "是" if form_data.get("has_zj") == "1" else "否",
        "缺员分": _float("qy_total"),
        "小时缺员分": _float("xqy_total"),
        "质量分": _float("quality"),
        "粉碎分": _float("fs_score"),
        "转片子批数": _int("zp_batches"),
        "捡药筐数": _int("jy_baskets"),
        "备注": form_data.get("remark", "").strip(),
    }

    # 文本字段（只写入非空的）
    text_map = [
        ("zp_persons", "转片子人员"),
        ("mg_change", "更换模具人员"),
        ("mg_confirm", "模具确认人员"),
        ("jy_persons", "捡药人员"),
        ("qy_names", "缺员人员"),
        ("qy_hours", "缺员小时数"),
        ("chong", "压片冲型"),
    ]
    for text_key, feishu_key in text_map:
        val = form_data.get(text_key, "").strip()
        if val:
            mapped_fields[feishu_key] = val

    # 移除值为 None 或空字符串的字段
    mapped_fields = {k: v for k, v in mapped_fields.items()
                     if v is not None and v != ""}

    # 写入飞书流水账
    result = client.create_records(table_id, [{"fields": mapped_fields}])
    return result


# ============================================================
# 自动绩效考核计算
# ============================================================

def convert_form_to_perf_data(form_data):
    """
    将表单数据转换为 perf_engine.calc_perf 需要的格式。
    参考 app_feishu_flow.py 中的 read_flow_records 和 convert_flow_to_task 逻辑。
    """
    person_names, person_coeffs = parse_persons_from_form(form_data)

    # 日期格式转换: 2026-05-05 → 05.05.2026
    date_str = form_data.get("date", "")
    try:
        dt = datetime.datetime.strptime(date_str, "%Y-%m-%d")
        perf_date = dt.strftime("%d.%m.%Y")
    except (ValueError, TypeError):
        perf_date = datetime.date.today().strftime("%d.%m.%Y")

    process = form_data.get("process", "").strip()
    hours = float(form_data.get("hours", 8) or 8)
    batch_no = form_data.get("batch_no", "").strip()
    boxes = float(form_data.get("boxes", 0) or 0)
    ban_count = float(form_data.get("ban_count", 0) or 0)
    batch_count = int(float(form_data.get("batch_count", 1) or 1))

    # 构建 batches
    batches = []
    if batch_no:
        main_batch = {"batch_no": batch_no, "box": boxes, "ban": ban_count}
        if batch_count > 1:
            # 相同批号的主批次 + 重复的副批次（第二个起 box=0 作为计数标记）
            # 注：perf_engine 中 box 决定批量，后续批次用 box=0 来代表同一批号的额外批次
            batches = [main_batch] + [
                {"batch_no": batch_no, "box": 0, "ban": 0} for _ in range(batch_count - 1)
            ]
        else:
            batches = [main_batch]
    else:
        batches = [{"batch_no": "", "box": boxes, "ban": ban_count}]

    # 构建 persons
    persons = []
    for i, name in enumerate(person_names):
        coeff = person_coeffs[i] if i < len(person_coeffs) else 1.0
        persons.append({"name": name, "coeff": coeff})

    # 标志位
    has_dq = form_data.get("has_dq") == "1"
    has_gs = form_data.get("has_gs") == "1"
    has_zj = form_data.get("has_zj") == "1"

    # 缺员信息
    qy_names_str = form_data.get("qy_names", "").strip()
    qy_hours_str = form_data.get("qy_hours", "").strip()
    qy_names = [n.strip() for n in qy_names_str.replace("，", ",").split(",") if n.strip()]
    qy_hours_vals = [h.strip() for h in qy_hours_str.replace("，", ",").split(",") if h.strip()]
    qy_persons = []
    for j, qn in enumerate(qy_names):
        hh = 0.0
        if j < len(qy_hours_vals):
            try:
                hh = float(qy_hours_vals[j])
            except (ValueError, TypeError):
                hh = 0.0
        qy_persons.append({"name": qn, "hours": hh})

    # 转片子
    zp_batches = int(float(form_data.get("zp_batches", 0) or 0))
    zp_persons_str = form_data.get("zp_persons", "").strip()
    zp_persons = [n.strip() for n in zp_persons_str.replace("，", ",").split(",") if n.strip()]

    # 模具
    mg_change_str = form_data.get("mg_change", "").strip()
    mg_change_persons = [n.strip() for n in mg_change_str.replace("，", ",").split(",") if n.strip()]
    mg_confirm_str = form_data.get("mg_confirm", "").strip()
    mg_confirm_persons = [n.strip() for n in mg_confirm_str.replace("，", ",").split(",") if n.strip()]

    # 捡药
    jy_baskets = int(float(form_data.get("jy_baskets", 0) or 0))
    jy_persons_str = form_data.get("jy_persons", "").strip()
    jy_persons = [n.strip() for n in jy_persons_str.replace("，", ",").split(",") if n.strip()]

    # 粉碎分
    fs_score = float(form_data.get("fs_score", 0) or 0)

    # 压片冲型
    chong = form_data.get("chong", "").strip()

    return {
        "日期": perf_date,
        "工序": process,
        "工时": hours,
        "batches": batches,
        "persons": persons,
        "hasDQ": has_dq,
        "hasGS": has_gs,
        "hasZJ": has_zj,
        "ceshi": False,
        "chong": chong,
        "qyPersons": qy_persons,
        "zpBatches": zp_batches,
        "zpPersons": zp_persons,
        "mgChangePersons": mg_change_persons,
        "mgConfirmPersons": mg_confirm_persons,
        "jyBaskets": jy_baskets,
        "jyPersons": jy_persons,
        "fsScore": fs_score,
    }


def auto_calc_performance(form_data, flow_record_id=None):
    """
    表单提交后自动计算绩效并写入飞书。
    返回: (detail_count, summary_count, error_msg)
    """
    try:
        client = FeishuClient()

        # 1. 获取规则
        rules = build_rules_map(client)

        # 2. 转换为 perf engine 格式
        perf_data = convert_form_to_perf_data(form_data)

        # 3. 执行计算
        results = calc_perf(perf_data, rules)
        if not results:
            return 0, 0, None

        # 4. 获取绩效考核表人员列表（用于更新汇总）
        person_list = build_person_list(client)
        person_coeff_map = build_person_coeff_map(client)

        # 5. 写入个人绩效明细表
        detail_table_id = TABLES.get("个人绩效明细表")
        if not detail_table_id:
            return 0, 0, "未找到个人绩效明细表"

        # 日期转换：05.05.2026 → 飞书时间戳
        # perf_date 已经是 DD.MM.YYYY 格式
        perf_date_str = perf_data.get("日期", "")
        try:
            dt = datetime.datetime.strptime(perf_date_str, "%d.%m.%Y")
            feishu_date_ts = int(dt.timestamp() * 1000)
        except (ValueError, TypeError):
            # Fallback: use today
            today = datetime.date.today()
            dt = datetime.datetime.combine(today, datetime.time.min)
            feishu_date_ts = int(dt.timestamp() * 1000)

        detail_records = []
        for r in results:
            coeff_score = r.get("系数分", 0.0) or 0.0
            jiajian_score = r.get("加减分", 0.0) or 0.0
            total_score = r.get("总分", 0.0) or 0.0
            base_score = r.get("基础分", 0.0) or 0.0
            person_coeff = r.get("系数", 1.0) or 1.0
            display_remark = r.get("显示备注", "") or ""

            # 解析批量数字：从 "10箱" 提取 10
            batch_str = r.get("批量", "") or ""
            batch_num = 0
            try:
                batch_num = float(batch_str.replace("箱", "").strip())
            except (ValueError, TypeError):
                batch_num = 0

            detail_fields = {
                "姓名": r.get("姓名", ""),
                "生产日期": feishu_date_ts,
                "工序": r.get("工序", ""),
                "批号": r.get("批号", ""),
                "批次量(箱)": batch_num,
                "基础分": base_score,
                "大清分": r.get("大清", 0.0) or 0.0,
                "过筛分": r.get("过筛", 0.0) or 0.0,
                "装机分": r.get("装机", 0.0) or 0.0,
                "缺员补偿分": (r.get("缺员补", 0.0) or 0.0) + (r.get("自动缺员补", 0.0) or 0.0),
                "小时缺员补偿分": r.get("时缺补", 0.0) or 0.0,
                "质量补偿分": r.get("质量", 0.0) or 0.0,
                "缺员扣减分": r.get("扣", 0.0) or 0.0,
                "粉碎分": r.get("粉碎", 0.0) or 0.0,
                "转片子分": r.get("转片", 0) or 0,
                "更换模具分": r.get("模具换", 0) or 0,
                "模具确认分": r.get("模具确", 0) or 0,
                "捡药分": r.get("捡药", 0) or 0,
                "系数得分": coeff_score,
                "不乘系数得分": jiajian_score,
                "个人总得分": total_score,
                "备注": display_remark,
                "工时": perf_data.get("工时", 8),
                "批次批数": len(perf_data.get("batches", [])),
            }
            # 关联流水账记录
            if flow_record_id:
                detail_fields["关联流水账"] = [flow_record_id]
            detail_records.append({"fields": detail_fields})

        # 批量写入
        created_count = 0
        if detail_records:
            client.create_records(detail_table_id, detail_records)
            created_count = len(detail_records)

        # 6. 更新绩效考核汇总表
        summary_table_id = TABLES.get("绩效考核表")
        updated_count = 0
        if summary_table_id and person_list:
            # 按姓名汇总本月绩效
            person_total = defaultdict(lambda: {"生产得分": 0.0, "加减分": 0.0, "备注": []})
            for r in results:
                name = r.get("姓名", "")
                coeff_score = r.get("系数分", 0.0) or 0.0
                jiajian_score = r.get("加减分", 0.0) or 0.0
                display_remark = r.get("显示备注", "") or ""
                person_total[name]["生产得分"] += coeff_score
                person_total[name]["加减分"] += jiajian_score
                if display_remark:
                    person_total[name]["备注"].append(display_remark)

            # 构建需要更新的记录
            update_records = []
            for p in person_list:
                name = p.get("name", "")
                if name in person_total:
                    record_id = p.get("record_id", "")
                    old_month_total = float(p.get("月累计", 0) or 0)
                    new_prod = person_total[name]["生产得分"]
                    new_jiajian = person_total[name]["加减分"]
                    new_month = old_month_total + new_prod + new_jiajian
                    new_remarks = "; ".join(person_total[name]["备注"])

                    # 读取旧备注，追加新备注
                    old_remark = p.get("备注", "") or ""

                    update_fields = {
                        "生产得分": new_prod,
                        "加减分": new_jiajian,
                        "月累计": new_month,
                    }
                    if new_remarks:
                        update_fields["备注"] = new_remarks if not old_remark else f"{old_remark}; {new_remarks}"

                    if record_id:
                        update_records.append({
                            "record_id": record_id,
                            "fields": update_fields
                        })

            if update_records:
                client.update_records(summary_table_id, update_records)
                updated_count = len(update_records)

        return created_count, updated_count, None

    except Exception as e:
        return 0, 0, f"{str(e)}\n{traceback.format_exc()}"


# ============================================================
# Flask 路由
# ============================================================

@app.route("/", methods=["GET", "POST"])
def index():
    message = None
    msg_type = "success"
    perf_msg = None

    if request.method == "POST":
        try:
            # 第一步：写入飞书流水账
            result = write_to_feishu(request.form)
            person_names, _ = parse_persons_from_form(request.form)
            record_id = result[0].get("record_id", "") if result else ""

            # 第二步：自动计算绩效并写入
            detail_count, summary_count, calc_error = auto_calc_performance(request.form, record_id)
            if calc_error:
                perf_msg = f"⚠️ 绩效自动计算失败：{calc_error[:200]}"
                msg_type = "error"
            else:
                perf_msg = f"📊 自动计算绩效：写入明细 {detail_count} 条，更新汇总 {summary_count} 人"

            message = (f"✅ 提交成功！{len(person_names)} 人，"
                       f"工序：{request.form.get('process')}，"
                       f"批号：{request.form.get('batch_no')} → 已写入飞书流水账\n{perf_msg}")
            msg_type = "success"

        except Exception as e:
            message = f"❌ 提交失败：{str(e)}"
            msg_type = "error"

    today = datetime.date.today().strftime("%Y-%m-%d")

    return render_template_string(
        FORM_TEMPLATE,
        today=today,
        processes=PROCESS_LIST,
        modes=MODE_OPTIONS,
        chong_opts=CHONG_OPTIONS,
        message=message,
        msg_type=msg_type,
        form_data=request.form if request.method == "POST" and message and msg_type == "error" else {},
        persons=[],
    )


@app.route("/health")
def health():
    return jsonify({"status": "ok", "timestamp": datetime.datetime.now().isoformat()})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("DEBUG", "true").lower() == "true"
    app.run(host="0.0.0.0", port=port, debug=debug)