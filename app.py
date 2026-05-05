# -*- coding: utf-8 -*-
"""
🏭 绩效报工 Web 表单系统 v2.0 — 完全重写
============================================
核心改进：
  1. 前工序→只显示批号+批数，不显示箱数/板数
  2. 内外包→显示批号+箱数/板数，备注自动生成
  3. 人员分组管理（像微信分组一样直观）
  4. 质量分→按人添加（默认最高5分），支持批量全加
  5. 转片子/转药→按人×批数×5分 自动计算
  6. 备注自动拼接 如"配料26260401-03"或"内包26260401（268箱）"
  7. 直接从绩效规则库调取规则自动计算
"""
import sys, io, json, time, requests, re, os, hmac, hashlib
from datetime import datetime
from collections import defaultdict
from flask import Flask, request, jsonify, redirect, url_for

from wps_client import WPS365DbsheetClient, WPS365Error
from rules_excel import load_merged_rules_from_excel, load_workers_from_excel
from perf_engine import calc_perf as _calc_perf_engine, find_prefix

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

app = Flask(__name__)
app.secret_key = os.urandom(24)

# ========== WPS 365 多维表格配置 ==========
# 凭证优先读环境变量；未设置时使用项目负责人提供的测试凭证（生产环境务必改为环境变量）。
WPS365_CLIENT_ID = os.environ.get("WPS365_CLIENT_ID", "AK20260504IRHSOH")
WPS365_CLIENT_SECRET = os.environ.get("WPS365_CLIENT_SECRET", "65d143a23cebc5ffdbd8fa5f1961b6c2")
# 多维表 file_id：在 WPS 多维表文件链接或开放平台中获取
WPS_DBSHEET_FILE_ID = os.environ.get("WPS_DBSHEET_FILE_ID", "").strip()
# 各数据表 sheet_id 为整数，需在 WPS 多维表或开放平台中对应填写
WPS_SHEETS = {
    "流水账": int(os.environ.get("WPS_SHEET_FLOW", "0") or 0),
    "绩效明细": int(os.environ.get("WPS_SHEET_PERF", "0") or 0),
    "品种规则": int(os.environ.get("WPS_SHEET_RULES", "0") or 0),
    "人员绑定": int(os.environ.get("WPS_SHEET_BIND", "0") or 0),
    "车间人员": int(os.environ.get("WPS_SHEET_WORKERS", "0") or 0),
    "系数规则": int(os.environ.get("WPS_SHEET_COEFF", "0") or 0),
}
HOUR_SCORE = 20  # 20分/小时/人

# 工序分类
QIAN_GONGXU = ['配料', '混合', '预混', '制粒', '压片', '包衣']
NEIWAI_BAO = ['内包', '外包']
ALL_PROCESS = QIAN_GONGXU + NEIWAI_BAO

# 工序配置（前端动态显隐用）
# has_guoshao:是否显示过筛 has_daqing:是否显示大清 has_zhuangji:干法制粒装机
# has_zhiliang:质量分(所有工序都有) has_queyuan:缺员 has_xiaoqueyuan:小时缺员
# has_muju:模具 has_jianyao:捡药 has_fensui:粉碎 has_zhuanpian:转片子/转药
# has_chongxing:压片冲型 show_batch_count:显批数 hide_box_ban:隐藏箱数板数
GONGXU_OPTS = {
    '配料': {'has_guoshao':True,'has_daqing':True,'has_zhuangji':False,'has_zhiliang':True,
             'has_queyuan':True,'has_xiaoqueyuan':True,'has_muju':False,'has_jianyao':False,
             'has_fensui':False,'has_zhuanpian':True,'has_chongxing':False,
             'show_batch_count':True,'hide_box_ban':True},
    '混合': {'has_guoshao':False,'has_daqing':True,'has_zhuangji':False,'has_zhiliang':True,
             'has_queyuan':True,'has_xiaoqueyuan':True,'has_muju':False,'has_jianyao':False,
             'has_fensui':False,'has_zhuanpian':True,'has_chongxing':False,
             'show_batch_count':True,'hide_box_ban':True},
    '预混': {'has_guoshao':False,'has_daqing':True,'has_zhuangji':False,'has_zhiliang':True,
             'has_queyuan':True,'has_xiaoqueyuan':True,'has_muju':False,'has_jianyao':False,
             'has_fensui':False,'has_zhuanpian':True,'has_chongxing':False,
             'show_batch_count':True,'hide_box_ban':True},
    '制粒': {'has_guoshao':False,'has_daqing':True,'has_zhuangji':True,'has_zhiliang':True,
             'has_queyuan':True,'has_xiaoqueyuan':True,'has_muju':False,'has_jianyao':False,
             'has_fensui':False,'has_zhuanpian':True,'has_chongxing':False,
             'show_batch_count':True,'hide_box_ban':True},
    '压片': {'has_guoshao':False,'has_daqing':True,'has_zhuangji':True,'has_zhiliang':True,
             'has_queyuan':True,'has_xiaoqueyuan':True,'has_muju':False,'has_jianyao':False,
             'has_fensui':False,'has_zhuanpian':True,'has_chongxing':True,
             'show_batch_count':True,'hide_box_ban':True},
    '包衣': {'has_guoshao':False,'has_daqing':True,'has_zhuangji':False,'has_zhiliang':True,
             'has_queyuan':True,'has_xiaoqueyuan':True,'has_muju':False,'has_jianyao':False,
             'has_fensui':False,'has_zhuanpian':True,'has_chongxing':False,
             'show_batch_count':True,'hide_box_ban':True},
    '内包': {'has_guoshao':False,'has_daqing':True,'has_zhuangji':False,'has_zhiliang':True,
             'has_queyuan':True,'has_xiaoqueyuan':True,'has_muju':True,'has_jianyao':True,
             'has_fensui':True,'has_zhuanpian':True,'has_chongxing':False,
             'show_batch_count':False,'hide_box_ban':False},
    '外包': {'has_guoshao':False,'has_daqing':True,'has_zhuangji':False,'has_zhiliang':True,
             'has_queyuan':True,'has_xiaoqueyuan':True,'has_muju':True,'has_jianyao':True,
             'has_fensui':True,'has_zhuanpian':True,'has_chongxing':False,
             'show_batch_count':False,'hide_box_ban':False},
}

# 分组管理文件
GROUPS_FILE = os.path.join(os.path.dirname(__file__), 'groups.json')

# ========== 工具函数 ==========
def num(v, d=0):
    if v is None: return d
    if isinstance(v, (int, float)): return float(v)
    try: return float(str(v).strip())
    except: return d

def sv(v):
    if v is None: return ''
    if isinstance(v, (int, float)): return str(int(v)) if v == int(v) else str(v)
    return str(v)


# ========== 本地分组管理 ==========
def load_groups():
    """加载分组配置"""
    if not os.path.exists(GROUPS_FILE):
        return {}
    try:
        with open(GROUPS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except:
        return {}

def save_groups(groups):
    """保存分组配置"""
    with open(GROUPS_FILE, 'w', encoding='utf-8') as f:
        json.dump(groups, f, ensure_ascii=False, indent=2)


# ========== WPS 多维表格 API ==========
def _wps_client() -> WPS365DbsheetClient:
    return WPS365DbsheetClient(client_id=WPS365_CLIENT_ID, client_secret=WPS365_CLIENT_SECRET)


def _wps_ready() -> bool:
    return bool(WPS_DBSHEET_FILE_ID and WPS365_CLIENT_ID and WPS365_CLIENT_SECRET)


def _wps_list_records(sheet_key: str):
    sid = WPS_SHEETS.get(sheet_key, 0)
    if not _wps_ready() or not sid:
        return []
    return _wps_client().list_records(WPS_DBSHEET_FILE_ID, sid)


def _wps_create_records(sheet_key: str, fields_rows: list) -> dict:
    sid = WPS_SHEETS.get(sheet_key, 0)
    if not _wps_ready() or not sid:
        raise WPS365Error("未配置 WPS_DBSHEET_FILE_ID 或对应 WPS_SHEET_* 工作表 ID")
    return _wps_client().create_records(WPS_DBSHEET_FILE_ID, sid, fields_rows)

# ========== 缓存 ==========
CACHE = {}
CACHE_TIME = {}
CACHE_TTL = 120

def get_cache(key, loader):
    now = time.time()
    if key in CACHE and now - CACHE_TIME.get(key, 0) < CACHE_TTL:
        return CACHE[key]
    try:
        data = loader()
        CACHE[key] = data
        CACHE_TIME[key] = now
        return data
    except Exception as e:
        if key in CACHE:
            return CACHE[key]
        raise e

def clear_cache():
    for k in list(CACHE.keys()):
        del CACHE[k]
        del CACHE_TIME[k]


# ========== 数据加载 ==========
def load_workers():
    workers: list = []
    try:
        if _wps_ready() and WPS_SHEETS.get("车间人员"):
            for rec in _wps_list_records("车间人员"):
                f = rec.get("fields", {})
                name = sv(f.get("姓名", ""))
                if name:
                    workers.append(name)
    except Exception:
        workers = []
    if not workers:
        workers = load_workers_from_excel()
    return sorted(set(workers))


def load_rules_wps():
    """从 WPS 品种规则表加载（专属号→规则行）"""
    rules = {}
    try:
        for rec in _wps_list_records("品种规则"):
            f = rec.get("fields", {})
            code = sv(f.get("专属号（批号前缀）", f.get("专属号", ""))).strip()
            if not code:
                continue
            clean = {}
            for k, v in f.items():
                if v is not None and v != "":
                    clean[k] = v
            rules[code] = clean
    except Exception:
        pass
    return rules


def load_rules():
    """本地 Excel 覆盖云端：先 WPS 再合并 Excel（同专属号以 Excel 为准）。"""
    rules = load_rules_wps()
    try:
        excel_rules = load_merged_rules_from_excel()
        for code, row in excel_rules.items():
            prev = rules.get(code, {})
            merged = dict(prev)
            merged.update(row)
            rules[code] = merged
    except Exception:
        pass
    return rules


def load_coeffs():
    """加载系数规则（WPS）"""
    coeffs = {}
    try:
        for rec in _wps_list_records("系数规则"):
            f = rec.get("fields", {})
            role = sv(f.get("岗位身份", ""))
            raw = f.get("系数", "1.0")
            try:
                coeffs[role] = float(raw) if raw else 1.0
            except Exception:
                coeffs[role] = 1.0
    except Exception:
        pass
    return coeffs


def calc_perf(form_data):
    """绩效计算（引擎见 perf_engine.py）。"""
    rules = get_cache('rules', load_rules)
    return _calc_perf_engine(form_data, rules, HOUR_SCORE)


def _wps_response_ok(resp):
    if not resp or not isinstance(resp, dict):
        return False
    if resp.get('data') is not None:
        return True
    return resp.get('code') in (0, '0')


# ========== 写入 WPS 多维表格 ==========
def write_to_dbsheet(form_data, results):
    """写入流水账 1 条 + 绩效明细（批量 create）。"""
    batches = form_data.get('batches', [])
    batch_strs = []
    for b in batches:
        bn = b.get('batch_no', '')
        bx = float(b.get('box', 0))
        if bx > 0:
            batch_strs.append('%s (%d箱)' % (bn, int(bx)))
        else:
            batch_strs.append(bn)
    batch_info = '; '.join(batch_strs)

    persons = form_data.get('persons', [])
    zp_batches = float(form_data.get('zpBatches', 0))
    zp_persons = form_data.get('zpPersons', [])
    qy_persons = form_data.get('qyPersons', [])
    qy_info = ','.join(['%s:%gh' % (qp['name'], float(qp.get('hours', 0))) for qp in qy_persons if float(qp.get('hours', 0)) > 0])
    jy_persons_list = form_data.get('jyPersons', [])
    first_result = results[0] if results else {}
    flow_remark = first_result.get('显示备注', batch_info)

    names_str = ','.join([p.get('name', '') for p in persons if p.get('name')])
    coeffs_str = ','.join([str(float(p.get('coeff', 1.0))) for p in persons if p.get('name')])
    qual_scores = ','.join([str(int(float(p.get('quality', 0)))) for p in persons if float(p.get('quality', 0)) > 0])

    date_cell = (form_data.get('日期') or str(datetime.now().date()))[:10]
    flow_fields = {
        '生产日期': date_cell,
        '工序': form_data.get('工序', ''),
        '批号': batch_info,
        '人员姓名': names_str,
        '人员系数': coeffs_str,
        '实际工时': float(form_data.get('工时', 8)),
        '是否有大清分': '是' if form_data.get('hasDQ') else '否',
        '缺员分': float(form_data.get('qyTotal', 0)),
        '小时缺员分': float(form_data.get('xqyTotal', 0)),
        '质量分': qual_scores,
        '缺员人员': qy_info,
        '粉碎分': float(form_data.get('fsScore', 0)),
        '转片子批数': float(form_data.get('zpBatches', 0)),
        '转片子人员': ','.join(zp_persons),
        '捡药筐数': float(form_data.get('jyBaskets', 0)),
        '捡药人员': ','.join(jy_persons_list),
        '更换模具人员': ','.join(form_data.get('mgChangePersons', [])),
        '模具确认人员': ','.join(form_data.get('mgConfirmPersons', [])),
        '备注': flow_remark,
    }

    try:
        flow_resp = _wps_create_records('流水账', [flow_fields])
    except WPS365Error:
        return False, 0
    flow_ok = _wps_response_ok(flow_resp)

    perf_rows = []
    for r in results:
        perf_rows.append({
            '生产日期': date_cell,
            '姓名': r['姓名'],
            '工序': r['工序'],
            '批号': batch_info,
            '基础分': r['基础分'],
            '大清分': r['大清'],
            '过筛分': r['过筛'],
            '装机分': r['装机'],
            '缺员补偿分': r['缺员补'],
            '小时缺员补偿分': r['时缺补'],
            '质量补偿分': r['质量'],
            '粉碎分': r['粉碎'],
            '转片子分': r['转片'],
            '更换模具分': r['模具换'],
            '模具确认分': r['模具确'],
            '捡药分': r['捡药'],
            '缺员扣减分': r['扣'],
            '系数得分': r['系数分'],
            '不乘系数得分': r['加减分'],
            '个人总得分': r['总分'],
            '备注': r['备注'],
        })

    written = 0
    try:
        perf_resp = _wps_create_records('绩效明细', perf_rows)
        if _wps_response_ok(perf_resp):
            pdata = perf_resp.get('data') or {}
            recs = pdata.get('records') or pdata.get('items') or []
            written = len(recs) if recs else len(perf_rows)
    except WPS365Error:
        written = 0

    return flow_ok, written


# ======================================================
# HTML 模板
# ======================================================
FORM_HTML = '''
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>🏭 绩效报工系统 v2</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Microsoft YaHei", sans-serif;
            background: #f0f2f5; color: #333; padding: 20px;
        }
        .container { max-width: 860px; margin: 0 auto; }
        .header {
            background: linear-gradient(135deg, #1a73e8, #1557b0);
            color: white; padding: 20px 28px; border-radius: 14px 14px 0 0;
        }
        .header h1 { font-size: 22px; margin: 0; }
        .header p { opacity: 0.85; margin-top: 4px; font-size: 13px; }
        .card {
            background: white; padding: 24px 28px;
            border-radius: 0 0 14px 14px; box-shadow: 0 2px 12px rgba(0,0,0,0.08);
            margin-bottom: 18px;
        }
        .form-section { margin-bottom: 24px; }
        .form-section:last-child { margin-bottom: 0; }
        .section-title {
            font-size: 15px; font-weight: 600; color: #1a73e8;
            padding-bottom: 8px; border-bottom: 2px solid #e8f0fe;
            margin-bottom: 14px; display: flex; align-items: center; gap: 8px;
        }
        .form-row { display: grid; gap: 10px; margin-bottom: 12px; }
        .form-row.cols-2 { grid-template-columns: 1fr 1fr; }
        .form-row.cols-3 { grid-template-columns: 1fr 1fr 1fr; }
        .form-row.cols-4 { grid-template-columns: 1fr 1fr 1fr 1fr; }
        label {
            display: block; font-size: 13px; font-weight: 500;
            color: #555; margin-bottom: 3px;
        }
        input, select, textarea {
            width: 100%; padding: 9px 12px; font-size: 14px;
            border: 1.5px solid #ddd; border-radius: 8px;
            background: #fafafa; transition: all 0.2s;
        }
        input:focus, select:focus, textarea:focus {
            outline: none; border-color: #1a73e8; background: white;
            box-shadow: 0 0 0 3px rgba(26,115,232,0.12);
        }
        textarea { min-height: 50px; resize: vertical; }
        .btn {
            padding: 10px 24px; font-size: 14px; font-weight: 600;
            border: none; border-radius: 8px; cursor: pointer;
            transition: all 0.2s;
        }
        .btn-primary { background: #1a73e8; color: white; }
        .btn-primary:hover { background: #1557b0; transform: translateY(-1px); }
        .btn-success { background: #34a853; color: white; }
        .btn-success:hover { background: #2d9249; transform: translateY(-1px); }
        .btn-outline {
            background: transparent; border: 1.5px solid #1a73e8;
            color: #1a73e8; padding: 7px 14px; font-size: 13px;
        }
        .btn-outline:hover { background: #e8f0fe; }
        .btn-danger { background: #ea4335; color: white; padding: 6px 12px; font-size: 12px; }
        .btn-danger:hover { background: #c5221f; }
        .btn-sm { padding: 5px 12px; font-size: 12px; }
        .btn-warning { background: #f9ab00; color: white; }
        .btn-warning:hover { background: #e09800; }
        .field-group {
            background: #f8f9fa; border-radius: 10px; padding: 14px;
            margin-bottom: 10px; display: none;
        }
        .field-group.visible { display: block; }
        .field-hint { font-size: 12px; color: #888; margin-top: 2px; }
        
        /* 左侧分组导航 */
        .layout-sidebar { display: flex; gap: 16px; }
        .sidebar {
            width: 160px; flex-shrink: 0;
            background: #f8f9fa; border-radius: 10px; padding: 10px;
            border: 1px solid #eee;
        }
        .sidebar .group-item {
            padding: 8px 12px; border-radius: 6px; cursor: pointer;
            font-size: 13px; margin-bottom: 3px; transition: all 0.15s;
        }
        .sidebar .group-item:hover { background: #e8f0fe; }
        .sidebar .group-item.active { background: #1a73e8; color: white; }
        .sidebar .group-item .badge {
            float: right; background: #ddd; border-radius: 10px;
            padding: 0 6px; font-size: 11px; color: #666;
        }
        .sidebar .group-item.active .badge { background: rgba(255,255,255,0.3); color: white; }
        .main-area { flex: 1; min-width: 0; }
        
        .person-table { width: 100%; border-collapse: collapse; font-size: 13px; }
        .person-table th {
            background: #e8f0fe; color: #1a73e8; padding: 8px;
            text-align: left; font-weight: 600; font-size: 12px;
        }
        .person-table td { padding: 6px 8px; border-bottom: 1px solid #f0f0f0; }
        .person-table input[type="number"] { width: 70px; padding: 4px 6px; }
        
        .batch-item {
            background: #f8f9fa; border: 1px solid #e8e8e8;
            border-radius: 10px; padding: 12px; margin-bottom: 8px;
            position: relative;
        }
        .batch-item .remove-btn {
            position: absolute; top: 6px; right: 6px;
        }
        
        /* 质量分输入 */
        .qual-input { width: 60px !important; padding: 4px 6px !important; text-align: center; }
        
        .alert {
            padding: 12px 16px; border-radius: 10px; margin-bottom: 14px;
            font-size: 14px; display: none;
        }
        .alert.success { display: block; background: #e6f4ea; color: #1e7e34; border: 1px solid #b7e1c6; }
        .alert.error { display: block; background: #fce8e6; color: #c5221f; border: 1px solid #f5c6cb; }
        .alert.loading { display: block; background: #e8f0fe; color: #1a73e8; border: 1px solid #c6dafc; }
        
        .result-table { width: 100%; border-collapse: collapse; font-size: 13px; margin-top: 10px; }
        .result-table th { background: #1a73e8; color: white; padding: 7px 8px; text-align: center; }
        .result-table td { padding: 7px 8px; border-bottom: 1px solid #eee; text-align: center; }
        .result-table tr:nth-child(even) { background: #f8f9fa; }
        .hidden { display: none !important; }
        .mt-8 { margin-top: 8px; }
        .mt-12 { margin-top: 12px; }
        .mb-8 { margin-bottom: 8px; }
        .inline-flex { display: inline-flex; align-items: center; gap: 6px; }
        .tag-badge {
            display: inline-block; background: #e8f0fe; color: #1a73e8;
            padding: 2px 10px; border-radius: 12px; font-size: 12px;
            margin: 2px; cursor: pointer; border: none;
        }
        .tag-badge:hover { background: #d2e3fc; }
        .flex-row { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
        @media (max-width: 720px) {
            .layout-sidebar { flex-direction: column; }
            .sidebar { width: 100%; display: flex; flex-wrap: wrap; }
            .sidebar .group-item { flex: 1; text-align: center; min-width: 80px; }
            .form-row.cols-2, .form-row.cols-3, .form-row.cols-4 { grid-template-columns: 1fr; }
            .card { padding: 14px; }
        }
    </style>
</head>
<body>
<div class="container">
    <div class="header">
        <h1>📋 每日绩效报工</h1>
        <p>选择工序 → 选组选人 → 填批号产量 → 自动计算</p>
    </div>
    <div class="card">
        <div id="alertBox" class="alert"></div>
        
        <!-- ===== 基本信息 ===== -->
        <div class="form-section">
            <div class="section-title">📅 基本信息</div>
            <div class="form-row cols-4">
                <div>
                    <label>生产日期 *</label>
                    <input type="date" id="date" value="''' + datetime.now().strftime('%Y-%m-%d') + '''" />
                </div>
                <div>
                    <label>工序 *</label>
                    <select id="process" onchange="onProcessChange()">
                        <option value="">-- 请选择 --</option>
                        <option value="配料">配料</option>
                        <option value="混合">混合</option>
                        <option value="预混">预混</option>
                        <option value="制粒">制粒</option>
                        <option value="压片">压片</option>
                        <option value="包衣">包衣</option>
                        <option value="内包">内包</option>
                        <option value="外包">外包</option>
                    </select>
                </div>
                <div>
                    <label>工时（小时）</label>
                    <input type="number" id="hours" value="8" min="0.5" max="24" step="0.5" />
                </div>
                <div>
                    <label>报工组长</label>
                    <input type="text" id="leader" placeholder="组长签名" />
                </div>
            </div>
        </div>
        
        <!-- ===== 品种批次 ===== -->
        <div class="form-section">
            <div class="section-title">
                🏷️ 品种与批号
                <button type="button" class="btn btn-outline btn-sm" onclick="addBatch()" style="margin-left:auto;display:none;" id="addBatchBtn">+ 添加品种</button>
            </div>
            <div id="batchContainer">
                <div class="batch-item" id="batch0">
                    <div class="form-row">
                        <div id="batchFrontField" style="grid-column:1/-1;">
                            <label>生产批号 <span style="color:#999;font-weight:normal;">（如 26260401-03，自动匹配品种规则）</span></label>
                            <input type="text" id="batch0_no" placeholder="如: 26260401-03" onchange="onBatchChange(0)" />
                        </div>
                    </div>
                    <div class="field-hint" id="batch0_hint" style="margin-top:-8px;margin-bottom:8px;color:#1a73e8;"></div>
                    <div class="form-row cols-2" id="batchBoxField" style="display:none;">
                        <div>
                            <label>箱数</label>
                            <input type="number" id="batch0_box" value="0" min="0" step="1" />
                        </div>
                        <div>
                            <label>板数</label>
                            <input type="number" id="batch0_ban" value="0" min="0" step="1" />
                        </div>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- ===== 人员选择 ===== -->
        <div class="form-section">
            <div class="section-title">
                👥 人员与系数
                <span style="font-size:12px;color:#999;font-weight:normal;margin-left:8px;">
                    （从左侧分组选择或手动添加）
                </span>
            </div>
            
            <div class="layout-sidebar">
                <!-- 左侧分组导航 -->
                <div class="sidebar" id="groupSidebar">
                    <div style="font-size:12px;color:#888;margin-bottom:6px;">📁 人员分组</div>
                    <div id="groupList"></div>
                    <div class="mt-8">
                        <button class="btn btn-outline btn-sm" onclick="window.open('/groups','_blank')" style="width:100%;">⚙️ 管理分组</button>
                    </div>
                </div>
                
                <!-- 右侧人员表格 -->
                <div class="main-area">
                    <table class="person-table">
                        <thead>
                            <tr>
                                <th style="width:24px;"><input type="checkbox" id="selectAll" onchange="toggleAllPersons()" checked /></th>
                                <th>姓名</th>
                                <th style="width:80px;">系数</th>
                                <th style="width:80px;">质量分</th>
                                <th style="width:36px;"></th>
                            </tr>
                        </thead>
                        <tbody id="personTableBody"></tbody>
                    </table>
                    
                    <div class="flex-row mt-8">
                        <button type="button" class="btn btn-outline btn-sm" onclick="addPerson()">+ 添加人员</button>
                        <button type="button" class="btn btn-outline btn-sm" onclick="setAllQuality(5)">⭐ 全加5分</button>
                        <button type="button" class="btn btn-outline btn-sm" onclick="setAllQuality(0)">✖ 全清零</button>
                        <span style="font-size:12px;color:#999;">默认质量分最高5分/人</span>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- ===== 额外分项（动态显示） ===== -->
        <div class="form-section">
            <div class="section-title">➕ 额外加减分</div>
            
            <!-- 大清分 -->
            <div class="field-group" id="dqGroup">
                <label class="inline-flex">
                    <input type="checkbox" id="hasDQ" /> 有大清分
                    <span style="font-size:12px;color:#999;margin-left:4px;">（从规则库自动取值，所有人平分）</span>
                </label>
            </div>
            
            <!-- 过筛分（配料）—— 自动根据品种规则判断，无需手动勾选 -->
            <div class="field-group" id="gsGroup" style="display:none;">
                <label class="inline-flex">
                    <span style="font-size:12px;color:#1a73e8;">ℹ️ 过筛分已根据品种规则自动计算</span>
                </label>
            </div>
            
            <!-- 干法制粒装机 -->
            <div class="field-group" id="zjGroup">
                <label class="inline-flex">
                    <input type="checkbox" id="hasZJ" /> 有干法制粒装机
                    <span style="font-size:12px;color:#999;margin-left:4px;">（从规则库自动取值，所有人平分）</span>
                </label>
            </div>
            
            <!-- 压片冲型 -->
            <div class="field-group" id="chongGroup">
                <label>压片冲型</label>
                <select id="chongType" onchange="onExtraChange()">
                    <option value="">-- 不装机 --</option>
                    <option value="77">77冲</option>
                    <option value="65">65冲</option>
                    <option value="40">40冲</option>
                </select>
            </div>
            
            <!-- 转片子/转药 -->
            <div class="field-group" id="zpGroup">
                <div class="form-row cols-2">
                    <div>
                        <label>转片子/转药 批数</label>
                        <input type="number" id="zpBatches" value="0" min="0" step="1" />
                    </div>
                    <div>
                        <label>指定人员（逗号分隔）</label>
                        <input type="text" id="zpPersons" placeholder="如: 王东松,姚娟" />
                        <div class="field-hint">自动按 批数×5分 ÷ 人数 计算</div>
                    </div>
                </div>
            </div>
            
            <!-- 模具 -->
            <div class="field-group" id="mjGroup">
                <div class="form-row cols-2">
                    <div>
                        <label>更换模具人员（逗号分隔）</label>
                        <input type="text" id="mgChangePersons" placeholder="如: 李勇,张勇攀" />
                        <div class="field-hint">每人 +30分</div>
                    </div>
                    <div>
                        <label>模具确认人员（逗号分隔）</label>
                        <input type="text" id="mgConfirmPersons" placeholder="如: 张勇攀" />
                        <div class="field-hint">每人 +10分</div>
                    </div>
                </div>
            </div>
            
            <!-- 捡药（内外包） -->
            <div class="field-group" id="jyGroup">
                <div class="form-row cols-2">
                    <div>
                        <label>捡药筐数</label>
                        <input type="number" id="jyBaskets" value="0" min="0" step="1" />
                    </div>
                    <div>
                        <label>捡药人员（逗号分隔）</label>
                        <input type="text" id="jyPersons" placeholder="如: 刘洁,刘金浩" />
                        <div class="field-hint">自动按 筐数×单价÷人数 计算</div>
                    </div>
                </div>
            </div>
            
            <!-- 粉碎（内外包） -->
            <div class="field-group" id="fsGroup">
                <div class="form-row cols-2">
                    <div>
                        <label>粉碎批数</label>
                        <input type="number" id="fsScore" value="0" min="0" step="0.5" />
                    </div>
                    <div>
                        <label>粉碎人员（逗号分隔）</label>
                        <input type="text" id="fsPersons" placeholder="如: 刘金浩,石法智" />
                        <div class="field-hint">自动平分</div>
                    </div>
                </div>
            </div>
            
            <!-- 缺员（自动从定员推导补偿分，此处仅填写缺员扣减名单） -->
            <div class="field-group" id="qyGroup">
                <div class="field-hint" style="margin-bottom:8px;color:#1a73e8;">ℹ️ 缺员补偿分已根据品种规则定员自动计算，此处仅填写缺员扣减名单</div>
                <div class="form-row cols-2">
                    <div>
                        <label>缺员名单（逗号分隔）</label>
                        <input type="text" id="qyNames" placeholder="如: 王磊,刘秀霞" />
                    </div>
                    <div>
                        <label>缺员小时数（逗号分隔，对应上述名单）</label>
                        <input type="text" id="qyHours" placeholder="如: 8,4" />
                        <div class="field-hint">扣减: 小时数 × 20分/小时</div>
                    </div>
                </div>
            </div>
            
            <!-- 小时缺员 -->
            <div class="field-group" id="xqyGroup" style="display:none;">
                <!-- 保留占位，与前端 onProcessChange 兼容 -->
            </div>
        </div>
        
        <!-- ===== 试验品种工时考核 ===== -->
        <div class="form-section" id="ceshiSection" style="display:none;">
            <div class="section-title">🧪 试验品种工时考核</div>
            <div class="field-group">
                <label class="inline-flex">
                    <input type="checkbox" id="ceshi" onchange="onExtraChange()" /> 对无规则匹配的试验品种，按工时（小时×20分）计算
                    <span style="font-size:12px;color:#999;margin-left:4px;">（关闭时无规则品种不产生得分）</span>
                </label>
            </div>
        </div>
        
        <!-- ===== 底部按钮 ===== -->
        <div class="form-section">
            <div class="flex-row" style="justify-content:center;">
                <button class="btn btn-primary" onclick="previewCalc()">🔍 预览计算</button>
                <button class="btn btn-success" onclick="submitForm()">📤 提交报工</button>
            </div>
            <div style="text-align:center;margin-top:8px;">
                <a href="/groups" target="_blank" style="color:#1a73e8;font-size:13px;">⚙️ 人员分组管理</a>
                <span style="color:#ddd;margin:0 8px;">|</span>
                <a href="/data" target="_blank" style="color:#1a73e8;font-size:13px;">📊 今日数据</a>
            </div>
        </div>
        
        <!-- ===== 预览区域 ===== -->
        <div id="previewArea" class="hidden">
            <div class="section-title">📊 计算结果预览</div>
            <div id="previewContent"></div>
        </div>
    </div>
</div>

<datalist id="workerList"></datalist>

<script>
// ========== 状态 ==========
let personCount = 0;
let batchCount = 1;
let groups = {};
let allWorkers = [];
let currentGroup = '';

// ========== 初始化 ==========
window.onload = function() {
    loadWorkerList();
    loadGroups();
    loadRuleList(); // 加载规则库品种
    document.getElementById('date').valueAsDate = new Date();
};

let allRules = [];
function loadRuleList() {
    fetch('/api/rules').then(r=>r.json()).then(data => {
        allRules = data.rules || [];
        renderProductOptions(0);
    });
}

function renderProductOptions(idx) {
    const sel = document.getElementById('batch'+idx+'_product');
    if (!sel) return;
    sel.innerHTML = '<option value="">-- 自动匹配 --</option>' + 
        allRules.map(r => '<option value="'+r.code+'">'+r.code+' | '+r.name+'</option>').join('');
}

function loadWorkerList() {
    fetch('/api/workers').then(r=>r.json()).then(data => {
        allWorkers = data.workers || [];
        const dl = document.getElementById('workerList');
        dl.innerHTML = allWorkers.map(w => '<option value="'+w+'">').join('');
    });
}

function loadGroups() {
    fetch('/api/groups').then(r=>r.json()).then(data => {
        groups = data.groups || {};
        renderGroupList();
        if (currentGroup && groups[currentGroup]) {
            renderGroupMembers(currentGroup);
        }
    });
}

// ========== 分组渲染 ==========
function renderGroupList() {
    const gl = document.getElementById('groupList');
    const keys = Object.keys(groups);
    if (keys.length === 0) {
        gl.innerHTML = '<div style="font-size:12px;color:#999;padding:8px;">暂无分组，请先管理</div>';
        return;
    }
    gl.innerHTML = keys.map(k => 
        '<div class="group-item' + (k === currentGroup ? ' active' : '') + '" onclick="selectGroup(\\''+k+'\\')">' +
        k + ' <span class="badge">' + (groups[k] ? groups[k].length : 0) + '</span></div>'
    ).join('');
}

function selectGroup(groupName) {
    currentGroup = groupName;
    renderGroupList();
    renderGroupMembers(groupName);
}

function renderGroupMembers(groupName) {
    const members = groups[groupName] || [];
    const tb = document.getElementById('personTableBody');
    
    // 清空现有
    tb.innerHTML = '';
    personCount = 0;
    
    if (members.length === 0) {
        tb.innerHTML = '<tr><td colspan="5" style="text-align:center;color:#999;padding:20px;">该分组暂无人员</td></tr>';
        return;
    }
    
    members.forEach((m, idx) => {
        addPersonRow(idx, m.name, m.coeff || 1.0, m.quality || 0);
    });
}

function addPersonRow(idx, name, coeff, quality) {
    const tb = document.getElementById('personTableBody');
    const tr = document.createElement('tr');
    tr.id = 'personRow'+idx;
    tr.innerHTML = 
        '<td><input type="checkbox" class="person-check" checked onchange="onPersonCheck('+idx+')" /></td>' +
        '<td><input type="text" value="'+name+'" style="border:none;background:transparent;font-weight:500;" readonly /></td>' +
        '<td><input type="number" class="coeff-input" value="'+coeff+'" min="0.5" max="2.0" step="0.01" /></td>' +
        '<td><input type="number" class="qual-input" value="'+quality+'" min="0" max="5" step="0.5" /></td>' +
        '<td><button class="btn btn-danger btn-sm" onclick="removePersonRow('+idx+')">✕</button></td>';
    tb.appendChild(tr);
    personCount = idx + 1;
}

// ========== 额外添加人员 ==========
function addPerson() {
    const idx = personCount;
    const tb = document.getElementById('personTableBody');
    const tr = document.createElement('tr');
    tr.id = 'personRow'+idx;
    tr.innerHTML = 
        '<td><input type="checkbox" class="person-check" checked /></td>' +
        '<td><input type="text" list="workerList" placeholder="输入姓名" style="border:none;background:#f5f5f5;" /></td>' +
        '<td><input type="number" class="coeff-input" value="1.0" min="0.5" max="2.0" step="0.01" /></td>' +
        '<td><input type="number" class="qual-input" value="0" min="0" max="5" step="0.5" /></td>' +
        '<td><button class="btn btn-danger btn-sm" onclick="removePersonRow('+idx+')">✕</button></td>';
    tb.appendChild(tr);
    personCount = idx + 1;
    
    // 自动聚焦输入
    const inp = tr.querySelector('input[type="text"]');
    if (inp) { inp.focus(); inp.select(); }
}

function removePersonRow(idx) {
    const tr = document.getElementById('personRow'+idx);
    if (tr) tr.remove();
}

function toggleAllPersons() {
    const checked = document.getElementById('selectAll').checked;
    document.querySelectorAll('.person-check').forEach(cb => cb.checked = checked);
}

function setAllQuality(score) {
    document.querySelectorAll('.qual-input').forEach(inp => inp.value = score);
}

// ========== 工序变更 ==========
const PROC_OPTS = ''' + json.dumps(GONGXU_OPTS, ensure_ascii=False) + ''';

function onProcessChange() {
    const proc = document.getElementById('process').value;
    
    // 先隐藏所有分组
    document.querySelectorAll('.field-group').forEach(g => g.classList.remove('visible'));
    document.getElementById('addBatchBtn').style.display = 'none';
    
    if (!proc) return;
    
    const opts = PROC_OPTS[proc];
    if (!opts) return;
    
    // 显示对应字段
    if (opts.has_daqing) showGroup('dqGroup');
    if (opts.has_guoshao) showGroup('gsGroup');
    if (opts.has_zhuangji) showGroup('zjGroup');
    if (opts.has_chongxing) showGroup('chongGroup');
    if (opts.has_zhuanpian) showGroup('zpGroup');
    if (opts.has_queyuan) showGroup('qyGroup');
    if (opts.has_xiaoqueyuan) showGroup('xqyGroup');
    if (opts.has_muju) showGroup('mjGroup');
    if (opts.has_jianyao) showGroup('jyGroup');
    if (opts.has_fensui) showGroup('fsGroup');
    
    // 试验品种工时考核开关（仅前工序显示）
    document.getElementById('ceshiSection').style.display = opts.hide_box_ban ? 'block' : 'none';
    document.getElementById('ceshi').checked = false;
    
    // 批号输入切换
    const frontField = document.getElementById('batchFrontField');
    const boxField = document.getElementById('batchBoxField');
    const hint = document.querySelector('#batch0_no + .field-hint');
    
    if (opts.hide_box_ban) {
        // 前工序：显示批号+批数，隐藏箱数/板数
        frontField.style.display = 'block';
        boxField.style.display = 'none';
        if (hint) hint.textContent = '如 26260401-03（自动解析批数）';
        document.querySelector('#batch0_no').placeholder = '如: 26260401-03';
    } else {
        // 内外包：显示批号+箱数/板数
        frontField.style.display = 'block';
        boxField.style.display = 'grid';
        if (hint) hint.textContent = '如 26260401（自动匹配品种规则）';
        document.querySelector('#batch0_no').placeholder = '如: 26260401';
    }
    
    // 内外包可以添加多个品种
    if (opts.hide_box_ban === false) {
        document.getElementById('addBatchBtn').style.display = 'inline-block';
    }
}

function showGroup(id) {
    document.getElementById(id).classList.add('visible');
}

function onExtraChange() {}

// ========== 批号输入 ==========
function onBatchChange(idx) {
    const no = document.getElementById('batch'+idx+'_no');
    const sel = document.getElementById('batch'+idx+'_product');
    const hint = document.getElementById('batch'+idx+'_hint');
    if (!no || !hint) return;
    
    // 如果没有手动选品种，则尝试自动匹配
    if (!sel || !sel.value) {
        fetch('/api/find_rule?batch=' + encodeURIComponent(no.value))
            .then(r => r.json())
            .then(data => {
                if (data.name) {
                    const proc = document.getElementById('process').value;
                    const isFront = !PROC_OPTS[proc] || PROC_OPTS[proc].hide_box_ban;
                    if (isFront) {
                        const bc = parseBatchCount(no.value);
                        hint.textContent = '✅ ' + data.name + ' | 解析: ' + bc.count + '批';
                    } else {
                        hint.textContent = '✅ ' + data.name;
                    }
                } else {
                    hint.textContent = '⚠️ 未匹配到品种，请手动选择 ->';
                }
            })
            .catch(() => hint.textContent = '');
    } else {
        const bc = parseBatchCount(no.value);
        const rule = allRules.find(r => r.code === sel.value);
        hint.textContent = '✅ 已选: ' + (rule ? rule.name : sel.value) + ' | 解析: ' + bc.count + '批';
    }
}

function parseBatchCount(s) {
    let count = 1;
    let clean = s;
    if (!s) return {clean: s, count: 1};
    
    // 括号批数: "26260401（2批）"
    const m = s.match(/[（(](\\d+)\\s*批[)）]/);
    if (m) {
        count = parseInt(m[1]);
        clean = s.replace(/[（(].*?[)）]/g, '').trim();
        return {clean, count};
    }
    
    const s2 = s.replace('批', '');
    
    // 范围: "26260401-03"
    if (s2.includes('-')) {
        const parts = s2.split('-');
        const l = parts[0].replace(/\\D/g, '');
        const r = parts[1].replace(/\\D/g, '');
        if (l && r) {
            const rv = parseInt(r);
            const lv = l.length >= r.length ? parseInt(l.slice(-r.length)) : parseInt(l);
            count = (rv >= lv) ? (rv - lv + 1) : 1;
            return {clean: l, count};
        }
        return {clean: l || parts[0], count: 1};
    }
    
    // 逗号分隔
    const parts = s2.split(/[,，+、\\s]+/).filter(p => p.trim());
    if (parts.length > 1) {
        count = parts.length;
        return {clean: parts[0].trim(), count};
    }
    
    return {clean: s, count: 1};
}

// ========== 品种添加/删除（内外包） ==========
function addBatch() {
    const container = document.getElementById('batchContainer');
    const idx = batchCount;
    const div = document.createElement('div');
    div.className = 'batch-item';
    div.id = 'batch'+idx;
    div.innerHTML = 
        '<button type="button" class="btn btn-danger btn-sm remove-btn" onclick="removeBatch('+idx+')">✕</button>' +
        '<div class="form-row">' +
            '<div style="grid-column:1/-1;">' +
                '<label>生产批号 <span style="color:#999;font-weight:normal;">（自动匹配品种规则）</span></label>' +
                '<input type="text" id="batch'+idx+'_no" placeholder="如: 26260401" onchange="onBatchChange('+idx+')" />' +
            '</div>' +
        '</div>' +
        '<div class="field-hint" id="batch'+idx+'_hint" style="margin-top:-8px;margin-bottom:8px;color:#1a73e8;"></div>' +
        '<div class="form-row cols-2">' +
            '<div><label>箱数</label><input type="number" id="batch'+idx+'_box" value="0" min="0" step="1" /></div>' +
            '<div><label>板数</label><input type="number" id="batch'+idx+'_ban" value="0" min="0" step="1" /></div>' +
        '</div>';
    container.appendChild(div);
    renderProductOptions(idx);
    batchCount++;
}

function removeBatch(idx) {
    const el = document.getElementById('batch'+idx);
    if (el) el.remove();
}

// ========== 收集表单数据 ==========
function getFormData() {
    // 批号数据
    const batches = [];
    for (let i = 0; i < batchCount; i++) {
        const no = document.getElementById('batch'+i+'_no');
        const sel = document.getElementById('batch'+i+'_product');
        if (!no || !no.value.trim()) continue;
        batches.push({
            batch_no: no.value,
            product_code: sel ? sel.value : '',
            box: parseFloat(document.getElementById('batch'+i+'_box')?.value || 0),
            ban: parseFloat(document.getElementById('batch'+i+'_ban')?.value || 0),
        });
    }
    if (batches.length === 0 && document.getElementById('batch0_no')?.value) {
        const sel = document.getElementById('batch0_product');
        batches.push({
            batch_no: document.getElementById('batch0_no').value,
            product_code: sel ? sel.value : '',
            box: parseFloat(document.getElementById('batch0_box')?.value || 0),
            ban: parseFloat(document.getElementById('batch0_ban')?.value || 0),
        });
    }
    
    // 人员数据
    const persons = [];
    const rows = document.querySelectorAll('#personTableBody tr');
    rows.forEach(row => {
        const cb = row.querySelector('.person-check');
        if (!cb || !cb.checked) return;
        const nameInput = row.querySelector('input[type="text"]');
        const coeffInput = row.querySelector('.coeff-input');
        const qualInput = row.querySelector('.qual-input');
        if (!nameInput || !nameInput.value.trim()) return;
        persons.push({
            name: nameInput.value.trim(),
            coeff: parseFloat(coeffInput?.value || 1.0),
            quality: parseFloat(qualInput?.value || 0),
        });
    });
    
    // 缺员
    const qyNamesStr = document.getElementById('qyNames')?.value || '';
    const qyHoursStr = document.getElementById('qyHours')?.value || '';
    const qyNames = qyNamesStr.split(',').map(s => s.trim()).filter(s => s);
    const qyHours = qyHoursStr.split(',').map(s => parseFloat(s.trim()) || 0);
    const qyPersons = [];
    qyNames.forEach((name, i) => {
        if (i < qyHours.length) qyPersons.push({name, hours: qyHours[i]});
    });
    
    // 转片子
    const zpPersonsStr = document.getElementById('zpPersons')?.value || '';
    const zpPersons = zpPersonsStr.split(',').map(s => s.trim()).filter(s => s);
    
    // 捡药
    const jyPersonsStr = document.getElementById('jyPersons')?.value || '';
    const jyPersons = jyPersonsStr.split(',').map(s => s.trim()).filter(s => s);
    
    // 模具
    const mgChangeStr = document.getElementById('mgChangePersons')?.value || '';
    const mgChangePersons = mgChangeStr.split(',').map(s => s.trim()).filter(s => s);
    const mgConfirmStr = document.getElementById('mgConfirmPersons')?.value || '';
    const mgConfirmPersons = mgConfirmStr.split(',').map(s => s.trim()).filter(s => s);
    
    // 粉碎
    const fsPersonsStr = document.getElementById('fsPersons')?.value || '';
    const fsPersons = fsPersonsStr.split(',').map(s => s.trim()).filter(s => s);
    
    return {
        日期: document.getElementById('date').value,
        工序: document.getElementById('process').value,
        工时: parseFloat(document.getElementById('hours').value || 8),
        batches: batches,
        persons: persons,
        hasDQ: document.getElementById('hasDQ')?.checked || false,
        hasGS: document.getElementById('process').value === '配料',
        hasZJ: document.getElementById('hasZJ')?.checked || false,
        chong: document.getElementById('chongType')?.value || '',
        ceshi: document.getElementById('ceshi')?.checked || false,
        qyTotal: 0,
        xqyTotal: 0,
        qyPersons: qyPersons,
        zpBatches: parseFloat(document.getElementById('zpBatches')?.value || 0),
        zpPersons: zpPersons,
        mgChangePersons: mgChangePersons,
        mgConfirmPersons: mgConfirmPersons,
        jyBaskets: parseFloat(document.getElementById('jyBaskets')?.value || 0),
        jyPersons: jyPersons,
        fsScore: parseFloat(document.getElementById('fsScore')?.value || 0),
    };
}

// ========== 预览计算 ==========
function previewCalc() {
    const data = getFormData();
        if (!data.工序) { showAlert('请选择工序', 'error'); return; }
        if (data.persons.length === 0) { showAlert('请添加至少一名人员', 'error'); return; }
        
        showAlert('正在计算...', 'loading');
    
    fetch('/api/calc', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify(data),
    })
    .then(r => r.json())
    .then(res => {
        hideAlert();
        document.getElementById('previewArea').classList.remove('hidden');
        if (res.error) {
            document.getElementById('previewContent').innerHTML = '<div class="alert error">'+res.error+'</div>';
            return;
        }
        let html = '<table class="result-table">' +
            '<tr><th>姓名</th><th>系数</th><th>生产得分</th><th>奖罚说明</th><th>扣减</th><th>总分</th><th>工序批号</th></tr>';
        res.results.forEach(r => {
            html += '<tr>' +
                '<td>'+r.姓名+'</td>' +
                '<td>'+r.系数+'</td>' +
                '<td>'+(r.生产得分||r.系数分)+'</td>' +
                '<td style="font-size:11px;text-align:left;max-width:180px;">'+(r.奖罚说明||'无')+'</td>' +
                '<td>'+(r.扣||0)+'</td>' +
                '<td><b>'+r.总分+'</b></td>' +
                '<td style="font-size:11px;text-align:left;max-width:180px;word-break:break-all;">'+(r.显示备注||r.备注)+'</td>' +
                '</tr>';
        });
        html += '</table>';
        document.getElementById('previewContent').innerHTML = html;
    })
    .catch(e => {
        hideAlert();
        showAlert('计算失败: '+e.message, 'error');
    });
}

// ========== 提交 ==========
function submitForm() {
    const data = getFormData();
        if (!data.工序) { showAlert('请选择工序', 'error'); return; }
        if (data.persons.length === 0) { showAlert('请添加至少一名人员', 'error'); return; }
        
        if (!confirm('确认提交 '+data.persons.length+' 人的报工？')) return;
    
    showAlert('⏳ 正在提交并计算绩效...', 'loading');
    
    fetch('/api/submit', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify(data),
    })
    .then(r => r.json())
    .then(res => {
        if (res.error) {
            showAlert('❌ '+res.error, 'error');
            return;
        }
        let msg = '✅ 提交成功！写入流水账1条，绩效明细'+res.written+'条';
        if (res.results) {
            let total = res.results.reduce((s, r) => s + r.总分, 0);
            let avg = total / res.results.length;
            msg += '<br/>📊 本批总分: '+total.toFixed(1)+'，人均: '+avg.toFixed(1);
        }
        showAlert(msg, 'success');
        
        setTimeout(() => {
            if (confirm('是否重置表单？')) { location.reload(); }
        }, 2000);
    })
    .catch(e => {
        showAlert('❌ 提交失败: '+e.message, 'error');
    });
}

// ========== 提示框 ==========
function showAlert(msg, type) {
    const box = document.getElementById('alertBox');
    box.className = 'alert ' + type;
    box.innerHTML = msg;
    box.style.display = 'block';
}

function hideAlert() {
    document.getElementById('alertBox').style.display = 'none';
}
</script>
</body>
</html>
'''


# ======================================================
# Flask 路由
# ======================================================
@app.route('/')
def index():
    return redirect(url_for('form'))

@app.route('/form')
def form():
    return FORM_HTML


# ========== API: 人员分组 ==========
@app.route('/api/groups')
def api_groups():
    groups = load_groups()
    return jsonify({'groups': groups})

@app.route('/api/save_group', methods=['POST'])
def api_save_group():
    data = request.json
    name = data.get('name', '').strip()
    members = data.get('members', [])
    if not name: return jsonify({'error': '分组名不能为空'})
    
    groups = load_groups()
    groups[name] = members
    save_groups(groups)
    return jsonify({'success': True})

@app.route('/api/delete_group', methods=['POST'])
def api_delete_group():
    data = request.json
    name = data.get('name', '').strip()
    groups = load_groups()
    if name in groups:
        del groups[name]
        save_groups(groups)
    return jsonify({'success': True})

@app.route('/api/add_member_to_group', methods=['POST'])
def api_add_member():
    data = request.json
    group_name = data.get('group', '').strip()
    member = data.get('member', {})
    if not group_name or not member.get('name'): return jsonify({'error': '参数不全'})
    
    groups = load_groups()
    if group_name not in groups: groups[group_name] = []
    
    # 检查是否已存在
    for m in groups[group_name]:
        if m.get('name') == member['name']:
            m['coeff'] = member.get('coeff', 1.0)
            save_groups(groups)
            return jsonify({'success': True, 'action': 'updated'})
    
    groups[group_name].append({
        'name': member['name'],
        'coeff': member.get('coeff', 1.0),
    })
    save_groups(groups)
    return jsonify({'success': True, 'action': 'added'})

@app.route('/api/remove_member_from_group', methods=['POST'])
def api_remove_member():
    data = request.json
    group_name = data.get('group', '').strip()
    member_name = data.get('name', '').strip()
    
    groups = load_groups()
    if group_name in groups:
        groups[group_name] = [m for m in groups[group_name] if m.get('name') != member_name]
        save_groups(groups)
    return jsonify({'success': True})


# ========== API: 其他 ==========
@app.route('/api/workers')
def api_workers():
    workers = get_cache('workers', load_workers)
    return jsonify({'workers': workers})

@app.route('/api/rules')
def api_rules():
    rules = get_cache('rules', load_rules)
    result = []
    for code, rule in rules.items():
        result.append({
            'code': code, 'name': rule.get('产品名称', ''),
            'product': rule.get('商品名', ''),
        })
    return jsonify({'rules': result})

@app.route('/api/find_rule')
def api_find_rule():
    batch = request.args.get('batch', '')
    rules = get_cache('rules', load_rules)
    prefix = find_prefix(batch, rules)
    rule = rules.get(prefix) if prefix else None
    if prefix and rule:
        return jsonify({
            'found': True, 'code': prefix, 'name': rule.get('产品名称', ''),
            'product': rule.get('商品名', ''),
        })
    return jsonify({'found': False})

def normalize_form_payload(d):
    """兼容旧版前端英文字段名。"""
    if not isinstance(d, dict):
        return {}
    o = dict(d)
    if (not o.get('日期')) and o.get('date'):
        o['日期'] = o['date']
    if (not o.get('工序')) and o.get('process'):
        o['工序'] = o['process']
    if o.get('工时') is None and o.get('hours') is not None:
        o['工时'] = o['hours']
    return o


@app.route('/api/calc', methods=['POST'])
def api_calc():
    try:
        data = normalize_form_payload(request.json or {})
        results = calc_perf(data)
        return jsonify({'results': results})
    except Exception as e:
        import traceback
        return jsonify({'error': str(e) + '\n' + traceback.format_exc()})


@app.route('/api/submit', methods=['POST'])
def api_submit():
    try:
        data = normalize_form_payload(request.json or {})
        results = calc_perf(data)
        flow_ok, written = write_to_dbsheet(data, results)
        if not flow_ok and written == 0:
            return jsonify({'error': '写入 WPS 多维表格失败，请检查 WPS_DBSHEET_FILE_ID、各 WPS_SHEET_* 及开放平台权限'})
        clear_cache()
        return jsonify({'success': True, 'written': written, 'results': results})
    except Exception as e:
        import traceback
        return jsonify({'error': str(e) + '\n' + traceback.format_exc()})


@app.route('/api/status')
def api_status():
    from rules_excel import discover_rule_workbooks, summarize_templates_for_api

    return jsonify({
        'wps_file_configured': bool(WPS_DBSHEET_FILE_ID),
        'wps_sheets': {k: (v > 0) for k, v in WPS_SHEETS.items()},
        'rules_excel_files': discover_rule_workbooks(),
        'excel_templates': summarize_templates_for_api(),
    })


@app.route('/data')
def data_page():
    return (
        '<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"/>'
        '<title>系统状态</title><style>body{font-family:Microsoft YaHei;padding:24px;}'
        'a{color:#1a73e8}</style></head><body>'
        '<h2>绩效报工 · 数据与状态</h2>'
        '<p>多维表写入与规则拉取依赖 WPS 配置。以下为接口状态（JSON）：</p>'
        '<pre id="j">加载中…</pre>'
        '<p><a href="/form">返回报工</a></p>'
        '<script>fetch("/api/status").then(r=>r.json()).then(j=>{document.getElementById("j").textContent=JSON.stringify(j,null,2);});</script>'
        '</body></html>'
    )


# ========== 分组管理页面 ==========
@app.route('/groups')
def groups_page():
    return '''
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>⚙️ 人员分组管理</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: "Microsoft YaHei", sans-serif; background: #f0f2f5; padding: 20px; }
        .container { max-width: 900px; margin: 0 auto; }
        .header {
            background: linear-gradient(135deg,#1a73e8,#1557b0);
            color: white; padding: 20px 28px; border-radius: 14px 14px 0 0;
        }
        .header h1 { font-size: 20px; }
        .card { background: white; padding: 20px 24px; border-radius: 0 0 14px 14px; 
                box-shadow: 0 2px 12px rgba(0,0,0,0.08); margin-bottom: 16px; }
        .layout-2col { display: flex; gap: 16px; }
        .col-left { width: 220px; flex-shrink: 0; }
        .col-right { flex: 1; min-width: 0; }
        .panel { border: 1px solid #e8e8e8; border-radius: 10px; overflow: hidden; }
        .panel-title { background: #f8f9fa; padding: 12px 14px; font-weight: 600; font-size: 14px; border-bottom: 1px solid #e8e8e8; }
        .panel-body { padding: 10px; }
        .group-item {
            padding: 10px 12px; border-radius: 6px; cursor: pointer;
            font-size: 14px; margin-bottom: 3px; transition: all 0.15s;
            display: flex; justify-content: space-between; align-items: center;
        }
        .group-item:hover { background: #e8f0fe; }
        .group-item.active { background: #1a73e8; color: white; }
        .group-item .badge { background: #ddd; border-radius: 10px; padding: 0 8px; font-size: 12px; }
        .group-item.active .badge { background: rgba(255,255,255,0.3); color: white; }
        .group-item .del-btn { color: #ea4335; cursor: pointer; font-size: 16px; opacity: 0.4; }
        .group-item .del-btn:hover { opacity: 1; }
        .group-item.active .del-btn { color: white; }
        
        input, select { padding: 8px 12px; border: 1.5px solid #ddd; border-radius: 6px; font-size: 14px; width: 100%; }
        input:focus { outline: none; border-color: #1a73e8; }
        .btn { padding: 8px 16px; border: none; border-radius: 6px; cursor: pointer; font-size: 13px; font-weight: 500; }
        .btn-primary { background: #1a73e8; color: white; }
        .btn-danger { background: #ea4335; color: white; }
        .btn-outline { background: transparent; border: 1.5px solid #1a73e8; color: #1a73e8; }
        .btn-sm { padding: 4px 10px; font-size: 12px; }
        .form-row { display: flex; gap: 8px; margin-bottom: 10px; align-items: center; }
        .member-row {
            display: flex; gap: 8px; align-items: center; padding: 8px 0;
            border-bottom: 1px solid #f0f0f0;
        }
        .member-row:last-child { border-bottom: none; }
        .member-row .name { flex: 1; font-weight: 500; }
        .member-row .coeff { width: 80px; text-align: center; }
        .mt-12 { margin-top: 12px; }
        .empty { text-align: center; color: #999; padding: 30px; }
        @media (max-width: 640px) { .layout-2col { flex-direction: column; } .col-left { width: 100%; } }
    </style>
</head>
<body>
<div class="container">
    <div class="header">
        <h1>⚙️ 人员分组管理</h1>
        <p style="opacity:0.85;margin-top:4px;font-size:13px;">像微信群一样管理人员分组，报工时一键选择整组人员</p>
    </div>
    
    <div class="card">
        <div class="layout-2col">
            <!-- 左栏：分组列表 -->
            <div class="col-left">
                <div class="panel">
                    <div class="panel-title">📁 分组列表</div>
                    <div class="panel-body">
                        <div class="form-row">
                            <input type="text" id="newGroupName" placeholder="新建分组名" style="flex:1;" />
                            <button class="btn btn-primary btn-sm" onclick="createGroup()">新建</button>
                        </div>
                        <div id="groupListPanel"></div>
                    </div>
                </div>
            </div>
            
            <!-- 右栏：成员管理 -->
            <div class="col-right">
                <div class="panel">
                    <div class="panel-title" id="currentGroupTitle">👤 请选择分组</div>
                    <div class="panel-body">
                        <div id="memberPanel">
                            <div class="empty">← 从左侧选择一个分组</div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        <div style="text-align:center;margin-top:12px;">
            <a href="/form" style="color:#1a73e8;font-size:13px;">← 返回报工表单</a>
        </div>
    </div>
</div>

<datalist id="workerList"></datalist>

<script>
let currentGroup = '';
let allWorkers = [];
let allGroups = {};

window.onload = function() {
    loadAll();
};

function loadAll() {
    fetch('/api/workers').then(r=>r.json()).then(data => {
        allWorkers = data.workers || [];
        document.getElementById('workerList').innerHTML = allWorkers.map(w => '<option value="'+w+'">').join('');
    });
    fetchGroups();
}

function fetchGroups() {
    fetch('/api/groups').then(r=>r.json()).then(data => {
        allGroups = data.groups || {};
        renderGroupList();
        if (currentGroup && allGroups[currentGroup]) {
            renderMembers(currentGroup);
        }
    });
}

function renderGroupList() {
    const panel = document.getElementById('groupListPanel');
    const keys = Object.keys(allGroups);
    if (keys.length === 0) {
        panel.innerHTML = '<div style="font-size:13px;color:#999;padding:10px;text-align:center;">暂无分组</div>';
        return;
    }
    panel.innerHTML = keys.map(k => 
        '<div class="group-item' + (k === currentGroup ? ' active' : '') + '" onclick="selectGroup(\\''+k+'\\')">' +
        '<span>' + k + ' <span class="badge">' + (allGroups[k] ? allGroups[k].length : 0) + '</span></span>' +
        '<span class="del-btn" onclick="event.stopPropagation();deleteGroup(\\''+k+'\\')">×</span>' +
        '</div>'
    ).join('');
}

function selectGroup(name) {
    currentGroup = name;
    renderGroupList();
    renderMembers(name);
}

function renderMembers(groupName) {
    const members = allGroups[groupName] || [];
    document.getElementById('currentGroupTitle').textContent = '👤 ' + groupName + '（' + members.length + '人）';
    
    const panel = document.getElementById('memberPanel');
    
    let html = '<div class="form-row">';
    html += '<input type="text" id="newMemberName" list="workerList" placeholder="输入姓名添加" style="flex:1;" />';
    html += '<input type="number" id="newMemberCoeff" value="1.0" min="0.5" max="2.0" step="0.01" style="width:80px;" />';
    html += '<button class="btn btn-primary btn-sm" onclick="addMember(\\''+groupName+'\\')">添加</button>';
    html += '</div>';
    
    if (members.length === 0) {
        html += '<div class="empty">该分组暂无人员，输入姓名添加</div>';
    } else {
        html += '<div style="margin-top:8px;">';
        members.forEach((m, idx) => {
            html += '<div class="member-row">' +
                '<span class="name">' + (idx+1) + '. ' + m.name + '</span>' +
                '<input type="number" class="coeff" value="' + (m.coeff || 1.0) + '" min="0.5" max="2.0" step="0.01" ' +
                'onchange="updateCoeff(\\''+groupName+'\\',\\''+m.name+'\\',this.value)" />' +
                '<button class="btn btn-danger btn-sm" onclick="removeMember(\\''+groupName+'\\',\\''+m.name+'\\')">✕</button>' +
                '</div>';
        });
        html += '</div>';
    }
    
    panel.innerHTML = html;
}

function createGroup() {
    const name = document.getElementById('newGroupName').value.trim();
    if (!name) { alert('请输入分组名'); return; }
    
    fetch('/api/save_group', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({name, members: []}),
    }).then(r=>r.json()).then(res => {
        if (res.success) {
            document.getElementById('newGroupName').value = '';
            fetchGroups();
            currentGroup = name;
        }
    });
}

function deleteGroup(name) {
    if (!confirm('确定删除分组「'+name+'」？人员不会被删除')) return;
    fetch('/api/delete_group', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({name}),
    }).then(r=>r.json()).then(res => {
        if (res.success) {
            if (currentGroup === name) currentGroup = '';
            fetchGroups();
            if (!currentGroup) document.getElementById('memberPanel').innerHTML = '<div class="empty">← 从左侧选择一个分组</div>';
        }
    });
}

function addMember(groupName) {
    const nameInput = document.getElementById('newMemberName');
    const coeffInput = document.getElementById('newMemberCoeff');
    const name = nameInput.value.trim();
    const coeff = parseFloat(coeffInput.value) || 1.0;
    if (!name) { alert('请输入姓名'); return; }
    
    fetch('/api/add_member_to_group', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({group: groupName, member: {name, coeff}}),
    }).then(r=>r.json()).then(res => {
        if (res.success) {
            nameInput.value = '';
            nameInput.focus();
            fetchGroups();
        } else {
            alert(res.error);
        }
    });
}

function removeMember(groupName, name) {
    if (!confirm('从「'+groupName+'」移除 '+name+'？')) return;
    fetch('/api/remove_member_from_group', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({group: groupName, name}),
    }).then(r=>r.json()).then(res => {
        if (res.success) fetchGroups();
    });
}

function updateCoeff(groupName, name, coeff) {
    fetch('/api/add_member_to_group', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({group: groupName, member: {name, coeff: parseFloat(coeff) || 1.0}}),
    }).then(r=>r.json());
}
</script>
</body>
</html>
'''

# ========== 启动 ==========
if __name__ == '__main__':
    port = int(os.environ.get('PORT', '5000'))
    print("=" * 55)
    print("  🏭 绩效报工 Web 表单系统 v2.0 — WPS 多维表 + Waitress")
    print("=" * 55)
    print(f"  📅 启动时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  🔗 报工表单: http://0.0.0.0:{port}/form")
    print(f"  ⚙️ 分组管理: http://0.0.0.0:{port}/groups")
    print(f"  ⏰ 工时单价: {HOUR_SCORE}分/小时/人")
    print("  📌 WPS：请设置环境变量 WPS_DBSHEET_FILE_ID 与各 WPS_SHEET_*（整数 sheet_id）。")
    print("     若列举/写入报 invalid_scope，请在开放平台为本应用开通多维表权限后设置 WPS365_OAUTH_SCOPE=kso.dbsheet.readwrite")
    print("=" * 55)
    try:
        from waitress import serve
        serve(app, host='0.0.0.0', port=port, threads=16, channel_timeout=300)
    except ImportError:
        app.run(host='0.0.0.0', port=port, debug=False, threaded=True)
