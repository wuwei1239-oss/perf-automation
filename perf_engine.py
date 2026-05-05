# -*- coding: utf-8 -*-
"""
绩效计算引擎（第一性原理，与 04月份绩效.xlsx / 绩效规则库.xlsx 语义对齐）

数学模型（每批 b）：
  基准分_b = T_b / D_b   （T_b=该批固定总分，D_b=规则基础定员）
  每人正常得分（系数分）累加 += 基准分_b × 个人系数
  缺员总池_b = max(0, D_b - nn) × 基准分_b = (D_b-nn)×T_b/D_b
  每人缺员补偿（加减分，不乘系数）+= 缺员总池_b / nn

其它加减分（不乘系数，来自表单勾选 + 规则）：大清、过筛、装机、手填缺员分、小时缺员、
粉碎、转片、模具、捡药、质量分（可正可负）。扣减：缺员人员按小时×单价。

04 月模板侧栏语义（process_performance）：奖罚说明/备注中的「缺员、质量分、请假、护理假、
休息、返回原…」等由组长在网页侧对应字段录入；非生产行在批量导入场景下可另处理。
"""
from __future__ import annotations

import re
from typing import Any, Dict, List, Optional, Tuple

QIAN_GONGXU = ('配料', '混合', '预混', '制粒', '压片', '包衣')
NEIWAI_BAO = ('内包', '外包')


def find_prefix(batch_no: str, rules: dict) -> Optional[str]:
    if not batch_no:
        return None
    digits = re.sub(r'\D', '', str(batch_no))
    if not digits:
        return None
    for l in (4, 3, 2):
        if len(digits) >= l:
            p = digits[:l]
            if p in rules:
                return p
    s = str(batch_no).strip()
    for l in (4, 3, 2):
        if len(s) >= l:
            p = s[:l]
            if p in rules:
                return p
    return None


def parse_batch_count(batch_str: Any) -> Tuple[str, int]:
    if not batch_str:
        return batch_str, 1
    s = str(batch_str).strip()
    m = re.search(r'[（(](\d+)\s*批[)）]', s)
    if m:
        cnt = int(m.group(1))
        s_clean = re.sub(r'[（(].*?[)）]', '', s).strip()
        return s_clean, cnt
    s = s.replace('批', '')
    if '-' in s:
        parts = s.split('-', 1)
        l, r = parts[0].strip(), parts[1].strip()
        ld = re.sub(r'\D', '', l)
        rd = re.sub(r'\D', '', r)
        if ld and rd:
            try:
                rv = int(rd)
                lv = int(ld[-len(rd) :]) if len(ld) >= len(rd) else int(ld)
                cnt = (rv - lv + 1) if rv >= lv else 1
                return ld, cnt
            except Exception:
                pass
        return ld if ld else l, 1
    parts = re.split(r'[,，+、\s]+', s)
    cnt = len([p for p in parts if p.strip()])
    return parts[0].strip() if parts else s, max(cnt, 1)


def _clean_num(s: str) -> str:
    return (
        str(s)
        .replace('分/批', '')
        .replace('分|批', '')
        .replace('分|次', '')
        .replace('分/次', '')
        .replace('分/筐', '')
        .replace('分|筐', '')
        .replace('分/人/次', '')
        .strip()
    )


def _scan_rule_float(rule: dict, substrings: Tuple[str, ...]) -> float:
    """在规则行中按列名子串模糊匹配取第一个可解析数字（补全列名不一致）。"""
    for k, v in rule.items():
        if not isinstance(k, str):
            continue
        if all(t in k for t in substrings):
            try:
                return float(_clean_num(str(v))) if v not in (None, '') else 0.0
            except Exception:
                continue
    return 0.0


def get_process_scores(rule: dict, process: str, chong_type: Optional[str] = None) -> dict:
    result: Dict[str, Any] = {'base': 0.0, 'daqing': 0.0, 'ding': None, 'zhuangji': 0.0, 'guoshao': 0.0}

    field_map = {
        '配料': {
            'base': ('配料基础分',),
            'daqing': ('配料大清分',),
            'ding': ('配料定员（人）', '配料定员'),
            'guoshao': ('过筛分',),
        },
        '混合': {
            'base': ('混合基础分|批', '混合基础分'),
            'daqing': ('混合大清分|次', '混合大清分'),
            'ding': ('混合定员',),
        },
        '预混': {
            'base': ('预混基础分|批', '预混基础分'),
            'daqing': ('预混大清分|次', '预混大清分'),
            'ding': ('预混定员',),
        },
        '制粒': {
            'base': ('制粒基础分|批', '制粒基础分'),
            'daqing': ('制粒大清|次', '制粒大清分'),
            'ding': ('制粒定员（人）', '制粒定员'),
            'zhuangji': ('干法制粒装机|次', '干法制粒装机'),
        },
        '压片': {
            'base': ('压片基础分|批', '压片基础分'),
            'daqing': ('压片大清分|次', '压片大清|次'),
            'ding': ('压片定员（人）', '压片定员'),
        },
        '包衣': {
            'base': ('包衣基础分|批', '包衣基础分'),
            'daqing': ('包衣大清分|批', '包衣大清分|次', '包衣大清分'),
            'ding': ('包衣定员（人）', '包衣定员'),
        },
        '内包': {
            'base': ('内包基础分',),
            'daqing': ('内包大清分|次', '内包大清分'),
            'ding': ('内包定员（人）', '内包定员'),
            'jianyao': ('捡药分|筐', '捡药分'),
        },
        '外包': {
            'base': ('外包基础分',),
            'daqing': ('外包大清分|次', '外包大清分'),
            'ding': ('外包定员（人）', '外包定员'),
            'jianyao': ('捡药分|筐', '捡药分'),
        },
    }

    fm = field_map.get(process, {})
    if not fm:
        return result

    def pick(candidates: Tuple[str, ...]) -> Optional[str]:
        for c in candidates:
            if c in rule:
                return c
        return None

    bf = pick(fm.get('base', ()))
    if bf:
        try:
            v = _clean_num(str(rule[bf]))
            if '基础分=' in v:
                result['base_formula'] = str(rule[bf])
                result['base'] = 0.0
            else:
                result['base'] = float(v) if v else 0.0
        except Exception:
            result['base'] = 0.0

    df = pick(fm.get('daqing', ()))
    if df:
        try:
            result['daqing'] = float(_clean_num(str(rule[df]))) if rule[df] not in (None, '') else 0.0
        except Exception:
            result['daqing'] = 0.0
    elif process == '压片':
        result['daqing'] = _scan_rule_float(rule, ('压片', '大清'))

    df_ = pick(fm.get('ding', ()))
    if df_:
        try:
            dv = str(rule[df_]).strip()
            m = re.search(r'(\d+)', dv)
            if m:
                result['ding'] = int(m.group(1))
        except Exception:
            pass

    if process == '配料':
        if '过筛分' in rule:
            try:
                result['guoshao'] = float(str(rule['过筛分']).strip()) if rule['过筛分'] not in (None, '') else 0.0
            except Exception:
                result['guoshao'] = 0.0

    zf = pick(fm.get('zhuangji', ())) if 'zhuangji' in fm else None
    if zf and zf in rule:
        try:
            result['zhuangji'] = float(str(rule[zf]).strip()) if rule[zf] not in (None, '') else 0.0
        except Exception:
            pass

    if process == '压片' and chong_type:
        ck = {'77': '压片大清分77冲|次', '65': '压片大清分65冲|次', '40': '压片大清分40冲|次'}
        alt = ck.get(chong_type)
        if alt and alt in rule:
            try:
                result['zhuangji'] = float(str(rule[alt]).strip()) if rule[alt] not in (None, '') else 0.0
            except Exception:
                pass

    if process in ('内包', '外包') and 'jianyao' in fm:
        jf = pick(fm['jianyao'])
        if jf:
            try:
                result['jianyao_per_box'] = float(_clean_num(str(rule[jf]))) if rule[jf] not in (None, '') else 0.0
            except Exception:
                result['jianyao_per_box'] = 0.0

    return result


def _ledger_batches(
    batches: List[dict],
    rules: dict,
    proc: str,
    chong: str,
    hours: float,
    hour_score: float,
    nn: int,
    has_dq: bool,
    has_gs: bool,
    has_zj: bool,
    ceshi: bool = False,
) -> Tuple[List[dict], float, float, float]:
    """返回 batch_rows, total_daqing(仅勾选大清时累加), gs_score, zj_sum。
    ceshi=True: 试验品种(无规则批)也按工时考核; ceshi=False: 跳过无规则批。
    """
    batch_rows: List[dict] = []
    total_daqing = 0.0
    gs_score = 0.0
    zj_sum = 0.0

    for batch in batches:
        batch_no = batch.get('batch_no', '')
        manual_code = batch.get('product_code', '')
        box = float(batch.get('box', 0))
        ban = float(batch.get('ban', 0))
        prefix = manual_code if manual_code else find_prefix(batch_no, rules)
        rule = rules.get(prefix) if prefix else None
        if rule:
            scores = get_process_scores(rule, proc, chong)
            if scores.get('base_formula'):
                try:
                    expr = scores['base_formula'].replace('基础分=', '')
                    expr = expr.replace('实际生产箱数', str(box))
                    expr = expr.replace('箱数', str(box))
                    expr = expr.replace('实际板数', str(ban))
                    expr = expr.replace('板数', str(ban))
                    T_b = float(eval(expr))
                except Exception:
                    T_b = float(scores.get('base') or 0)
            else:
                bn = 1
                if proc in QIAN_GONGXU:
                    _, bn = parse_batch_count(batch_no)
                T_b = float(scores.get('base') or 0) * bn
            D_b = int(scores['ding']) if scores.get('ding') else nn
            if D_b <= 0:
                D_b = nn
            batch_rows.append({'T': T_b, 'D': D_b, 'batch_no': batch_no, 'scores': scores})
            if has_dq:
                total_daqing += float(scores.get('daqing') or 0)
            if has_gs and scores.get('guoshao'):
                _, bn = parse_batch_count(batch_no)
                gs_score += float(scores['guoshao']) * bn
            if has_zj:
                zj_sum += float(scores.get('zhuangji') or 0)
        else:
            if ceshi:
                T_b = hours * hour_score
            else:
                T_b = 0.0
            D_b = nn if nn > 0 else 1
            batch_rows.append({'T': T_b, 'D': D_b, 'batch_no': batch_no, 'scores': {}})

    if not batch_rows and nn > 0:
        T_fallback = hours * hour_score if ceshi else 0.0
        batch_rows.append({'T': T_fallback, 'D': nn, 'batch_no': '', 'scores': {}})

    return batch_rows, total_daqing, gs_score, zj_sum


def calc_perf(form_data: dict, rules: dict, hour_score: float = 20.0) -> List[dict]:
    proc = form_data.get('工序', '')
    date_str = form_data.get('日期', '')
    hours = float(form_data.get('工时', 8))
    batches = form_data.get('batches', [])
    persons = form_data.get('persons', [])
    names = [p.get('name', '') for p in persons if p.get('name')]
    coeffs = [float(p.get('coeff', 1.0)) for p in persons if p.get('name')]
    quality_map = {p['name']: float(p.get('quality', 0)) for p in persons if p.get('name')}
    if not names:
        return []
    nn = len(names)

    flags = {
        'has_dq': bool(form_data.get('hasDQ')),
        'has_gs': bool(form_data.get('hasGS')),
        'has_zj': bool(form_data.get('hasZJ')),
    }
    qy_total = 0.0  # 缺员补偿已自动计算，不再接受手动输入
    xqy_total = 0.0
    ceshi = bool(form_data.get('ceshi', False))  # 试验品种工时考核开关
    chong = form_data.get('chong', '') or ''
    qy_persons = form_data.get('qyPersons', [])
    zp_batches = float(form_data.get('zpBatches', 0))
    zp_persons = form_data.get('zpPersons', [])
    mg_change_persons = form_data.get('mgChangePersons', [])
    mg_confirm_persons = form_data.get('mgConfirmPersons', [])
    jy_baskets = float(form_data.get('jyBaskets', 0))
    jy_persons = form_data.get('jyPersons', [])
    fs_score = float(form_data.get('fsScore', 0))

    batch_rows, total_daqing_raw, gs_score, zj_sum = _ledger_batches(
        batches, rules, proc, chong, hours, hour_score, nn,
        flags['has_dq'], flags['has_gs'], flags['has_zj'], ceshi,
    )

    wc_by_name = {nm: 0.0 for nm in names}
    auto_short_by_name = {nm: 0.0 for nm in names}
    shortage_notes: List[str] = []

    for br in batch_rows:
        T, D = float(br['T']), int(br['D'])
        if D <= 0:
            D = nn
        if nn <= 0:
            continue
        jizhun = T / D
        for i, nm in enumerate(names):
            wc_by_name[nm] += jizhun * coeffs[i]
        if nn < D:
            gap = D - nn
            pool = gap * jizhun
            share = pool / nn
            for nm in names:
                auto_short_by_name[nm] += share
            shortage_notes.append('缺员%d人(%s)' % (gap, br.get('batch_no') or '本批'))

    pp_dq = (total_daqing_raw / nn) if (total_daqing_raw > 0 and nn) else 0.0
    pp_gs = (gs_score / nn) if (gs_score > 0 and nn) else 0.0
    pp_zj = (zj_sum / nn) if (zj_sum > 0 and nn) else 0.0
    qcomp = qy_total / nn if qy_total > 0 and nn else 0.0
    xqcomp = xqy_total / nn if xqy_total > 0 and nn else 0.0
    pp_fs = fs_score / nn if fs_score > 0 and nn else 0.0
    zp_total = 5.0 * zp_batches
    pp_zp = zp_total / len(zp_persons) if (zp_batches > 0 and zp_persons) else 0.0

    pp_jy = 0.0
    if jy_baskets > 0 and jy_persons:
        jy_price = 0.0
        for batch in batches:
            batch_no = batch.get('batch_no', '')
            prefix = find_prefix(batch_no, rules)
            rule = rules.get(prefix) if prefix else None
            if rule:
                scores = get_process_scores(rule, proc, chong)
                jy_price = float(scores.get('jianyao_per_box', 0) or 0)
                break
        if jy_price == 0:
            jy_price = 1.1
        pp_jy = (jy_baskets * jy_price) / len(jy_persons)

    ded = {}
    for qp in qy_persons:
        nm = qp.get('name', '')
        hours_q = float(qp.get('hours', 0))
        if nm and hours_q > 0:
            ded[nm] = hours_q * hour_score

    base_slot_sum = sum(float(br['T']) / max(int(br['D']), 1) for br in batch_rows) if batch_rows else 0.0

    results: List[dict] = []
    for i, name in enumerate(names):
        c = coeffs[i]
        wc = wc_by_name[name]
        # 奖罚明细（不包含系数的加减项）
        award_parts: List[str] = []
        nc = auto_short_by_name[name] + pp_dq + pp_gs + pp_zj + qcomp + xqcomp + pp_fs
        if name in zp_persons:
            nc += pp_zp
        if name in jy_persons:
            nc += pp_jy
        if name in mg_change_persons:
            nc += 30.0
        if name in mg_confirm_persons:
            nc += 10.0
        qual_score = quality_map.get(name, 0.0)
        nc += qual_score
        d = ded.get(name, 0.0)
        total = wc + nc - d

        batch_strs = []
        for b in batches:
            bn = b.get('batch_no', '')
            bx = float(b.get('box', 0))
            if proc in NEIWAI_BAO and bx > 0:
                batch_strs.append('%s（%d箱）' % (bn, int(bx)))
            else:
                batch_strs.append(str(bn))

        remark_parts: List[str] = []
        if proc in QIAN_GONGXU and batches:
            remark_parts.append('%s%s' % (proc, batches[0].get('batch_no', '')))
        elif proc in NEIWAI_BAO and batch_strs:
            remark_parts.append('%s%s' % (proc, ' + '.join(batch_strs)))
        else:
            for b in batches:
                bn = b.get('batch_no', '')
                if bn:
                    remark_parts.append('%s%s' % (proc, bn))
        if shortage_notes:
            remark_parts.append(' '.join(shortage_notes))

        auto_short = auto_short_by_name.get(name, 0.0)
        detail_parts = [
            '基准分Σ(T/D)=%.2f×系数%g' % (base_slot_sum, c),
        ]
        if auto_short > 0:
            detail_parts.append('缺员补偿(不乘系数)+%.2f' % auto_short)
        if pp_dq > 0:
            detail_parts.append('大清+%.1f' % pp_dq)
        if pp_gs > 0:
            detail_parts.append('过筛+%.1f' % pp_gs)
        if pp_zj > 0:
            detail_parts.append('装机+%.1f' % pp_zj)
        if qual_score != 0:
            detail_parts.append('质量%+s' % qual_score)
        if qcomp > 0:
            detail_parts.append('缺员补(手填)+%.1f' % qcomp)
        if xqcomp > 0:
            detail_parts.append('时缺补+%.1f' % xqcomp)
        if pp_fs > 0:
            detail_parts.append('粉碎+%.1f' % pp_fs)
        if name in zp_persons and pp_zp > 0:
            detail_parts.append('转片%d批+%.1f' % (int(zp_batches), pp_zp))
        if name in mg_change_persons:
            detail_parts.append('模具换+30')
        if name in mg_confirm_persons:
            detail_parts.append('模具确+10')
        if name in jy_persons and pp_jy > 0:
            detail_parts.append('捡药%d筐+%.1f' % (int(jy_baskets), pp_jy))
        if d > 0:
            detail_parts.append('扣%.0f' % d)

        # 构建奖罚说明
        if auto_short > 0:
            award_parts.append('缺员补+%.2f' % auto_short)
        if pp_dq > 0:
            award_parts.append('大清+%.1f' % pp_dq)
        if pp_gs > 0:
            award_parts.append('过筛+%.1f' % pp_gs)
        if pp_zj > 0:
            award_parts.append('装机+%.1f' % pp_zj)
        if qual_score != 0:
            award_parts.append('质量%+.1f' % qual_score)
        if qcomp > 0:
            award_parts.append('缺员手补+%.1f' % qcomp)
        if xqcomp > 0:
            award_parts.append('时缺补+%.1f' % xqcomp)
        if pp_fs > 0:
            award_parts.append('粉碎+%.1f' % pp_fs)
        if name in zp_persons and pp_zp > 0:
            award_parts.append('转片%d批+%.1f' % (int(zp_batches), pp_zp))
        if name in mg_change_persons:
            award_parts.append('换模具+30')
        if name in mg_confirm_persons:
            award_parts.append('模具确认+10')
        if name in jy_persons and pp_jy > 0:
            award_parts.append('捡药%d筐+%.1f' % (int(jy_baskets), pp_jy))
        if d > 0:
            award_parts.append('扣%.0f' % d)

        results.append(
            {
                '日期': date_str,
                '姓名': name,
                '工序': proc,
                '系数': c,
                '批号': ','.join([b.get('batch_no', '') for b in batches]),
                '批量': ','.join(['%d箱' % int(b.get('box', 0)) for b in batches if b.get('box', 0) > 0]),
                '基础分': round(base_slot_sum, 2),
                '生产得分': round(wc, 2),
                '奖罚说明': '; '.join(award_parts) if award_parts else '无',
                '大清': round(pp_dq, 2),
                '过筛': round(pp_gs, 2),
                '装机': round(pp_zj, 2),
                '缺员补': round(qcomp, 2),
                '自动缺员补': round(auto_short, 2),
                '时缺补': round(xqcomp, 2),
                '质量': qual_score,
                '粉碎': round(pp_fs, 2),
                '转片': round(pp_zp if name in zp_persons else 0, 2),
                '模具换': 30 if name in mg_change_persons else 0,
                '模具确': 10 if name in mg_confirm_persons else 0,
                '捡药': round(pp_jy if name in jy_persons else 0, 2),
                '扣': round(d, 2),
                '系数分': round(wc, 2),
                '加减分': round(nc, 2),
                '总分': round(total, 2),
                '备注': '; '.join(detail_parts),
                '显示备注': ' '.join(remark_parts),
            }
        )

    return results
