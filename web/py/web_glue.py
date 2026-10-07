# -*- coding: utf-8 -*-
"""网页版胶合层：以 JSON 字符串为接口，把 src/core 的核心逻辑暴露给前端 JS。

运行于 Pyodide 中。question_bank 原本写入用户目录 JSON 文件，
此处将读写重定向到浏览器 localStorage（键名前缀 srh_web_）。
"""

import json
import re

import js

from srh.core import question_bank as qb
from srh.core.parser import mask_blank_question_text
from srh.core.session import weighted_random_pick

_RECORDS_KEY = 'srh_web_records'
_EDITS_KEY = 'srh_web_edits'
_STATE_KEY = 'srh_web_state'

_TYPE_NAMES = {'single': '单选', 'multi': '多选', 'judge': '判断', 'blank': '填空', 'short': '简答'}


def _ls_load(key, default):
    raw = js.localStorage.getItem(key)
    if not raw:
        return default
    try:
        return json.loads(raw)
    except Exception:
        return default


def _ls_save(key, obj):
    js.localStorage.setItem(key, json.dumps(obj, ensure_ascii=False))


def _install_storage_override():
    qb.load_records = lambda: _ls_load(_RECORDS_KEY, {})
    qb.save_records = lambda records: _ls_save(_RECORDS_KEY, records)
    qb.load_manual_question_edits = lambda: _ls_load(_EDITS_KEY, {})
    qb.save_manual_question_edits = lambda edits: _ls_save(_EDITS_KEY, edits)
    qb.load_app_state = lambda: _ls_load(_STATE_KEY, {})
    qb.save_app_state = lambda state: _ls_save(_STATE_KEY, state)


_install_storage_override()

_state = {
    'questions': [],
    'by_id': {},
    'records': {},
    'current': None,
    'submitted': False,
    'answer_revealed': False,
    'recent_signatures': [],
    'duplicate_groups': [],
    'duplicate_sig_set': set(),
    'recent_limit': 6,
}


def _question_signature(q):
    """题目归一签名，用于重复题检测（与桌面版逻辑一致）。"""
    text = str(q.get('text', '') or '')
    text = re.sub(r'（\s*\d+\s*）\s*[_＿﹍]+', '（）', text)
    text = re.sub(r'[_＿﹍]+', '', text)
    text = re.sub(r'[\s，,。；;：:、（）()\[\]【】]+', '', text)

    options = q.get('options') or {}
    option_sig = []
    for k in sorted(options.keys()):
        v = re.sub(r'\s+', '', str(options.get(k, '') or ''))
        option_sig.append(f'{k}:{v}')

    return (q.get('type', ''), text, '|'.join(option_sig))


def _build_duplicate_state(questions):
    sig_map = {}
    for q in questions:
        sig_map.setdefault(_question_signature(q), []).append(q['id'])

    groups = [ids for ids in sig_map.values() if len(ids) >= 2]
    groups.sort(key=lambda x: (-len(x), x[0]))
    _state['duplicate_groups'] = groups
    _state['duplicate_sig_set'] = {sig for sig, ids in sig_map.items() if len(ids) >= 2}


def _is_recent_duplicate_pick(q):
    sig = _question_signature(q)
    return sig in _state['duplicate_sig_set'] and sig in _state['recent_signatures']


def _record_view(rec):
    return {'attempts': int(rec.get('attempts', 0) or 0), 'errors': int(rec.get('errors', 0) or 0)}


def _answer_text(q):
    ans = q.get('answer')
    if isinstance(ans, list):
        ans = ''.join(ans)
    return str(ans or '')


def _dump_question(q):
    display_text = q.get('text', '')
    if q.get('type') == 'blank':
        display_text = mask_blank_question_text(q.get('text', ''), q.get('answer', ''))
    return {
        'id': q.get('id'),
        'type': q.get('type'),
        'type_name': _TYPE_NAMES.get(q.get('type'), q.get('type', '')),
        'text': str(q.get('text', '') or ''),
        'display_text': str(display_text or ''),
        'options': {k: str(v) for k, v in (q.get('options') or {}).items()},
        'has_answer': bool(q.get('answer')),
        'record': _record_view(qb.get_record(_state['records'], q)),
    }


def api_init(bank_json):
    """初始化：加载题库、历史记录并应用手动修改。"""
    data = json.loads(bank_json)
    questions = list(data.get('questions') or [])
    edits = qb.load_manual_question_edits()
    for q in questions:
        qb.ensure_question_identity_fields(q)
    qb.apply_manual_question_edits(questions, edits)

    _state['questions'] = questions
    _state['by_id'] = {q['id']: q for q in questions}
    _state['records'] = qb.load_records()
    _build_duplicate_state(questions)

    dup_questions = sum(len(g) for g in _state['duplicate_groups'])
    return json.dumps({
        'title': str(data.get('title', '')),
        'total': len(questions),
        'duplicate_groups': len(_state['duplicate_groups']),
        'duplicate_questions': dup_questions,
    })


def api_pick_next():
    """加权抽题（含近期重复题规避），返回题目数据。"""
    _state['submitted'] = False
    _state['answer_revealed'] = False

    questions = _state['questions']
    records = _state['records']
    picked = weighted_random_pick(questions, records)

    if len(questions) > 1 and _state['duplicate_groups']:
        for _ in range(30):
            if not _is_recent_duplicate_pick(picked):
                break
            candidate = weighted_random_pick(questions, records)
            if not _is_recent_duplicate_pick(candidate):
                picked = candidate
                break
            picked = candidate

    _state['current'] = picked
    sig = _question_signature(picked)
    _state['recent_signatures'].append(sig)
    limit = _state['recent_limit']
    if len(_state['recent_signatures']) > limit:
        _state['recent_signatures'] = _state['recent_signatures'][-limit:]

    return json.dumps(_dump_question(picked))


def api_submit_objective(selected_json):
    """客观题判分并记录。"""
    q = _state['current']
    selected = set(json.loads(selected_json or '[]'))
    correct = set(q.get('answer') or [])
    is_correct = selected == correct

    qb.update_record(_state['records'], q, is_correct)
    _state['submitted'] = True

    return json.dumps({
        'correct': bool(is_correct),
        'answer': ''.join(sorted(correct)),
        'record': _record_view(qb.get_record(_state['records'], q)),
    })


def api_reveal_answer():
    """主观题：展示参考答案（进入自评阶段）。"""
    q = _state['current']
    _state['answer_revealed'] = True
    return json.dumps({'answer': _answer_text(q)})


def api_record_subjective(is_correct):
    """主观题：记录自评结果。"""
    q = _state['current']
    qb.update_record(_state['records'], q, bool(is_correct))
    _state['submitted'] = True
    return json.dumps({'record': _record_view(qb.get_record(_state['records'], q))})


def api_overall():
    """顶部进度：已做题目数 / 总答题数 / 总正确率。"""
    records = _state['records']
    attempted = 0
    total_attempts = 0
    total_errors = 0
    for q in _state['questions']:
        rec = qb.get_record(records, q)
        attempts = int(rec.get('attempts', 0) or 0)
        errors = int(rec.get('errors', 0) or 0)
        if attempts > 0:
            attempted += 1
        total_attempts += attempts
        total_errors += errors

    accuracy = (1 - total_errors / total_attempts) * 100 if total_attempts > 0 else 0.0
    return json.dumps({
        'total': len(_state['questions']),
        'attempted': attempted,
        'attempts': total_attempts,
        'accuracy': round(accuracy, 1),
    })


def api_stats():
    """考频统计：汇总与逐题行数据。"""
    records = _state['records']
    rows = []
    attempted_count = 0
    total_attempts = 0
    total_errors = 0

    for q in _state['questions']:
        rec = qb.get_record(records, q)
        attempts = int(rec.get('attempts', 0) or 0)
        errors = int(rec.get('errors', 0) or 0)
        if attempts > 0:
            attempted_count += 1
        total_attempts += attempts
        total_errors += errors
        rows.append({
            'id': q.get('id'),
            'type': _TYPE_NAMES.get(q.get('type'), q.get('type', '')),
            'attempts': attempts,
            'errors': errors,
            'error_rate': (errors / attempts * 100) if attempts > 0 else 0.0,
            'text': str(q.get('text', '')).replace('\n', ' ').strip(),
        })

    overall_rate = (total_errors / total_attempts * 100) if total_attempts > 0 else 0.0
    return json.dumps({
        'summary': {
            'total': len(_state['questions']),
            'attempted': attempted_count,
            'attempts': total_attempts,
            'errors': total_errors,
            'error_rate': overall_rate,
        },
        'duplicate_groups': len(_state['duplicate_groups']),
        'duplicate_questions': sum(len(g) for g in _state['duplicate_groups']),
        'rows': rows,
    })


def api_detail(qid):
    """题目详情（供统计表点击查看）。"""
    q = _state['by_id'].get(int(qid))
    if not q:
        return json.dumps(None)
    options = q.get('options') or {}
    return json.dumps({
        'id': q.get('id'),
        'type': _TYPE_NAMES.get(q.get('type'), q.get('type', '')),
        'text': str(q.get('text', '') or ''),
        'options': [{'key': k, 'text': str(options[k])} for k in sorted(options.keys())],
        'answer': _answer_text(q),
        'record': _record_view(qb.get_record(_state['records'], q)),
    })


def api_reset_records():
    """清空全部作答记录。"""
    _state['records'] = {}
    qb.save_records(_state['records'])
    return json.dumps({'ok': True})
