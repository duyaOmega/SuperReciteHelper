#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""应用状态、题目身份键、手动修改与做题记录持久化模块。"""

import tkinter as tk
from tkinter import messagebox, font as tkfont, filedialog, ttk
import re
import json
import os
import random
import math
import hashlib
from datetime import datetime
import ctypes
import tempfile
import zipfile
import xml.etree.ElementTree as ET

APP_NAME = 'SuperReciteHelper'

def _get_app_storage_dir():
    """获取可执行程序友好的持久化目录。"""
    if os.name == 'nt':
        base = os.getenv('APPDATA') or os.path.expanduser('~')
    else:
        base = os.path.join(os.path.expanduser('~'), '.local', 'share')

    path = os.path.join(base, APP_NAME)
    os.makedirs(path, exist_ok=True)
    return path

APP_STORAGE_DIR = _get_app_storage_dir()

RECORD_FILE = os.path.join(APP_STORAGE_DIR, 'error_record.json')

STATE_FILE = os.path.join(APP_STORAGE_DIR, 'app_state.json')

QUESTION_EDITS_FILE = os.path.join(APP_STORAGE_DIR, 'question_edits.json')

def load_app_state():
    """读取应用级状态（如最近打开文件），读取失败时返回空字典。"""
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if isinstance(data, dict):
                    return data
        except Exception:
            pass
    return {}

def save_app_state(state):
    """保存应用级状态到本地 JSON 文件。"""
    try:
        with open(STATE_FILE, 'w', encoding='utf-8') as f:
            json.dump(state or {}, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def load_manual_question_edits():
    """读取用户手动编辑过的题目覆盖数据。"""
    if os.path.exists(QUESTION_EDITS_FILE):
        try:
            with open(QUESTION_EDITS_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if isinstance(data, dict):
                    return data
        except Exception:
            pass
    return {}

def save_manual_question_edits(edits):
    """持久化题目手动修改映射。"""
    try:
        with open(QUESTION_EDITS_FILE, 'w', encoding='utf-8') as f:
            json.dump(edits or {}, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def _ensure_question_identity_fields(question):
    """确保题目具备稳定键与原始内容快照。"""
    # _base_key: 由“原始题干+选项+题型”计算出的稳定主键，手动编辑前后都不变。
    if '_base_key' not in question or not question.get('_base_key'):
        question['_base_key'] = _build_question_record_key(question)
    # _record_key: 当前用于答题记录统计的键，默认与 _base_key 一致。
    if '_record_key' not in question or not question.get('_record_key'):
        question['_record_key'] = question['_base_key']

    if '_orig_text' not in question:
        question['_orig_text'] = str(question.get('text', '') or '')
    if '_orig_answer' not in question:
        if isinstance(question.get('answer'), list):
            question['_orig_answer'] = list(question.get('answer', []))
        else:
            question['_orig_answer'] = str(question.get('answer', '') or '')
    if '_orig_type' not in question:
        question['_orig_type'] = str(question.get('type', '') or '')
    if '_orig_options' not in question:
        src_options = question.get('options') or {}
        question['_orig_options'] = dict(src_options)

def apply_manual_question_edits(questions, edits):
    """将持久化编辑应用到当前题集（严格按 base_key 匹配）。"""
    for q in questions:
        _ensure_question_identity_fields(q)
        # base_key: 当前题目的稳定身份标识，用于命中 edits 中的覆盖数据。
        base_key = q.get('_base_key')
        if not base_key or base_key not in edits:
            continue

        payload = edits.get(base_key) or {}
        if 'type' in payload and payload.get('type'):
            q['type'] = str(payload.get('type'))
        if 'options' in payload and isinstance(payload.get('options'), dict):
            q['options'] = dict(payload.get('options') or {})
        if 'text' in payload:
            q['text'] = str(payload.get('text', '') or '')
        if 'answer' in payload:
            q['answer'] = payload.get('answer')

def upsert_manual_question_edit(edits, question):
    """写入某题的手动修改（以原始 base_key 为主键）。"""
    _ensure_question_identity_fields(question)
    # key: 手动编辑记录的主键，固定使用 _base_key 防止题号变化导致丢记录。
    key = question['_base_key']

    payload = {
        'type': question.get('type', ''),
        'options': dict(question.get('options') or {}),
        'text': str(question.get('text', '') or ''),
        'answer': question.get('answer'),
        'orig_type': question.get('_orig_type', ''),
        'orig_options': dict(question.get('_orig_options') or {}),
        'orig_text': str(question.get('_orig_text', '') or ''),
        'orig_answer': question.get('_orig_answer', ''),
        'updated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'preview': str(question.get('text', '') or '').replace('\n', ' ').strip()[:80]
    }
    edits[key] = payload
    save_manual_question_edits(edits)
    return key

def _build_question_record_key(question):
    """生成稳定题目键：跨会话/跨题号可复用。"""
    q = question or {}
    text = re.sub(r'\s+', '', str(q.get('text', '') or ''))
    text = re.sub(r'[_＿﹍]+', '_', text)

    options = q.get('options') or {}
    option_items = []
    for k in sorted(options.keys()):
        v = re.sub(r'\s+', '', str(options.get(k, '') or ''))
        option_items.append(f'{k}:{v}')

    # payload: 参与哈希的最小身份数据，确保同一道题跨会话生成同一键。
    payload = {
        'type': str(q.get('type', '') or ''),
        'text': text,
        'options': option_items,
    }
    raw = json.dumps(payload, ensure_ascii=False, sort_keys=True)
    digest = hashlib.sha1(raw.encode('utf-8')).hexdigest()
    return f'q:{digest}'

def _record_key(q_or_id):
    """统一记录键入口：既支持题目对象，也支持题号字符串。"""
    if isinstance(q_or_id, dict):
        # 优先使用题目对象上的稳定键；缺失时即时回算，兼容旧数据。
        key = q_or_id.get('_record_key')
        if key:
            return key
        return _build_question_record_key(q_or_id)
    return str(q_or_id)

def load_records():
    """加载历史错误记录"""
    # 先读新路径；若不存在则尝试从旧版本路径迁移。
    if os.path.exists(RECORD_FILE):
        with open(RECORD_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)

    # 兼容旧版本：首次升级时迁移脚本目录下的记录文件。
    legacy_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'error_record.json')
    if os.path.exists(legacy_file):
        try:
            with open(legacy_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if isinstance(data, dict):
                save_records(data)
                return data
        except Exception:
            pass
    return {}

def save_records(records):
    """保存错误记录"""
    with open(RECORD_FILE, 'w', encoding='utf-8') as f:
        json.dump(records, f, ensure_ascii=False, indent=2)

def get_record(records, qid):
    """获取某题的记录"""
    # key: 统一后的记录索引键（可能是稳定哈希，也可能是旧版题号字符串）。
    key = _record_key(qid)
    if key in records:
        return records[key]

    # 兼容历史版本（按题号存储）
    if isinstance(qid, dict):
        legacy_id = qid.get('id')
        if legacy_id is not None and str(legacy_id) in records:
            return records[str(legacy_id)]

    return {'attempts': 0, 'errors': 0}

def update_record(records, qid, is_correct):
    """更新某题的记录"""
    # key: 当前这道题在 records 字典中的归档位置。
    key = _record_key(qid)
    if key not in records:
        records[key] = {'attempts': 0, 'errors': 0}
    records[key]['attempts'] += 1
    if not is_correct:
        records[key]['errors'] += 1
    save_records(records)
