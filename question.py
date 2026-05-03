#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""题目编辑对话框与答案格式化模块。"""

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

from parser import _extract_choice_answer, _normalize_answer_text

def _format_answer_text(answer_value):
    """将答案值统一格式化为可展示字符串。"""
    if isinstance(answer_value, list):
        return ''.join(answer_value)
    return str(answer_value or '')

def _parse_manual_answer_for_question(question, raw_text, target_type=None, options_override=None):
    """按题型校验并解析手动输入答案。"""
    # q_type: 本次按什么题型解析答案（可由 target_type 临时覆盖原题型）。
    q_type = target_type or question.get('type')
    # text: 用户在弹窗中输入的原始答案文本（去前后空白）。
    text = str(raw_text or '').strip()

    if q_type in ('blank', 'short'):
        if not text:
            return None, '主观题答案不能为空。'
        return text, None

    # options: 当前判题所依据的选项字典；编辑时优先使用“刚输入的新选项”。
    options = options_override if options_override is not None else (question.get('options') or {})
    valid_keys = sorted(options.keys())
    if not valid_keys:
        return None, '该题没有可用选项，无法按客观题规则修改。'

    # letters: 从输入中提取出的答案字母（去重且保持输入顺序）。
    letters = []
    normalized = _normalize_answer_text(text)
    for ch in re.findall(r'[A-H]', normalized):
        if ch in valid_keys and ch not in letters:
            letters.append(ch)

    if not letters:
        # inferred: 当用户输入的是“选项内容文本”而非字母时，反向推断字母答案。
        inferred = _extract_choice_answer(text, options)
        for ch in inferred:
            if ch in valid_keys and ch not in letters:
                letters.append(ch)

    if q_type in ('single', 'judge'):
        if len(letters) != 1:
            label = '判断' if q_type == 'judge' else '单选'
            return None, f'该题为{label}题，请输入 1 个选项字母（如 A）。'
        return letters, None

    if q_type == 'multi':
        if not letters:
            return None, '多选题请输入至少 1 个选项字母（如 AC）。'
        return letters, None

    return None, '暂不支持该题型的答案编辑。'

def _format_options_for_edit(options):
    """把选项字典转成多行可编辑文本。"""
    opt = options or {}
    lines = []
    for k in sorted(opt.keys()):
        lines.append(f'{k}: {opt.get(k, "")}')
    return '\n'.join(lines)

def _parse_manual_options_text(raw_text):
    """解析手动输入的选项文本，格式示例：A: xxx"""
    lines = [l.strip() for l in str(raw_text or '').splitlines() if l.strip()]
    if not lines:
        return {}, '请先输入选项，格式如：A: 选项内容'

    # out: 解析后的选项映射，如 {'A': 'xxx', 'B': 'yyy'}。
    out = {}
    for line in lines:
        m = re.match(r'^([A-HＡ-Ｈ])\s*[.、．:：\)）\-]?\s*(.*)$', line)
        if not m:
            return {}, f'选项格式错误：{line}（示例：A: 选项内容）'
        # letter: 统一转成半角大写字母，避免全角输入导致后续判题失败。
        letter = m.group(1).translate(str.maketrans('ＡＢＣＤＥＦＧＨ', 'ABCDEFGH'))
        text = m.group(2).strip()
        if not text:
            return {}, f'选项 {letter} 内容不能为空。'
        out[letter] = text

    if len(out) < 2:
        return {}, '客观题至少需要 2 个选项。'

    return dict(sorted(out.items())), None

def _show_question_edit_dialog(parent, question, title='编辑题目与答案'):
    """多行编辑弹窗：适合长题干。"""
    q = question or {}
    q_type = str(q.get('type', '') or '')
    type_map = {
        'single': '单选',
        'multi': '多选',
        'judge': '判断',
        'blank': '填空',
        'short': '简答'
    }
    # label_to_type: 下拉框中文显示值 -> 内部题型编码。
    label_to_type = {v: k for k, v in type_map.items()}
    type_text = type_map.get(q_type, str(q_type))

    win = tk.Toplevel(parent)
    win.title(title)
    win.transient(parent)
    win.grab_set()

    p_w = parent.winfo_width() if parent.winfo_width() > 1 else 1200
    p_h = parent.winfo_height() if parent.winfo_height() > 1 else 800
    p_x = parent.winfo_x()
    p_y = parent.winfo_y()

    win_w = max(780, int(p_w * 0.78))
    win_h = max(560, int(p_h * 0.78))
    x = p_x + max(0, (p_w - win_w) // 2)
    y = p_y + max(0, (p_h - win_h) // 2)
    win.geometry(f'{win_w}x{win_h}+{x}+{y}')
    win.configure(bg='#f5f5f5')

    info = tk.Label(
        win,
        text=f'题号：{q.get("id", "")}  |  题型：{type_text}',
        bg='#2c3e50', fg='white', anchor='w',
        font=('Microsoft YaHei', 10)
    )
    info.pack(fill='x', padx=0, pady=0, ipady=10)

    type_wrap = tk.Frame(win, bg='#f5f5f5')
    type_wrap.pack(fill='x', padx=12, pady=(10, 0))
    tk.Label(type_wrap, text='题型：', bg='#f5f5f5', fg='#2c3e50').pack(side='left')
    type_var = tk.StringVar(value=type_map.get(q_type, '简答'))
    type_combo = ttk.Combobox(
        type_wrap,
        state='readonly',
        textvariable=type_var,
        values=[type_map[k] for k in ('single', 'multi', 'judge', 'blank', 'short')],
        width=12
    )
    type_combo.pack(side='left', padx=(6, 0))

    content = tk.Frame(win, bg='#f5f5f5')
    content.pack(fill='both', expand=True, padx=12, pady=10)

    tk.Label(content, text='题目文本（可多行）：', bg='#f5f5f5', fg='#2c3e50', anchor='w').pack(fill='x')
    q_text_box = tk.Text(content, height=11, wrap='word', font=('Microsoft YaHei', 10), relief='groove', bd=1)
    q_scroll = ttk.Scrollbar(content, orient='vertical', command=q_text_box.yview)
    q_text_box.configure(yscrollcommand=q_scroll.set)
    q_text_box.pack(side='left', fill='both', expand=True, pady=(4, 10))
    q_scroll.pack(side='left', fill='y', pady=(4, 10))
    q_text_box.insert('1.0', str(q.get('text', '') or ''))

    ans_wrap = tk.Frame(win, bg='#f5f5f5')
    ans_wrap.pack(fill='both', padx=12, pady=(0, 8))

    answer_hint_var = tk.StringVar(value='答案：')
    tk.Label(ans_wrap, textvariable=answer_hint_var, bg='#f5f5f5', fg='#2c3e50', anchor='w').pack(fill='x')
    ans_text_box = tk.Text(ans_wrap, height=4, wrap='word', font=('Microsoft YaHei', 10), relief='groove', bd=1)
    ans_scroll = ttk.Scrollbar(ans_wrap, orient='vertical', command=ans_text_box.yview)
    ans_text_box.configure(yscrollcommand=ans_scroll.set)
    ans_text_box.pack(side='left', fill='both', expand=True, pady=(4, 0))
    ans_scroll.pack(side='left', fill='y', pady=(4, 0))
    ans_text_box.insert('1.0', _format_answer_text(q.get('answer')))

    opt_wrap = tk.Frame(win, bg='#f5f5f5')
    opt_wrap.pack(fill='both', padx=12, pady=(0, 8))
    options_title_var = tk.StringVar(value='选项（仅客观题需要，可多行）：')
    tk.Label(opt_wrap, textvariable=options_title_var, bg='#f5f5f5', fg='#2c3e50', anchor='w').pack(fill='x')
    opt_text_box = tk.Text(opt_wrap, height=5, wrap='word', font=('Microsoft YaHei', 10), relief='groove', bd=1)
    opt_scroll = ttk.Scrollbar(opt_wrap, orient='vertical', command=opt_text_box.yview)
    opt_text_box.configure(yscrollcommand=opt_scroll.set)
    opt_text_box.pack(side='left', fill='both', expand=True, pady=(4, 0))
    opt_scroll.pack(side='left', fill='y', pady=(4, 0))
    opt_text_box.insert('1.0', _format_options_for_edit(q.get('options') or {}))

    options_hint_var = tk.StringVar(value='')
    options_label = tk.Label(
        win,
        textvariable=options_hint_var,
        bg='#f5f5f5', fg='#7f8c8d', anchor='w',
        font=('Microsoft YaHei', 9)
    )
    options_label.pack(fill='x', padx=12, pady=(0, 8))

    # result 作为闭包共享容器，用于窗口关闭后向外返回编辑结果。
    # result: 弹窗关闭后返回给调用方的编辑结果容器。
    result = {'ok': False, 'text': '', 'answer': None, 'type': q_type, 'options': dict(q.get('options') or {})}

    def _resolve_target_options(target_type):
        """按题型推导默认选项，判断题固定为“正确/错误”。"""
        src = dict(q.get('options') or {})
        if target_type == 'judge':
            return {'A': '正确', 'B': '错误'}
        if target_type in ('single', 'multi'):
            return src
        return {}

    def _refresh_type_hints(_event=None):
        """根据当前题型刷新答案与选项输入提示。"""
        target_type = label_to_type.get(type_var.get(), q_type)
        if target_type in ('single', 'judge'):
            answer_hint_var.set('答案（输入 1 个选项字母，如 A）：')
        elif target_type == 'multi':
            answer_hint_var.set('答案（输入多选字母，如 AC）：')
        else:
            answer_hint_var.set('答案：')

        opt = _resolve_target_options(target_type)
        if target_type in ('single', 'multi', 'judge'):
            if opt:
                opt_txt = '  '.join(f'{k}:{opt.get(k, "")}' for k in sorted(opt.keys()))
                options_hint_var.set(f'可选项：{opt_txt}')
            else:
                options_hint_var.set('当前题无选项，请在下方“选项”框手动录入（如 A: xxx）。')
        else:
            options_hint_var.set('')

        if target_type in ('single', 'multi', 'judge'):
            options_title_var.set('选项（客观题必填，多行，示例：A: 选项内容）：')
            opt_text_box.config(state='normal')
        else:
            options_title_var.set('选项（当前题型不需要）：')
            opt_text_box.config(state='disabled')

    type_combo.bind('<<ComboboxSelected>>', _refresh_type_hints)
    _refresh_type_hints()

    def on_save():
        """校验输入并提交编辑结果。"""
        new_question_text = q_text_box.get('1.0', 'end').strip()
        if not new_question_text:
            messagebox.showerror('格式错误', '题目文本不能为空。', parent=win)
            return

        target_type = label_to_type.get(type_var.get(), q_type)
        if target_type == 'judge':
            target_options = {'A': '正确', 'B': '错误'}
        elif target_type in ('single', 'multi'):
            opt_raw = opt_text_box.get('1.0', 'end').strip()
            target_options, opt_err = _parse_manual_options_text(opt_raw)
            if opt_err:
                messagebox.showerror('格式错误', opt_err, parent=win)
                return
        else:
            target_options = {}

        new_answer_raw = ans_text_box.get('1.0', 'end').strip()
        parsed_answer, err = _parse_manual_answer_for_question(
            q,
            new_answer_raw,
            target_type=target_type,
            options_override=target_options
        )
        if err:
            messagebox.showerror('格式错误', err, parent=win)
            return

        result['ok'] = True
        result['text'] = new_question_text
        result['answer'] = parsed_answer
        result['type'] = target_type
        result['options'] = target_options
        win.destroy()

    def on_cancel():
        """取消编辑并关闭弹窗。"""
        win.destroy()

    btn_bar = tk.Frame(win, bg='#f5f5f5')
    btn_bar.pack(fill='x', padx=12, pady=(0, 12))

    tk.Button(
        btn_bar, text='取消',
        bg='#e0e0e0', fg='#2c3e50',
        relief='flat', padx=14, pady=6,
        command=on_cancel
    ).pack(side='right', padx=(8, 0))

    tk.Button(
        btn_bar, text='保存修改',
        bg='#2ecc71', fg='white', activebackground='#27ae60',
        relief='flat', padx=14, pady=6,
        command=on_save
    ).pack(side='right')

    win.protocol('WM_DELETE_WINDOW', on_cancel)
    parent.wait_window(win)

    if result['ok']:
        return result['text'], result['answer'], result['type'], result['options']
    return None
