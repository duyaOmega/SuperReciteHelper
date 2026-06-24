#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""题目编辑对话框与答案格式化模块。"""

import re

from .parser import _extract_choice_answer, _normalize_answer_text

def format_answer_text(answer_value):
    """将答案值统一格式化为可展示字符串。"""
    if isinstance(answer_value, list):
        return ''.join(answer_value)
    return str(answer_value or '')

def parse_manual_answer_for_question(question, raw_text, target_type=None, options_override=None):
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

def format_options_for_edit(options):
    """把选项字典转成多行可编辑文本。"""
    opt = options or {}
    lines = []
    for k in sorted(opt.keys()):
        lines.append(f'{k}: {opt.get(k, "")}')
    return '\n'.join(lines)

def parse_manual_options_text(raw_text):
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
