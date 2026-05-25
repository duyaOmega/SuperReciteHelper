#!/usr/bin/env python3
"""PyQt 界面辅助函数。"""

import re


TYPE_LABELS = {
    "single": "单选题",
    "multi": "多选题",
    "judge": "判断题",
    "blank": "填空题",
    "short": "简答题",
}


def normalize_keyboard_text(text):
    """规范化键盘输入，兼容全角字母。"""
    normalized = (text or "").strip().upper()
    return normalized.translate(str.maketrans("ＡＢＣＤＥＦＧＨ，。、；：　", "ABCDEFGH,,,,  "))


def question_signature(question):
    """生成题目归一签名，用于重复题检测。"""
    q = question or {}
    # 题干签名忽略空白、标点和填空占位差异。
    text = str(q.get("text", "") or "")
    text = re.sub(r"（\s*\d+\s*）\s*[_＿﹍]+", "（）", text)
    text = re.sub(r"[_＿﹍]+", "", text)
    text = re.sub(r"[\s，,。；;：:、（）()\[\]【】]+", "", text)

    options = q.get("options") or {}
    option_sig = []
    # 选项也参与签名，避免题干相同但选项不同的题被误判重复。
    for key in sorted(options.keys()):
        value = re.sub(r"\s+", "", str(options.get(key, "") or ""))
        option_sig.append(f"{key}:{value}")
    return (q.get("type", ""), text, "|".join(option_sig))


def build_duplicate_groups(questions):
    """构建重复题分组。"""
    sig_map = {}
    for q in questions:
        sig_map.setdefault(question_signature(q), []).append(q.get("id"))
    groups = [ids for ids in sig_map.values() if len(ids) >= 2]
    groups.sort(key=lambda ids: (-len(ids), ids[0]))
    return groups


def build_duplicate_signature_set(duplicate_groups, question_map):
    """提取重复题签名集合。"""
    sigs = set()
    for group in duplicate_groups:
        for qid in group:
            q = question_map.get(qid)
            if q:
                sigs.add(question_signature(q))
    return sigs


def is_recent_duplicate_pick(question, duplicate_signature_set, recent_signatures):
    """判断候选题是否属于近期重复组。"""
    sig = question_signature(question)
    return sig in duplicate_signature_set and sig in recent_signatures
