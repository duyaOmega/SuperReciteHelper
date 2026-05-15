#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""刷题会话与抽题策略模块。"""

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

from question_bank import get_record

def weighted_random_pick(questions, records):
    """根据错误次数和错误率进行加权随机抽题"""
    weights = []
    for q in questions:
        rec = get_record(records, q)
        # attempts/errors: 该题历史作答次数与错误次数。
        attempts = rec['attempts']
        errors = rec['errors']
        # error_rate: 历史错误率，用于提升“高风险题”抽中概率。
        error_rate = errors / attempts if attempts > 0 else 0

        # 基础权重
        weight = 1.0

        # 错误次数越多权重越大
        if errors >= 5:
            weight += 2.5
        elif errors >= 3:
            weight += 1.0

        # 错误率高的加权
        if attempts > 0:
            if error_rate > 0.5:
                weight += 3.0
            elif error_rate > 0.3:
                weight += 1.5

        # 从未做过的题目：剩余越少，加权越高，确保能刷完题库
        if attempts == 0:
            weight += 2.0

        weights.append(weight)

    # 根据未做题目占比动态提升未做题的加权
    # ratio = unseen_count / total_count — 未做题目占比
    # boost = 1.0 / ratio — 越少未做，boost 越大
    unseen_count = sum(1 for q in questions if get_record(records, q)['attempts'] == 0)
    total_count = len(questions)
    if unseen_count > 0 and unseen_count < total_count:
        # ratio 越小（剩余未做越少），boost 越大
        ratio = unseen_count / total_count
        boost = 1.0 / ratio
        for i in range(len(weights)):
            if get_record(records, questions[i])['attempts'] == 0:
                weights[i] *= boost

    # 加权随机选择
    total = sum(weights)
    # r: 在 [0, total] 上均匀采样的随机阈值。
    r = random.uniform(0, total)
    cumulative = 0
    for i, w in enumerate(weights):
        # cumulative: 累积权重区间，命中即返回对应题目。
        cumulative += w
        if r <= cumulative:
            return questions[i]
    return questions[-1]

class PracticeSession:
    """轻量刷题会话对象。"""

    def __init__(self, questions, records):
        """初始化会话：持有题集、记录与当前题指针。"""
        self.questions = questions
        self.records = records
        self.current_question = None

    def pick_next(self):
        """按加权策略抽取下一题并更新当前题。"""
        self.current_question = weighted_random_pick(self.questions, self.records)
        return self.current_question
