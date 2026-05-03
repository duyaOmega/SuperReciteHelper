#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""程序入口。"""

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

from parser import build_parse_candidates
from question_bank import load_app_state, save_app_state
from ui_main import (
    QuizApp,
    _choose_startup_files,
    _dedupe_existing_paths,
    _enable_windows_high_dpi,
    show_import_preview,
)

def main():
    """应用主入口：选择题库、解析预览并启动刷题界面。"""
    _enable_windows_high_dpi()

    # 确定脚本目录
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # 文件选择阶段使用独立隐藏窗口，避免主窗口状态异常导致不显示
    selector = tk.Tk()
    selector.withdraw()
    selector.update_idletasks()

    # selected_paths: 用户在文件选择阶段挑选的“原始路径列表”（可能有重复/失效路径）。
    selected_paths = _choose_startup_files(selector, script_dir)
    selector.destroy()

    # file_paths: 去重且确认存在后的最终导入列表。
    file_paths = _dedupe_existing_paths(selected_paths)
    if not file_paths:
        messagebox.showinfo('已取消', '未选择题库文件，程序将退出。')
        return

    print(f"正在解析题库，共 {len(file_paths)} 个文件。")

    # 单文件走标准解析流程；多文件时先分别解析再合并。
    if len(file_paths) == 1:
        txt_path = file_paths[0]
        try:
            candidates = build_parse_candidates(txt_path)
        except Exception as e:
            messagebox.showerror('解析失败', f'文件解析失败：\n{e}')
            return

        if not candidates or not candidates[0][1]:
            messagebox.showerror('错误', '未能解析出任何题目，请检查题库格式。')
            return

        source_label = txt_path
    else:
        # merged_questions: 多文件合并后的统一题目池。
        merged_questions = []
        # failed_files: 记录失败文件及失败原因，便于一次性反馈给用户。
        failed_files = []

        for path in file_paths:
            try:
                file_candidates = build_parse_candidates(path)
                if not file_candidates or not file_candidates[0][1]:
                    failed_files.append((path, '未解析出题目'))
                    continue

                # file_candidates[0] 约定为“自动识别（推荐）”方案。
                file_questions = file_candidates[0][1]
                for q in file_questions:
                    copied = dict(q)
                    # source_file: 题目来源文件名，便于后续排查解析问题。
                    copied['source_file'] = os.path.basename(path)
                    merged_questions.append(copied)
                print(f"已解析：{os.path.basename(path)} -> {len(file_questions)} 题")
            except Exception as e:
                failed_files.append((path, str(e)))

        if not merged_questions:
            detail = '\n'.join(f"- {os.path.basename(p)}: {msg}" for p, msg in failed_files) if failed_files else '未知错误'
            messagebox.showerror('解析失败', f'所有文件均解析失败：\n{detail}')
            return

        # 合并后重排 id，确保题号在当前会话内连续可读。
        for idx, q in enumerate(merged_questions, 1):
            q['id'] = idx

        if failed_files:
            detail = '\n'.join(f"- {os.path.basename(p)}: {msg}" for p, msg in failed_files[:8])
            messagebox.showwarning('部分文件解析失败', f'以下文件未成功导入（其余文件已合并）：\n{detail}')

        candidates = [
            (
                f'自动识别（多文件合并，{len(file_paths)}个）',
                merged_questions,
                '每个文件采用自动识别方案后合并为同一题库'
            )
        ]
        source_label = f'多文件({len(file_paths)})'

    # 主窗口单独创建，确保可见
    root = tk.Tk()
    root.title('刷题工具')
    root.lift()
    root.focus_force()

    # 导入预览窗口：先确认解析结果再进入刷题。
    preview_source = file_paths[0] if len(file_paths) == 1 else f'多文件导入（{len(file_paths)}个）'
    selected_questions = show_import_preview(root, candidates, preview_source)
    if not selected_questions:
        root.destroy()
        return

    state = load_app_state()
    state['last_open_files'] = file_paths
    save_app_state(state)

    app = QuizApp(root, selected_questions, source_path=source_label)
    root.mainloop()

if __name__ == '__main__':
    main()
