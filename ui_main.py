#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""图形界面模块。"""

# 变量缩写约定（阅读 UI 交互代码时可快速对照）：
# q: 当前题目字典（question）
# rec: 该题历史作答记录（record）
# rows: 统计表行数据（dict 列表）
# sig: 题目归一化签名（用于重复题分组）
# selected: 用户当前选中的选项字母集合

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

from parser import _clean_pdf_judge_text_noise, _mask_blank_question_text
from question import _format_answer_text, _show_question_edit_dialog
from question_bank import (
    _ensure_question_identity_fields,
    apply_manual_question_edits,
    get_record,
    load_app_state,
    load_manual_question_edits,
    load_records,
    save_app_state,
    save_manual_question_edits,
    save_records,
    update_record,
    upsert_manual_question_edit,
)
from session import weighted_random_pick

class QuizApp:
    def __init__(self, root, questions, source_path=''):
        """初始化刷题主界面、题库状态与交互变量。"""
        self.root = root
        # questions: 当前会话的完整题库（已应用手动编辑覆盖）。
        self.questions = questions
        self.manual_edits = load_manual_question_edits()
        for q in self.questions:
            _ensure_question_identity_fields(q)
        apply_manual_question_edits(self.questions, self.manual_edits)
        # question_map: 题号 -> 题目对象，供统计窗口“按题号快速回查详情”。
        self.question_map = {q['id']: q for q in questions}
        # 兼容单文件字符串与多文件显示标签。
        self.source_path = source_path
        if isinstance(source_path, (list, tuple)):
            source_name = f'多文件({len(source_path)})'
        else:
            source_name = os.path.basename(source_path) if source_path else '未命名'
        self.source_name = source_name
        self.records = load_records()
        self.current_questions = list(self.questions)
        self.current_q = None
        # selected: 客观题当前选中的字母集合（单选也统一用 set，便于复用比较逻辑）。
        self.selected = set()
        self.submitted = False
        self.answer_revealed = False
        self.option_buttons = {}
        self.judge_buttons = []
        self.keyboard_var = tk.StringVar()
        # recent_signatures: 最近抽到题目的签名队列，用于减少重复题短期复现。
        self.recent_signatures = []
        self.recent_signature_limit = 6
        self.duplicate_groups = self._build_duplicate_groups()
        self.duplicate_signature_set = self._build_duplicate_signature_set()

        self.root.title("刷题工具")
        self.ui_scale = self._get_ui_scale()
        self._apply_adaptive_window_geometry()
        self.root.configure(bg='#f5f5f5')
        min_w = max(760, int(self.window_width * 0.72))
        min_h = max(560, int(self.window_height * 0.72))
        self.root.minsize(min_w, min_h)

        # 字体
        self.title_font = tkfont.Font(family='Microsoft YaHei', size=self._scale_font(16), weight='bold')
        self.text_font = tkfont.Font(family='Microsoft YaHei', size=self._scale_font(13))
        self.option_font = tkfont.Font(family='Microsoft YaHei', size=self._scale_font(12))
        self.question_display_font = tkfont.Font(family='Consolas', size=self._scale_font(13))
        self.option_display_font = tkfont.Font(family='Consolas', size=self._scale_font(12))
        self.small_font = tkfont.Font(family='Microsoft YaHei', size=self._scale_font(10))
        self.btn_font = tkfont.Font(family='Microsoft YaHei', size=self._scale_font(11), weight='bold')

        self.question_wrap = max(680, int(self.window_width * 0.84))

        self.build_ui()
        self.root.bind('<Return>', self._on_global_enter)
        self.root.after(120, self._refresh_layout)
        self.show_welcome()

    def _get_ui_scale(self):
        """基于屏幕分辨率估算 UI 缩放倍率。"""
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        scale_w = screen_w / 1920
        scale_h = screen_h / 1080
        return max(1.0, min(1.5, (scale_w + scale_h) / 2))

    def _scale_font(self, base_size):
        """按当前缩放倍率换算字体大小。"""
        return max(10, int(round(base_size * self.ui_scale)))

    def _apply_adaptive_window_geometry(self):
        """根据屏幕尺寸设置窗口初始大小与居中位置。"""
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()

        self.window_width = max(920, min(1600, int(screen_w * 0.78)))
        self.window_height = max(720, min(1100, int(screen_h * 0.84)))

        x = max(0, (screen_w - self.window_width) // 2)
        y = max(0, (screen_h - self.window_height) // 2)
        self.root.geometry(f"{self.window_width}x{self.window_height}+{x}+{y}")

    def build_ui(self):
        """构建主界面控件：题干区、选项区、统计栏与按钮区。"""
        # 顶部信息栏
        top_frame = tk.Frame(self.root, bg='#2c3e50', height=50)
        top_frame.pack(fill='x')
        top_frame.pack_propagate(False)

        self.info_label = tk.Label(
            top_frame,
            text=f"题库：{self.source_name} | 共 {len(self.questions)} 题",
            font=self.small_font, fg='white', bg='#2c3e50'
        )
        self.info_label.pack(side='left', padx=15, pady=10)

        self.stats_label = tk.Label(
            top_frame, text="",
            font=self.small_font, fg='#ecf0f1', bg='#2c3e50'
        )
        self.stats_label.pack(side='right', padx=15, pady=10)

        # 主内容区域（使用 Canvas + Scrollbar 实现滚动）
        main_container = tk.Frame(self.root, bg='#f5f5f5')
        main_container.pack(fill='both', expand=True, padx=20, pady=10)

        self.canvas = tk.Canvas(main_container, bg='#f5f5f5', highlightthickness=0)
        self.scrollbar = tk.Scrollbar(main_container, orient='vertical', command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg='#f5f5f5')

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side='left', fill='both', expand=True)
        self.scrollbar.pack(side='right', fill='y')

        # 让 scrollable_frame 宽度跟随 canvas
        self.canvas.bind('<Configure>', self._on_canvas_configure)

        # 鼠标滚轮绑定
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Button-4>", self._on_mousewheel_linux)
        self.canvas.bind_all("<Button-5>", self._on_mousewheel_linux)

        # 题号标签
        self.q_number_label = tk.Label(
            self.scrollable_frame, text="", font=self.title_font,
            fg='#2c3e50', bg='#f5f5f5', anchor='w'
        )
        self.q_number_label.pack(fill='x', pady=(10, 2))

        # 题目类型标签
        self.q_type_label = tk.Label(
            self.scrollable_frame, text="", font=self.small_font,
            fg='#7f8c8d', bg='#f5f5f5', anchor='w'
        )
        self.q_type_label.pack(fill='x', pady=(0, 5))

        # 题目文本
        self.q_text_label = tk.Label(
            self.scrollable_frame, text="", font=self.question_display_font,
            fg='#2c3e50', bg='#ffffff', anchor='w', justify='left',
            wraplength=self.question_wrap, padx=15, pady=15, relief='ridge', bd=1
        )
        self.q_text_label.pack(fill='x', pady=(0, 15))

        # 选项容器
        self.options_frame = tk.Frame(self.scrollable_frame, bg='#f5f5f5')
        self.options_frame.pack(fill='x', pady=5)

        # 结果显示
        self.result_label = tk.Label(
            self.scrollable_frame, text="", font=self.text_font,
            fg='#27ae60', bg='#f5f5f5', anchor='w', justify='left',
            wraplength=self.question_wrap
        )
        self.result_label.pack(fill='x', pady=10)

        # 历史记录显示
        self.history_label = tk.Label(
            self.scrollable_frame, text="", font=self.small_font,
            fg='#95a5a6', bg='#f5f5f5', anchor='w'
        )
        self.history_label.pack(fill='x', pady=(0, 5))

        # 底部按钮栏（双行布局，避免高 DPI 或窄窗口下按钮被挤扁/截断）
        btn_frame = tk.Frame(self.root, bg='#ecf0f1')
        btn_frame.pack(fill='x', side='bottom')

        action_row = tk.Frame(btn_frame, bg='#ecf0f1')
        action_row.pack(fill='x', padx=10, pady=(8, 2))

        utility_row = tk.Frame(btn_frame, bg='#ecf0f1')
        utility_row.pack(fill='x', padx=10, pady=(0, 8))

        self.submit_btn = tk.Button(
            action_row, text="提交答案", font=self.btn_font,
            bg='#3498db', fg='white', activebackground='#2980b9',
            relief='flat', padx=20, pady=8, command=self.submit_answer
        )
        self.submit_btn.pack(side='left', padx=(10, 8), pady=2)

        self.next_btn = tk.Button(
            action_row, text="下一题 ▶", font=self.btn_font,
            bg='#2ecc71', fg='white', activebackground='#27ae60',
            relief='flat', padx=20, pady=8, command=self.next_question
        )
        self.next_btn.pack(side='left', padx=(0, 8), pady=2)

        # 键盘输入区：支持 ABC 选项输入与 t/f 主观自评。
        input_frame = tk.Frame(action_row, bg='#ecf0f1')
        input_frame.pack(side='left', padx=(8, 0), pady=2)

        self.keyboard_hint_label = tk.Label(
            input_frame,
            text='键盘输入：A/B/C 或 ABC；主观题用 t/f',
            font=self.small_font,
            fg='#34495e',
            bg='#ecf0f1'
        )
        self.keyboard_hint_label.pack(side='left', padx=(0, 8))

        self.keyboard_entry = tk.Entry(
            input_frame,
            textvariable=self.keyboard_var,
            font=self.option_font,
            width=16,
            relief='groove'
        )
        self.keyboard_entry.pack(side='left', padx=(0, 6))
        self.keyboard_entry.bind('<Return>', self._on_entry_enter)

        self.reset_btn = tk.Button(
            utility_row, text="重置记录", font=self.small_font,
            bg='#e74c3c', fg='white', activebackground='#c0392b',
            relief='flat', padx=10, pady=5, command=self.reset_records
        )
        self.reset_btn.pack(side='right', padx=(8, 0), pady=2)

        self.manage_edits_btn = tk.Button(
            utility_row, text="管理题目修改", font=self.small_font,
            bg='#16a085', fg='white', activebackground='#138d75',
            relief='flat', padx=10, pady=5, command=self.manage_manual_edits
        )
        self.manage_edits_btn.pack(side='right', padx=(8, 0), pady=2)

        self.edit_current_btn = tk.Button(
            utility_row, text="编辑当前题", font=self.small_font,
            bg='#f39c12', fg='white', activebackground='#d68910',
            relief='flat', padx=10, pady=5, command=self.edit_current_question
        )
        self.edit_current_btn.pack(side='right', padx=(8, 0), pady=2)

        self.freq_btn = tk.Button(
            utility_row, text="考频统计", font=self.small_font,
            bg='#8e44ad', fg='white', activebackground='#7d3c98',
            relief='flat', padx=10, pady=5, command=self.show_frequency_stats
        )
        self.freq_btn.pack(side='right', padx=(8, 0), pady=2)

    def _on_canvas_configure(self, event):
        """Canvas 尺寸变化时同步内部容器宽度并刷新换行布局。"""
        self.canvas.itemconfig(self.canvas_window, width=event.width)
        self._refresh_layout()

    def _refresh_layout(self):
        """根据当前窗口宽度动态更新题干与选项的换行宽度。"""
        canvas_w = self.canvas.winfo_width()
        if canvas_w <= 1:
            canvas_w = max(600, self.root.winfo_width() - 90)

        content_width = max(420, canvas_w - 45)
        self.q_text_label.config(wraplength=content_width)
        self.result_label.config(wraplength=content_width)

        option_wrap = max(380, content_width - 36)
        for btn in self.option_buttons.values():
            btn.config(wraplength=option_wrap)

    def _on_mousewheel(self, event):
        """Windows/macOS 鼠标滚轮滚动处理。"""
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_mousewheel_linux(self, event):
        """Linux 鼠标滚轮滚动处理。"""
        if event.num == 4:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self.canvas.yview_scroll(1, "units")

    def show_welcome(self):
        """显示欢迎页与题库总体提示信息。"""
        self.q_number_label.config(text="欢迎使用刷题工具！")
        self.q_type_label.config(text="")
        duplicate_hint = ''
        if self.duplicate_groups:
            dup_questions = sum(len(g) for g in self.duplicate_groups)
            duplicate_hint = f"\n\n检测到重复题组 {len(self.duplicate_groups)} 组（共 {dup_questions} 题），系统会自动尽量避免短时间重复抽到同组题。"

        self.q_text_label.config(
            text=f"题库已加载 {len(self.questions)} 道题目。\n\n"
                 f"点击「下一题」开始刷题！\n\n"
                 f"系统会根据你的错误率自动加权抽题，\n"
                 f"错得越多的题越容易被抽到哦～"
                 f"{duplicate_hint}"
        )
        self.result_label.config(text="")
        self.history_label.config(text="")
        self.keyboard_var.set('')
        self.keyboard_hint_label.config(text='键盘输入：A/B/C 或 ABC；主观题用 t/f')
        self.submit_btn.config(state='disabled')
        self.update_stats()

    def update_stats(self):
        """汇总并刷新顶部作答统计信息。"""
        total_attempts = 0
        total_errors = 0
        attempted = 0

        for q in self.current_questions:
            rec = get_record(self.records, q)
            attempts = int(rec.get('attempts', 0) or 0)
            errors = int(rec.get('errors', 0) or 0)
            total_attempts += attempts
            total_errors += errors
            if attempts > 0:
                attempted += 1

        # acc: 当前题库口径的总体正确率（不是单题正确率）。
        acc = (1 - total_errors / total_attempts) * 100 if total_attempts > 0 else 0
        self.stats_label.config(
            text=f"已做: {attempted}/{len(self.questions)} | "
                 f"当前题库总答题: {total_attempts} | 当前题库总正确率: {acc:.1f}%"
        )

    def next_question(self):
        """抽取下一题，并尽量规避短时间命中重复题组。"""
        self.submitted = False
        self.answer_revealed = False
        self.selected = set()
        self.keyboard_var.set('')
        # picked: 本轮抽中的候选题，后续可能因“近期重复”而重抽。
        picked = weighted_random_pick(self.questions, self.records)

        # 若命中近期重复题组，尝试重抽，减少“同题反复出现”的体感。
        if len(self.questions) > 1 and self.duplicate_groups:
            # 最多重抽 30 次，避免极端情况下进入长循环。
            for _ in range(30):
                if not self._is_recent_duplicate_pick(picked):
                    break
                candidate = weighted_random_pick(self.questions, self.records)
                if not self._is_recent_duplicate_pick(candidate):
                    picked = candidate
                    break
                picked = candidate

        self.current_q = picked
        sig = self._question_signature(self.current_q)
        self.recent_signatures.append(sig)
        if len(self.recent_signatures) > self.recent_signature_limit:
            self.recent_signatures = self.recent_signatures[-self.recent_signature_limit:]
        self.display_question()

    def display_question(self):
        """渲染当前题目、历史记录与对应作答控件。"""
        q = self.current_q
        type_map = {
            'single': '【单选题】',
            'multi': '【多选题】',
            'judge': '【判断题】',
            'blank': '【填空题】',
            'short': '【简答题】'
        }

        self.q_number_label.config(text=f"第 {q['id']} 题")
        self.q_type_label.config(text=type_map.get(q['type'], ''))
        display_text = q['text']
        if q.get('type') == 'blank':
            display_text = _mask_blank_question_text(q['text'], q.get('answer', ''))
        self.q_text_label.config(text=display_text)
        self.result_label.config(text="")

        # 显示历史记录
        rec = get_record(self.records, q)
        if rec['attempts'] > 0:
            rate = rec['errors'] / rec['attempts'] * 100
            self.history_label.config(
                text=f"历史记录：答过 {rec['attempts']} 次，错误 {rec['errors']} 次，错误率 {rate:.0f}%"
            )
        else:
            self.history_label.config(text="历史记录：首次作答")

        # 清除旧选项
        for widget in self.options_frame.winfo_children():
            widget.destroy()
        self.option_buttons = {}
        self.judge_buttons = []

        if q['type'] in ('single', 'multi', 'judge'):
            # 客观题：创建选项按钮
            sorted_keys = sorted(q['options'].keys())
            for key in sorted_keys:
                btn = tk.Button(
                    self.options_frame,
                    text=f"  {key}. {q['options'][key]}",
                    font=self.option_display_font,
                    bg='white', fg='#2c3e50',
                    activebackground='#d5e8d4',
                    relief='ridge', bd=1,
                    anchor='w', justify='left',
                    padx=15, pady=8,
                    wraplength=700,
                    command=lambda k=key: self.toggle_option(k)
                )
                btn.pack(fill='x', pady=3, ipady=3)
                self.option_buttons[key] = btn

            self.submit_btn.config(text='提交答案', state='normal')
            if q['type'] == 'multi':
                self.keyboard_hint_label.config(text='键盘输入：如 ABC；回车提交，已判分后回车下一题')
            else:
                self.keyboard_hint_label.config(text='键盘输入：如 A 或 B；回车提交，已判分后回车下一题')
        else:
            # 主观题：先不显示任何作答选项，先看答案再自判
            self.result_label.config(
                text='请先自己作答，点击「显示正确答案」后再进行自评。',
                fg='#7f8c8d'
            )
            self.submit_btn.config(text='显示正确答案', state='normal')
            self.keyboard_hint_label.config(text='主观题：先回车显示答案，再输入 t/f 回车自评；再回车下一题')

        self._refresh_layout()
        self.keyboard_entry.focus_set()

        # 滚动到顶部
        self.canvas.yview_moveto(0)

    def toggle_option(self, key):
        """处理客观题选项点击：单选覆盖、多选切换。"""
        if self.submitted:
            return

        q = self.current_q
        if q['type'] == 'single' or q['type'] == 'judge':
            # 单选：清除其他选中
            self.selected = {key}
            for k, btn in self.option_buttons.items():
                if k == key:
                    btn.config(bg='#3498db', fg='white')
                else:
                    btn.config(bg='white', fg='#2c3e50')
        else:
            # 多选：切换选中状态
            if key in self.selected:
                self.selected.discard(key)
                self.option_buttons[key].config(bg='white', fg='#2c3e50')
            else:
                self.selected.add(key)
                self.option_buttons[key].config(bg='#3498db', fg='white')

    def submit_answer(self):
        """提交当前答案：客观题自动判分，主观题先展示参考答案。"""
        if self.submitted:
            return

        q = self.current_q

        if q['type'] in ('blank', 'short'):
            if self.answer_revealed:
                return

            self.answer_revealed = True
            self.result_label.config(
                text=f"参考答案：{q['answer']}\n\n请根据你的作答进行自评：",
                fg='#2c3e50'
            )

            correct_btn = tk.Button(
                self.options_frame,
                text='我答对了',
                font=self.btn_font,
                bg='#2ecc71', fg='white',
                activebackground='#27ae60',
                relief='flat', padx=15, pady=8,
                command=lambda: self.submit_subjective_result(True)
            )
            correct_btn.pack(side='left', padx=(0, 10), pady=5)

            wrong_btn = tk.Button(
                self.options_frame,
                text='我答错了',
                font=self.btn_font,
                bg='#e74c3c', fg='white',
                activebackground='#c0392b',
                relief='flat', padx=15, pady=8,
                command=lambda: self.submit_subjective_result(False)
            )
            wrong_btn.pack(side='left', pady=5)

            self.judge_buttons = [correct_btn, wrong_btn]
            self.submit_btn.config(state='disabled')
            self.keyboard_hint_label.config(text='输入 t(对)/f(错) 后回车自评，已记录后回车下一题')
            return

        if not self.selected:
            messagebox.showwarning("提示", "请先选择一个答案！")
            return

        if q['type'] in ('single', 'multi', 'judge') and not q.get('answer'):
            messagebox.showwarning("提示", "本题未识别出标准答案，暂无法自动判分。")
            return

        self.submitted = True
        correct = set(q['answer'])
        is_correct = self.selected == correct

        # 更新记录
        update_record(self.records, q, is_correct)
        self.update_stats()

        # 高亮显示正确/错误选项
        for k, btn in self.option_buttons.items():
            if k in correct:
                btn.config(bg='#27ae60', fg='white')  # 正确选项绿色
            elif k in self.selected and k not in correct:
                btn.config(bg='#e74c3c', fg='white')  # 错选红色
            else:
                btn.config(bg='#ecf0f1', fg='#95a5a6')  # 未选灰色

        if is_correct:
            self.result_label.config(
                text="✓ 回答正确！", fg='#27ae60'
            )
        else:
            self.result_label.config(
                text=f"✗ 回答错误！正确答案是：{''.join(sorted(correct))}",
                fg='#e74c3c'
            )

        # 更新历史显示
        rec = get_record(self.records, q)
        rate = rec['errors'] / rec['attempts'] * 100
        self.history_label.config(
            text=f"历史记录：答过 {rec['attempts']} 次，错误 {rec['errors']} 次，错误率 {rate:.0f}%"
        )

        self.submit_btn.config(state='disabled')
        self.keyboard_hint_label.config(text='已判分，按回车进入下一题')

    def submit_subjective_result(self, is_correct):
        """记录主观题自评结果并刷新统计。"""
        if self.submitted:
            return

        self.submitted = True
        q = self.current_q

        update_record(self.records, q, is_correct)
        self.update_stats()

        for btn in self.judge_buttons:
            btn.config(state='disabled')

        if is_correct:
            self.result_label.config(
                text=f"参考答案：{q['answer']}\n\n✓ 已记录：你判定本题答对。",
                fg='#27ae60'
            )
        else:
            self.result_label.config(
                text=f"参考答案：{q['answer']}\n\n✗ 已记录：你判定本题答错。",
                fg='#e74c3c'
            )

        rec = get_record(self.records, q)
        rate = rec['errors'] / rec['attempts'] * 100
        self.history_label.config(
            text=f"历史记录：答过 {rec['attempts']} 次，错误 {rec['errors']} 次，错误率 {rate:.0f}%"
        )

        self.submit_btn.config(state='disabled')
        self.keyboard_hint_label.config(text='已记录，按回车进入下一题')

    def reset_records(self):
        """清空全部作答记录。"""
        if messagebox.askyesno("确认", "确定要重置所有错误记录吗？\n此操作不可撤销！"):
            self.records = {}
            save_records(self.records)
            self.update_stats()
            messagebox.showinfo("完成", "所有记录已重置。")

    def _ask_edit_question_and_answer(self, q):
        """弹窗编辑题干与答案。"""
        edited = _show_question_edit_dialog(self.root, q, title='编辑当前题（题目+答案）')
        if not edited:
            return False

        new_text, parsed_answer, new_type, new_options = edited
        q['text'] = new_text
        q['answer'] = parsed_answer
        q['type'] = new_type
        q['options'] = dict(new_options or {})
        upsert_manual_question_edit(self.manual_edits, q)
        return True

    def edit_current_question(self):
        """打开当前题编辑弹窗并持久化修改。"""
        if not self.current_q:
            messagebox.showwarning('提示', '请先点击“下一题”抽取题目。', parent=self.root)
            return

        if not self._ask_edit_question_and_answer(self.current_q):
            return

        # 编辑后重置提交状态并刷新展示
        self.submitted = False
        self.answer_revealed = False
        self.selected = set()
        self.display_question()
        messagebox.showinfo('完成', '当前题修改已保存，下次启动仍会生效。', parent=self.root)

    def manage_manual_edits(self):
        """管理已保存的手动改题记录（恢复单条或清空全部）。"""
        win = tk.Toplevel(self.root)
        win.title('管理题目修改')
        win_w = max(900, int(self.root.winfo_width() * 0.88))
        win_h = max(520, int(self.root.winfo_height() * 0.72))
        x = self.root.winfo_x() + max(0, (self.root.winfo_width() - win_w) // 2)
        y = self.root.winfo_y() + max(0, (self.root.winfo_height() - win_h) // 2)
        win.geometry(f'{win_w}x{win_h}+{x}+{y}')
        win.configure(bg='#f5f5f5')
        win.transient(self.root)

        top = tk.Frame(win, bg='#2c3e50', height=52)
        top.pack(fill='x')
        top.pack_propagate(False)

        title_label = tk.Label(
            top,
            text=f'已保存修改：{len(self.manual_edits)} 项',
            font=self.small_font,
            fg='white',
            bg='#2c3e50',
            anchor='w'
        )
        title_label.pack(fill='x', padx=12, pady=14)

        body = tk.Frame(win, bg='#f5f5f5')
        body.pack(fill='both', expand=True, padx=12, pady=10)

        columns = ('no', 'type', 'answer', 'updated_at', 'preview')
        tree = ttk.Treeview(body, columns=columns, show='headings')
        tree.heading('no', text='序号')
        tree.heading('type', text='题型')
        tree.heading('answer', text='答案')
        tree.heading('updated_at', text='更新时间')
        tree.heading('preview', text='题干预览')
        tree.column('no', width=60, anchor='center')
        tree.column('type', width=70, anchor='center')
        tree.column('answer', width=120, anchor='w')
        tree.column('updated_at', width=140, anchor='center')
        tree.column('preview', width=620, anchor='w')

        yscroll = ttk.Scrollbar(body, orient='vertical', command=tree.yview)
        xscroll = ttk.Scrollbar(body, orient='horizontal', command=tree.xview)
        tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        tree.grid(row=0, column=0, sticky='nsew')
        yscroll.grid(row=0, column=1, sticky='ns')
        xscroll.grid(row=1, column=0, sticky='ew')
        body.grid_columnconfigure(0, weight=1)
        body.grid_rowconfigure(0, weight=1)

        def refresh_table():
            for item in tree.get_children():
                tree.delete(item)

            for i, (k, v) in enumerate(sorted(self.manual_edits.items(), key=lambda x: str(x[1].get('updated_at', '')), reverse=True), 1):
                ans = _format_answer_text(v.get('answer', ''))
                preview = str(v.get('preview', '') or '')
                tree.insert('', 'end', iid=k, values=(
                    i,
                    v.get('type', ''),
                    (ans[:22] + '...') if len(ans) > 22 else ans,
                    v.get('updated_at', ''),
                    preview
                ))
            title_label.config(text=f'已保存修改：{len(self.manual_edits)} 项')

        btn_bar = tk.Frame(win, bg='#f5f5f5')
        btn_bar.pack(fill='x', padx=12, pady=(0, 12))

        def restore_selected():
            selected = tree.selection()
            if not selected:
                messagebox.showwarning('提示', '请先选择一条修改记录。', parent=win)
                return

            key = selected[0]
            payload = self.manual_edits.get(key)
            if not payload:
                return

            if not messagebox.askyesno('确认', '确定恢复该题为默认解析结果吗？', parent=win):
                return

            del self.manual_edits[key]
            save_manual_question_edits(self.manual_edits)

            for q in self.questions:
                _ensure_question_identity_fields(q)
                if q.get('_base_key') == key:
                    q['type'] = str(payload.get('orig_type', q.get('_orig_type', q.get('type', ''))))
                    q['options'] = dict(payload.get('orig_options', q.get('_orig_options', q.get('options', {}))) or {})
                    q['text'] = str(payload.get('orig_text', q.get('_orig_text', q.get('text', ''))))
                    q['answer'] = payload.get('orig_answer', q.get('_orig_answer', q.get('answer', '')))
            if self.current_q and self.current_q.get('_base_key') == key:
                self.display_question()

            refresh_table()

        def clear_all():
            if not self.manual_edits:
                messagebox.showinfo('提示', '当前没有可清空的修改。', parent=win)
                return
            if not messagebox.askyesno('确认', '确定清空全部题目修改吗？\n该操作不可撤销。', parent=win):
                return

            deleted = dict(self.manual_edits)
            self.manual_edits = {}
            save_manual_question_edits(self.manual_edits)

            for q in self.questions:
                _ensure_question_identity_fields(q)
                key = q.get('_base_key')
                if key in deleted:
                    payload = deleted[key]
                    q['type'] = str(payload.get('orig_type', q.get('_orig_type', q.get('type', ''))))
                    q['options'] = dict(payload.get('orig_options', q.get('_orig_options', q.get('options', {}))) or {})
                    q['text'] = str(payload.get('orig_text', q.get('_orig_text', q.get('text', ''))))
                    q['answer'] = payload.get('orig_answer', q.get('_orig_answer', q.get('answer', '')))

            if self.current_q:
                self.display_question()

            refresh_table()

        tk.Button(
            btn_bar, text='恢复所选默认',
            font=('Microsoft YaHei', 9),
            bg='#f39c12', fg='white', activebackground='#d68910',
            relief='flat', padx=12, pady=6,
            command=restore_selected
        ).pack(side='right', padx=(8, 0))

        tk.Button(
            btn_bar, text='清空全部修改',
            font=('Microsoft YaHei', 9),
            bg='#e74c3c', fg='white', activebackground='#c0392b',
            relief='flat', padx=12, pady=6,
            command=clear_all
        ).pack(side='right')

        refresh_table()

    def show_frequency_stats(self):
        """展示题目考频统计（基于作答记录）。"""
        win = tk.Toplevel(self.root)
        win.title('考频统计')
        win_w = max(900, int(self.root.winfo_width() * 0.9))
        win_h = max(560, int(self.root.winfo_height() * 0.8))
        x = self.root.winfo_x() + max(0, (self.root.winfo_width() - win_w) // 2)
        y = self.root.winfo_y() + max(0, (self.root.winfo_height() - win_h) // 2)
        win.geometry(f'{win_w}x{win_h}+{x}+{y}')
        win.configure(bg='#f5f5f5')
        win.transient(self.root)

        # 汇总数据
        # rows: TreeView 的原始数据源，每行对应一道题的统计快照。
        rows = []
        attempted_count = 0
        total_attempts = 0
        total_errors = 0
        type_map = {
            'single': '单选',
            'multi': '多选',
            'judge': '判断',
            'blank': '填空',
            'short': '简答'
        }

        for q in self.questions:
            rec = get_record(self.records, q)
            attempts = rec['attempts']
            errors = rec['errors']
            # error_rate: 单题错误率百分比。
            error_rate = (errors / attempts * 100) if attempts > 0 else 0.0
            if attempts > 0:
                attempted_count += 1
            total_attempts += attempts
            total_errors += errors

            rows.append({
                'id': q['id'],
                'type': type_map.get(q.get('type', ''), q.get('type', '')),
                'attempts': attempts,
                'errors': errors,
                'error_rate': error_rate,
                'text': str(q.get('text', '')).replace('\n', ' ').strip()
            })

        total_questions = len(self.questions)
        overall_error_rate = (total_errors / total_attempts * 100) if total_attempts > 0 else 0.0

        top_frame = tk.Frame(win, bg='#2c3e50', height=56)
        top_frame.pack(fill='x')
        top_frame.pack_propagate(False)

        summary_label = tk.Label(
            top_frame,
            text=(
                f'总题数 {total_questions}  |  已作答 {attempted_count}  |  总作答次数 {total_attempts}  '
                f'|  总错误次数 {total_errors}  |  总错误率 {overall_error_rate:.1f}%'
            ),
            font=self.small_font,
            fg='white',
            bg='#2c3e50',
            anchor='w'
        )
        summary_label.pack(fill='x', padx=12, pady=16)

        if self.duplicate_groups:
            dup_questions = sum(len(g) for g in self.duplicate_groups)
            dup_label = tk.Label(
                win,
                text=f'重复题检测：{len(self.duplicate_groups)} 组（共 {dup_questions} 题），系统抽题时会自动规避短时间重复同组。',
                font=self.small_font,
                fg='#8e44ad',
                bg='#f5f5f5',
                anchor='w'
            )
            dup_label.pack(fill='x', padx=12, pady=(6, 0))

        control_frame = tk.Frame(win, bg='#f5f5f5')
        control_frame.pack(fill='x', padx=12, pady=(10, 6))

        tk.Label(
            control_frame, text='排序方式：',
            font=self.small_font, bg='#f5f5f5', fg='#2c3e50'
        ).pack(side='left')

        sort_var = tk.StringVar(value='按作答次数')
        sort_combo = ttk.Combobox(
            control_frame,
            state='readonly',
            textvariable=sort_var,
            values=['按作答次数', '按错误次数', '按错误率', '按题号'],
            width=14
        )
        sort_combo.pack(side='left', padx=(6, 10))

        only_attempted_var = tk.BooleanVar(value=True)
        only_attempted_cb = tk.Checkbutton(
            control_frame,
            text='仅看已作答题目',
            variable=only_attempted_var,
            bg='#f5f5f5',
            fg='#2c3e50',
            font=self.small_font,
            activebackground='#f5f5f5'
        )
        only_attempted_cb.pack(side='left')

        def open_selected_question_detail(_event=None):
            selected = tree.selection()
            if not selected:
                messagebox.showwarning('提示', '请先选择一题。', parent=win)
                return

            try:
                qid = int(selected[0])
            except Exception:
                return

            q = self.question_map.get(qid)
            if not q:
                messagebox.showerror('错误', '未找到题目详情。', parent=win)
                return

            detail_win = tk.Toplevel(win)
            detail_win.title(f'题目详情 - 第 {qid} 题')
            d_w = max(760, int(win.winfo_width() * 0.75))
            d_h = max(520, int(win.winfo_height() * 0.75))
            x = win.winfo_x() + max(0, (win.winfo_width() - d_w) // 2)
            y = win.winfo_y() + max(0, (win.winfo_height() - d_h) // 2)
            detail_win.geometry(f'{d_w}x{d_h}+{x}+{y}')
            detail_win.configure(bg='#f5f5f5')
            detail_win.transient(win)

            type_map_rev = {
                'single': '单选',
                'multi': '多选',
                'judge': '判断',
                'blank': '填空',
                'short': '简答'
            }
            rec = get_record(self.records, q)
            error_rate = (rec['errors'] / rec['attempts'] * 100) if rec['attempts'] > 0 else 0.0

            head = tk.Label(
                detail_win,
                text=(
                    f"第 {qid} 题  |  题型：{type_map_rev.get(q.get('type', ''), q.get('type', ''))}"
                    f"  |  作答 {rec['attempts']} 次  错误 {rec['errors']} 次  错误率 {error_rate:.0f}%"
                ),
                bg='#2c3e50', fg='white', anchor='w',
                font=('Microsoft YaHei', 10)
            )
            head.pack(fill='x', ipady=10)

            text_box = tk.Text(detail_win, wrap='word', font=('Microsoft YaHei', 11), relief='flat')
            yscroll_detail = ttk.Scrollbar(detail_win, orient='vertical', command=text_box.yview)
            text_box.configure(yscrollcommand=yscroll_detail.set)
            text_box.pack(side='left', fill='both', expand=True, padx=(12, 0), pady=12)
            yscroll_detail.pack(side='left', fill='y', pady=12, padx=(0, 12))

            text_box.insert('end', '题干：\n')
            text_box.insert('end', str(q.get('text', '') or '') + '\n\n')

            options = q.get('options') or {}
            if options:
                text_box.insert('end', '选项：\n')
                for k in sorted(options.keys()):
                    text_box.insert('end', f"{k}. {options.get(k, '')}\n")
                text_box.insert('end', '\n')

            answer = q.get('answer', '')
            answer_text = ''.join(answer) if isinstance(answer, list) else str(answer)
            text_box.insert('end', f'答案：\n{answer_text}\n')
            text_box.config(state='disabled')

        tk.Button(
            control_frame,
            text='查看所选题',
            font=self.small_font,
            bg='#3498db', fg='white', activebackground='#2980b9',
            relief='flat', padx=10, pady=4,
            command=open_selected_question_detail
        ).pack(side='right')

        body = tk.Frame(win, bg='#f5f5f5')
        body.pack(fill='both', expand=True, padx=12, pady=(4, 10))

        columns = ('rank', 'id', 'type', 'attempts', 'errors', 'error_rate', 'text')
        tree = ttk.Treeview(body, columns=columns, show='headings')
        tree.heading('rank', text='排名')
        tree.heading('id', text='题号')
        tree.heading('type', text='题型')
        tree.heading('attempts', text='作答次数')
        tree.heading('errors', text='错误次数')
        tree.heading('error_rate', text='错误率')
        tree.heading('text', text='题干预览')

        tree.column('rank', width=56, anchor='center')
        tree.column('id', width=68, anchor='center')
        tree.column('type', width=72, anchor='center')
        tree.column('attempts', width=88, anchor='center')
        tree.column('errors', width=88, anchor='center')
        tree.column('error_rate', width=84, anchor='center')
        tree.column('text', width=620, anchor='w')

        yscroll = ttk.Scrollbar(body, orient='vertical', command=tree.yview)
        xscroll = ttk.Scrollbar(body, orient='horizontal', command=tree.xview)
        tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        tree.grid(row=0, column=0, sticky='nsew')
        yscroll.grid(row=0, column=1, sticky='ns')
        xscroll.grid(row=1, column=0, sticky='ew')
        body.grid_columnconfigure(0, weight=1)
        body.grid_rowconfigure(0, weight=1)

        def apply_sort(items):
            mode = sort_var.get()
            if mode == '按作答次数':
                return sorted(items, key=lambda x: (-x['attempts'], -x['errors'], x['id']))
            if mode == '按错误次数':
                return sorted(items, key=lambda x: (-x['errors'], -x['attempts'], x['id']))
            if mode == '按错误率':
                return sorted(items, key=lambda x: (-x['error_rate'], -x['attempts'], x['id']))
            return sorted(items, key=lambda x: x['id'])

        def refresh_table(_event=None):
            for item in tree.get_children():
                tree.delete(item)

            # display_rows: 当前筛选+排序后用于渲染表格的行。
            display_rows = rows
            if only_attempted_var.get():
                display_rows = [r for r in display_rows if r['attempts'] > 0]

            display_rows = apply_sort(display_rows)
            for i, r in enumerate(display_rows, 1):
                text_preview = r['text']
                if len(text_preview) > 90:
                    text_preview = text_preview[:90] + '...'
                tree.insert('', 'end', iid=str(r['id']), values=(
                    i,
                    r['id'],
                    r['type'],
                    r['attempts'],
                    r['errors'],
                    f"{r['error_rate']:.0f}%",
                    text_preview
                ))

        sort_combo.bind('<<ComboboxSelected>>', refresh_table)
        only_attempted_cb.config(command=refresh_table)
        tree.bind('<Double-1>', open_selected_question_detail)
        refresh_table()

    def _normalize_keyboard_text(self, text):
        """规范化键盘输入，兼容全角字母与中文标点。"""
        normalized = (text or '').strip().upper()
        return normalized.translate(str.maketrans('ＡＢＣＤＥＦＧＨ，。、；：　', 'ABCDEFGH,,,,  '))

    def _question_signature(self, q):
        """生成题目归一签名，用于重复题检测。"""
        text = str(q.get('text', '') or '')
        text = re.sub(r'（\s*\d+\s*）\s*[_＿﹍]+', '（）', text)
        text = re.sub(r'[_＿﹍]+', '', text)
        text = re.sub(r'[\s，,。；;：:、（）()\[\]【】]+', '', text)

        options = q.get('options') or {}
        option_sig = []
        for k in sorted(options.keys()):
            v = re.sub(r'\s+', '', str(options.get(k, '') or ''))
            option_sig.append(f'{k}:{v}')

        q_type = q.get('type', '')
        return (q_type, text, '|'.join(option_sig))

    def _build_duplicate_groups(self):
        """构建重复题分组，用于抽题去重策略。"""
        sig_map = {}
        for q in self.questions:
            sig = self._question_signature(q)
            sig_map.setdefault(sig, []).append(q['id'])

        groups = [ids for ids in sig_map.values() if len(ids) >= 2]
        groups.sort(key=lambda x: (-len(x), x[0]))
        return groups

    def _is_recent_duplicate_pick(self, q):
        """判断候选题是否与近期已抽题属于同一重复签名组。"""
        sig = self._question_signature(q)
        # 仅当该签名属于重复题组，且近期出现过，才判定为重复抽到。
        return sig in self.duplicate_signature_set and sig in self.recent_signatures

    def _build_duplicate_signature_set(self):
        """提取重复题分组的签名集合，便于快速查重。"""
        sigs = set()
        for g in self.duplicate_groups:
            for qid in g:
                q = self.question_map.get(qid)
                if q:
                    sigs.add(self._question_signature(q))
        return sigs

    def _select_objective_by_keyboard(self, token):
        """将键盘输入映射为客观题选项选中状态。"""
        if not self.current_q or self.submitted:
            return False

        q = self.current_q
        if q['type'] not in ('single', 'multi', 'judge'):
            return False

        valid_keys = sorted(q['options'].keys())
        letters = [ch for ch in token if ch in valid_keys]
        if not letters:
            return False

        if q['type'] in ('single', 'judge'):
            self.selected = {letters[-1]}
        else:
            self.selected = set(letters)

        for k, btn in self.option_buttons.items():
            if k in self.selected:
                btn.config(bg='#3498db', fg='white')
            else:
                btn.config(bg='white', fg='#2c3e50')
        return True

    def _submit_subjective_by_keyboard(self, token):
        """将 t/f 等键盘输入映射为主观题自评提交。"""
        if not self.current_q or self.submitted:
            return False

        q = self.current_q
        if q['type'] not in ('blank', 'short'):
            return False

        if not self.answer_revealed:
            return False

        true_tokens = {'T', 'TRUE', 'Y', 'YES', '对', '正确'}
        false_tokens = {'F', 'FALSE', 'N', 'NO', '错', '错误'}
        if token in true_tokens:
            self.submit_subjective_result(True)
            return True
        if token in false_tokens:
            self.submit_subjective_result(False)
            return True
        return False

    def _process_keyboard_enter(self):
        """统一处理回车行为：提交答案、主观自评或切换下一题。"""
        if self.current_q is None:
            self.next_question()
            return

        token = self._normalize_keyboard_text(self.keyboard_var.get())
        self.keyboard_var.set('')

        if token:
            q_type = self.current_q['type']
            if q_type in ('single', 'multi', 'judge'):
                if not self._select_objective_by_keyboard(token):
                    self.result_label.config(text='未识别到有效选项，请输入题目存在的字母（如 A/ABC）。', fg='#e67e22')
                    return
                if not self.submitted:
                    self.submit_answer()
                return

            if q_type in ('blank', 'short'):
                if self._submit_subjective_by_keyboard(token):
                    return
                self.result_label.config(text='主观题请输入 t/f 后回车（t=答对，f=答错）。', fg='#e67e22')
                return

        # 无输入时：未提交则提交/显示答案；已提交则下一题。
        if not self.submitted:
            self.submit_answer()
        else:
            self.next_question()

    def _on_entry_enter(self, _event=None):
        """输入框内回车事件。"""
        self._process_keyboard_enter()
        return 'break'

    def _on_global_enter(self, event=None):
        """全局回车事件（输入框已处理时不重复触发）。"""
        widget = self.root.focus_get()
        if widget is self.keyboard_entry:
            return
        self._process_keyboard_enter()

def _dedupe_existing_paths(paths):
    """去重并过滤不存在的文件路径，返回绝对路径列表。"""
    out = []
    seen = set()
    for p in paths or []:
        ap = os.path.abspath(str(p))
        if not os.path.exists(ap):
            continue
        if ap in seen:
            continue
        seen.add(ap)
        out.append(ap)
    return out

def _choose_files_incrementally(parent, initial_dir, seed_paths=None):
    """支持多次“增加文件”选择，便于跨目录导入。"""
    # chosen: 已确认存在且已去重的文件列表，支持“多轮追加选择”。
    chosen = _dedupe_existing_paths(seed_paths or [])
    current_dir = initial_dir
    if chosen:
        current_dir = os.path.dirname(chosen[-1])

    while True:
        selected_paths = filedialog.askopenfilenames(
            title='选择题库文件（可多选）',
            initialdir=current_dir,
            filetypes=[
                ('题库文件', '*.txt *.pdf *.doc *.docx'),
                ('文本文件', '*.txt'),
                ('PDF 文件', '*.pdf'),
                ('Word 文件', '*.doc *.docx'),
                ('所有文件', '*.*')
            ],
            parent=parent
        )

        if selected_paths:
            for p in selected_paths:
                ap = os.path.abspath(p)
                if ap not in chosen:
                    chosen.append(ap)
            current_dir = os.path.dirname(chosen[-1])

        if chosen:
            ans = messagebox.askyesnocancel(
                '增加文件',
                f'当前已选择 {len(chosen)} 个文件。\n是否继续增加文件？\n\n是：继续添加\n否：完成选择\n取消：放弃本次导入',
                parent=parent
            )
            if ans is True:
                continue
            if ans is False:
                return chosen
            return []

        retry = messagebox.askyesno(
            '未选择文件',
            '尚未选择任何文件，是否继续选择？',
            parent=parent
        )
        if not retry:
            return []

def _choose_startup_files(parent, script_dir):
    """启动时优先支持继续上次文件；否则进入多次添加模式。"""
    state = load_app_state()
    last_files = _dedupe_existing_paths(state.get('last_open_files', []))

    if last_files:
        preview = '\n'.join(os.path.basename(p) for p in last_files[:5])
        if len(last_files) > 5:
            preview += f'\n... 另 {len(last_files) - 5} 个'

        use_last = messagebox.askyesnocancel(
            '继续上次题库',
            f'检测到上次打开的文件：\n{preview}\n\n是否直接继续使用上次文件？\n\n是：直接继续\n否：选择是否在上次基础上追加\n取消：退出',
            parent=parent
        )
        if use_last is None:
            return []
        if use_last:
            return last_files

        append_last = messagebox.askyesno(
            '追加方式',
            '是否在上次文件基础上继续追加新文件？\n\n是：在上次基础上追加\n否：从空列表重新选择',
            parent=parent
        )
        if append_last:
            return _choose_files_incrementally(parent, script_dir, seed_paths=last_files)
        return _choose_files_incrementally(parent, script_dir)

    return _choose_files_incrementally(parent, script_dir)

def show_import_preview(root, candidates, source_path):
    """展示解析预览，用户确认后开始刷题。"""
    # 导入预览阶段也应用手动改题，确保“预览所见即实战所用”。
    manual_edits = load_manual_question_edits()
    for _, qs, _ in candidates:
        for q in qs:
            _ensure_question_identity_fields(q)
        apply_manual_question_edits(qs, manual_edits)

    win = tk.Toplevel(root)
    win.title('题库导入预览')
    screen_w = root.winfo_screenwidth()
    screen_h = root.winfo_screenheight()
    win_w = max(900, min(1700, int(screen_w * 0.82)))
    win_h = max(580, min(1100, int(screen_h * 0.8)))
    x = max(0, (screen_w - win_w) // 2)
    y = max(0, (screen_h - win_h) // 2)
    win.geometry(f'{win_w}x{win_h}+{x}+{y}')
    win.minsize(max(850, int(win_w * 0.75)), max(520, int(win_h * 0.72)))
    win.configure(bg='#f5f5f5')
    win.transient(root)
    win.grab_set()

    top = tk.Frame(win, bg='#2c3e50', height=52)
    top.pack(fill='x')
    top.pack_propagate(False)

    top_label = tk.Label(
        top, text='', fg='white', bg='#2c3e50',
        font=('Microsoft YaHei', 10), anchor='w'
    )
    top_label.pack(fill='x', padx=12, pady=14)

    body = tk.Frame(win, bg='#f5f5f5')
    body.pack(fill='both', expand=True, padx=12, pady=10)

    columns = ('id', 'type', 'answer', 'text')
    tree = ttk.Treeview(body, columns=columns, show='headings', height=22)
    tree.heading('id', text='题号')
    tree.heading('type', text='题型')
    tree.heading('answer', text='答案')
    tree.heading('text', text='题干预览')
    tree.column('id', width=70, anchor='center')
    tree.column('type', width=90, anchor='center')
    tree.column('answer', width=170, anchor='w')
    tree.column('text', width=600, anchor='w')

    type_name = {
        'single': '单选',
        'multi': '多选',
        'judge': '判断',
        'blank': '填空',
        'short': '简答'
    }

    # selected_idx: 当前选择的解析方案索引（对应 candidates 下标）。
    selected_idx = tk.IntVar(value=0)

    strategy_frame = tk.Frame(win, bg='#f5f5f5')
    strategy_frame.pack(fill='x', padx=12, pady=(0, 8))

    tk.Label(
        strategy_frame,
        text='识别方案：',
        bg='#f5f5f5', fg='#2c3e50', font=('Microsoft YaHei', 10)
    ).pack(side='left')

    strategy_names = [name for name, _, _ in candidates]
    strategy_combo = ttk.Combobox(
        strategy_frame,
        values=strategy_names,
        state='readonly',
        width=34
    )
    strategy_combo.current(0)
    strategy_combo.pack(side='left', padx=(6, 10))

    strategy_desc_label = tk.Label(
        strategy_frame,
        text='',
        bg='#f5f5f5', fg='#7f8c8d', font=('Microsoft YaHei', 9), anchor='w'
    )
    strategy_desc_label.pack(side='left', fill='x', expand=True)

    edit_bar = tk.Frame(win, bg='#f5f5f5')
    edit_bar.pack(fill='x', padx=12, pady=(0, 6))

    edit_hint_label = tk.Label(
        edit_bar,
        text='可选中题目后手动修改题目和答案（多行编辑，双击行也可编辑）。',
        bg='#f5f5f5', fg='#7f8c8d', font=('Microsoft YaHei', 9), anchor='w'
    )
    edit_hint_label.pack(side='left')

    def _get_selected_question():
        selected = tree.selection()
        if not selected:
            return None, None, None
        iid = selected[0]
        try:
            q_idx = int(iid)
        except Exception:
            return None, None, None

        # c_idx / q_idx: 当前方案下标与当前题在该方案中的题目下标。
        c_idx = selected_idx.get()
        if c_idx < 0 or c_idx >= len(candidates):
            return None, None, None
        questions = candidates[c_idx][1]
        if q_idx < 0 or q_idx >= len(questions):
            return None, None, None
        return questions[q_idx], c_idx, q_idx

    def edit_selected_question(_event=None):
        q, c_idx, q_idx = _get_selected_question()
        if q is None:
            messagebox.showwarning('提示', '请先在预览表中选择一题。', parent=win)
            return

        edited = _show_question_edit_dialog(win, q, title='修改所选题（题目+答案）')
        if not edited:
            return

        new_question_text, parsed, new_type, new_options = edited
        q['text'] = new_question_text
        q['answer'] = parsed
        q['type'] = new_type
        q['options'] = dict(new_options or {})
        upsert_manual_question_edit(manual_edits, q)
        refresh_summary(c_idx)
        tree.selection_set(str(q_idx))
        tree.focus(str(q_idx))
        messagebox.showinfo('完成', '该题修改已保存，下次启动仍会生效。', parent=win)

    def fill_tree(questions):
        for item in tree.get_children():
            tree.delete(item)

        for q_idx, q in enumerate(questions):
            if isinstance(q.get('answer'), list):
                answer_preview = ''.join(q['answer'])
            else:
                answer_preview = str(q.get('answer', ''))
            if not answer_preview.strip():
                answer_preview = '（未识别）'
            answer_preview = answer_preview.replace('\n', ' / ').strip()
            text_preview = str(q.get('text', '')).replace('\n', ' ').strip()
            if len(text_preview) > 80:
                text_preview = text_preview[:80] + '...'
            if len(answer_preview) > 32:
                answer_preview = answer_preview[:32] + '...'

            tree.insert('', 'end', iid=str(q_idx), values=(
                q.get('id', ''),
                type_name.get(q.get('type'), q.get('type', '')),
                answer_preview,
                text_preview
            ))

    def refresh_summary(idx):
        name, questions, desc = candidates[idx]
        type_count = {'single': 0, 'multi': 0, 'judge': 0, 'blank': 0, 'short': 0}
        for q in questions:
            t = q.get('type')
            if t in type_count:
                type_count[t] += 1

        summary_text = (
            f"文件：{os.path.basename(source_path)}  |  方案：{name}  |  共 {len(questions)} 题"
            f"  |  单选 {type_count['single']}  多选 {type_count['multi']}  判断 {type_count['judge']}"
            f"  填空 {type_count['blank']}  简答 {type_count['short']}"
        )
        top_label.config(text=summary_text)
        strategy_desc_label.config(text=desc)
        fill_tree(questions)

    def on_strategy_change(_event=None):
        idx = strategy_combo.current()
        if idx < 0:
            idx = 0
        selected_idx.set(idx)
        refresh_summary(idx)

    strategy_combo.bind('<<ComboboxSelected>>', on_strategy_change)
    tree.bind('<Double-1>', edit_selected_question)

    tk.Button(
        edit_bar, text='修改所选题（题目+答案）',
        font=('Microsoft YaHei', 9),
        bg='#f39c12', fg='white', activebackground='#d68910',
        relief='flat', padx=10, pady=4,
        command=edit_selected_question
    ).pack(side='right')

    refresh_summary(0)

    yscroll = ttk.Scrollbar(body, orient='vertical', command=tree.yview)
    xscroll = ttk.Scrollbar(body, orient='horizontal', command=tree.xview)
    tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

    tree.grid(row=0, column=0, sticky='nsew')
    yscroll.grid(row=0, column=1, sticky='ns')
    xscroll.grid(row=1, column=0, sticky='ew')
    body.grid_columnconfigure(0, weight=1)
    body.grid_rowconfigure(0, weight=1)

    hint = tk.Label(
        win,
        text='请检查题型与答案是否合理。确认无误后点击「开始刷题」。',
        bg='#f5f5f5', fg='#7f8c8d', font=('Microsoft YaHei', 10)
    )
    hint.pack(fill='x', padx=12, pady=(0, 8))

    result = {'ok': False, 'questions': None}

    btn_bar = tk.Frame(win, bg='#f5f5f5')
    btn_bar.pack(fill='x', padx=12, pady=(0, 12))

    def on_cancel():
        result['ok'] = False
        win.destroy()

    def on_confirm():
        result['ok'] = True
        idx = selected_idx.get()
        if idx < 0 or idx >= len(candidates):
            idx = 0
        result['questions'] = candidates[idx][1]
        win.destroy()

    tk.Button(
        btn_bar, text='取消',
        font=('Microsoft YaHei', 10),
        bg='#e0e0e0', fg='#2c3e50',
        relief='flat', padx=16, pady=6,
        command=on_cancel
    ).pack(side='right', padx=(8, 0))

    tk.Button(
        btn_bar, text='开始刷题',
        font=('Microsoft YaHei', 10, 'bold'),
        bg='#2ecc71', fg='white',
        activebackground='#27ae60',
        relief='flat', padx=16, pady=6,
        command=on_confirm
    ).pack(side='right')

    win.protocol('WM_DELETE_WINDOW', on_cancel)
    root.wait_window(win)
    if result['ok']:
        return result['questions']
    return None

def _enable_windows_high_dpi():
    """启用 Windows 高 DPI 感知，避免缩放模糊。"""
    if os.name != 'nt':
        return

    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass
