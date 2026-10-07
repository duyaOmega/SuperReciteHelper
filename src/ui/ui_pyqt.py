#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PyQt6 主刷题窗口。"""

import os
import re
from functools import partial
from string import Template

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QApplication,
    QButtonGroup,
    QCheckBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from src.core import *
from src.ui import *
from src.ui import theme


def normalize_keyboard_text(text): # 规范化键盘输入，兼容全角字母。
    normalized = (text or "").strip().upper()
    return normalized.translate(str.maketrans("ＡＢＣＤＥＦＧＨ，。、；：　", "ABCDEFGH,,,,  "))

#--------------------定义一堆 PyQt 样式模板（按主题 token 生成）-------------------------------
_T = Template

_STYLE_TEMPLATES = {
    "btn_header": _T(
        "QPushButton { background: transparent; color: $muted; border: none;"
        " font-size: 13px; padding: 5px 10px; border-radius: 6px; }"
        "QPushButton:hover { color: $text_secondary; background: $surface_alt; }"
    ),
    "btn_ghost": _T(
        "QPushButton { background: $surface_alt; color: $text_secondary; border: 1px solid $input_border;"
        " border-radius: 8px; padding: 7px 16px; font-size: 13px; font-weight: 500; }"
        "QPushButton:hover { background: $surface_hover; }"
        "QPushButton:disabled { color: $faint; }"
    ),
    "btn_primary": _T(
        "QPushButton { background: $primary; color: white; border: none;"
        " border-radius: 8px; padding: 8px 22px; font-size: 14px; font-weight: 600; }"
        "QPushButton:hover { background: $primary_hover; }"
        "QPushButton:disabled { background: $primary_disabled; color: white; }"
    ),
    "card_default": _T("QFrame { background: $surface; border: 1.5px solid $border; border-radius: 10px; }"),
    "card_selected": _T("QFrame { background: $surface; border: 2px solid $primary; border-radius: 10px; }"),
    "card_correct": _T("QFrame { background: $ok_bg; border: 1.5px solid $ok_border; border-radius: 10px; }"),
    "card_wrong": _T("QFrame { background: $err_bg; border: 1.5px solid $err_border; border-radius: 10px; }"),
    "card_dimmed": _T("QFrame { background: $bg; border: 1.5px solid $border; border-radius: 10px; }"),
    "header_bar": _T("background: $surface; border-bottom: 1px solid $border;"),
    "bottom_bar": _T("background: $surface; border-top: 1px solid $border;"),
    "prog_row": _T("background: $surface; padding-bottom: 2px;"),
    "content_area": _T("background: $bg;"),
    "progress": _T(
        "QProgressBar { background: $surface_alt; border-radius: 3px; border: none; }"
        "QProgressBar::chunk { background: $primary; border-radius: 3px; }"
    ),
    "source_label": _T("font-size: 13px; color: $muted;"),
    "progress_label": _T("font-size: 13px; color: $text_secondary;"),
    "accuracy_label": _T("font-size: 13px; color: $muted;"),
    "title_label": _T("font-size: 20px; font-weight: 700; color: $text;"),
    "type_badge": _T(
        "background: $accent_bg; color: $accent_text; font-size: 12px; font-weight: 600;"
        " padding: 2px 10px; border-radius: 10px;"
    ),
    "history_label": _T("font-size: 13px; color: $faint;"),
    "question_text": _T(
        "background: $surface_alt; border-radius: 10px; padding: 16px 18px;"
        " font-size: 15px; color: $text;"
    ),
    "result_base": _T("font-size: 14px;"),
    "result_info": _T("font-size: 14px; color: $muted;"),
    "result_reveal": _T("font-size: 14px; color: $text_secondary;"),
    "result_ok": _T("font-size: 14px; color: $ok_text; font-weight: 700;"),
    "result_err": _T("font-size: 14px; color: $err_text; font-weight: 700;"),
    "result_warn": _T("font-size: 14px; color: $warn_text;"),
    "hint_label": _T("font-size: 13px; color: $subtle;"),
    "keyboard_entry": _T(
        "border: 1px solid $input_border; border-radius: 6px; padding: 5px 10px;"
        " font-size: 13px; color: $text_secondary; background: $surface;"
    ),
    "opt_indicator": _T(
        "QRadioButton::indicator { width: 17px; height: 17px; }"
        " QCheckBox::indicator { width: 17px; height: 17px; }"
    ),
    "opt_label": _T("font-size: 14px; color: $text; background: transparent; border: none;"),
    "opt_label_correct": _T("font-size: 14px; color: $ok_text; font-weight: 600; background: transparent; border: none;"),
    "opt_label_wrong": _T("font-size: 14px; color: $err_text; font-weight: 600; background: transparent; border: none;"),
    "opt_label_dimmed": _T("font-size: 14px; color: $faint; background: transparent; border: none;"),
    "subj_ok_btn": _T(
        "QPushButton { background: $ok_bg; color: $ok_text; border: 1.5px solid $ok_border;"
        " border-radius: 8px; padding: 7px 18px; font-size: 13px; font-weight: 600; }"
        "QPushButton:hover { background: $ok_bg_hover; }"
    ),
    "subj_err_btn": _T(
        "QPushButton { background: $err_bg; color: $err_text; border: 1.5px solid $err_border;"
        " border-radius: 8px; padding: 7px 18px; font-size: 13px; font-weight: 600; }"
        "QPushButton:hover { background: $err_bg_hover; }"
    ),
}


def window_styles(tk):
    """按当前主题 token 生成主窗口使用的全部样式字符串。"""
    return {key: tpl.substitute(tk) for key, tpl in _STYLE_TEMPLATES.items()}

#-------------------------PyQt主刷题窗口--------------------------
class QuizWindow(QMainWindow):
    def __init__(self, questions, source_path=""):
        super().__init__() #调用父类 QMainWindow 的初始化函数
        self.questions = list(questions or [])
        self.manual_edits = load_manual_question_edits()
        for q in self.questions:
            ensure_question_identity_fields(q)
        apply_manual_question_edits(self.questions, self.manual_edits)

        self.question_map = {q.get("id"): q for q in self.questions}
        self.source_path = source_path
        self.source_name = os.path.basename(source_path) if source_path else "未命名题库"
        self.records = load_records()

        self.current_q = None
        self.submitted = False
        self.answer_revealed = False
        self.option_widgets = {}
        self.option_cards = {}
        self.option_group = None
        self._graded = None       # 客观题判分结果，切主题重绘时需要
        self._subj_state = None   # 主观题展示状态（revealed / graded）

        self.setWindowTitle("SuperReciteHelper")
        self.ui_scale = self._get_ui_scale()
        self._apply_scaled_window_geometry()
        self._build_ui()
        self._show_welcome()

    def _get_ui_scale(self):
        screen = QApplication.primaryScreen()
        if not screen:
            return 1.0
        geo = screen.availableGeometry()
        dpr = screen.devicePixelRatio()
        physical_w = geo.width() * dpr
        physical_h = geo.height() * dpr
        scale_w = physical_w / 1920
        scale_h = physical_h / 1080
        resolution_scale = (scale_w + scale_h) / 2
        dpi_scale = screen.logicalDotsPerInch() / 96
        return max(1.0, min(1.5, max(resolution_scale, dpi_scale)))

    def _px(self, value):
        return max(1, int(round(value * self.ui_scale)))

    def _style(self, style_text):
        def repl(match):
            return f"{self._px(float(match.group(1)))}px"

        return re.sub(r'(\d+(?:\.\d+)?)px', repl, style_text)

    def _apply_scaled_window_geometry(self):
        self.resize(960, 720)
        self.setMinimumSize(760, 560)

    def _build_ui(self):
        self.styles = window_styles(theme.tokens())

        root = QWidget()
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        # ── Header bar ──────────────────────────────────────────────
        header = QWidget()
        header.setStyleSheet(self.styles["header_bar"])
        hl = QHBoxLayout(header)
        hl.setContentsMargins(self._px(20), self._px(10), self._px(16), self._px(10))
        hl.setSpacing(self._px(4))

        self.source_label = QLabel(f"题库：{self.source_name}")
        self.source_label.setStyleSheet(self._style(self.styles["source_label"]))
        hl.addWidget(self.source_label)
        hl.addStretch(1)

        self.edit_btn = QPushButton("✏  编辑")
        self.manage_edits_btn = QPushButton("管理修改")
        self.stats_btn = QPushButton("📊  统计")
        self.theme_btn = QPushButton("☀️  浅色" if theme.active_name() == "dark" else "🌙  深色")
        self.reset_btn = QPushButton("↺  重置")
        for btn in (self.edit_btn, self.manage_edits_btn, self.stats_btn, self.theme_btn, self.reset_btn):
            btn.setStyleSheet(self._style(self.styles["btn_header"]))
            hl.addWidget(btn)
        self.edit_btn.clicked.connect(self.edit_current_question)
        self.manage_edits_btn.clicked.connect(self.manage_manual_edits)
        self.stats_btn.clicked.connect(self.show_frequency_stats)
        self.theme_btn.clicked.connect(self.switch_theme)
        self.reset_btn.clicked.connect(self.reset_records)
        root_layout.addWidget(header)

        # ── Progress row ─────────────────────────────────────────────
        prog_row = QWidget()
        prog_row.setStyleSheet(self._style(self.styles["prog_row"]))
        pl = QHBoxLayout(prog_row)
        pl.setContentsMargins(self._px(20), self._px(6), self._px(20), self._px(10))
        pl.setSpacing(self._px(10))

        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedHeight(self._px(5))
        self.progress_bar.setStyleSheet(self._style(self.styles["progress"]))
        pl.addWidget(self.progress_bar, 1)

        self.progress_label = QLabel("0/0")
        self.progress_label.setStyleSheet(self._style(self.styles["progress_label"]))
        pl.addWidget(self.progress_label)

        self.accuracy_label = QLabel("正确率 —")
        self.accuracy_label.setStyleSheet(self._style(self.styles["accuracy_label"]))
        pl.addWidget(self.accuracy_label)
        root_layout.addWidget(prog_row)

        # ── Scroll area ──────────────────────────────────────────────
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFrameShape(QScrollArea.Shape.NoFrame)
        self.scroll_area.setStyleSheet(self.styles["content_area"])

        content = QWidget()
        content.setStyleSheet(self.styles["content_area"])
        self.content_layout = QVBoxLayout(content)
        self.content_layout.setContentsMargins(self._px(28), self._px(28), self._px(28), self._px(28))
        self.content_layout.setSpacing(self._px(14))

        # Question header row
        q_header = QHBoxLayout()
        q_header.setSpacing(self._px(8))
        self.title_label = QLabel()
        self.title_label.setStyleSheet(self._style(self.styles["title_label"]))
        q_header.addWidget(self.title_label)

        self.type_badge = QLabel()
        self.type_badge.setStyleSheet(self._style(self.styles["type_badge"]))
        q_header.addWidget(self.type_badge)

        self.history_label = QLabel()
        self.history_label.setStyleSheet(self._style(self.styles["history_label"]))
        q_header.addWidget(self.history_label)
        q_header.addStretch(1)
        self.content_layout.addLayout(q_header)

        # Question text card
        self.question_label = QLabel()
        self.question_label.setWordWrap(True)
        self.question_label.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        self.question_label.setStyleSheet(self._style(self.styles["question_text"]))
        self.content_layout.addWidget(self.question_label)

        # Options container
        self.options_container = QWidget()
        self.options_container.setStyleSheet("background: transparent;")
        self.options_layout = QVBoxLayout(self.options_container)
        self.options_layout.setContentsMargins(0, 0, 0, 0)
        self.options_layout.setSpacing(self._px(8))
        self.content_layout.addWidget(self.options_container)

        # Result label
        self.result_label = QLabel()
        self.result_label.setWordWrap(True)
        self.result_label.setStyleSheet(self._style(self.styles["result_base"]))
        self.content_layout.addWidget(self.result_label)

        self.content_layout.addStretch(1)
        self.scroll_area.setWidget(content)
        root_layout.addWidget(self.scroll_area, 1)

        # ── Bottom bar ───────────────────────────────────────────────
        bottom = QWidget()
        bottom.setStyleSheet(self.styles["bottom_bar"])
        bl = QHBoxLayout(bottom)
        bl.setContentsMargins(self._px(20), self._px(10), self._px(20), self._px(10))
        bl.setSpacing(self._px(8))

        self.hint_label = QLabel("按 A–D 选择，Enter 提交")
        self.hint_label.setStyleSheet(self._style(self.styles["hint_label"]))
        bl.addWidget(self.hint_label)

        self.keyboard_entry = QLineEdit()
        self.keyboard_entry.setPlaceholderText("键盘输入")
        self.keyboard_entry.setFixedWidth(self._px(88))
        self.keyboard_entry.setStyleSheet(self._style(self.styles["keyboard_entry"]))
        self.keyboard_entry.returnPressed.connect(self._process_keyboard_enter)
        bl.addWidget(self.keyboard_entry)

        bl.addStretch(1)

        self.next_btn = QPushButton("下一题")
        self.next_btn.setStyleSheet(self._style(self.styles["btn_ghost"]))
        self.next_btn.clicked.connect(self.next_question)
        bl.addWidget(self.next_btn)

        self.submit_btn = QPushButton("提交答案")
        self.submit_btn.setStyleSheet(self._style(self.styles["btn_primary"]))
        self.submit_btn.clicked.connect(self.submit_answer)
        bl.addWidget(self.submit_btn)

        root_layout.addWidget(bottom)
        self.setCentralWidget(root)

    def _show_welcome(self):
        self.current_q = None
        self.submitted = False
        self.answer_revealed = False
        self._graded = None
        self._subj_state = None
        self._clear_options()

        self.title_label.setText("SuperReciteHelper")
        self.type_badge.setText("")
        self.type_badge.setVisible(False)
        self.question_label.setText(f'共 {len(self.questions)} 题，点击“下一题”开始作答。')
        self.result_label.setText("")
        self.history_label.setText("")
        self.keyboard_entry.clear()
        self.submit_btn.setEnabled(False)
        self._update_stats()

    def _update_stats(self):
        attempted = 0
        total_attempts = 0
        total_errors = 0
        for q in self.questions:
            rec = get_record(self.records, q)
            attempts = int(rec.get("attempts", 0) or 0)
            errors = int(rec.get("errors", 0) or 0)
            if attempts > 0:
                attempted += 1
            total_attempts += attempts
            total_errors += errors

        total = len(self.questions)
        self.progress_bar.setMaximum(max(total, 1))
        self.progress_bar.setValue(attempted)
        self.progress_label.setText(f"{attempted}/{total}")
        if total_attempts:
            accuracy = (1 - total_errors / total_attempts) * 100
            self.accuracy_label.setText(f"正确率 {accuracy:.0f}%")
        else:
            self.accuracy_label.setText("正确率 —")

    def next_question(self):
        if not self.questions:
            QMessageBox.warning(self, "提示", "当前没有可用题目。")
            return

        picked = weighted_random_pick(self.questions, self.records)
        self.current_q = picked
        self.submitted = False
        self.answer_revealed = False
        self.keyboard_entry.clear()
        self._display_question()

    def _make_option_card(self, key, text, q_type):
        card = QFrame()
        card.setStyleSheet(self._style(self.styles["card_default"]))
        card.setCursor(Qt.CursorShape.PointingHandCursor)

        card_layout = QHBoxLayout(card)
        card_layout.setContentsMargins(self._px(14), self._px(10), self._px(14), self._px(10))
        card_layout.setSpacing(self._px(12))

        if q_type in ("single", "judge"):
            btn = QRadioButton()
        else:
            btn = QCheckBox()
        btn.setStyleSheet(self._style(self.styles["opt_indicator"]))
        card_layout.addWidget(btn)

        lbl = QLabel(f"{key}. {text}")
        lbl.setWordWrap(True)
        lbl.setStyleSheet(self._style(self.styles["opt_label"]))
        card_layout.addWidget(lbl, 1)

        btn.toggled.connect(partial(self._on_option_toggled, key))

        if q_type in ("single", "judge"):
            card.mousePressEvent = lambda _e, b=btn: b.setChecked(True)
        else:
            card.mousePressEvent = lambda _e, b=btn: b.setChecked(not b.isChecked())

        return card, btn

    def _on_option_toggled(self, key, checked):
        card = self.option_cards.get(key)
        if card and not self.submitted:
            name = "card_selected" if checked else "card_default"
            card.setStyleSheet(self._style(self.styles[name]))

    def _display_question(self):
        q = self.current_q
        self._clear_options()
        self._graded = None
        self._subj_state = None

        q_type = q.get("type", "")
        self.title_label.setText(f"第 {q.get('id', '')} 题")

        type_text = TYPE_LABELS.get(q_type, q_type)
        self.type_badge.setText(type_text)
        self.type_badge.setVisible(bool(type_text))

        question_text = str(q.get("text", "") or "")
        if q_type == "blank":
            question_text = mask_blank_question_text(question_text, q.get("answer", ""))
        self.question_label.setText(question_text)

        rec = get_record(self.records, q)
        if rec.get("attempts", 0):
            attempts = int(rec.get("attempts", 0) or 0)
            errors = int(rec.get("errors", 0) or 0)
            self.history_label.setText(f"已做 {attempts} 次 · 正确 {attempts - errors} 次")
        else:
            self.history_label.setText("首次作答")

        if q_type in ("single", "judge"):
            self.option_group = QButtonGroup(self)
            self.option_group.setExclusive(True)
            for key, text in sorted((q.get("options") or {}).items()):
                card, btn = self._make_option_card(key, text, q_type)
                self.options_layout.addWidget(card)
                self.option_group.addButton(btn)
                self.option_widgets[key] = btn
                self.option_cards[key] = card
            self.submit_btn.setText("提交答案")
            self.submit_btn.setEnabled(True)
            self.result_label.setText("")
            self.hint_label.setText("按 A–D 选择，Enter 提交")
            self.keyboard_entry.setPlaceholderText("输入 A/B/C/D")
        elif q_type == "multi":
            for key, text in sorted((q.get("options") or {}).items()):
                card, btn = self._make_option_card(key, text, q_type)
                self.options_layout.addWidget(card)
                self.option_widgets[key] = btn
                self.option_cards[key] = card
            self.submit_btn.setText("提交答案")
            self.submit_btn.setEnabled(True)
            self.result_label.setText("")
            self.hint_label.setText("按字母多选（如 ABC），Enter 提交")
            self.keyboard_entry.setPlaceholderText("输入 ABC")
        else:
            self.result_label.setStyleSheet(self._style(self.styles["result_info"]))
            self.result_label.setText('先自行作答，然后点击"显示答案"。')
            self.submit_btn.setText("显示答案")
            self.submit_btn.setEnabled(True)
            self.hint_label.setText("Enter 显示答案，再输入 t/f 自评")
            self.keyboard_entry.setPlaceholderText("t / f 自评")

        self.scroll_area.verticalScrollBar().setValue(0)
        self.keyboard_entry.setFocus()

    def submit_answer(self):
        if not self.current_q or self.submitted:
            return

        q = self.current_q
        q_type = q.get("type")

        if q_type in ("blank", "short"):
            if not self.answer_revealed:
                self.answer_revealed = True
                self._subj_state = ("revealed",)
                self._render_subjective_result(self._subj_state)
            return

        selected = self._selected_options()
        if not selected:
            QMessageBox.information(self, "提示", "请先选择答案。")
            return

        answer = q.get("answer") or []
        correct = set(answer if isinstance(answer, list) else [answer])
        if not correct:
            QMessageBox.warning(self, "提示", "本题没有标准答案，暂无法自动判分。")
            return

        is_correct = selected == correct
        update_record(self.records, q, is_correct)
        self.submitted = True
        self.submit_btn.setEnabled(False)
        self._update_stats()
        self._graded = {"correct": correct, "selected": selected, "correct_flag": is_correct}
        self._mark_objective_result(correct, selected, is_correct)

    def _selected_options(self):
        return {key for key, widget in self.option_widgets.items() if widget.isChecked()}

    def _mark_objective_result(self, correct, selected, is_correct):
        for key, card in self.option_cards.items():
            lbl = card.findChild(QLabel)
            if key in correct:
                card.setStyleSheet(self._style(self.styles["card_correct"]))
                if lbl:
                    lbl.setStyleSheet(self._style(self.styles["opt_label_correct"]))
            elif key in selected:
                card.setStyleSheet(self._style(self.styles["card_wrong"]))
                if lbl:
                    lbl.setStyleSheet(self._style(self.styles["opt_label_wrong"]))
            else:
                card.setStyleSheet(self._style(self.styles["card_dimmed"]))
                if lbl:
                    lbl.setStyleSheet(self._style(self.styles["opt_label_dimmed"]))

        if is_correct:
            self.result_label.setStyleSheet(self._style(self.styles["result_ok"]))
            self.result_label.setText("回答正确！")
        else:
            self.result_label.setStyleSheet(self._style(self.styles["result_err"]))
            self.result_label.setText(f"回答错误。正确答案：{''.join(sorted(correct))}")

        rec = get_record(self.records, self.current_q)
        attempts = int(rec.get("attempts", 0) or 0)
        errors = int(rec.get("errors", 0) or 0)
        self.history_label.setText(f"已做 {attempts} 次 · 正确 {attempts - errors} 次")

    def _add_subjective_buttons(self):
        row = QWidget()
        row.setStyleSheet("background: transparent;")
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, self._px(4), 0, 0)
        layout.setSpacing(self._px(10))

        correct_btn = QPushButton("✓  我答对了")
        correct_btn.setStyleSheet(self._style(self.styles["subj_ok_btn"]))
        correct_btn.clicked.connect(lambda: self._submit_subjective_result(True))
        layout.addWidget(correct_btn)

        wrong_btn = QPushButton("✗  我答错了")
        wrong_btn.setStyleSheet(self._style(self.styles["subj_err_btn"]))
        wrong_btn.clicked.connect(lambda: self._submit_subjective_result(False))
        layout.addWidget(wrong_btn)

        layout.addStretch(1)
        self.options_layout.addWidget(row)

    def _render_subjective_result(self, state):
        """渲染主观题的展示状态：("revealed",) 已显示答案 / ("graded", bool) 已自评。"""
        answer = format_answer_text(self.current_q.get("answer"))
        if state[0] == "revealed":
            self.result_label.setStyleSheet(self._style(self.styles["result_reveal"]))
            self.result_label.setText(f"参考答案：{answer}")
            self._add_subjective_buttons()
            self.submit_btn.setEnabled(False)
        else:
            is_correct = bool(state[1])
            name = "result_ok" if is_correct else "result_err"
            self.result_label.setStyleSheet(self._style(self.styles[name]))
            self.result_label.setText(f"参考答案：{answer}\n已记录：{'答对' if is_correct else '答错'}。")
            self.submit_btn.setEnabled(False)
            for widget in self.option_widgets.values():
                widget.setEnabled(False)

    def _submit_subjective_result(self, is_correct):
        if not self.current_q or self.submitted:
            return

        update_record(self.records, self.current_q, is_correct)
        self.submitted = True
        self._update_stats()
        self._subj_state = ("graded", is_correct)
        self._render_subjective_result(self._subj_state)

    def reset_records(self):
        reply = QMessageBox.question(
            self,
            "确认",
            "确定要重置所有做题记录吗？\n此操作不可撤销。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        self.records = {}
        save_records(self.records)
        self._update_stats()
        if self.current_q:
            self._display_question()
        QMessageBox.information(self, "完成", "所有记录已重置。")

    def edit_current_question(self):
        if not self.current_q:
            QMessageBox.information(self, "提示", '请先点击"下一题"抽取题目。')
            return

        edited = show_question_edit_dialog(self, self.current_q, "编辑当前题")
        if not edited:
            return

        new_text, parsed_answer, new_type, new_options = edited
        self.current_q["text"] = new_text
        self.current_q["answer"] = parsed_answer
        self.current_q["type"] = new_type
        self.current_q["options"] = dict(new_options or {})
        upsert_manual_question_edit(self.manual_edits, self.current_q)
        self.submitted = False
        self.answer_revealed = False
        self._display_question()
        QMessageBox.information(self, "完成", "当前题修改已保存。")

    def manage_manual_edits(self):
        show_manual_edits_dialog(
            self,
            self.questions,
            self.manual_edits,
            current_q=self.current_q,
            on_refresh_current=self._display_question if self.current_q else None,
        )

    def show_frequency_stats(self):
        show_frequency_stats_dialog(
            self,
            self.questions,
            self.records,
            self.question_map,
        )

    def switch_theme(self):
        app = QApplication.instance()
        if app is None:
            return
        name = "light" if theme.active_name() == "dark" else "dark"
        theme.apply(app, name)
        state = load_app_state()
        state["theme"] = name
        save_app_state(state)
        self._rebuild_for_theme()

    def _rebuild_for_theme(self):
        """主题切换后重建界面，并保留当前题目的作答状态。"""
        snapshot = self._view_snapshot()
        if self.option_group is not None:
            self.option_group.setParent(None)
        self._build_ui()
        self._restore_view(snapshot)

    def _view_snapshot(self):
        if self.current_q is None:
            return None
        return {
            "selected": self._selected_options(),
            "graded": self._graded,
            "subj": self._subj_state,
        }

    def _restore_view(self, snapshot):
        self._update_stats()
        if snapshot is None:
            self.current_q = None
            self._show_welcome()
            return

        self.submitted = False
        self.answer_revealed = False
        self._display_question()
        for key, widget in self.option_widgets.items():
            widget.setChecked(key in snapshot["selected"])

        self._graded = snapshot["graded"]
        self._subj_state = snapshot["subj"]
        if self._graded:
            self.submitted = True
            self.submit_btn.setEnabled(False)
            g = self._graded
            self._mark_objective_result(g["correct"], g["selected"], g["correct_flag"])
        elif self._subj_state:
            self.answer_revealed = True
            self.submitted = self._subj_state[0] == "graded"
            self._render_subjective_result(self._subj_state)

    def _select_objective_by_keyboard(self, token):
        if not self.current_q or self.submitted:
            return False
        if self.current_q.get("type") not in ("single", "multi", "judge"):
            return False

        valid_keys = sorted((self.current_q.get("options") or {}).keys())
        letters = [ch for ch in token if ch in valid_keys]
        if not letters:
            return False

        if self.current_q.get("type") in ("single", "judge"):
            target = {letters[-1]}
        else:
            target = set(letters)

        for key, widget in self.option_widgets.items():
            widget.setChecked(key in target)
        return True

    def _submit_subjective_by_keyboard(self, token):
        if not self.current_q or self.submitted:
            return False
        if self.current_q.get("type") not in ("blank", "short") or not self.answer_revealed:
            return False

        true_tokens = {"T", "TRUE", "Y", "YES", "对", "正确"}
        false_tokens = {"F", "FALSE", "N", "NO", "错", "错误"}
        if token in true_tokens:
            self._submit_subjective_result(True)
            return True
        if token in false_tokens:
            self._submit_subjective_result(False)
            return True
        return False

    def _process_keyboard_enter(self):
        if self.current_q is None:
            self.next_question()
            return

        token = normalize_keyboard_text(self.keyboard_entry.text())
        self.keyboard_entry.clear()

        if token:
            q_type = self.current_q.get("type")
            if q_type in ("single", "multi", "judge"):
                if not self._select_objective_by_keyboard(token):
                    self.result_label.setStyleSheet(self._style(self.styles["result_warn"]))
                    self.result_label.setText("未识别到有效选项，请输入题目存在的字母。")
                    return
                self.submit_answer()
                return
            if q_type in ("blank", "short"):
                if self._submit_subjective_by_keyboard(token):
                    return
                self.result_label.setStyleSheet(self._style(self.styles["result_warn"]))
                self.result_label.setText("主观题请在显示答案后输入 t/f 自评。")
                return

        if not self.submitted:
            self.submit_answer()
        else:
            self.next_question()

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            focused = QApplication.focusWidget()
            if focused is not self.keyboard_entry:
                self._process_keyboard_enter()
                return
        super().keyPressEvent(event)

    def _clear_options(self):
        while self.options_layout.count():
            item = self.options_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()
        self.option_widgets = {}
        self.option_cards = {}
        self.option_group = None
        self.result_label.setStyleSheet(self._style(self.styles["result_base"]))
