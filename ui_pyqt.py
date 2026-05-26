#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PyQt6 主刷题窗口。"""

import os
from functools import partial

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

from parser import _mask_blank_question_text
from question import _format_answer_text
from question_bank import (
    _ensure_question_identity_fields,
    apply_manual_question_edits,
    get_record,
    load_manual_question_edits,
    load_records,
    save_records,
    update_record,
    upsert_manual_question_edit,
)
from session import weighted_random_pick
from ui_pyqt_dialogs import (
    show_frequency_stats_dialog,
    show_manual_edits_dialog,
    show_question_edit_dialog,
)
from ui_pyqt_dialogs import TYPE_LABELS


def normalize_keyboard_text(text): # 规范化键盘输入，兼容全角字母。
    normalized = (text or "").strip().upper()
    return normalized.translate(str.maketrans("ＡＢＣＤＥＦＧＨ，。、；：　", "ABCDEFGH,,,,  "))

#--------------------定义一堆 PyQt 样式字符串-------------------------------
_BTN_HEADER = (
    "QPushButton { background: transparent; color: #667085; border: none;"
    " font-size: 13px; padding: 5px 10px; border-radius: 6px; }"
    "QPushButton:hover { color: #344054; background: #f2f4f7; }"
)
_BTN_GHOST = (
    "QPushButton { background: #f2f4f7; color: #344054; border: 1px solid #d0d5dd;"
    " border-radius: 8px; padding: 7px 16px; font-size: 13px; font-weight: 500; }"
    "QPushButton:hover { background: #e4e7ec; }"
    "QPushButton:disabled { color: #98a2b3; }"
)
_BTN_PRIMARY = (
    "QPushButton { background: #1570ef; color: white; border: none;"
    " border-radius: 8px; padding: 8px 22px; font-size: 14px; font-weight: 600; }"
    "QPushButton:hover { background: #175cd3; }"
    "QPushButton:disabled { background: #b2ccff; color: white; }"
)
_CARD_DEFAULT = "QFrame { background: white; border: 1.5px solid #e4e7ec; border-radius: 10px; }"
_CARD_SELECTED = "QFrame { background: white; border: 2px solid #1570ef; border-radius: 10px; }"
_CARD_CORRECT  = "QFrame { background: #ecfdf3; border: 1.5px solid #6ce9a6; border-radius: 10px; }"
_CARD_WRONG    = "QFrame { background: #fff1f0; border: 1.5px solid #fca5a5; border-radius: 10px; }"
_CARD_DIMMED   = "QFrame { background: #f9fafb; border: 1.5px solid #e4e7ec; border-radius: 10px; }"

#-------------------------PyQt主刷题窗口--------------------------
class QuizWindow(QMainWindow):
    def __init__(self, questions, source_path=""):
        super().__init__() #调用父类 QMainWindow 的初始化函数
        self.questions = list(questions or [])
        self.manual_edits = load_manual_question_edits()
        for q in self.questions:
            _ensure_question_identity_fields(q)
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

        self.setWindowTitle("SuperReciteHelper")
        self.resize(960, 720)
        self._build_ui()
        self._show_welcome()

    def _build_ui(self):
        root = QWidget()
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        # ── Header bar ──────────────────────────────────────────────
        header = QWidget()
        header.setStyleSheet("background: white; border-bottom: 1px solid #e4e7ec;")
        hl = QHBoxLayout(header)
        hl.setContentsMargins(20, 10, 16, 10)
        hl.setSpacing(4)

        self.source_label = QLabel(f"题库：{self.source_name}")
        self.source_label.setStyleSheet("font-size: 13px; color: #667085;")
        hl.addWidget(self.source_label)
        hl.addStretch(1)

        self.edit_btn = QPushButton("✏  编辑")
        self.manage_edits_btn = QPushButton("管理修改")
        self.stats_btn = QPushButton("📊  统计")
        self.reset_btn = QPushButton("↺  重置")
        for btn in (self.edit_btn, self.manage_edits_btn, self.stats_btn, self.reset_btn):
            btn.setStyleSheet(_BTN_HEADER)
            hl.addWidget(btn)
        self.edit_btn.clicked.connect(self.edit_current_question)
        self.manage_edits_btn.clicked.connect(self.manage_manual_edits)
        self.stats_btn.clicked.connect(self.show_frequency_stats)
        self.reset_btn.clicked.connect(self.reset_records)
        root_layout.addWidget(header)

        # ── Progress row ─────────────────────────────────────────────
        prog_row = QWidget()
        prog_row.setStyleSheet("background: white; padding-bottom: 2px;")
        pl = QHBoxLayout(prog_row)
        pl.setContentsMargins(20, 6, 20, 10)
        pl.setSpacing(10)

        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedHeight(5)
        self.progress_bar.setStyleSheet(
            "QProgressBar { background: #f2f4f7; border-radius: 3px; border: none; }"
            "QProgressBar::chunk { background: #1570ef; border-radius: 3px; }"
        )
        pl.addWidget(self.progress_bar, 1)

        self.progress_label = QLabel("0/0")
        self.progress_label.setStyleSheet("font-size: 13px; color: #344054;")
        pl.addWidget(self.progress_label)

        self.accuracy_label = QLabel("正确率 —")
        self.accuracy_label.setStyleSheet("font-size: 13px; color: #667085;")
        pl.addWidget(self.accuracy_label)
        root_layout.addWidget(prog_row)

        # ── Scroll area ──────────────────────────────────────────────
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFrameShape(QScrollArea.Shape.NoFrame)
        self.scroll_area.setStyleSheet("background: #f9fafb;")

        content = QWidget()
        content.setStyleSheet("background: #f9fafb;")
        self.content_layout = QVBoxLayout(content)
        self.content_layout.setContentsMargins(28, 28, 28, 28)
        self.content_layout.setSpacing(14)

        # Question header row
        q_header = QHBoxLayout()
        q_header.setSpacing(8)
        self.title_label = QLabel()
        self.title_label.setStyleSheet("font-size: 20px; font-weight: 700; color: #101828;")
        q_header.addWidget(self.title_label)

        self.type_badge = QLabel()
        self.type_badge.setStyleSheet(
            "background: #eff8ff; color: #1570ef; font-size: 12px; font-weight: 600;"
            " padding: 2px 10px; border-radius: 10px;"
        )
        q_header.addWidget(self.type_badge)

        self.history_label = QLabel()
        self.history_label.setStyleSheet("font-size: 13px; color: #98a2b3;")
        q_header.addWidget(self.history_label)
        q_header.addStretch(1)
        self.content_layout.addLayout(q_header)

        # Question text card
        self.question_label = QLabel()
        self.question_label.setWordWrap(True)
        self.question_label.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        self.question_label.setStyleSheet(
            "background: #f2f4f7; border-radius: 10px; padding: 16px 18px;"
            " font-size: 15px; color: #101828;"
        )
        self.content_layout.addWidget(self.question_label)

        # Options container
        self.options_container = QWidget()
        self.options_container.setStyleSheet("background: transparent;")
        self.options_layout = QVBoxLayout(self.options_container)
        self.options_layout.setContentsMargins(0, 0, 0, 0)
        self.options_layout.setSpacing(8)
        self.content_layout.addWidget(self.options_container)

        # Result label
        self.result_label = QLabel()
        self.result_label.setWordWrap(True)
        self.result_label.setStyleSheet("font-size: 14px;")
        self.content_layout.addWidget(self.result_label)

        self.content_layout.addStretch(1)
        self.scroll_area.setWidget(content)
        root_layout.addWidget(self.scroll_area, 1)

        # ── Bottom bar ───────────────────────────────────────────────
        bottom = QWidget()
        bottom.setStyleSheet("background: white; border-top: 1px solid #e4e7ec;")
        bl = QHBoxLayout(bottom)
        bl.setContentsMargins(20, 10, 20, 10)
        bl.setSpacing(8)

        self.hint_label = QLabel("按 A–D 选择，Enter 提交")
        self.hint_label.setStyleSheet("font-size: 13px; color: #b0b7c3;")
        bl.addWidget(self.hint_label)

        self.keyboard_entry = QLineEdit()
        self.keyboard_entry.setPlaceholderText("键盘输入")
        self.keyboard_entry.setFixedWidth(88)
        self.keyboard_entry.setStyleSheet(
            "border: 1px solid #d0d5dd; border-radius: 6px; padding: 5px 10px;"
            " font-size: 13px; color: #344054; background: white;"
        )
        self.keyboard_entry.returnPressed.connect(self._process_keyboard_enter)
        bl.addWidget(self.keyboard_entry)

        bl.addStretch(1)

        self.next_btn = QPushButton("下一题")
        self.next_btn.setStyleSheet(_BTN_GHOST)
        self.next_btn.clicked.connect(self.next_question)
        bl.addWidget(self.next_btn)

        self.submit_btn = QPushButton("提交答案")
        self.submit_btn.setStyleSheet(_BTN_PRIMARY)
        self.submit_btn.clicked.connect(self.submit_answer)
        bl.addWidget(self.submit_btn)

        root_layout.addWidget(bottom)
        self.setCentralWidget(root)

    def _show_welcome(self):
        self.current_q = None
        self.submitted = False
        self.answer_revealed = False
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
        card.setStyleSheet(_CARD_DEFAULT)
        card.setCursor(Qt.CursorShape.PointingHandCursor)

        card_layout = QHBoxLayout(card)
        card_layout.setContentsMargins(14, 10, 14, 10)
        card_layout.setSpacing(12)

        if q_type in ("single", "judge"):
            btn = QRadioButton()
        else:
            btn = QCheckBox()
        btn.setStyleSheet(
            "QRadioButton::indicator { width: 17px; height: 17px; }"
            " QCheckBox::indicator { width: 17px; height: 17px; }"
        )
        card_layout.addWidget(btn)

        lbl = QLabel(f"{key}. {text}")
        lbl.setWordWrap(True)
        lbl.setStyleSheet("font-size: 14px; color: #101828; background: transparent; border: none;")
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
            card.setStyleSheet(_CARD_SELECTED if checked else _CARD_DEFAULT)

    def _display_question(self):
        q = self.current_q
        self._clear_options()

        q_type = q.get("type", "")
        self.title_label.setText(f"第 {q.get('id', '')} 题")

        type_text = TYPE_LABELS.get(q_type, q_type)
        self.type_badge.setText(type_text)
        self.type_badge.setVisible(bool(type_text))

        question_text = str(q.get("text", "") or "")
        if q_type == "blank":
            question_text = _mask_blank_question_text(question_text, q.get("answer", ""))
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
            self.result_label.setStyleSheet("font-size: 14px; color: #667085;")
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
                self.result_label.setStyleSheet("font-size: 14px; color: #344054;")
                self.result_label.setText(f"参考答案：{_format_answer_text(q.get('answer'))}")
                self._add_subjective_buttons()
                self.submit_btn.setEnabled(False)
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
        self._mark_objective_result(correct, selected, is_correct)

    def _selected_options(self):
        return {key for key, widget in self.option_widgets.items() if widget.isChecked()}

    def _mark_objective_result(self, correct, selected, is_correct):
        for key, card in self.option_cards.items():
            lbl = card.findChild(QLabel)
            if key in correct:
                card.setStyleSheet(_CARD_CORRECT)
                if lbl:
                    lbl.setStyleSheet("font-size: 14px; color: #027a48; font-weight: 600; background: transparent; border: none;")
            elif key in selected:
                card.setStyleSheet(_CARD_WRONG)
                if lbl:
                    lbl.setStyleSheet("font-size: 14px; color: #b42318; font-weight: 600; background: transparent; border: none;")
            else:
                card.setStyleSheet(_CARD_DIMMED)
                if lbl:
                    lbl.setStyleSheet("font-size: 14px; color: #98a2b3; background: transparent; border: none;")

        if is_correct:
            self.result_label.setStyleSheet("font-size: 14px; color: #027a48; font-weight: 700;")
            self.result_label.setText("回答正确！")
        else:
            self.result_label.setStyleSheet("font-size: 14px; color: #b42318; font-weight: 700;")
            self.result_label.setText(f"回答错误。正确答案：{''.join(sorted(correct))}")

        rec = get_record(self.records, self.current_q)
        attempts = int(rec.get("attempts", 0) or 0)
        errors = int(rec.get("errors", 0) or 0)
        self.history_label.setText(f"已做 {attempts} 次 · 正确 {attempts - errors} 次")

    def _add_subjective_buttons(self):
        row = QWidget()
        row.setStyleSheet("background: transparent;")
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 4, 0, 0)
        layout.setSpacing(10)

        correct_btn = QPushButton("✓  我答对了")
        correct_btn.setStyleSheet("""
            QPushButton {
                background: #ecfdf3; color: #027a48; border: 1.5px solid #6ce9a6;
                border-radius: 8px; padding: 7px 18px; font-size: 13px; font-weight: 600;
            }
            QPushButton:hover { background: #d1fae5; }
        """)
        correct_btn.clicked.connect(lambda: self._submit_subjective_result(True))
        layout.addWidget(correct_btn)

        wrong_btn = QPushButton("✗  我答错了")
        wrong_btn.setStyleSheet("""
            QPushButton {
                background: #fff1f0; color: #b42318; border: 1.5px solid #fca5a5;
                border-radius: 8px; padding: 7px 18px; font-size: 13px; font-weight: 600;
            }
            QPushButton:hover { background: #ffe4e6; }
        """)
        wrong_btn.clicked.connect(lambda: self._submit_subjective_result(False))
        layout.addWidget(wrong_btn)

        layout.addStretch(1)
        self.options_layout.addWidget(row)

    def _submit_subjective_result(self, is_correct):
        if not self.current_q or self.submitted:
            return

        update_record(self.records, self.current_q, is_correct)
        self.submitted = True
        self._update_stats()

        answer = _format_answer_text(self.current_q.get("answer"))
        if is_correct:
            self.result_label.setStyleSheet("font-size: 14px; color: #027a48; font-weight: 700;")
            self.result_label.setText(f"参考答案：{answer}\n已记录：答对。")
        else:
            self.result_label.setStyleSheet("font-size: 14px; color: #b42318; font-weight: 700;")
            self.result_label.setText(f"参考答案：{answer}\n已记录：答错。")

        for widget in self.option_widgets.values():
            widget.setEnabled(False)

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
                    self.result_label.setStyleSheet("font-size: 14px; color: #b54708;")
                    self.result_label.setText("未识别到有效选项，请输入题目存在的字母。")
                    return
                self.submit_answer()
                return
            if q_type in ("blank", "short"):
                if self._submit_subjective_by_keyboard(token):
                    return
                self.result_label.setStyleSheet("font-size: 14px; color: #b54708;")
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
        self.result_label.setStyleSheet("font-size: 14px;")
