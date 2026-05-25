#!/usr/bin/env python3
"""PyQt6 主刷题窗口。"""

import os

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QApplication,
    QButtonGroup,
    QCheckBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
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
from ui_pyqt_utils import (
    TYPE_LABELS,
    build_duplicate_groups,
    build_duplicate_signature_set,
    is_recent_duplicate_pick,
    normalize_keyboard_text,
    question_signature,
)


class QuizWindow(QMainWindow):
    """PyQt 主刷题窗口。"""

    def __init__(self, questions, source_path=""):
        super().__init__()
        # 题库数据由 parser 解析后传入，界面层只消费题目字典。
        self.questions = list(questions or [])
        self.manual_edits = load_manual_question_edits()
        for q in self.questions:
            _ensure_question_identity_fields(q)
        apply_manual_question_edits(self.questions, self.manual_edits)

        # question_map 用于统计/详情弹窗按题号回查题目。
        self.question_map = {q.get("id"): q for q in self.questions}
        self.source_path = source_path
        self.source_name = os.path.basename(source_path) if source_path else "未命名题库"
        self.records = load_records()

        # 当前作答状态。
        self.current_q = None
        self.submitted = False
        self.answer_revealed = False
        self.option_widgets = {}
        self.option_group = None
        # 近期题目签名用于减少重复题短时间连续出现。
        self.recent_signatures = []
        self.recent_signature_limit = 6
        self.duplicate_groups = build_duplicate_groups(self.questions)
        self.duplicate_signature_set = build_duplicate_signature_set(self.duplicate_groups, self.question_map)

        self.setWindowTitle("SuperReciteHelper - PyQt")
        self.resize(960, 720)
        self._build_ui()
        self._show_welcome()

    def _build_ui(self):
        """搭建主窗口布局。"""
        root = QWidget()
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(18, 14, 18, 14)
        root_layout.setSpacing(12)

        self.info_label = QLabel()
        self.info_label.setWordWrap(True)
        root_layout.addWidget(self.info_label)

        self.stats_label = QLabel()
        root_layout.addWidget(self.stats_label)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFrameShape(QScrollArea.Shape.NoFrame)

        content = QWidget()
        self.content_layout = QVBoxLayout(content)
        self.content_layout.setContentsMargins(0, 0, 0, 0)
        self.content_layout.setSpacing(10)

        self.title_label = QLabel()
        self.title_label.setWordWrap(True)
        self.title_label.setStyleSheet("font-size: 20px; font-weight: 700; color: #243447;")
        self.content_layout.addWidget(self.title_label)

        self.type_label = QLabel()
        self.type_label.setStyleSheet("color: #667085;")
        self.content_layout.addWidget(self.type_label)

        self.question_label = QLabel()
        self.question_label.setWordWrap(True)
        self.question_label.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        self.question_label.setStyleSheet(
            "background: white; border: 1px solid #d0d5dd; border-radius: 6px; "
            "padding: 14px; font-size: 16px; color: #101828;"
        )
        self.content_layout.addWidget(self.question_label)

        self.options_container = QWidget()
        self.options_layout = QVBoxLayout(self.options_container)
        self.options_layout.setContentsMargins(0, 0, 0, 0)
        self.options_layout.setSpacing(8)
        self.content_layout.addWidget(self.options_container)

        self.result_label = QLabel()
        self.result_label.setWordWrap(True)
        self.result_label.setStyleSheet("font-size: 15px;")
        self.content_layout.addWidget(self.result_label)

        self.history_label = QLabel()
        self.history_label.setWordWrap(True)
        self.history_label.setStyleSheet("color: #667085;")
        self.content_layout.addWidget(self.history_label)

        self.content_layout.addStretch(1)
        self.scroll_area.setWidget(content)
        root_layout.addWidget(self.scroll_area, 1)

        # 底部操作区：刷题、键盘输入和工具入口。
        button_row = QHBoxLayout()
        self.next_btn = QPushButton("下一题")
        self.next_btn.clicked.connect(self.next_question)
        button_row.addWidget(self.next_btn)

        self.submit_btn = QPushButton("提交答案")
        self.submit_btn.clicked.connect(self.submit_answer)
        button_row.addWidget(self.submit_btn)

        self.keyboard_entry = QLineEdit()
        self.keyboard_entry.setPlaceholderText("键盘输入：A / ABC / t / f，回车提交")
        self.keyboard_entry.returnPressed.connect(self._process_keyboard_enter)
        button_row.addWidget(self.keyboard_entry, 1)

        self.edit_btn = QPushButton("编辑当前题")
        self.edit_btn.clicked.connect(self.edit_current_question)
        button_row.addWidget(self.edit_btn)

        self.manage_edits_btn = QPushButton("管理修改")
        self.manage_edits_btn.clicked.connect(self.manage_manual_edits)
        button_row.addWidget(self.manage_edits_btn)

        self.stats_btn = QPushButton("考频统计")
        self.stats_btn.clicked.connect(self.show_frequency_stats)
        button_row.addWidget(self.stats_btn)

        self.reset_btn = QPushButton("重置记录")
        self.reset_btn.clicked.connect(self.reset_records)
        button_row.addWidget(self.reset_btn)

        button_row.addStretch(1)
        root_layout.addLayout(button_row)

        self.setCentralWidget(root)

    def _show_welcome(self):
        """显示初始欢迎页。"""
        self.current_q = None
        self.submitted = False
        self.answer_revealed = False
        self._clear_options()

        self.info_label.setText(f"题库：{self.source_name} | 共 {len(self.questions)} 题")
        self.title_label.setText("欢迎使用 PyQt 最小刷题界面")
        self.type_label.setText("")
        self.question_label.setText("点击“下一题”开始。当前版本只保留最核心的刷题流程。")
        self.result_label.setText("")
        self.history_label.setText("")
        self.keyboard_entry.clear()
        self.submit_btn.setEnabled(False)
        self._update_stats()

    def _update_stats(self):
        """刷新顶部统计信息。"""
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

        accuracy = (1 - total_errors / total_attempts) * 100 if total_attempts else 0
        self.stats_label.setText(
            f"已做：{attempted}/{len(self.questions)} | "
            f"总答题：{total_attempts} | 正确率：{accuracy:.1f}%"
        )

    def next_question(self):
        """按当前抽题策略抽取下一题。"""
        if not self.questions:
            QMessageBox.warning(self, "提示", "当前没有可用题目。")
            return

        picked = weighted_random_pick(self.questions, self.records)
        if len(self.questions) > 1 and self.duplicate_groups:
            for _ in range(30):
                if not is_recent_duplicate_pick(picked, self.duplicate_signature_set, self.recent_signatures):
                    break
                candidate = weighted_random_pick(self.questions, self.records)
                picked = candidate
                if not is_recent_duplicate_pick(candidate, self.duplicate_signature_set, self.recent_signatures):
                    break

        self.current_q = picked
        self.recent_signatures.append(question_signature(self.current_q))
        if len(self.recent_signatures) > self.recent_signature_limit:
            self.recent_signatures = self.recent_signatures[-self.recent_signature_limit:]
        self.submitted = False
        self.answer_revealed = False
        self.keyboard_entry.clear()
        self._display_question()

    def _display_question(self):
        """把当前题目渲染到界面。"""
        q = self.current_q
        self._clear_options()

        q_type = q.get("type", "")
        self.title_label.setText(f"第 {q.get('id', '')} 题")
        self.type_label.setText(TYPE_LABELS.get(q_type, q_type))

        question_text = str(q.get("text", "") or "")
        if q_type == "blank":
            question_text = _mask_blank_question_text(question_text, q.get("answer", ""))
        self.question_label.setText(question_text)

        rec = get_record(self.records, q)
        if rec.get("attempts", 0):
            errors = int(rec.get("errors", 0) or 0)
            attempts = int(rec.get("attempts", 0) or 0)
            rate = errors / attempts * 100 if attempts else 0
            self.history_label.setText(f"历史记录：答过 {attempts} 次，错误 {errors} 次，错误率 {rate:.0f}%")
        else:
            self.history_label.setText("历史记录：首次作答")

        if q_type in ("single", "judge"):
            # 单选和判断题使用互斥按钮组。
            self.option_group = QButtonGroup(self)
            self.option_group.setExclusive(True)
            for key, text in sorted((q.get("options") or {}).items()):
                btn = QRadioButton(f"{key}. {text}")
                btn.setStyleSheet("font-size: 15px; padding: 6px;")
                self.options_layout.addWidget(btn)
                self.option_group.addButton(btn)
                self.option_widgets[key] = btn
            self.submit_btn.setText("提交答案")
            self.submit_btn.setEnabled(True)
            self.result_label.setText("")
            self.keyboard_entry.setPlaceholderText("输入 A 或 B，回车提交；已判分后回车下一题")
        elif q_type == "multi":
            # 多选题使用复选框。
            for key, text in sorted((q.get("options") or {}).items()):
                btn = QCheckBox(f"{key}. {text}")
                btn.setStyleSheet("font-size: 15px; padding: 6px;")
                self.options_layout.addWidget(btn)
                self.option_widgets[key] = btn
            self.submit_btn.setText("提交答案")
            self.submit_btn.setEnabled(True)
            self.result_label.setText("")
            self.keyboard_entry.setPlaceholderText("输入 ABC，回车提交；已判分后回车下一题")
        else:
            # 主观题先显示答案，再由用户自评。
            self.result_label.setText("请先自行作答，然后点击“显示答案”。")
            self.submit_btn.setText("显示答案")
            self.submit_btn.setEnabled(True)
            self.keyboard_entry.setPlaceholderText("先回车显示答案，再输入 t/f 回车自评")

        self.scroll_area.verticalScrollBar().setValue(0)
        self.keyboard_entry.setFocus()

    def submit_answer(self):
        """提交当前题答案。"""
        if not self.current_q or self.submitted:
            return

        q = self.current_q
        q_type = q.get("type")

        if q_type in ("blank", "short"):
            if not self.answer_revealed:
                self.answer_revealed = True
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
        """读取当前选中的客观题选项。"""
        return {key for key, widget in self.option_widgets.items() if widget.isChecked()}

    def _mark_objective_result(self, correct, selected, is_correct):
        """标记客观题判题结果。"""
        for key, widget in self.option_widgets.items():
            if key in correct:
                widget.setStyleSheet("font-size: 15px; padding: 6px; color: #027a48; font-weight: 700;")
            elif key in selected:
                widget.setStyleSheet("font-size: 15px; padding: 6px; color: #b42318; font-weight: 700;")
            else:
                widget.setStyleSheet("font-size: 15px; padding: 6px; color: #98a2b3;")

        if is_correct:
            self.result_label.setStyleSheet("font-size: 15px; color: #027a48; font-weight: 700;")
            self.result_label.setText("回答正确！")
        else:
            self.result_label.setStyleSheet("font-size: 15px; color: #b42318; font-weight: 700;")
            self.result_label.setText(f"回答错误。正确答案：{''.join(sorted(correct))}")

        rec = get_record(self.records, self.current_q)
        attempts = int(rec.get("attempts", 0) or 0)
        errors = int(rec.get("errors", 0) or 0)
        rate = errors / attempts * 100 if attempts else 0
        self.history_label.setText(f"历史记录：答过 {attempts} 次，错误 {errors} 次，错误率 {rate:.0f}%")

    def _add_subjective_buttons(self):
        """给主观题添加自评按钮。"""
        row = QWidget()
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 0, 0, 0)

        correct_btn = QPushButton("我答对了")
        correct_btn.clicked.connect(lambda: self._submit_subjective_result(True))
        layout.addWidget(correct_btn)

        wrong_btn = QPushButton("我答错了")
        wrong_btn.clicked.connect(lambda: self._submit_subjective_result(False))
        layout.addWidget(wrong_btn)

        layout.addStretch(1)
        self.options_layout.addWidget(row)

    def _submit_subjective_result(self, is_correct):
        """记录主观题自评结果。"""
        if not self.current_q or self.submitted:
            return

        update_record(self.records, self.current_q, is_correct)
        self.submitted = True
        self._update_stats()

        answer = _format_answer_text(self.current_q.get("answer"))
        if is_correct:
            self.result_label.setStyleSheet("font-size: 15px; color: #027a48; font-weight: 700;")
            self.result_label.setText(f"参考答案：{answer}\n已记录：答对。")
        else:
            self.result_label.setStyleSheet("font-size: 15px; color: #b42318; font-weight: 700;")
            self.result_label.setText(f"参考答案：{answer}\n已记录：答错。")

        for widget in self.option_widgets.values():
            widget.setEnabled(False)

    def reset_records(self):
        """清空全部做题记录。"""
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
        """编辑当前题并持久化。"""
        if not self.current_q:
            QMessageBox.information(self, "提示", "请先点击“下一题”抽取题目。")
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
        self._refresh_duplicate_cache()
        self._display_question()
        QMessageBox.information(self, "完成", "当前题修改已保存。")

    def manage_manual_edits(self):
        """打开手动修改管理弹窗。"""
        show_manual_edits_dialog(
            self,
            self.questions,
            self.manual_edits,
            current_q=self.current_q,
            on_refresh_current=self._display_question if self.current_q else None,
        )
        self._refresh_duplicate_cache()

    def show_frequency_stats(self):
        """打开考频统计弹窗。"""
        show_frequency_stats_dialog(
            self,
            self.questions,
            self.records,
            self.duplicate_groups,
            self.question_map,
        )

    def _select_objective_by_keyboard(self, token):
        """将键盘输入映射为客观题选项。"""
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
        """将 t/f 等键盘输入映射为主观题自评。"""
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
        """回车快捷操作：下一题、提交答案、主观自评。"""
        if self.current_q is None:
            self.next_question()
            return

        token = normalize_keyboard_text(self.keyboard_entry.text())
        self.keyboard_entry.clear()

        if token:
            q_type = self.current_q.get("type")
            if q_type in ("single", "multi", "judge"):
                if not self._select_objective_by_keyboard(token):
                    self.result_label.setStyleSheet("font-size: 15px; color: #b54708;")
                    self.result_label.setText("未识别到有效选项，请输入题目存在的字母。")
                    return
                self.submit_answer()
                return
            if q_type in ("blank", "short"):
                if self._submit_subjective_by_keyboard(token):
                    return
                self.result_label.setStyleSheet("font-size: 15px; color: #b54708;")
                self.result_label.setText("主观题请在显示答案后输入 t/f 自评。")
                return

        if not self.submitted:
            self.submit_answer()
        else:
            self.next_question()

    def keyPressEvent(self, event):
        """全局回车快捷键。"""
        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            focused = QApplication.focusWidget()
            if focused is not self.keyboard_entry:
                self._process_keyboard_enter()
                return
        super().keyPressEvent(event)

    def _refresh_duplicate_cache(self):
        """刷新重复题缓存。"""
        self.duplicate_groups = build_duplicate_groups(self.questions)
        self.duplicate_signature_set = build_duplicate_signature_set(self.duplicate_groups, self.question_map)

    def _clear_options(self):
        """清空上一题的选项控件。"""
        while self.options_layout.count():
            item = self.options_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()
        self.option_widgets = {}
        self.option_group = None
        self.result_label.setStyleSheet("font-size: 15px;")
