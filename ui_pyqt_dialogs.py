#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PyQt 弹窗集合。"""

import re

from PyQt6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from question import (
    _format_answer_text,
    _format_options_for_edit,
    _parse_manual_answer_for_question,
    _parse_manual_options_text,
)
from question_bank import (
    _ensure_question_identity_fields,
    get_record,
    save_manual_question_edits,
)
TYPE_LABELS = {
    "single": "单选题",
    "multi":  "多选题",
    "judge":  "判断题",
    "blank":  "填空题",
    "short":  "简答题",
}

_DIALOG_BG = "background: white;"
_INPUT_STYLE = (
    "border: 1px solid #d0d5dd; border-radius: 8px; padding: 8px 12px;"
    " font-size: 14px; color: #101828; background: white;"
)
_LABEL_STYLE = "font-size: 13px; font-weight: 600; color: #344054;"
_BTN_CANCEL = (
    "QPushButton { background: white; color: #344054; border: 1px solid #d0d5dd;"
    " border-radius: 8px; padding: 8px 20px; font-size: 14px; font-weight: 500; }"
    "QPushButton:hover { background: #f9fafb; }"
)
_BTN_PRIMARY = (
    "QPushButton { background: #1570ef; color: white; border: none;"
    " border-radius: 8px; padding: 8px 24px; font-size: 14px; font-weight: 600; }"
    "QPushButton:hover { background: #175cd3; }"
)
_BTN_GHOST = (
    "QPushButton { background: #f2f4f7; color: #344054; border: 1px solid #d0d5dd;"
    " border-radius: 8px; padding: 7px 14px; font-size: 13px; }"
    "QPushButton:hover { background: #e4e7ec; }"
)
_BTN_DEL = (
    "QPushButton { background: #fff1f0; color: #b42318; border: none;"
    " border-radius: 6px; font-size: 18px; font-weight: 700; }"
    "QPushButton:hover { background: #ffe4e6; }"
)
_BTN_ADD_OPT = (
    "QPushButton { background: transparent; color: #1570ef;"
    " border: 1.5px dashed #b2ccff; border-radius: 8px;"
    " padding: 7px 14px; font-size: 13px; font-weight: 600; }"
    "QPushButton:hover { background: #eff8ff; }"
)


def _section(label_text, widget, parent_layout):
    """Add a labeled section to a vertical layout."""
    lbl = QLabel(label_text)
    lbl.setStyleSheet(_LABEL_STYLE)
    parent_layout.addWidget(lbl)
    parent_layout.addWidget(widget)


def show_question_edit_dialog(parent, question, title="编辑当前题"):
    """编辑题目、题型、选项和答案。"""
    q = question or {}
    dialog = QDialog(parent)
    dialog.setWindowTitle(title)
    dialog.resize(580, 680)
    dialog.setStyleSheet(_DIALOG_BG)

    outer = QVBoxLayout(dialog)
    outer.setContentsMargins(0, 0, 0, 0)
    outer.setSpacing(0)

    # Title bar
    title_bar = QWidget()
    title_bar.setStyleSheet("background: white; border-bottom: 1px solid #e4e7ec;")
    tb = QHBoxLayout(title_bar)
    tb.setContentsMargins(20, 14, 16, 14)
    title_lbl = QLabel(title)
    title_lbl.setStyleSheet("font-size: 16px; font-weight: 700; color: #101828;")
    tb.addWidget(title_lbl)
    tb.addStretch(1)
    close_x = QPushButton("×")
    close_x.setFixedSize(28, 28)
    close_x.setStyleSheet(
        "QPushButton { background: transparent; color: #98a2b3; border: none; font-size: 20px; }"
        "QPushButton:hover { color: #344054; }"
    )
    close_x.clicked.connect(dialog.reject)
    tb.addWidget(close_x)
    outer.addWidget(title_bar)

    # Scrollable form
    scroll = QScrollArea()
    scroll.setWidgetResizable(True)
    scroll.setFrameShape(QScrollArea.Shape.NoFrame)
    form_w = QWidget()
    form_w.setStyleSheet(_DIALOG_BG)
    form = QVBoxLayout(form_w)
    form.setContentsMargins(20, 20, 20, 16)
    form.setSpacing(14)

    # ── 题型 ──
    type_combo = QComboBox()
    type_items = [("单选", "single"), ("多选", "multi"), ("判断", "judge"), ("填空", "blank"), ("简答", "short")]
    for lbl_t, val in type_items:
        type_combo.addItem(lbl_t, val)
    current_type = str(q.get("type", "") or "short")
    type_index = next((i for i, (_, v) in enumerate(type_items) if v == current_type), 4)
    type_combo.setCurrentIndex(type_index)
    type_combo.setFixedWidth(160)
    type_combo.setStyleSheet(
        "QComboBox { border: 1px solid #d0d5dd; border-radius: 8px; padding: 7px 12px;"
        " font-size: 14px; background: white; }"
        "QComboBox::drop-down { border: none; width: 24px; }"
    )
    _section("题型", type_combo, form)

    # ── 题目 ──
    question_text = QPlainTextEdit()
    question_text.setPlainText(str(q.get("text", "") or ""))
    question_text.setMinimumHeight(96)
    question_text.setStyleSheet(_INPUT_STYLE)
    _section("题目", question_text, form)

    # ── 答案 ──
    answer_line = QLineEdit()
    answer_line.setText(_format_answer_text(q.get("answer")))
    answer_line.setStyleSheet(_INPUT_STYLE)
    _section("答案", answer_line, form)

    # ── 选项 ──
    opts_section = QWidget()
    opts_section.setStyleSheet("background: transparent;")
    opts_v = QVBoxLayout(opts_section)
    opts_v.setContentsMargins(0, 0, 0, 0)
    opts_v.setSpacing(8)

    opts_lbl = QLabel("选项")
    opts_lbl.setStyleSheet(_LABEL_STYLE)
    opts_v.addWidget(opts_lbl)

    opts_rows_w = QWidget()
    opts_rows_w.setStyleSheet("background: transparent;")
    opts_rows_layout = QVBoxLayout(opts_rows_w)
    opts_rows_layout.setContentsMargins(0, 0, 0, 0)
    opts_rows_layout.setSpacing(8)
    opts_v.addWidget(opts_rows_w)

    option_rows = []  # list of (QLineEdit, QPushButton, QWidget)

    def _refresh_labels():
        letters = "ABCDEFGH"
        for i, (*_, row_w) in enumerate(option_rows):
            lbl = row_w.findChild(QLabel)
            if lbl:
                lbl.setText((letters[i] if i < len(letters) else "?") + ".")

    def _clear_all_rows():
        while option_rows:
            *_, row_w = option_rows.pop(0)
            row_w.setParent(None)
            row_w.deleteLater()

    def add_option_row(text="", enabled=True):
        idx = len(option_rows)
        letters = "ABCDEFGH"
        letter = (letters[idx] if idx < len(letters) else "?") + "."

        row_w = QWidget()
        row_w.setStyleSheet("background: transparent;")
        rl = QHBoxLayout(row_w)
        rl.setContentsMargins(0, 0, 0, 0)
        rl.setSpacing(8)

        prefix = QLabel(letter)
        prefix.setFixedWidth(22)
        prefix.setStyleSheet("font-size: 14px; color: #667085; font-weight: 600;")
        rl.addWidget(prefix)

        edit = QLineEdit(text)
        edit.setEnabled(enabled)
        edit.setStyleSheet(_INPUT_STYLE)
        rl.addWidget(edit, 1)

        del_btn = QPushButton("−")
        del_btn.setFixedSize(34, 34)
        del_btn.setEnabled(enabled)
        del_btn.setStyleSheet(_BTN_DEL)

        def on_remove(*_, rw=row_w):
            for i, (*_, w) in enumerate(option_rows):
                if w is rw:
                    option_rows.pop(i)
                    break
            rw.setParent(None)
            rw.deleteLater()
            _refresh_labels()

        del_btn.clicked.connect(on_remove)
        rl.addWidget(del_btn)

        opts_rows_layout.addWidget(row_w)
        option_rows.append((edit, del_btn, row_w))

    add_opt_btn = QPushButton("＋  添加选项")
    add_opt_btn.setStyleSheet(_BTN_ADD_OPT)
    add_opt_btn.clicked.connect(lambda: add_option_row())
    opts_v.addWidget(add_opt_btn)

    hint_label = QLabel()
    hint_label.setWordWrap(True)
    hint_label.setStyleSheet("font-size: 12px; color: #98a2b3;")
    opts_v.addWidget(hint_label)

    form.addWidget(opts_section)
    form.addStretch(1)
    scroll.setWidget(form_w)
    outer.addWidget(scroll, 1)

    # Footer
    footer = QWidget()
    footer.setStyleSheet("background: white; border-top: 1px solid #e4e7ec;")
    fl = QHBoxLayout(footer)
    fl.setContentsMargins(20, 12, 20, 12)
    fl.setSpacing(10)
    fl.addStretch(1)
    cancel_btn = QPushButton("取消")
    cancel_btn.setStyleSheet(_BTN_CANCEL)
    cancel_btn.clicked.connect(dialog.reject)
    fl.addWidget(cancel_btn)
    save_btn = QPushButton("保存")
    save_btn.setStyleSheet(_BTN_PRIMARY)
    fl.addWidget(save_btn)
    outer.addWidget(footer)

    # ── Type-change logic ──
    prev_type = [current_type]

    def refresh_hint():
        target_type = type_combo.currentData()
        is_obj = target_type in ("single", "multi", "judge")
        opts_section.setVisible(is_obj)

        if target_type == "judge":
            if prev_type[0] != "judge":
                _clear_all_rows()
                add_option_row("正确", enabled=False)
                add_option_row("错误", enabled=False)
            else:
                for edit, del_btn, _ in option_rows:
                    edit.setEnabled(False)
                    del_btn.setEnabled(False)
            add_opt_btn.setEnabled(False)
            hint_label.setText("判断题固定使用 A: 正确 / B: 错误，答案请输入 A 或 B。")
        elif target_type in ("single", "multi"):
            if prev_type[0] == "judge":
                _clear_all_rows()
            add_opt_btn.setEnabled(True)
            for edit, del_btn, _ in option_rows:
                edit.setEnabled(True)
                del_btn.setEnabled(True)
            hint_label.setText("单选答案如 A，多选答案如 AC。")
        else:
            hint_label.setText("主观题不需要选项。")

        prev_type[0] = target_type

    # Init existing options
    existing = _format_options_for_edit(q.get("options") or {})
    for line in existing.split("\n"):
        line = line.strip()
        if not line:
            continue
        m = re.match(r'^[A-Za-z][:.]\s*(.*)', line)
        add_option_row(m.group(1) if m else line)

    type_combo.currentIndexChanged.connect(refresh_hint)
    refresh_hint()

    result = {"value": None}

    def save():
        new_q_text = question_text.toPlainText().strip()
        if not new_q_text:
            QMessageBox.warning(dialog, "格式错误", "题目文本不能为空。")
            return

        target_type = type_combo.currentData()
        if target_type == "judge":
            target_options = {"A": "正确", "B": "错误"}
        elif target_type in ("single", "multi"):
            letters = "ABCDEFGH"
            values = [edit.text().strip() for edit, _, _ in option_rows if edit.text().strip()]
            if not values:
                QMessageBox.warning(dialog, "格式错误", "请至少添加一个选项。")
                return
            opts_text = "\n".join(f"{letters[i]}: {v}" for i, v in enumerate(values))
            target_options, opt_err = _parse_manual_options_text(opts_text)
            if opt_err:
                QMessageBox.warning(dialog, "格式错误", opt_err)
                return
        else:
            target_options = {}

        parsed_answer, err = _parse_manual_answer_for_question(
            q,
            answer_line.text().strip(),
            target_type=target_type,
            options_override=target_options,
        )
        if err:
            QMessageBox.warning(dialog, "格式错误", err)
            return

        result["value"] = (new_q_text, parsed_answer, target_type, target_options)
        dialog.accept()

    save_btn.clicked.connect(save)
    dialog.exec()
    return result["value"]


def show_manual_edits_dialog(parent, questions, manual_edits, current_q=None, on_refresh_current=None):
    """管理已保存的题目修改。"""
    dialog = QDialog(parent)
    dialog.setWindowTitle("管理题目修改")
    dialog.resize(900, 520)
    dialog.setStyleSheet(_DIALOG_BG)

    layout = QVBoxLayout(dialog)
    layout.setContentsMargins(16, 16, 16, 16)
    layout.setSpacing(10)

    title = QLabel()
    title.setStyleSheet("font-size: 14px; color: #344054;")
    layout.addWidget(title)

    table = QTableWidget()
    table.setColumnCount(5)
    table.setHorizontalHeaderLabels(["序号", "题型", "答案", "更新时间", "题干预览"])
    table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
    table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
    table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
    table.setStyleSheet(
        "QTableWidget { border: 1px solid #e4e7ec; border-radius: 8px; gridline-color: #f2f4f7; }"
        "QHeaderView::section { background: #f9fafb; color: #667085; font-size: 12px;"
        " padding: 6px; border: none; border-bottom: 1px solid #e4e7ec; }"
    )
    layout.addWidget(table, 1)

    key_by_row = []

    def refresh():
        key_by_row.clear()
        items = sorted(
            manual_edits.items(),
            key=lambda item: str((item[1] or {}).get("updated_at", "")),
            reverse=True,
        )
        title.setText(f"已保存修改：{len(items)} 项")
        table.setRowCount(len(items))
        for row, (key, payload) in enumerate(items):
            key_by_row.append(key)
            answer = _format_answer_text((payload or {}).get("answer", ""))
            values = [
                str(row + 1),
                str((payload or {}).get("type", "")),
                answer[:24] + "..." if len(answer) > 24 else answer,
                str((payload or {}).get("updated_at", "")),
                str((payload or {}).get("preview", "")),
            ]
            for col, value in enumerate(values):
                table.setItem(row, col, QTableWidgetItem(value))
        table.resizeColumnsToContents()
        table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)

    def selected_key():
        row = table.currentRow()
        if row < 0 or row >= len(key_by_row):
            return None
        return key_by_row[row]

    def restore_question_from_payload(key, payload):
        for q in questions:
            _ensure_question_identity_fields(q)
            if q.get("_base_key") == key:
                q["type"] = str(payload.get("orig_type", q.get("_orig_type", q.get("type", ""))))
                q["options"] = dict(payload.get("orig_options", q.get("_orig_options", q.get("options", {}))) or {})
                q["text"] = str(payload.get("orig_text", q.get("_orig_text", q.get("text", ""))))
                q["answer"] = payload.get("orig_answer", q.get("_orig_answer", q.get("answer", "")))

    def restore_selected():
        key = selected_key()
        if not key:
            QMessageBox.information(dialog, "提示", "请先选择一条修改记录。")
            return
        payload = manual_edits.get(key) or {}
        if QMessageBox.question(dialog, "确认", "确定恢复该题为默认解析结果吗？") != QMessageBox.StandardButton.Yes:
            return
        del manual_edits[key]
        save_manual_question_edits(manual_edits)
        restore_question_from_payload(key, payload)
        if current_q and current_q.get("_base_key") == key and on_refresh_current:
            on_refresh_current()
        refresh()

    def clear_all():
        if not manual_edits:
            QMessageBox.information(dialog, "提示", "当前没有可清空的修改。")
            return
        if QMessageBox.question(dialog, "确认", "确定清空全部题目修改吗？") != QMessageBox.StandardButton.Yes:
            return
        deleted = dict(manual_edits)
        manual_edits.clear()
        save_manual_question_edits(manual_edits)
        for key, payload in deleted.items():
            restore_question_from_payload(key, payload)
        if current_q and on_refresh_current:
            on_refresh_current()
        refresh()

    button_row = QHBoxLayout()
    restore_btn = QPushButton("恢复所选默认")
    restore_btn.setStyleSheet(_BTN_GHOST)
    restore_btn.clicked.connect(restore_selected)
    button_row.addWidget(restore_btn)

    clear_btn = QPushButton("清空全部修改")
    clear_btn.setStyleSheet(_BTN_GHOST)
    clear_btn.clicked.connect(clear_all)
    button_row.addWidget(clear_btn)

    close_btn = QPushButton("关闭")
    close_btn.setStyleSheet(_BTN_GHOST)
    close_btn.clicked.connect(dialog.accept)
    button_row.addWidget(close_btn)
    button_row.addStretch(1)
    layout.addLayout(button_row)

    refresh()
    dialog.exec()


def show_frequency_stats_dialog(parent, questions, records, question_map):
    """展示当前题库的作答统计。"""
    dialog = QDialog(parent)
    dialog.setWindowTitle("考频统计")
    dialog.resize(980, 620)
    dialog.setStyleSheet(_DIALOG_BG)

    layout = QVBoxLayout(dialog)
    layout.setContentsMargins(16, 16, 16, 16)
    layout.setSpacing(10)

    summary = QLabel()
    summary.setWordWrap(True)
    summary.setStyleSheet("font-size: 13px; color: #344054;")
    layout.addWidget(summary)

    control_row = QHBoxLayout()
    sort_combo = QComboBox()
    sort_combo.addItems(["按作答次数", "按错误次数", "按错误率", "按题号"])
    sort_combo.setStyleSheet(
        "QComboBox { border: 1px solid #d0d5dd; border-radius: 6px; padding: 5px 10px;"
        " font-size: 13px; background: white; }"
        "QComboBox::drop-down { border: none; }"
    )
    ctrl_lbl = QLabel("排序方式：")
    ctrl_lbl.setStyleSheet("font-size: 13px; color: #667085;")
    control_row.addWidget(ctrl_lbl)
    control_row.addWidget(sort_combo)

    only_attempted = QCheckBox("仅看已作答题目")
    only_attempted.setChecked(True)
    only_attempted.setStyleSheet("font-size: 13px; color: #344054;")
    control_row.addWidget(only_attempted)
    control_row.addStretch(1)
    layout.addLayout(control_row)

    table = QTableWidget()
    table.setColumnCount(7)
    table.setHorizontalHeaderLabels(["排名", "题号", "题型", "作答次数", "错误次数", "错误率", "题干预览"])
    table.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeMode.Stretch)
    table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
    table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
    table.setStyleSheet(
        "QTableWidget { border: 1px solid #e4e7ec; border-radius: 8px; gridline-color: #f2f4f7; }"
        "QHeaderView::section { background: #f9fafb; color: #667085; font-size: 12px;"
        " padding: 6px; border: none; border-bottom: 1px solid #e4e7ec; }"
    )
    layout.addWidget(table, 1)

    rows = []
    attempted_count = 0
    total_attempts = 0
    total_errors = 0
    for q in questions:
        rec = get_record(records, q)
        attempts = int(rec.get("attempts", 0) or 0)
        errors = int(rec.get("errors", 0) or 0)
        if attempts > 0:
            attempted_count += 1
        total_attempts += attempts
        total_errors += errors
        error_rate = errors / attempts * 100 if attempts else 0.0
        rows.append({
            "id": q.get("id", ""),
            "type": TYPE_LABELS.get(q.get("type", ""), q.get("type", "")),
            "attempts": attempts,
            "errors": errors,
            "error_rate": error_rate,
            "text": str(q.get("text", "") or "").replace("\n", " ").strip(),
        })

    total_questions = len(questions)
    overall_error_rate = total_errors / total_attempts * 100 if total_attempts else 0.0
    summary.setText(
        f"总题数 {total_questions} | 已作答 {attempted_count} | 总作答 {total_attempts}"
        f" | 总错误 {total_errors} | 总错误率 {overall_error_rate:.1f}%"
    )

    def sorted_rows():
        data = rows
        if only_attempted.isChecked():
            data = [row for row in data if row["attempts"] > 0]
        mode = sort_combo.currentText()
        if mode == "按作答次数":
            return sorted(data, key=lambda row: (-row["attempts"], -row["errors"], row["id"]))
        if mode == "按错误次数":
            return sorted(data, key=lambda row: (-row["errors"], -row["attempts"], row["id"]))
        if mode == "按错误率":
            return sorted(data, key=lambda row: (-row["error_rate"], -row["attempts"], row["id"]))
        return sorted(data, key=lambda row: row["id"])

    def refresh_table():
        data = sorted_rows()
        table.setRowCount(len(data))
        for row_index, row in enumerate(data):
            preview = row["text"]
            if len(preview) > 90:
                preview = preview[:90] + "..."
            values = [
                str(row_index + 1),
                str(row["id"]),
                str(row["type"]),
                str(row["attempts"]),
                str(row["errors"]),
                f"{row['error_rate']:.0f}%",
                preview,
            ]
            for col, value in enumerate(values):
                table.setItem(row_index, col, QTableWidgetItem(value))
        table.resizeColumnsToContents()
        table.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeMode.Stretch)

    def show_selected_detail():
        row = table.currentRow()
        if row < 0:
            QMessageBox.information(dialog, "提示", "请先选择一题。")
            return
        item = table.item(row, 1)
        if not item:
            return
        q = question_map.get(int(item.text()))
        if not q:
            QMessageBox.warning(dialog, "错误", "未找到题目详情。")
            return
        answer = _format_answer_text(q.get("answer"))
        options = q.get("options") or {}
        option_lines = "\n".join(f"{key}. {options[key]}" for key in sorted(options.keys()))
        detail = f"题干：\n{q.get('text', '')}\n\n"
        if option_lines:
            detail += f"选项：\n{option_lines}\n\n"
        detail += f"答案：\n{answer}"
        QMessageBox.information(dialog, f"第 {q.get('id', '')} 题", detail)

    sort_combo.currentIndexChanged.connect(refresh_table)
    only_attempted.stateChanged.connect(refresh_table)
    table.cellDoubleClicked.connect(lambda _row, _col: show_selected_detail())

    button_row = QHBoxLayout()
    detail_btn = QPushButton("查看所选题")
    detail_btn.setStyleSheet(_BTN_GHOST)
    detail_btn.clicked.connect(show_selected_detail)
    button_row.addWidget(detail_btn)

    close_btn = QPushButton("关闭")
    close_btn.setStyleSheet(_BTN_GHOST)
    close_btn.clicked.connect(dialog.accept)
    button_row.addWidget(close_btn)
    button_row.addStretch(1)
    layout.addLayout(button_row)

    refresh_table()
    dialog.exec()
