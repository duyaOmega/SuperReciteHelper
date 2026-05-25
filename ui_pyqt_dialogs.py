#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PyQt 弹窗集合。"""

from PyQt6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
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
from ui_pyqt_utils import TYPE_LABELS


def show_question_edit_dialog(parent, question, title="编辑当前题"):
    """编辑题目、题型、选项和答案。"""
    q = question or {}
    dialog = QDialog(parent)
    dialog.setWindowTitle(title)
    dialog.resize(760, 620)

    layout = QVBoxLayout(dialog)
    form = QFormLayout()
    layout.addLayout(form)

    type_combo = QComboBox()
    type_items = [
        ("单选", "single"),
        ("多选", "multi"),
        ("判断", "judge"),
        ("填空", "blank"),
        ("简答", "short"),
    ]
    for label, value in type_items:
        type_combo.addItem(label, value)
    current_type = str(q.get("type", "") or "short")
    type_index = next((i for i, (_, value) in enumerate(type_items) if value == current_type), 4)
    type_combo.setCurrentIndex(type_index)
    form.addRow("题型：", type_combo)

    question_text = QPlainTextEdit()
    question_text.setPlainText(str(q.get("text", "") or ""))
    question_text.setMinimumHeight(150)
    form.addRow("题目文本：", question_text)

    answer_text = QPlainTextEdit()
    answer_text.setPlainText(_format_answer_text(q.get("answer")))
    answer_text.setMinimumHeight(80)
    form.addRow("答案：", answer_text)

    options_text = QPlainTextEdit()
    options_text.setPlainText(_format_options_for_edit(q.get("options") or {}))
    options_text.setMinimumHeight(120)
    form.addRow("选项：", options_text)

    hint_label = QLabel()
    hint_label.setWordWrap(True)
    hint_label.setStyleSheet("color: #667085;")
    layout.addWidget(hint_label)

    def refresh_hint():
        """根据题型切换选项输入状态。"""
        target_type = type_combo.currentData()
        if target_type == "judge":
            options_text.setPlainText("A: 正确\nB: 错误")
            options_text.setEnabled(False)
            hint_label.setText("判断题固定使用 A: 正确 / B: 错误，答案请输入 A 或 B。")
        elif target_type in ("single", "multi"):
            options_text.setEnabled(True)
            hint_label.setText("客观题选项格式示例：A: 选项内容。单选答案如 A，多选答案如 AC。")
        else:
            options_text.setEnabled(False)
            hint_label.setText("主观题不需要选项，答案可填写文本。")

    type_combo.currentIndexChanged.connect(refresh_hint)
    refresh_hint()

    buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
    layout.addWidget(buttons)

    result = {"value": None}

    def save():
        """校验输入并返回编辑结果。"""
        new_question_text = question_text.toPlainText().strip()
        if not new_question_text:
            QMessageBox.warning(dialog, "格式错误", "题目文本不能为空。")
            return

        target_type = type_combo.currentData()
        if target_type == "judge":
            target_options = {"A": "正确", "B": "错误"}
        elif target_type in ("single", "multi"):
            target_options, opt_err = _parse_manual_options_text(options_text.toPlainText().strip())
            if opt_err:
                QMessageBox.warning(dialog, "格式错误", opt_err)
                return
        else:
            target_options = {}

        parsed_answer, err = _parse_manual_answer_for_question(
            q,
            answer_text.toPlainText().strip(),
            target_type=target_type,
            options_override=target_options,
        )
        if err:
            QMessageBox.warning(dialog, "格式错误", err)
            return

        result["value"] = (new_question_text, parsed_answer, target_type, target_options)
        dialog.accept()

    buttons.accepted.connect(save)
    buttons.rejected.connect(dialog.reject)
    dialog.exec()
    return result["value"]


def show_manual_edits_dialog(parent, questions, manual_edits, current_q=None, on_refresh_current=None):
    """管理已保存的题目修改。"""
    dialog = QDialog(parent)
    dialog.setWindowTitle("管理题目修改")
    dialog.resize(900, 520)

    layout = QVBoxLayout(dialog)
    title = QLabel()
    layout.addWidget(title)

    table = QTableWidget()
    table.setColumnCount(5)
    table.setHorizontalHeaderLabels(["序号", "题型", "答案", "更新时间", "题干预览"])
    table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
    table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
    table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
    layout.addWidget(table, 1)

    key_by_row = []

    def refresh():
        """刷新修改记录表格。"""
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
        """获取当前选中的修改记录键。"""
        row = table.currentRow()
        if row < 0 or row >= len(key_by_row):
            return None
        return key_by_row[row]

    def restore_question_from_payload(key, payload):
        """按保存的原始快照恢复题目。"""
        for q in questions:
            _ensure_question_identity_fields(q)
            if q.get("_base_key") == key:
                q["type"] = str(payload.get("orig_type", q.get("_orig_type", q.get("type", ""))))
                q["options"] = dict(payload.get("orig_options", q.get("_orig_options", q.get("options", {}))) or {})
                q["text"] = str(payload.get("orig_text", q.get("_orig_text", q.get("text", ""))))
                q["answer"] = payload.get("orig_answer", q.get("_orig_answer", q.get("answer", "")))

    def restore_selected():
        """恢复选中的单条修改。"""
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
        """清空全部手动修改。"""
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
    restore_btn.clicked.connect(restore_selected)
    button_row.addWidget(restore_btn)
    clear_btn = QPushButton("清空全部修改")
    clear_btn.clicked.connect(clear_all)
    button_row.addWidget(clear_btn)
    close_btn = QPushButton("关闭")
    close_btn.clicked.connect(dialog.accept)
    button_row.addWidget(close_btn)
    button_row.addStretch(1)
    layout.addLayout(button_row)

    refresh()
    dialog.exec()


def show_frequency_stats_dialog(parent, questions, records, duplicate_groups, question_map):
    """展示当前题库的作答统计。"""
    dialog = QDialog(parent)
    dialog.setWindowTitle("考频统计")
    dialog.resize(980, 620)

    layout = QVBoxLayout(dialog)
    summary = QLabel()
    summary.setWordWrap(True)
    layout.addWidget(summary)

    control_row = QHBoxLayout()
    sort_combo = QComboBox()
    sort_combo.addItems(["按作答次数", "按错误次数", "按错误率", "按题号"])
    control_row.addWidget(QLabel("排序方式："))
    control_row.addWidget(sort_combo)

    only_attempted = QCheckBox("仅看已作答题目")
    only_attempted.setChecked(True)
    control_row.addWidget(only_attempted)
    control_row.addStretch(1)
    layout.addLayout(control_row)

    table = QTableWidget()
    table.setColumnCount(7)
    table.setHorizontalHeaderLabels(["排名", "题号", "题型", "作答次数", "错误次数", "错误率", "题干预览"])
    table.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeMode.Stretch)
    table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
    table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
    layout.addWidget(table, 1)

    rows = []
    attempted_count = 0
    total_attempts = 0
    total_errors = 0
    # 先把统计数据整理成 rows，后续排序/筛选只操作这个列表。
    for q in questions:
        rec = get_record(records, q)
        attempts = int(rec.get("attempts", 0) or 0)
        errors = int(rec.get("errors", 0) or 0)
        if attempts > 0:
            attempted_count += 1
        total_attempts += attempts
        total_errors += errors
        error_rate = errors / attempts * 100 if attempts else 0.0
        rows.append(
            {
                "id": q.get("id", ""),
                "type": TYPE_LABELS.get(q.get("type", ""), q.get("type", "")),
                "attempts": attempts,
                "errors": errors,
                "error_rate": error_rate,
                "text": str(q.get("text", "") or "").replace("\n", " ").strip(),
            }
        )

    total_questions = len(questions)
    overall_error_rate = total_errors / total_attempts * 100 if total_attempts else 0.0
    duplicate_hint = ""
    if duplicate_groups:
        duplicate_count = sum(len(group) for group in duplicate_groups)
        duplicate_hint = f" | 重复题组 {len(duplicate_groups)} 组，共 {duplicate_count} 题"
    summary.setText(
        f"总题数 {total_questions} | 已作答 {attempted_count} | 总作答 {total_attempts} "
        f"| 总错误 {total_errors} | 总错误率 {overall_error_rate:.1f}%{duplicate_hint}"
    )

    def sorted_rows():
        """按当前筛选和排序方式返回表格行。"""
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
        """刷新统计表格。"""
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
        """弹窗展示所选题目的完整内容。"""
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
    detail_btn.clicked.connect(show_selected_detail)
    button_row.addWidget(detail_btn)
    close_btn = QPushButton("关闭")
    close_btn.clicked.connect(dialog.accept)
    button_row.addWidget(close_btn)
    button_row.addStretch(1)
    layout.addLayout(button_row)

    refresh_table()
    dialog.exec()
