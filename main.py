#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PyQt6 程序入口。"""

import os
import sys

from PyQt6.QtWidgets import QApplication, QFileDialog, QMessageBox

from parser import build_parse_candidates
from question_bank import load_app_state, save_app_state
from ui_pyqt import QuizWindow


def _choose_question_file(parent=None):
    """选择一个结构化文本题库文件。"""
    state = load_app_state()
    last_files = state.get("last_open_files") or []
    if last_files:
        initial_dir = os.path.dirname(last_files[0])
    else:
        initial_dir = os.path.dirname(os.path.abspath(__file__))

    file_path, _ = QFileDialog.getOpenFileName(
        parent,
        "选择题库文件",
        initial_dir,
        "题库文件 (*.txt *.docx);;文本文件 (*.txt);;Word 文档 (*.docx)",
    )
    return file_path


def _load_questions(file_path):
    """调用 parser，取默认解析方案中的题目列表。"""
    candidates = build_parse_candidates(file_path)
    if not candidates or not candidates[0][1]:
        raise ValueError("未能解析出任何题目，请检查题库格式。")
    return candidates[0][1]


def main():
    """应用主入口：选择题库、解析并启动 PyQt 刷题界面。"""
    app = QApplication(sys.argv)

    file_path = _choose_question_file()
    if not file_path:
        return 0

    try:
        questions = _load_questions(file_path)
    except Exception as exc:
        QMessageBox.critical(None, "解析失败", f"文件解析失败：\n{exc}")
        return 1

    state = load_app_state()
    state["last_open_files"] = [file_path]
    save_app_state(state)

    window = QuizWindow(questions, source_path=file_path)
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
