#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""应用主题定义与切换（浅色 / 深色）。

所有界面样式统一从 token 生成，禁止在 UI 代码里硬编码颜色，
否则会在另一套主题下出现看不清的文字。
"""

from PyQt6.QtCore import Qt

LIGHT = {
    "bg": "#f9fafb",
    "surface": "#ffffff",
    "surface_alt": "#f2f4f7",
    "surface_hover": "#e4e7ec",
    "border": "#e4e7ec",
    "input_border": "#d0d5dd",
    "text": "#101828",
    "text_secondary": "#344054",
    "muted": "#667085",
    "faint": "#98a2b3",
    "subtle": "#b0b7c3",
    "primary": "#1570ef",
    "primary_hover": "#175cd3",
    "primary_disabled": "#b2ccff",
    "accent_bg": "#eff8ff",
    "accent_text": "#1570ef",
    "ok_bg": "#ecfdf3",
    "ok_bg_hover": "#d1fae5",
    "ok_border": "#6ce9a6",
    "ok_text": "#027a48",
    "err_bg": "#fff1f0",
    "err_bg_hover": "#ffe4e6",
    "err_border": "#fca5a5",
    "err_text": "#b42318",
    "warn_text": "#b54708",
    "grid": "#f2f4f7",
    "header_bg": "#f9fafb",
    "add_border": "#b2ccff",
}

DARK = {
    "bg": "#0e1116",
    "surface": "#171b21",
    "surface_alt": "#1e242c",
    "surface_hover": "#272e38",
    "border": "#262d37",
    "input_border": "#39414d",
    "text": "#e8edf3",
    "text_secondary": "#c6cfda",
    "muted": "#98a2b3",
    "faint": "#7c8794",
    "subtle": "#606a78",
    "primary": "#3b82f6",
    "primary_hover": "#5b97f8",
    "primary_disabled": "#2c466b",
    "accent_bg": "#182741",
    "accent_text": "#7cb0ff",
    "ok_bg": "#10231b",
    "ok_bg_hover": "#173324",
    "ok_border": "#2e6f4e",
    "ok_text": "#63d99b",
    "err_bg": "#291417",
    "err_bg_hover": "#381a1f",
    "err_border": "#7f3d3d",
    "err_text": "#ff9d9d",
    "warn_text": "#eab168",
    "grid": "#222933",
    "header_bg": "#1c222b",
    "add_border": "#33507e",
}

THEMES = {"light": LIGHT, "dark": DARK}

_active = "light"


def active_name():
    """返回当前主题名（light / dark）。"""
    return _active


def tokens():
    """返回当前主题的颜色 token 表。"""
    return THEMES[_active]


def set_active(name):
    """仅切换内部状态，不触碰应用级配色。"""
    global _active
    if name in THEMES:
        _active = name


def apply(app, name):
    """切换主题：更新内部 token 并同步应用级配色方案。

    配色方案影响未显式着色的控件（消息框、滚动条、下拉弹层等）。
    """
    set_active(name)
    scheme = Qt.ColorScheme.Dark if name == "dark" else Qt.ColorScheme.Light
    app.styleHints().setColorScheme(scheme)


def initial(app, saved_name):
    """决定启动主题：优先用户上次选择，否则跟随系统。"""
    if saved_name in THEMES:
        return saved_name
    return "dark" if app.styleHints().colorScheme() == Qt.ColorScheme.Dark else "light"
