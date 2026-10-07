#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""构建网页版静态资源。

用法（任意目录下均可执行）：
    <python> web/tools/build_assets.py [题库文档路径]

产出：
    web/data/bank.json   题库数据（解析自 docx，需随仓库提交）
    web/theme.css        主题变量（由 src/ui/theme.py 的 token 生成）
    web/pycore/*.py      src/core 运行时模块副本（question_bank/session/parser）
"""

import json
import os
import shutil
import sys
from datetime import datetime

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
WEB_DIR = os.path.join(REPO_ROOT, 'web')
sys.path.insert(0, REPO_ROOT)

from src.core.parser import parse_questions  # noqa: E402

DEFAULT_BANK = r'D:\Study\大二上\高党\高党往年真题汇总\高党往年真题汇总（SuperReciteHelper格式）.docx'
CORE_COPY_FILES = ('parser.py', 'question_bank.py', 'session.py')


def build_bank(bank_path):
    questions = parse_questions(bank_path)
    payload = {
        'title': os.path.splitext(os.path.basename(bank_path))[0],
        'source': os.path.basename(bank_path),
        'generated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'count': len(questions),
        'questions': questions,
    }
    dst_dir = os.path.join(WEB_DIR, 'data')
    os.makedirs(dst_dir, exist_ok=True)
    dst = os.path.join(dst_dir, 'bank.json')
    with open(dst, 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False, indent=1)
    return dst, len(questions)


def build_theme_css():
    from src.ui.theme import LIGHT, DARK

    def block(name, tokens):
        lines = [f':root[data-theme="{name}"] {{']
        lines += [f'  --{k.replace("_", "-")}: {v};' for k, v in tokens.items()]
        lines.append('}')
        return '\n'.join(lines)

    css = (
        '/* 本文件由 web/tools/build_assets.py 自动生成（token 来源：src/ui/theme.py），请勿手改。 */\n\n'
        + block('light', LIGHT) + '\n\n' + block('dark', DARK) + '\n'
    )
    os.makedirs(WEB_DIR, exist_ok=True)
    dst = os.path.join(WEB_DIR, 'theme.css')
    with open(dst, 'w', encoding='utf-8') as f:
        f.write(css)
    return dst


def copy_core_modules():
    src_dir = os.path.join(REPO_ROOT, 'src', 'core')
    dst_dir = os.path.join(WEB_DIR, 'pycore')
    os.makedirs(dst_dir, exist_ok=True)
    out = []
    for name in CORE_COPY_FILES:
        dst = os.path.join(dst_dir, name)
        shutil.copyfile(os.path.join(src_dir, name), dst)
        out.append(dst)
    return out


def main():
    bank_path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_BANK
    if not os.path.exists(bank_path):
        raise SystemExit(f'题库文件不存在：{bank_path}')

    dst, n = build_bank(bank_path)
    print(f'[bank]  {n} 题 -> {os.path.relpath(dst, REPO_ROOT)}')
    print(f'[theme] -> {os.path.relpath(build_theme_css(), REPO_ROOT)}')
    for p in copy_core_modules():
        print(f'[core]  -> {os.path.relpath(p, REPO_ROOT)}')


if __name__ == '__main__':
    main()
