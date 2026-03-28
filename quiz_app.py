#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
题库刷题工具
功能：
1. 手动选择题库文件，提取所有题目和答案，重新编号
2. 图形化界面，随机抽题、选择作答、显示正确答案
3. 记录错误次数到 error_record.json，支持断点续记
4. 根据错误次数和错误率进行加权随机抽题
"""

import tkinter as tk
from tkinter import messagebox, font as tkfont, filedialog, ttk
import re
import json
import os
import random
import math
import tempfile
import zipfile
import xml.etree.ElementTree as ET

# ============ 题目解析模块 ============

ANSWER_LABEL_RE = re.compile(r'^(?:正确答案|答案|参考答案|标准答案|【答案】|\[答案\]|答|参考解答)\s*[：:]?\s*(.*)\s*$')
QUESTION_START_RE = re.compile(r'^\s*(?:\d{1,4}(?:[、．\)]|[.](?!\d))|[一二三四五六七八九十百零]+[.、．\)])\s*')
OPTION_PREFIX_RE = re.compile(r'^\s*([A-HＡ-Ｈ])[.、．,，\)）:：]\s*')
OPTION_TOKEN_RE = re.compile(r'([A-HＡ-Ｈ])[.、．,，\)）:：]\s*')


def _normalize_extracted_text(text):
    """规范化不同来源文本，提升题块边界与选项识别稳定性。"""
    if not text:
        return ''

    t = text.replace('\r\n', '\n').replace('\r', '\n')
    t = t.replace('\u3000', ' ')
    # 标准答案标签前强制换行
    t = re.sub(r'(?<!\n)(正确答案|参考答案|标准答案|答案|答)[：:]', r'\n\1：', t)
    # 题号前强制换行，兼容 OCR/PDF 黏连；避免把小数(如4.5)误判成题号。
    t = re.sub(r'(?<!\n)(\d{1,4}(?:[、．\)]|[.](?!\d)))', r'\n\1', t)
    t = re.sub(r'\n{3,}', '\n\n', t)
    return t.strip()


def _looks_like_option_line(text):
    if not text:
        return False
    if OPTION_PREFIX_RE.match(text):
        return True
    markers = OPTION_TOKEN_RE.findall(text)
    return len(markers) >= 2


def _has_option_structure(lines):
    """跨多行判断是否包含客观题选项结构。"""
    if not lines:
        return False

    prefixed = sum(1 for line in lines if OPTION_PREFIX_RE.match(line))
    if prefixed >= 2:
        return True

    joined = ' '.join(lines)
    markers = OPTION_TOKEN_RE.findall(joined)
    return len(markers) >= 2


def _looks_like_answer_token(text):
    """判断文本是否像标准答案（而不是选项行）。"""
    if not text:
        return False
    if _looks_like_option_line(text):
        return False

    normalized = _normalize_answer_text(text)
    cleaned = re.sub(r'[\s,，、;；/\\]+', '', normalized)
    if re.fullmatch(r'[A-H]+', cleaned or ''):
        return True

    answer_keywords = ('正确', '错误', '对', '错', '是', '否')
    return any(k == text.strip() for k in answer_keywords)


def _read_text_file(filepath):
    for enc in ('utf-8-sig', 'utf-8', 'gb18030', 'gbk'):
        try:
            with open(filepath, 'r', encoding=enc) as f:
                return f.read()
        except Exception:
            continue
    raise ValueError(f'无法读取文本文件编码：{filepath}')


def _extract_docx_text_with_style(filepath):
    """优先用 python-docx 提取，尽量利用字体样式识别答案。"""
    try:
        from docx import Document
    except Exception:
        return None

    def run_is_styled(run):
        font = run.font
        color = None
        if font and font.color is not None and font.color.rgb is not None:
            color = str(font.color.rgb)
        return bool(
            run.bold or run.italic or run.underline or
            (font and font.highlight_color is not None) or
            color
        )

    lines = []
    doc = Document(filepath)

    def consume_para(para):
        raw = ''.join(r.text for r in para.runs).strip()
        if not raw:
            return
        styled = ''.join(r.text for r in para.runs if run_is_styled(r)).strip()
        lines.append(raw)

        # 若答案仅以不同字体标记，则注入标准答案行
        if styled and styled != raw and (QUESTION_START_RE.match(raw) or _is_blank_question(raw)):
            lines.append(f'答案：{styled}')

    for para in doc.paragraphs:
        consume_para(para)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    consume_para(para)

    return _normalize_extracted_text('\n'.join(lines).strip())


def _extract_docx_text_fallback(filepath):
    """不依赖第三方库的 docx 文本提取回退方案。"""
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    with zipfile.ZipFile(filepath, 'r') as zf:
        xml_data = zf.read('word/document.xml')

    root = ET.fromstring(xml_data)
    lines = []
    for p in root.findall('.//w:p', ns):
        texts = []
        for t in p.findall('.//w:t', ns):
            texts.append(t.text or '')
        line = ''.join(texts).strip()
        if line:
            lines.append(line)
    return _normalize_extracted_text('\n'.join(lines).strip())


def _extract_doc_text_windows(filepath):
    """使用 Windows Word COM 将 .doc/.docx 转成 txt（若本机可用）。"""
    try:
        import win32com.client  # type: ignore
    except Exception:
        return None

    temp_txt = None
    word = None
    doc = None
    try:
        fd, temp_txt = tempfile.mkstemp(suffix='.txt')
        os.close(fd)

        word = win32com.client.Dispatch('Word.Application')
        word.Visible = False
        doc = word.Documents.Open(os.path.abspath(filepath))
        # 7 = wdFormatUnicodeText
        doc.SaveAs(os.path.abspath(temp_txt), FileFormat=7)
        doc.Close(False)
        doc = None
        word.Quit()
        word = None
        return _normalize_extracted_text(_read_text_file(temp_txt))
    except Exception:
        return None
    finally:
        try:
            if doc is not None:
                doc.Close(False)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass
        if temp_txt and os.path.exists(temp_txt):
            try:
                os.remove(temp_txt)
            except Exception:
                pass


def _pdf_span_is_styled(span):
    font_name = str(span.get('font', '')).lower()
    color = span.get('color', 0)
    flags = int(span.get('flags', 0))
    if any(k in font_name for k in ('bold', 'black', 'heavy', 'demi', 'medium')):
        return True
    if color not in (0, None):
        return True
    # 低位标志中某些位常用于强调样式，作为启发式
    if flags & 2 or flags & 16:
        return True
    return False


def _extract_styled_segments_from_spans(spans):
    """从PDF行内span提取连续样式片段（红字/下划线/粗体等）。"""
    segments = []
    current = []

    for s in spans:
        txt = (s.get('text', '') or '').strip()
        if not txt:
            continue
        if _pdf_span_is_styled(s):
            current.append(txt)
        else:
            if current:
                seg = ''.join(current).strip(' ，,。；;：:、')
                if seg:
                    segments.append(seg)
                current = []

    if current:
        seg = ''.join(current).strip(' ，,。；;：:、')
        if seg:
            segments.append(seg)

    # 去重（保持顺序）
    seen = set()
    unique = []
    for seg in segments:
        if seg in seen:
            continue
        # 过滤明显噪声
        if len(seg) < 2 and not re.fullmatch(r'[\u4e00-\u9fff]', seg):
            continue
        if re.fullmatch(r'[\W_]+', seg):
            continue
        seen.add(seg)
        unique.append(seg)
    return unique


def _dedupe_keep_order(items):
    seen = set()
    out = []
    for x in items:
        if x in seen:
            continue
        seen.add(x)
        out.append(x)
    return out


def _normalize_pdf_sentence(text):
    t = text or ''
    # 修复中文文本因换行导致的断词空格
    t = re.sub(r'(?<=[\u4e00-\u9fff])\s+(?=[\u4e00-\u9fff])', '', t)
    t = re.sub(r'\s{2,}', ' ', t)
    return t.strip()


def _merge_split_segments(paragraph_text, segments):
    """合并被换行拆开的相邻答案片段（如“习近平新”+“时代...”）。"""
    if not segments:
        return segments

    merged = []
    i = 0
    while i < len(segments):
        current = segments[i]
        j = i + 1
        while j < len(segments):
            candidate = current + segments[j]
            if candidate in paragraph_text:
                current = candidate
                j += 1
            else:
                break
        merged.append(current)
        i = j

    return _dedupe_keep_order(merged)


def _build_blank_question_from_line(line_text, segments):
    """将样式答案片段替换为填空位，生成“多空填空题”。"""
    question = line_text
    answers = []

    for idx, seg in enumerate(segments, 1):
        placeholder = f'（{idx}）______'
        if seg in question:
            question = question.replace(seg, placeholder, 1)
            answers.append(f'（{idx}）{seg}')

    if not answers:
        return None, None
    return question, _clean_answer_text('；'.join(answers))


def _extract_pdf_text(filepath):
    """提取 PDF 文本；若可获取字体信息，则尝试把样式答案转成“答案：”行。"""
    # 优先 PyMuPDF（支持 span 样式）
    try:
        import fitz  # type: ignore

        lines = []
        styled_pairs = []
        blank_pairs = []
        paragraph_pairs = []
        stream_pairs = []
        all_line_items = []
        with fitz.open(filepath) as doc:
            for page in doc:
                data = page.get_text('dict')
                for block in data.get('blocks', []):
                    if block.get('type') != 0:
                        continue

                    block_lines = []
                    block_segments = []
                    for line in block.get('lines', []):
                        spans = line.get('spans', [])
                        line_text = ''.join(s.get('text', '') for s in spans).strip()
                        if not line_text:
                            continue

                        block_lines.append(line_text)
                        line_segments = _extract_styled_segments_from_spans(spans)
                        block_segments.extend(line_segments)
                        all_line_items.append((line_text, line_segments))

                        styled = ''.join(s.get('text', '') for s in spans if _pdf_span_is_styled(s)).strip()
                        lines.append(line_text)
                        if styled and styled != line_text and (QUESTION_START_RE.match(line_text) or _is_blank_question(line_text)):
                            lines.append(f'答案：{styled}')

                        # 优先生成“多空填空”题：将行内红字/样式片段替换为空位。
                        segments = _extract_styled_segments_from_spans(spans)
                        if (
                            segments and
                            len(line_text) <= 220 and
                            not _has_option_structure([line_text]) and
                            ('，' in line_text or '。' in line_text or ',' in line_text)
                        ):
                            q_line, a_line = _build_blank_question_from_line(line_text, segments)
                            if q_line and a_line and q_line != line_text:
                                blank_pairs.append((q_line, a_line))

                        # 兼容“红色/下划线大字即答案”这类PDF样式标注。
                        styled_ratio = len(styled) / max(len(line_text), 1)
                        if (
                            styled and styled != line_text and
                            0.05 <= styled_ratio <= 0.7 and
                            len(line_text) <= 180 and len(styled) <= 80 and
                            not _has_option_structure([line_text]) and
                            ('，' in line_text or '。' in line_text or ',' in line_text)
                        ):
                            styled_pairs.append((line_text, styled))

                    # 逐段多空：将一个文本块中的样式片段统一生成一道填空题。
                    if block_lines and block_segments:
                        block_text = _normalize_pdf_sentence(' '.join(block_lines))
                        block_segments = _dedupe_keep_order(block_segments)
                        if (
                            20 <= len(block_text) <= 800 and
                            1 <= len(block_segments) <= 16 and
                            not _has_option_structure([block_text]) and
                            ('，' in block_text or '。' in block_text or ',' in block_text)
                        ):
                            q_line, a_line = _build_blank_question_from_line(block_text, block_segments)
                            if q_line and a_line and q_line != block_text:
                                paragraph_pairs.append((q_line, a_line))

        # 跨页聚合：按阅读顺序拼接行，构造“整句/整段多空”题。
        buf_lines = []
        buf_segments = []
        for line_text, line_segments in all_line_items:
            buf_lines.append(line_text)
            buf_segments.extend(line_segments)

            paragraph_text = _normalize_pdf_sentence(' '.join(buf_lines))
            line_tail = line_text.rstrip()
            should_flush = line_tail.endswith(('。', '！', '？', '.', '!', '?')) or len(paragraph_text) >= 560
            if not should_flush:
                continue

            segs = _dedupe_keep_order(buf_segments)
            segs = _merge_split_segments(paragraph_text, segs)
            if (
                segs and
                30 <= len(paragraph_text) <= 1500 and
                1 <= len(segs) <= 24 and
                not _has_option_structure([paragraph_text]) and
                ('，' in paragraph_text or '。' in paragraph_text or ',' in paragraph_text)
            ):
                q_line, a_line = _build_blank_question_from_line(paragraph_text, segs)
                if q_line and a_line and q_line != paragraph_text:
                    stream_pairs.append((q_line, a_line))

            buf_lines = []
            buf_segments = []

        if buf_lines:
            paragraph_text = _normalize_pdf_sentence(' '.join(buf_lines))
            segs = _dedupe_keep_order(buf_segments)
            segs = _merge_split_segments(paragraph_text, segs)
            if (
                segs and
                30 <= len(paragraph_text) <= 1500 and
                1 <= len(segs) <= 24 and
                not _has_option_structure([paragraph_text]) and
                ('，' in paragraph_text or '。' in paragraph_text or ',' in paragraph_text)
            ):
                q_line, a_line = _build_blank_question_from_line(paragraph_text, segs)
                if q_line and a_line and q_line != paragraph_text:
                    stream_pairs.append((q_line, a_line))

        # 优先跨行整段多空，其次逐段/逐行，最后回退到样式答案对。
        generated_pairs = stream_pairs if stream_pairs else (paragraph_pairs if paragraph_pairs else (blank_pairs if blank_pairs else styled_pairs))

        if generated_pairs:
            lines.append('')
            for i, (q_text, ans_text) in enumerate(generated_pairs, 1):
                lines.append(f'{i}. {q_text}')
                lines.append(f'答案：{ans_text}')
        return _normalize_extracted_text('\n'.join(lines).strip())
    except Exception:
        pass

    # 回退到 pypdf / PyPDF2（仅文本）
    try:
        try:
            from pypdf import PdfReader  # type: ignore
        except Exception:
            from PyPDF2 import PdfReader  # type: ignore

        reader = PdfReader(filepath)
        pages = []
        for p in reader.pages:
            pages.append((p.extract_text() or '').strip())
        return _normalize_extracted_text('\n'.join(x for x in pages if x))
    except Exception:
        return None


def extract_text_by_filetype(filepath):
    """按扩展名读取题库文本，支持 txt/pdf/doc/docx。"""
    ext = os.path.splitext(filepath)[1].lower()

    if ext == '.txt':
        return _read_text_file(filepath)

    if ext == '.docx':
        text = _extract_docx_text_with_style(filepath)
        if text:
            return text
        return _extract_docx_text_fallback(filepath)

    if ext == '.doc':
        text = _extract_doc_text_windows(filepath)
        if text:
            return text
        raise ValueError('读取 .doc 失败：请安装 Microsoft Word（COM）或先另存为 .docx/.txt 再导入。')

    if ext == '.pdf':
        text = _extract_pdf_text(filepath)
        if text:
            return text
        raise ValueError('读取 PDF 失败：请先安装 PyMuPDF 或 pypdf，或先转为 .txt/.docx。')

    raise ValueError(f'暂不支持的文件类型：{ext}')


def _normalize_answer_text(answer_text):
    text = re.sub(r'\s+', '', answer_text or '').upper()
    # 全角字母转半角，兼容 DOC/PDF 中的答案格式。
    text = text.translate(str.maketrans('ＡＢＣＤＥＦＧＨ', 'ABCDEFGH'))
    return text


def _clean_answer_text(answer_text):
    """清理答案末尾孤立数字噪声（如换行后的 1/2/3）。"""
    t = (answer_text or '').strip()
    t = re.sub(r'\n\s*\d+\s*$', '', t)
    t = re.sub(r'[；;]\s*\d+\s*$', '', t)
    return t.strip()


def _extract_choice_answer(answer_text, options):
    """尽量从答案文本中提取选项字母（支持 A/B/C、AB、A,C 等格式）。"""
    normalized = _normalize_answer_text(answer_text)
    letters = re.findall(r'[A-H]', normalized)
    if letters:
        return sorted(set(letters), key=letters.index)

    if len(options) == 2:
        # 判断题常见写法：答案为“正确/错误、对/错、是/否”
        true_words = ('正确', '对', '是')
        false_words = ('错误', '错', '否', '不正确')
        for k, v in options.items():
            opt_text = v.strip()
            if any(w in answer_text for w in true_words) and any(w in opt_text for w in true_words):
                return [k]
            if any(w in answer_text for w in false_words) and any(w in opt_text for w in false_words):
                return [k]

    # 答案直接写了选项文本时，尝试反查
    for k, v in options.items():
        opt = v.strip()
        if opt and (opt in answer_text or answer_text in opt):
            return [k]

    return []


def _is_judge_options(options):
    """根据选项文本判断是否为判断题（不依赖答案行）。"""
    if set(options.keys()) != {'A', 'B'}:
        return False

    a_text = options.get('A', '')
    b_text = options.get('B', '')
    judge_words = ('正确', '错误', '对', '错', '是', '否')
    return any(w in a_text for w in judge_words) and any(w in b_text for w in judge_words)


def _is_blank_question(question_text):
    blank_patterns = [
        r'[_＿﹍]{2,}',
        r'（\s*）',
        r'\(\s*\)',
        r'【\s*】',
        r'\[\s*\]',
        r'（\s*[_＿﹍\s]+\s*）',
        r'\(\s*[_＿﹍\s]+\s*\)',
        r'填空'
    ]
    return any(re.search(p, question_text) for p in blank_patterns)


def _looks_like_choice_answer_text(text):
    normalized = _normalize_answer_text(text)
    cleaned = re.sub(r'[\s,，、;；/\\]+', '', normalized)
    if re.fullmatch(r'[A-H]+', cleaned or ''):
        return True
    return any(k in text for k in ('正确', '错误', '对', '错', '是', '否'))


def _split_content_and_answer(lines):
    """将题块行拆分为题干行与答案行，支持无前缀答案在下一行的极端格式。"""
    if len(lines) < 2:
        return None, None

    # 1) 优先匹配显式答案行（支持答案内容在下一行或多行）
    for i, line in enumerate(lines):
        m = ANSWER_LABEL_RE.match(line)
        if not m:
            continue

        head = m.group(1).strip()
        answer_lines = []
        if head:
            answer_lines.append(head)
        if i + 1 < len(lines):
            answer_lines.extend(l.strip() for l in lines[i + 1:] if l.strip())

        answer_text = '\n'.join(answer_lines).strip()
        content_lines = lines[:i]
        if content_lines and answer_text:
            return content_lines, answer_text

    # 2) 无答案前缀时的启发式拆分
    # 2.1 客观题常见：最后一行为标准答案 token（不是选项行）
    last_line = lines[-1].strip()
    content_lines = lines[:-1]
    if content_lines and _looks_like_answer_token(last_line):
        return content_lines, last_line

    # 2.2 填空/简答常见：最后一行为答案（尤其存在挖空标记）
    joined = ' '.join(lines)
    if (
        _is_blank_question(joined)
        and not _has_option_structure(content_lines)
        and not _has_option_structure([last_line])
    ):
        return content_lines, last_line

    # 2.25 简答常见：首行是问题（问号结尾），后续行为答案。
    has_option_in_tail = _has_option_structure(lines[1:])
    if len(lines) >= 2 and ('？' in lines[0] or '?' in lines[0]) and not has_option_in_tail:
        return [lines[0]], '\n'.join(lines[1:]).strip()

    # 2.3 默认视为未识别答案，不强拆最后一行，避免把选项误判为答案。
    return lines, ''


def _parse_questions_loose_qa(text):
    """宽松回退：从连续文本中抽取“问题行 + 后续答案行”的简答/填空。"""
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    if not lines:
        return []

    def is_question_line(line):
        return ('？' in line or '?' in line) and len(line) <= 160

    questions = []
    i = 0
    while i < len(lines):
        if not is_question_line(lines[i]):
            i += 1
            continue

        q_line = re.sub(r'^\s*(?:\d{1,4}|[一二三四五六七八九十百零]+)[.、．\)]\s*', '', lines[i]).strip()
        i += 1
        ans = []
        while i < len(lines) and not is_question_line(lines[i]):
            # 避免把章节标题吞进答案
            if QUESTION_START_RE.match(lines[i]) and len(lines[i]) < 25 and '？' not in lines[i] and '?' not in lines[i]:
                break
            ans.append(lines[i])
            i += 1

        answer_text = '\n'.join(ans).strip()
        if not q_line:
            continue

        questions.append({
            'id': 0,
            'text': q_line,
            'options': {},
            'answer': answer_text,
            'type': 'blank' if _is_blank_question(q_line) else 'short'
        })

    for idx, q in enumerate(questions):
        q['id'] = idx + 1

    return questions

def parse_questions(filepath):
    """解析题库文件，提取所有题目、选项、正确答案（支持 txt/pdf/doc/docx）。"""
    text = extract_text_by_filetype(filepath)
    text = _normalize_extracted_text(text)

    questions = []

    # 优先按题号分块，兼容题目后答案、题目后选项+答案等不同结构。
    lines = text.split('\n')
    current_block = []
    blocks = []
    has_question_boundary = False

    for line in lines:
        if QUESTION_START_RE.match(line.strip()):
            has_question_boundary = True
            if current_block and any(l.strip() for l in current_block):
                blocks.append('\n'.join(current_block).strip())
                current_block = []
        current_block.append(line)

    if current_block and any(l.strip() for l in current_block):
        blocks.append('\n'.join(current_block).strip())

    # 如果按题号分块失败，则退化为按答案行分块。
    if not has_question_boundary:
        current_block = []
        blocks = []
        for line in lines:
            current_block.append(line)
            if ANSWER_LABEL_RE.match(line.strip()):
                blocks.append('\n'.join(current_block).strip())
                current_block = []

    for block in blocks:
        q = parse_single_block(block)
        if q:
            questions.append(q)

    # 文档类文件零命中时，回退到宽松问答抽取（尤其针对 PDF/OCR 场景）。
    ext = os.path.splitext(filepath)[1].lower()
    if not questions and ext in ('.pdf', '.doc', '.docx'):
        questions = _parse_questions_loose_qa(text)

    # 重新编号
    for i, q in enumerate(questions):
        q['id'] = i + 1

    return questions


def parse_single_block(block):
    """解析单个题块，兼容客观题与主观题。"""
    lines = [l.strip() for l in block.strip().split('\n') if l.strip()]
    if not lines:
        return None

    content_lines, answer_text = _split_content_and_answer(lines)
    if not content_lines:
        return None
    answer_text = _clean_answer_text(answer_text)

    # 纠错：如果“答案”本身呈现选项结构，说明拆分过早，应回并到题干继续按客观题解析。
    if answer_text and _has_option_structure([answer_text]) and not _has_option_structure(content_lines):
        content_lines = content_lines + [answer_text]
        answer_text = ''

    # 将所有内容拼成一个长字符串，用换行分隔
    full_text = '\n'.join(content_lines)

    # 预处理：把紧挨在一起的选项拆开（如 "A.正确B.错误" → "A.正确\nB.错误"）
    # 匹配 非选项字母/非空白 后紧跟 选项字母+标点
    full_text = re.sub(r'(?<=[^\n])([A-H])[.、．,，]\s*', r'\n\1. ', full_text)
    # 清理可能产生的多余换行
    full_text = re.sub(r'\n{3,}', '\n\n', full_text)

    # 策略：先用正则拆出所有选项位置
    # 匹配 A. A、A，A) A）A: 等写法
    option_pattern = re.compile(r'(?:^|\n|[ \t]+)([A-H])[.、．,，\)）:]\s*', re.MULTILINE)

    options = {}
    option_positions = []

    for m in option_pattern.finditer(full_text):
        option_positions.append((m.group(1), m.start(), m.end()))

    if not option_positions:
        # 尝试更宽松的匹配：A.xxx 紧贴前文
        option_pattern2 = re.compile(r'([A-H])[.、．,，\)）:]\s*')
        for m in option_pattern2.finditer(full_text):
            option_positions.append((m.group(1), m.start(), m.end()))

    if option_positions:
        # 去重：同一个字母只保留第一次出现（在题目文本之后的）
        seen = set()
        unique_positions = []
        for letter, start, end in option_positions:
            if letter not in seen:
                seen.add(letter)
                unique_positions.append((letter, start, end))
        option_positions = unique_positions

        # 题目文本 = 第一个选项之前的内容
        first_opt_start = option_positions[0][1]
        question_text = full_text[:first_opt_start].strip()

        # 提取每个选项的文本
        for i, (letter, start, end) in enumerate(option_positions):
            if i + 1 < len(option_positions):
                next_start = option_positions[i + 1][1]
                opt_text = full_text[end:next_start].strip()
            else:
                opt_text = full_text[end:].strip()
            # 清理选项文本
            opt_text = opt_text.replace('\n', ' ').strip()
            # 去掉末尾的标点
            opt_text = opt_text.rstrip('。.，,')
            options[letter] = opt_text
    else:
        # 主观题：没有选项，整个内容就是题干
        question_text = full_text.strip()

    # 清理题目文本
    question_text = question_text.replace('\n', ' ')
    # 移除开头的编号
    question_text = re.sub(r'^[\d]+[.、．\s]+', '', question_text)
    # 移除 【单选题】【多选题】 等标签
    question_text = re.sub(r'【[^】]*】\s*', '', question_text)
    # 移除章节标题如 "二.认真认识党史国史..." "三. 一起读马克思"
    question_text = re.sub(r'^[一二三四五六七八九十]+[.、．]\s*[\u4e00-\u9fff]+\s*', '', question_text)
    question_text = question_text.strip()

    if not question_text:
        return None

    # 过滤噪声：既无选项也无答案时，不作为可答题目。
    if not options and not (answer_text or '').strip():
        return None

    # 确定题目类型
    if options:
        # 关键修复：只要有选项，就按客观题判型，不再降级成简答/填空。
        correct_answers = _extract_choice_answer(answer_text or '', options)
        if _is_judge_options(options):
            q_type = 'judge'
        elif len(correct_answers) > 1:
            q_type = 'multi'
        else:
            q_type = 'single'

        # 兜底：若答案未识别出字母，尝试从原始答案文本里提取；仍失败则保留为空列表。
        answer_value = correct_answers
        if not answer_value:
            normalized = _normalize_answer_text(answer_text or '')
            fallback_letters = re.findall(r'[A-H]', normalized)
            if fallback_letters:
                answer_value = sorted(set(fallback_letters), key=fallback_letters.index)
    else:
        q_type = 'blank' if _is_blank_question(question_text) else 'short'
        answer_value = answer_text or ''

    return {
        'id': 0,
        'text': question_text,
        'options': options,
        'answer': answer_value,
        'type': q_type
    }


# ============ 错误记录模块 ============

RECORD_FILE = 'error_record.json'


def load_records():
    """加载历史错误记录"""
    if os.path.exists(RECORD_FILE):
        with open(RECORD_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def save_records(records):
    """保存错误记录"""
    with open(RECORD_FILE, 'w', encoding='utf-8') as f:
        json.dump(records, f, ensure_ascii=False, indent=2)


def get_record(records, qid):
    """获取某题的记录"""
    key = str(qid)
    if key in records:
        return records[key]
    return {'attempts': 0, 'errors': 0}


def update_record(records, qid, is_correct):
    """更新某题的记录"""
    key = str(qid)
    if key not in records:
        records[key] = {'attempts': 0, 'errors': 0}
    records[key]['attempts'] += 1
    if not is_correct:
        records[key]['errors'] += 1
    save_records(records)


# ============ 加权随机抽题 ============

def weighted_random_pick(questions, records):
    """根据错误次数和错误率进行加权随机抽题"""
    weights = []
    for q in questions:
        rec = get_record(records, q['id'])
        attempts = rec['attempts']
        errors = rec['errors']
        error_rate = errors / attempts if attempts > 0 else 0

        # 基础权重
        weight = 1.0

        # 错误次数越多权重越大
        if errors >= 5:
            weight += 3.0
        elif errors >= 3:
            weight += 2.0
        elif errors >= 1:
            weight += 1.0

        # 错误率高的加权
        if attempts > 0:
            if error_rate > 0.5:
                weight += 3.0
            elif error_rate > 0.3:
                weight += 1.5
            elif error_rate > 0:
                weight += 0.5

        # 从未做过的题目也适当提高权重
        if attempts == 0:
            weight += 0.5

        weights.append(weight)

    # 加权随机选择
    total = sum(weights)
    r = random.uniform(0, total)
    cumulative = 0
    for i, w in enumerate(weights):
        cumulative += w
        if r <= cumulative:
            return questions[i]
    return questions[-1]


# ============ GUI 模块 ============

class QuizApp:
    def __init__(self, root, questions, source_path=''):
        self.root = root
        self.questions = questions
        self.source_path = source_path
        self.records = load_records()
        self.current_q = None
        self.selected = set()
        self.submitted = False
        self.answer_revealed = False
        self.option_buttons = {}
        self.judge_buttons = []

        self.root.title("刷题工具")
        self.root.geometry("900x700")
        self.root.configure(bg='#f5f5f5')
        self.root.minsize(700, 550)

        # 字体
        self.title_font = tkfont.Font(family='Microsoft YaHei', size=14, weight='bold')
        self.text_font = tkfont.Font(family='Microsoft YaHei', size=12)
        self.option_font = tkfont.Font(family='Microsoft YaHei', size=11)
        self.small_font = tkfont.Font(family='Microsoft YaHei', size=10)
        self.btn_font = tkfont.Font(family='Microsoft YaHei', size=11, weight='bold')

        self.build_ui()
        self.show_welcome()

    def build_ui(self):
        # 顶部信息栏
        top_frame = tk.Frame(self.root, bg='#2c3e50', height=50)
        top_frame.pack(fill='x')
        top_frame.pack_propagate(False)

        self.info_label = tk.Label(
            top_frame,
            text=f"题库：{os.path.basename(self.source_path) if self.source_path else '未命名'} | 共 {len(self.questions)} 题",
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
            self.scrollable_frame, text="", font=self.text_font,
            fg='#2c3e50', bg='#ffffff', anchor='w', justify='left',
            wraplength=750, padx=15, pady=15, relief='ridge', bd=1
        )
        self.q_text_label.pack(fill='x', pady=(0, 15))

        # 选项容器
        self.options_frame = tk.Frame(self.scrollable_frame, bg='#f5f5f5')
        self.options_frame.pack(fill='x', pady=5)

        # 结果显示
        self.result_label = tk.Label(
            self.scrollable_frame, text="", font=self.text_font,
            fg='#27ae60', bg='#f5f5f5', anchor='w', justify='left',
            wraplength=750
        )
        self.result_label.pack(fill='x', pady=10)

        # 历史记录显示
        self.history_label = tk.Label(
            self.scrollable_frame, text="", font=self.small_font,
            fg='#95a5a6', bg='#f5f5f5', anchor='w'
        )
        self.history_label.pack(fill='x', pady=(0, 5))

        # 底部按钮栏
        btn_frame = tk.Frame(self.root, bg='#ecf0f1', height=60)
        btn_frame.pack(fill='x', side='bottom')
        btn_frame.pack_propagate(False)

        self.submit_btn = tk.Button(
            btn_frame, text="提交答案", font=self.btn_font,
            bg='#3498db', fg='white', activebackground='#2980b9',
            relief='flat', padx=20, pady=8, command=self.submit_answer
        )
        self.submit_btn.pack(side='left', padx=20, pady=10)

        self.next_btn = tk.Button(
            btn_frame, text="下一题 ▶", font=self.btn_font,
            bg='#2ecc71', fg='white', activebackground='#27ae60',
            relief='flat', padx=20, pady=8, command=self.next_question
        )
        self.next_btn.pack(side='left', padx=5, pady=10)

        self.reset_btn = tk.Button(
            btn_frame, text="重置记录", font=self.small_font,
            bg='#e74c3c', fg='white', activebackground='#c0392b',
            relief='flat', padx=10, pady=5, command=self.reset_records
        )
        self.reset_btn.pack(side='right', padx=20, pady=10)

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_mousewheel_linux(self, event):
        if event.num == 4:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self.canvas.yview_scroll(1, "units")

    def show_welcome(self):
        self.q_number_label.config(text="欢迎使用刷题工具！")
        self.q_type_label.config(text="")
        self.q_text_label.config(
            text=f"题库已加载 {len(self.questions)} 道题目。\n\n"
                 f"点击「下一题」开始刷题！\n\n"
                 f"系统会根据你的错误率自动加权抽题，\n"
                 f"错得越多的题越容易被抽到哦～"
        )
        self.result_label.config(text="")
        self.history_label.config(text="")
        self.submit_btn.config(state='disabled')
        self.update_stats()

    def update_stats(self):
        total_attempts = sum(r['attempts'] for r in self.records.values())
        total_errors = sum(r['errors'] for r in self.records.values())
        attempted = len([r for r in self.records.values() if r['attempts'] > 0])
        acc = (1 - total_errors / total_attempts) * 100 if total_attempts > 0 else 0
        self.stats_label.config(
            text=f"已做: {attempted}/{len(self.questions)} | "
                 f"总答题: {total_attempts} | 总正确率: {acc:.1f}%"
        )

    def next_question(self):
        self.submitted = False
        self.answer_revealed = False
        self.selected = set()
        self.current_q = weighted_random_pick(self.questions, self.records)
        self.display_question()

    def display_question(self):
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
        self.q_text_label.config(text=q['text'])
        self.result_label.config(text="")

        # 显示历史记录
        rec = get_record(self.records, q['id'])
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
                    font=self.option_font,
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
        else:
            # 主观题：先不显示任何作答选项，先看答案再自判
            self.result_label.config(
                text='请先自己作答，点击「显示正确答案」后再进行自评。',
                fg='#7f8c8d'
            )
            self.submit_btn.config(text='显示正确答案', state='normal')

        # 滚动到顶部
        self.canvas.yview_moveto(0)

    def toggle_option(self, key):
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
        update_record(self.records, q['id'], is_correct)
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
        rec = get_record(self.records, q['id'])
        rate = rec['errors'] / rec['attempts'] * 100
        self.history_label.config(
            text=f"历史记录：答过 {rec['attempts']} 次，错误 {rec['errors']} 次，错误率 {rate:.0f}%"
        )

        self.submit_btn.config(state='disabled')

    def submit_subjective_result(self, is_correct):
        if self.submitted:
            return

        self.submitted = True
        q = self.current_q

        update_record(self.records, q['id'], is_correct)
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

        rec = get_record(self.records, q['id'])
        rate = rec['errors'] / rec['attempts'] * 100
        self.history_label.config(
            text=f"历史记录：答过 {rec['attempts']} 次，错误 {rec['errors']} 次，错误率 {rate:.0f}%"
        )

        self.submit_btn.config(state='disabled')

    def reset_records(self):
        if messagebox.askyesno("确认", "确定要重置所有错误记录吗？\n此操作不可撤销！"):
            self.records = {}
            save_records(self.records)
            self.update_stats()
            messagebox.showinfo("完成", "所有记录已重置。")


# ============ 主程序 ============

def main():
    # 确定脚本目录
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # 文件选择阶段使用独立隐藏窗口，避免主窗口状态异常导致不显示
    selector = tk.Tk()
    selector.withdraw()
    selector.update_idletasks()

    # 手动选择题库文件
    txt_path = filedialog.askopenfilename(
        title='请选择题库文件',
        initialdir=script_dir,
        filetypes=[
            ('题库文件', '*.txt *.pdf *.doc *.docx'),
            ('文本文件', '*.txt'),
            ('PDF 文件', '*.pdf'),
            ('Word 文件', '*.doc *.docx'),
            ('所有文件', '*.*')
        ],
        parent=selector
    )
    selector.destroy()

    if not txt_path:
        messagebox.showinfo('已取消', '未选择题库文件，程序将退出。')
        return

    print(f"正在解析题库：{txt_path}")
    try:
        questions = parse_questions(txt_path)
    except Exception as e:
        messagebox.showerror('解析失败', f'文件解析失败：\n{e}')
        return
    print(f"成功解析 {len(questions)} 道题目。")

    if not questions:
        messagebox.showerror('错误', '未能解析出任何题目，请检查题库格式。')
        return

    # 主窗口单独创建，确保可见
    root = tk.Tk()
    root.title('刷题工具')
    root.lift()
    root.focus_force()

    # 导入预览窗口：先确认解析结果再进入刷题
    if not show_import_preview(root, questions, txt_path):
        root.destroy()
        return

    # 切换工作目录到脚本目录（使 error_record.json 保存在同一位置）
    os.chdir(script_dir)

    app = QuizApp(root, questions, source_path=txt_path)
    root.mainloop()


def show_import_preview(root, questions, source_path):
    """展示解析预览，用户确认后开始刷题。"""
    win = tk.Toplevel(root)
    win.title('题库导入预览')
    win.geometry('980x640')
    win.minsize(850, 520)
    win.configure(bg='#f5f5f5')
    win.transient(root)
    win.grab_set()

    top = tk.Frame(win, bg='#2c3e50', height=52)
    top.pack(fill='x')
    top.pack_propagate(False)

    type_count = {'single': 0, 'multi': 0, 'judge': 0, 'blank': 0, 'short': 0}
    for q in questions:
        t = q.get('type')
        if t in type_count:
            type_count[t] += 1

    summary = (
        f"文件：{os.path.basename(source_path)}  |  共 {len(questions)} 题"
        f"  |  单选 {type_count['single']}  多选 {type_count['multi']}  判断 {type_count['judge']}"
        f"  填空 {type_count['blank']}  简答 {type_count['short']}"
    )
    tk.Label(
        top, text=summary, fg='white', bg='#2c3e50',
        font=('Microsoft YaHei', 10), anchor='w'
    ).pack(fill='x', padx=12, pady=14)

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

    for q in questions:
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

        tree.insert('', 'end', values=(
            q.get('id', ''),
            type_name.get(q.get('type'), q.get('type', '')),
            answer_preview,
            text_preview
        ))

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

    result = {'ok': False}

    btn_bar = tk.Frame(win, bg='#f5f5f5')
    btn_bar.pack(fill='x', padx=12, pady=(0, 12))

    def on_cancel():
        result['ok'] = False
        win.destroy()

    def on_confirm():
        result['ok'] = True
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
    return result['ok']


if __name__ == '__main__':
    main()
