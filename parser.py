#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""题库解析模块（仅识别 txt / docx 纯文本）。

只读取文本内容，不使用字体颜色 / 下划线 / 高亮 / 加粗等样式信息，
也不支持 PDF / .doc 等格式。函数签名保持向后兼容，故名字带
“pdf”/“docx”等历史词的函数仍然保留。
"""
'''
2026-05-10 扔给ai做了缩减，我手动测试了一下能跑。反正我先传到dev分支，后续可以再改。
下面是测试题库之一，识别无误。
1.【单选题】张静老师在《漫谈青年知识分子的成长》中讲到，为人.为学.行事三者中   是根本。
 A.为人
 B.为学
 C.行事
 D.立志
正确答案：A
'''

import re
import os
import zipfile
import xml.etree.ElementTree as ET

# --- 正则与常量 ------------------------------------------------------------
ANSWER_LABEL_RE = re.compile(
    r'^(?:正确答案|答案|参考答案|标准答案|【答案】|\[答案\]|答|参考解答)\s*[：:]?\s*(.*)\s*$')
QUESTION_START_RE = re.compile(
    r'^\s*(?:\d{1,4}(?:[、．\)]|[.](?!\d))|[一二三四五六七八九十百零]+[.、．\)])\s*')
OPTION_PREFIX_RE = re.compile(r'^\s*([A-HＡ-Ｈ])[.、．,，\)）:：]\s*')
OPTION_TOKEN_RE = re.compile(r'([A-HＡ-Ｈ])[.、．,，\)）:：]\s*')

SECTION_HEADING_PATTERNS = (
    ('single', re.compile(r'^\s*单选题\s*$')),
    ('multi',  re.compile(r'^\s*多选题\s*$')),
    ('judge',  re.compile(r'^\s*判断题\s*$')),
    ('blank',  re.compile(r'^\s*填空题\s*$')),
    ('short',  re.compile(r'^\s*简答题\s*$')),
)

JUDGE_WORDS = ('正确', '错误', '对', '错', '是', '否')

# --- 文本规范化与轻量判别 --------------------------------------------------
def _clean_option_text(text):
    t = re.sub(r'[\uf0b7\u2022]+', ' ', str(text or ''))
    return re.sub(r'\s{2,}', ' ', t).strip()

def _detect_section_heading(text):
    t = (text or '').strip()
    for sec, pat in SECTION_HEADING_PATTERNS:
        if pat.match(t):
            return sec
    return None

def _normalize_extracted_text(text):
    """规范化文本，提升题块边界与选项识别稳定性。"""
    if not text:
        return ''
    t = text.replace('\r\n', '\n').replace('\r', '\n').replace('\u3000', ' ')
    # 下面一行将答案行单独分开。
    t = re.sub(
    r'(?<!\n)(?<!正确)(?<!参考)(?<!标准)(正确答案|参考答案|标准答案|答案|答)[：:]',
    r'\n\1：',t)
    # 题号前换行；保留小数不被误拆。
    t = re.sub(r'(?<![\n\d])(\d{1,3}(?:[、．\)]|[.](?!\d)))', r'\n\1', t)
    return re.sub(r'\n{3,}', '\n\n', t).strip()

def _looks_like_option_line(text):
    if not text:
        return False
    return bool(OPTION_PREFIX_RE.match(text)) or len(OPTION_TOKEN_RE.findall(text)) >= 2

def _has_option_structure(lines):
    if not lines:
        return False
    if sum(1 for l in lines if OPTION_PREFIX_RE.match(l)) >= 2:
        return True
    return len(OPTION_TOKEN_RE.findall(' '.join(lines))) >= 2

def _looks_like_question_start_line(text):
    """判断是否像“新题开始”，避免把代码行误判成题号。"""
    t = (text or '').strip()
    if not QUESTION_START_RE.match(t):
        return False
    body = QUESTION_START_RE.sub('', t, count=1).strip()
    if not body:
        return False
    has_cjk = bool(re.search(r'[\u4e00-\u9fff]', body))
    if not has_cjk and any(s in body for s in ('{', '}', ';', '#include', '::')):
        return False
    return True

def _looks_like_answer_token(text):
    if not text or _looks_like_option_line(text):
        return False
    cleaned = re.sub(r'[\s,，、;；/\\]+', '', _normalize_answer_text(text))
    if re.fullmatch(r'[A-H]+', cleaned or ''):
        return True
    return text.strip() in JUDGE_WORDS

# --- 文本读取（仅 txt / docx）---------------------------------------------
def _read_text_file(filepath):
    for enc in ('utf-8-sig', 'utf-8', 'utf-16', 'utf-16-le', 'utf-16-be', 'gb18030', 'gbk'):
        try:
            with open(filepath, 'r', encoding=enc) as f:
                return f.read()
        except Exception:
            continue
    raise ValueError(f'无法读取文本文件编码：{filepath}')

def _extract_docx_text(filepath):
    """提取 docx 纯文本（含表格、页眉、文本框等所有 w:t 节点）。"""
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    lines = []
    with zipfile.ZipFile(filepath, 'r') as zf:
        root = ET.fromstring(zf.read('word/document.xml'))
    # 段落级聚合：每个 w:p 对应一行，避免把同段不同 run 拆成多行。
    for p in root.findall('.//w:p', ns):
        line = ''.join((t.text or '') for t in p.findall('.//w:t', ns)).strip()
        if line:
            lines.append(line)
    return '\n'.join(lines)

def extract_text_by_filetype(filepath):
    """按扩展名读取题库文本，仅支持 txt / docx。"""
    ext = os.path.splitext(filepath)[1].lower()
    if ext == '.txt':
        return _read_text_file(filepath)
    if ext == '.docx':
        return _extract_docx_text(filepath)
    raise ValueError(f'暂不支持的文件类型：{ext}（仅支持 .txt / .docx）')

# --- 答案文本与选项处理 ----------------------------------------------------
def _normalize_answer_text(answer_text):
    """标准化答案文本：去空白、转大写、全角字母转半角。"""
    text = re.sub(r'\s+', '', answer_text or '').upper()
    return text.translate(str.maketrans('ＡＢＣＤＥＦＧＨ', 'ABCDEFGH'))

def _clean_answer_text(answer_text):
    t = (answer_text or '').strip()
    t = re.sub(r'\n\s*\d+\s*$', '', t)
    t = re.sub(r'[；;]\s*\d+\s*$', '', t)
    return t.strip()

def _extract_choice_answer(answer_text, options):
    """从答案文本中提取选项字母（支持 A/B/C、AB、A,C 等格式）。"""
    letters = re.findall(r'[A-H]', _normalize_answer_text(answer_text))
    if letters:
        return sorted(set(letters), key=letters.index)
    if len(options) == 2:
        true_w = ('正确', '对', '是')
        false_w = ('错误', '错', '否', '不正确')
        for k, v in options.items():
            opt = v.strip()
            if any(w in answer_text for w in true_w) and any(w in opt for w in true_w):
                return [k]
            if any(w in answer_text for w in false_w) and any(w in opt for w in false_w):
                return [k]
    clean = (answer_text or '').strip()
    if clean:
        for k, v in options.items():
            opt = v.strip()
            if opt and (opt in clean or clean in opt):
                return [k]
    return []

def _is_judge_options(options):
    if set(options.keys()) != {'A', 'B'}:
        return False
    a, b = options.get('A', ''), options.get('B', '')
    return any(w in a for w in JUDGE_WORDS) and any(w in b for w in JUDGE_WORDS)

def _is_blank_question(question_text):
    patterns = (r'[_＿﹍]{2,}', r'（\s*）', r'\(\s*\)', r'【\s*】', r'\[\s*\]',
                r'（\s*[_＿﹍\s]+\s*）', r'\(\s*[_＿﹍\s]+\s*\)', r'填空')
    return any(re.search(p, question_text) for p in patterns)

# --- 文末答案区识别与回填 --------------------------------------------------
def _extract_answer_keys_from_text(text):
    """从文末答案区抽取“题号 -> 答案”映射，按单选/多选/判断分区。

    只处理常见格式：题号与答案在同一行（如 12.A 或 12：正确）。
    """
    lines = [l.strip() for l in str(text or '').split('\n') if l.strip()]
    result = {k: {} for k in ('single', 'multi', 'judge', 'blank', 'short', 'generic')}
    section, started = None, False

    title_re = re.compile(r'^\s*(?:(?:单选|多选|判断|填空|简答)题\s*)?'
                          r'(?:参考答案|标准答案|答案)\s*[：:]?\s*$')
    pair_re = re.compile(r'(\d{1,4})\s*[.、:：]\s*([A-HＡ-Ｈ]{1,8}|正确|错误|对|错)')
    sec_kw = {'single': '单选', 'multi': '多选', 'judge': '判断',
              'blank': '填空', 'short': '简答'}

    for line in lines:
        if title_re.match(line):
            started = True
            continue
        if not started:
            continue
        switched = False
        for sk, kw in sec_kw.items():
            if kw in line and '答案' in line:
                section = sk
                switched = True
                break
        if switched:
            continue
        if _detect_section_heading(line) and '答案' not in line:
            section = None
            continue
        for m in pair_re.finditer(line):
            qno = int(m.group(1))
            ans = _normalize_answer_text(m.group(2))
            if section in result:
                result[section][qno] = ans
            result['generic'].setdefault(qno, ans)
    return result

def _fill_answers_from_answer_keys(questions, answer_keys):
    """把文末答案区映射回题目列表。"""
    section_index = {'single': 0, 'multi': 0, 'judge': 0}
    for q in questions:
        options = q.get('options') or {}
        if not options or q.get('answer'):
            continue
        hint = q.get('section_hint')
        ans_token = ''
        if hint in section_index:
            section_index[hint] += 1
            ans_token = answer_keys.get(hint, {}).get(section_index[hint], '')
        if not ans_token:
            src_no = q.get('source_no')
            if src_no is not None:
                ans_token = answer_keys.get('generic', {}).get(src_no, '')
        if not ans_token:
            continue
        letters = _extract_choice_answer(ans_token, options)
        if not letters:
            letters = re.findall(r'[A-H]', _normalize_answer_text(ans_token))
        if letters:
            q['answer'] = sorted(set(letters), key=letters.index)
            if hint == 'multi':
                q['type'] = 'multi'
            elif hint == 'judge':
                q['type'] = 'judge'
            elif hint == 'single' and q.get('type') != 'judge':
                q['type'] = 'single'

# --- 题块拆分与单题解析 ----------------------------------------------------
def _split_content_and_answer(lines):
    """将题块行拆分为题干行与答案行。

    主要场景：
      1) 显式答案标签（“答案：xxx”）；
      2) 末行为 A/B/C/D 或对错的客观题答案；
      3) 填空题最后一行为答案；
      4) 首行问句、其余为简答答案。
    """
    if len(lines) < 2:
        return None, None

    # 1) 显式答案标签
    for i, line in enumerate(lines):
        m = ANSWER_LABEL_RE.match(line)
        if not m:
            continue
        head = m.group(1).strip()
        tail = [l.strip() for l in lines[i + 1:] if l.strip()]
        content = list(lines[:i])
        if not content:
            continue
        ans_text = '\n'.join([head] + tail).strip() if head else '\n'.join(tail).strip()
        if ans_text:
            return content, ans_text

    # 2) 末行像客观题答案 token
    last = lines[-1].strip()
    content = lines[:-1]
    if content and _looks_like_answer_token(last):
        return content, last

    # 3) 填空题：含挖空标记，末行通常是答案
    if (_is_blank_question(' '.join(lines)) and not _has_option_structure(content)
            and not _has_option_structure([last])):
        return content, last

    # 4) 首行问句、其余为答案
    if (('？' in lines[0] or '?' in lines[0])
            and not _has_option_structure(lines[1:])):
        return [lines[0]], '\n'.join(lines[1:]).strip()

    return lines, ''

def _parse_numbered_qa_blocks(text):
    """解析“每题一行问题 + 后续多行答案”的编号简答题格式。"""
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    if not lines:
        return []
    blocks, current = [], []
    for line in lines:
        if _looks_like_question_start_line(line):
            if current:
                blocks.append(current)
            current = [line]
        elif current:
            current.append(line)
    if current:
        blocks.append(current)

    questions = []
    for block in blocks:
        if len(block) < 2 or _has_option_structure(block):
            continue
        q_text = QUESTION_START_RE.sub('', block[0]).strip()
        ans_text = _clean_answer_text('\n'.join(block[1:]).strip())
        if not q_text or not ans_text:
            continue
        questions.append({'id': 0, 'text': q_text, 'options': {}, 'answer': ans_text,
                          'type': 'blank' if _is_blank_question(q_text) else 'short'})
    for idx, q in enumerate(questions, 1):
        q['id'] = idx
    return questions

# --- 公共：填空题题干自动挖空（被 ui_main.py 直接调用）---------------------
def _mask_blank_question_text(question_text, answer_text):
    """若填空题题干未挖空，则按答案片段自动替换为空位。"""
    text = (question_text or '').strip()
    if not text or _is_blank_question(text):
        return text
    ans = (answer_text or '').strip()
    if not ans:
        return text
    # 优先匹配“（1）答案；（2）答案”
    pairs = re.findall(r'（\s*(\d+)\s*）\s*([^；;\n]+)', ans)
    if pairs:
        out, replaced = text, False
        for idx, (_, seg) in enumerate(pairs, 1):
            seg = seg.strip()
            if seg and seg in out:
                out = out.replace(seg, f'（{idx}）______', 1)
                replaced = True
        if replaced:
            return out
    chunks = [c.strip() for c in re.split(r'[；;]+', ans) if c.strip()]
    out, replaced = text, 0
    for i, chunk in enumerate(chunks, 1):
        chunk = re.sub(r'^（\s*\d+\s*）', '', chunk).strip()
        if chunk and chunk in out:
            out = out.replace(chunk, f'（{i}）______', 1)
            replaced += 1
    return out if replaced > 0 else text

# --- 公共：历史兼容名（仅做通用文本噪声清理）-------------------------------
def _clean_pdf_judge_text_noise(text):
    """通用文本噪声清理。函数名保留以兼容历史调用方，已不依赖 PDF。"""
    t = re.sub(r'\s+', ' ', str(text or '')).strip()
    if not t:
        return t
    t = re.sub(r'(?<=[\u4e00-\u9fff，,、。；;：:])\s+\d{1,3}\s+(?=[\u4e00-\u9fff])', ' ', t)
    t = re.sub(r'\s*(?:对|错|正确|错误)\s+\d{1,3}\s+[\u4e00-\u9fff].*$', '', t)
    return re.sub(r'\s{2,}', ' ', t).strip(' .．。 ，,；;')

# --- 主流程 ----------------------------------------------------------------
def parse_single_block(block):
    """解析单个题块，兼容客观题与主观题。"""
    lines = [l.strip() for l in block.strip().split('\n') if l.strip()]
    if not lines:
        return None

    source_no = None
    m_no = re.match(r'^\s*(\d{1,4})\s*[.、．\)]', lines[0])
    if m_no:
        source_no = int(m_no.group(1))

    content_lines, answer_text = _split_content_and_answer(lines)
    if not content_lines:
        return None
    answer_text = _clean_answer_text(answer_text)

    # 答案本身像选项结构时，回并到题干继续按客观题解析。
    if (answer_text and _has_option_structure([answer_text])
            and not _has_option_structure(content_lines)):
        content_lines = content_lines + [answer_text]
        answer_text = ''

    full_text = '\n'.join(content_lines)
    # 把紧挨在一起的选项拆开（如 "A.正确B.错误" -> 分两行）
    full_text = re.sub(r'(?<=[^\n])([A-H])[.、．,，]\s*', r'\n\1. ', full_text)
    full_text = re.sub(r'\n{3,}', '\n\n', full_text)

    opt_pat = re.compile(r'(?:^|\n|[ \t]+)([A-H])[.、．,，\)）:]\s*', re.MULTILINE)
    positions = [(m.group(1), m.start(), m.end()) for m in opt_pat.finditer(full_text)]
    if not positions:
        positions = [(m.group(1), m.start(), m.end())
                     for m in re.finditer(r'([A-H])[.、．,，\)）:]\s*', full_text)]

    options = {}
    if positions:
        seen, unique = set(), []
        for letter, start, end in positions:
            if letter not in seen:
                seen.add(letter)
                unique.append((letter, start, end))
        positions = unique
        question_text = full_text[:positions[0][1]].strip()
        for i, (letter, start, end) in enumerate(positions):
            next_start = positions[i + 1][1] if i + 1 < len(positions) else len(full_text)
            opt = full_text[end:next_start].replace('\n', ' ').strip().rstrip('。.，,')
            options[letter] = _clean_option_text(opt)
    else:
        question_text = full_text.strip()

    # 清理题干
    question_text = question_text.replace('\n', ' ')
    question_text = re.sub(r'[\uf0b7\u2022]+', ' ', question_text)
    question_text = re.sub(r'^[\d]+[.、．\s]+', '', question_text)
    question_text = re.sub(r'【[^】]*】\s*', '', question_text)
    question_text = re.sub(r'^[一二三四五六七八九十]+[.、．]\s*[\u4e00-\u9fff]+\s*', '', question_text)
    question_text = question_text.strip()

    if not question_text:
        return None
    if not options and not (answer_text or '').strip():
        return None

    if options:
        correct = _extract_choice_answer(answer_text or '', options)
        if _is_judge_options(options):
            q_type = 'judge'
        elif len(correct) > 1:
            q_type = 'multi'
        else:
            q_type = 'single'
        answer_value = correct
    else:
        q_type = 'blank' if _is_blank_question(question_text) else 'short'
        answer_value = answer_text or ''

    return {'id': 0, 'text': question_text, 'options': options,
            'answer': answer_value, 'type': q_type, 'source_no': source_no}

def parse_questions(filepath):
    """解析题库文件，提取题目、选项、正确答案（仅 txt / docx）。"""
    text = _normalize_extracted_text(extract_text_by_filetype(filepath))

    questions = []
    lines = text.split('\n')
    current_block, blocks = [], []
    has_boundary, current_section = False, None

    for line in lines:
        sec = _detect_section_heading(line.strip())
        if sec:
            current_section = sec
            continue
        if _looks_like_question_start_line(line.strip()):
            has_boundary = True
            if current_block and any(l.strip() for l in current_block):
                blocks.append(('\n'.join(current_block).strip(), current_section))
                current_block = []
        current_block.append(line)
    if current_block and any(l.strip() for l in current_block):
        blocks.append(('\n'.join(current_block).strip(), current_section))

    # 按题号分块失败时，退化为按答案行分块。
    if not has_boundary:
        current_block, blocks, current_section = [], [], None
        for line in lines:
            sec = _detect_section_heading(line.strip())
            if sec:
                current_section = sec
                continue
            current_block.append(line)
            if ANSWER_LABEL_RE.match(line.strip()):
                blocks.append(('\n'.join(current_block).strip(), current_section))
                current_block = []

    for block, hint in blocks:
        q = parse_single_block(block)
        if q:
            if hint:
                q['section_hint'] = hint
            questions.append(q)

    _fill_answers_from_answer_keys(questions, _extract_answer_keys_from_text(text))

    # 没解析出题目时再尝试编号简答回退。
    if not questions:
        questions = _parse_numbered_qa_blocks(text)

    # 主观题答案中含“（1）xxx；（2）yyy”形式时，转为填空题并自动挖空。
    for q in questions:
        if q.get('type') == 'short':
            ans = str(q.get('answer', '') or '')
            if re.search(r'（\s*\d+\s*）\s*[^；;\n]+', ans):
                q['type'] = 'blank'
                q['text'] = _mask_blank_question_text(q.get('text', ''), ans)

    for i, q in enumerate(questions):
        q['id'] = i + 1
    return questions

def build_parse_candidates(filepath):
    """构建题库解析候选结果，供预览时切换比对。

    简化后只保留单一“自动识别”方案：原本的“红色客观 / 样式填空 /
    强制判断”候选都依赖 docx 颜色或 PDF，不再适用。
    """
    return [('自动识别（推荐）', parse_questions(filepath), '基于纯文本的综合解析策略')]
