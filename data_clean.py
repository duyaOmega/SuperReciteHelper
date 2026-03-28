#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从 Word 文档中提取选择题，红色字体标注的选项即为正确答案。
输出格式与 output.txt 一致：
  题号.题干
  A.选项内容
  B.选项内容
  ...
  正确答案：X
"""

import re
import sys
import os
from docx import Document
from docx.shared import RGBColor

def is_red_color(run):
    """判断一个 run 是否为红色字体"""
    try:
        color = run.font.color
        if color and color.rgb:
            r, g, b = color.rgb[0], color.rgb[1], color.rgb[2]
            # 红色：R 值高，G 和 B 值低
            if r > 180 and g < 100 and b < 100:
                return True
        # 也检查 XML 中的颜色属性（有些情况 python-docx 无法直接读取）
        rpr = run._element.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color')
        if rpr is not None:
            val = rpr.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '')
            if val:
                val = val.upper().strip('#')
                if len(val) == 6:
                    r_val = int(val[0:2], 16)
                    g_val = int(val[2:4], 16)
                    b_val = int(val[4:6], 16)
                    if r_val > 180 and g_val < 100 and b_val < 100:
                        return True
                # 常见红色名称
                if val in ('FF0000', 'FF0000', 'CC0000', 'DD0000', 'EE0000',
                           'C00000', 'FF3333', 'FF1111', 'RED'):
                    return True
    except Exception:
        pass
    return False


def extract_paragraphs_with_color(docx_path):
    """
    提取文档中每个段落的文字内容及每个字符是否为红色。
    返回列表，每个元素为 (full_text, [(text, is_red), ...])
    """
    doc = Document(docx_path)
    paragraphs = []
    for para in doc.paragraphs:
        full_text = para.text.strip()
        if not full_text:
            continue
        runs_info = []
        for run in para.runs:
            text = run.text
            if text:
                red = is_red_color(run)
                runs_info.append((text, red))
        paragraphs.append((full_text, runs_info))
    return paragraphs


def parse_questions_from_docx(docx_path):
    """
    解析 Word 文档，提取题目、选项和红色标注的正确答案。
    """
    paragraphs = extract_paragraphs_with_color(docx_path)
    
    questions = []
    current_question = None
    
    for full_text, runs_info in paragraphs:
        text = full_text.strip()
        if not text:
            continue
        
        # 检测是否是选项行（以 A. B. C. D. 开头）
        option_match = re.match(r'^([A-D])[.、．\s]+(.+)$', text)
        
        # 检测是否是题目行（以数字编号开头，或者包含问号等）
        question_match = re.match(r'^[\d]+[.、．)\s]+(.+)$', text)
        
        # 检测同一行内有多个选项的情况（如 "A.正确 B.错误"）
        multi_option_pattern = re.findall(r'([A-D])[.、．]\s*([^A-D.、．]+?)(?=\s*[A-D][.、．]|$)', text)
        
        if option_match and not multi_option_pattern or (option_match and len(multi_option_pattern) <= 1):
            # 单个选项行
            if current_question is not None:
                letter = option_match.group(1)
                opt_text = option_match.group(2).strip()
                
                # 判断此选项是否为红色
                is_red = any(red for t, red in runs_info if t.strip())
                # 更精确：检查选项文字部分是否为红色
                # 有时候只有字母是红色，有时候整行是红色
                red_text = ''.join(t for t, red in runs_info if red)
                if letter in red_text or opt_text[:3] in red_text or len(red_text) > 0:
                    # 这个选项含有红色文字
                    current_question['red_options'].add(letter)
                
                current_question['options'][letter] = opt_text
                
        elif len(multi_option_pattern) >= 2:
            # 同一行多个选项（如判断题 "A.正确B.错误" 或 "A.正确 B.错误"）
            if current_question is not None:
                for letter, opt_text in multi_option_pattern:
                    current_question['options'][letter] = opt_text.strip()
                
                # 检查哪些选项字母对应红色文字
                # 逐个 run 分析
                current_letter = None
                for run_text, is_red in runs_info:
                    # 找到 run 中的选项字母
                    for m in re.finditer(r'([A-D])[.、．]', run_text):
                        current_letter = m.group(1)
                    if is_red and current_letter:
                        current_question['red_options'].add(current_letter)
                    # 如果整个 run 是红色且包含选项字母
                    if is_red:
                        for m in re.finditer(r'([A-D])', run_text):
                            current_question['red_options'].add(m.group(1))
        
        elif question_match or (not option_match and current_question is None) or \
             (re.match(r'^[\d]+', text) and '?' in text or '？' in text) or \
             (re.match(r'^[\d]+[.、．)\s]', text)):
            # 新题目
            if current_question is not None:
                questions.append(current_question)
            
            q_text = text
            # 去掉开头的编号
            q_text = re.sub(r'^[\d]+[.、．)\s]+', '', q_text).strip()
            # 去掉【单选题】【多选题】【判断题】等标签
            q_text = re.sub(r'【[^】]*】\s*', '', q_text).strip()
            
            current_question = {
                'text': q_text,
                'options': {},
                'red_options': set(),
            }
        else:
            # 可能是题目的续行，或者不认识的行
            # 尝试判断是否属于当前题目
            if current_question is not None and not current_question['options']:
                # 题目还没有选项，可能是题干续行
                current_question['text'] += ' ' + text
            elif current_question is None:
                # 可能是新题目（没有标准编号）
                # 检查下面是否紧跟选项
                current_question = {
                    'text': text,
                    'options': {},
                    'red_options': set(),
                }
    
    # 别忘了最后一题
    if current_question is not None:
        questions.append(current_question)
    
    # 过滤掉没有选项的条目
    questions = [q for q in questions if q['options'] and q['red_options']]
    
    return questions


def format_output(questions):
    """格式化输出为 output.txt 的格式"""
    lines = []
    for i, q in enumerate(questions, 1):
        lines.append(f"{i}.{q['text']}")
        for letter in sorted(q['options'].keys()):
            lines.append(f"{letter}.{q['options'][letter]}")
        answer = ''.join(sorted(q['red_options']))
        lines.append(f"正确答案：{answer}")
        lines.append("")  # 空行分隔
    return '\n'.join(lines)


def main():
    # 查找 docx 文件
    docx_path = "必修总题库.docx"
    

    # 命令行参数
    if len(sys.argv) > 1:
        docx_path = sys.argv[1]
    
    if not docx_path or not os.path.exists(docx_path):
        print("错误：未找到 Word 文档！")
        print("用法：python3 extract_questions.py <文件路径.docx>")
        print("或将 .docx 文件放在当前目录下。")
        return
    
    print(f"正在处理：{docx_path}")
    
    # 提取题目
    questions = parse_questions_from_docx(docx_path)
    print(f"成功提取 {len(questions)} 道题目。")
    
    if not questions:
        print("警告：未提取到任何有效题目！")
        print("请确认：")
        print("  1. 文档中包含选择题（A/B/C/D 选项）")
        print("  2. 正确答案用红色字体标注")
        return
    
    # 输出
    output_text = format_output(questions)
    output_path = 'D:\Study\大一下\初党\output.txt'
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(output_text)
    
    print(f"已保存到：{output_path}")
    
    # 也打印前几题预览
    print("\n===== 前3题预览 =====")
    for q in questions[:3]:
        print(f"题目：{q['text'][:60]}...")
        print(f"选项：{q['options']}")
        print(f"答案：{''.join(sorted(q['red_options']))}")
        print()


if __name__ == '__main__':
    main()

