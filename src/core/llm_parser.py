import json
import os
import re
from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, List

from src.utils import CACHE_DIR


class DeepSeekParser:
    """DeepSeek 题库解析器。

    负责读取 txt/docx 文本并通过 DeepSeek API 解析为结构化题目列表。
    """

    def __init__(self, deepseek_api_key: str):
        self.deepseek_api_key = deepseek_api_key
        try:
            from openai import OpenAI
        except ImportError as exc:
            raise ImportError(
                "DeepSeekParser requires the openai package. "
                "Install it with: pip install openai"
            ) from exc
        self.client = OpenAI(api_key=self.deepseek_api_key, base_url="https://api.deepseek.com")

    def read_text_file(self, filepath: str) -> str:
        for enc in ('utf-8-sig', 'utf-8', 'utf-16', 'utf-16-le', 'utf-16-be', 'gb18030', 'gbk'):
            try:
                with open(filepath, 'r', encoding=enc) as f:
                    return f.read()
            except Exception:
                continue
        raise ValueError(f'无法读取文本文件编码：{filepath}')

    def read_docx_file(self, filepath: str) -> str:
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        lines = []
        with zipfile.ZipFile(filepath, 'r') as zf:
            root = ET.fromstring(zf.read('word/document.xml'))
        for p in root.findall('.//w:p', ns):
            line = ''.join((t.text or '') for t in p.findall('.//w:t', ns)).strip()
            if line:
                lines.append(line)
        return '\n'.join(lines)

    def read_file(self, filepath: str) -> str:
        ext = Path(filepath).suffix
        if ext == '.txt':
            return self.read_text_file(filepath)
        if ext == '.docx':
            return self.read_docx_file(filepath)
        raise ValueError(f'暂不支持的文件类型：{ext}（仅支持 .txt / .docx）')

    def parse_text(self, raw_content: str) -> List[Dict]:
        prompt = f"""
            你是一个专业的试题解析专家。请将以下文档中的每道题解析为严格的JSON格式。

            【要求】
            1. 自动识别题型：单选题、多选题、填空题、简答题；多选题会单独标注，否则选择题默认单选题；判断题转化为单选题
            2. 每道题必须包含字段：
            - "type": 题型，包括单选题、多选题、填空题、简答题
            - "question": 题干，仅包含题目文本，不含答案，可能题干存在代码段；填空题的填空区域可能需要根据上下文识别，模糊情况则默认将填空区域设在题目文本的最后
            - "answer": 答案，即提取的正确答案，源文件中的答案存在形式包括题末的答案、文末专门的答案区等，也可能无答案
            题目的可选字段：
            - "options": 仅单选和多选题需要，数组格式 ["A. xxx", "B. xxx", ...]，如果是判断题转换得到的单选题，选项固定为 "A. 正确", "B. 错误"
            3. 如果题干和答案混合在一起，请智能分离
            4. 输出格式：只输出JSON数组，不要任何额外解释，具体格式要求：
            - "type": 四种关键词，单选题、多选题、填空题、简答题
            - "question"：题干字符串，填空题的填空部分用 "____" 标注出来
            - "answer": 多选题答案用英文逗号加空格分隔；无答案时保留该字段，默认填充“本题暂未提供答案”
            - "options": 仅在`type`字段为单选题或多选题时，保留该字段；选项用ABCD等大写英文字母标识，后面加英文句号和空格作分隔

            【文档内容】
            {raw_content}

            【输出示例】
            [
            {{
                "type": "单选题",
                "question": "有STL容器，设其对象为 obj，其类型为 T，容器的元素为整型。该容器不能直接使用标准库算法 std::sort(obj.begin(), obj.end()) 进行排序。则 T 可以是：1. vector，2. list，3. deque，4. set",
                "options": ["A. 1和2", "B. 1和4", "C. 2和4", "D. 1和3"],
                "answer": "C"
            }},
            {{
                "type": "多选题",
                "question": "【多选题】张静老师在《漫谈青年知识分子的成长》中讲到，优秀的历史学家的基本素质包括：",
                "options": ["A. 无谄于权贵，无惧于恶世", "B. 无争于名利，无企于奇俏", "C. 无悲于己身，无动于衷情", "D. 做到冷眼观世事"],
                "answer": "A, B, C, D"
            }},
            {{
                "type": "填空题",
                "question": "跳出历史周期率的第二个答案是____",
                "answer": "党的自我革命"
            }},
            {{
                "type": "简答题",
                "question": "全面依法治国的总目标。",
                "answer": "建设中国特色社会主义法治体系，建设社会主义法治国家。"
            }}
            ]

            请开始解析：
        """
        response = self.client.chat.completions.create(
            model="deepseek-v4-flash",
            messages=[
                {"role": "system", "content": "你是一个专业的试题解析助手，只输出JSON格式，不输出其他内容。"},
                {"role": "user", "content": prompt}
            ],
            temperature=0.1,
            response_format={"type": "json_object"}
        )
        return json.loads(response.choices[0].message.content)

    def parse_file(self, filepath: str) -> List[Dict]:
        raw_content = self.read_file(filepath)
        parsed_content = self.parse_text(raw_content)

        file_name = Path(filepath).stem
        with open(CACHE_DIR / f'{file_name}.json', 'w', encoding='utf-8') as f:
            json.dump(parsed_content, f, ensure_ascii=False, indent=2)
