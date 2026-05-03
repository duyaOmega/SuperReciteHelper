# SuperReciteHelper 项目文档

## 1. 项目概述

**SuperReciteHelper（超级背诵助手）** 是一个基于 Python + Tkinter 的桌面刷题工具，支持从多种文件格式（TXT、PDF、DOC、DOCX）导入题库，并提供智能抽题、错题记录、考频统计等功能。

### 1.1 核心特性

- 多格式题库导入：支持 `.txt`、`.pdf`、`.doc`、`.docx`
- 智能答案识别：Word 红色标记、PDF 下划线/样式、文末答案区自动回填
- 五种题型支持：单选题、多选题、判断题、填空题、简答题
- 加权随机抽题：错得越多的题越容易被抽到
- 重复题检测与规避：自动识别重复题组，避免短时间重复出现
- 手动编辑功能：支持在导入前和刷题过程中修改题目与答案
- 数据持久化：做题记录、手动修改、应用状态自动保存

---

## 2. 文件结构与职责

```
SuperReciteHelper/
├── main.py              # 程序入口，负责启动流程与文件选择
├── parser.py            # 题库解析模块，核心解析引擎
├── question.py          # 题目编辑对话框与答案格式化
├── question_bank.py     # 应用状态、题目身份键、记录持久化
├── session.py           # 刷题会话与加权抽题策略
├── ui_main.py           # 图形界面模块（Tkinter GUI）
├── __init__.py           # Python 包标记文件
└── PROJECT_DOCUMENTATION.md  # 本文档
```

---

## 3. 启动流程 (`main.py`)

```
程序启动
  │
  ├─ 启用 Windows 高 DPI 感知
  │
  ├─ 隐藏主窗口，创建文件选择器
  │
  ├─ 用户选择题库文件（支持多文件、增量添加）
  │   └─ 可选择继续上次文件 / 追加新文件 / 全新选择
  │
  ├─ 解析题库文件
  │   ├─ 单文件：直接调用 build_parse_candidates()
  │   └─ 多文件：逐个解析后合并，重排题号
  │
  ├─ 显示导入预览窗口
  │   └─ 用户可切换解析方案、手动编辑题目
  │
  └─ 启动刷题主界面 QuizApp
```

### 3.1 关键函数

| 函数 | 作用 |
|------|------|
| `main()` | 程序主入口，协调整个启动流程 |
| `_choose_startup_files()` | 启动时文件选择逻辑，支持继续上次文件 |
| `_choose_files_incrementally()` | 支持多次"增加文件"的增量选择 |
| `_dedupe_existing_paths()` | 去重并过滤不存在的文件路径 |
| `show_import_preview()` | 解析预览窗口，确认后开始刷题 |

---

## 4. 题库解析模块 (`parser.py`)

这是项目最复杂的模块，负责从各种文件格式中提取题目、选项和答案。

### 4.1 核心解析流程

```
文件输入
  │
  ├─ 按文件类型提取原始文本
  │   ├─ TXT: 直接读取（自动检测编码）
  │   ├─ DOCX: python-docx 提取 / XML 回退
  │   ├─ DOC: Windows COM 转换
  │   └─ PDF: PyMuPDF 样式提取 / pypdf 文本回退
  │
  ├─ 文本规范化（_normalize_extracted_text）
  │
  ├─ 题块分割
  │   ├─ 按题号分割（优先）
  │   └─ 按答案行分割（回退）
  │
  ├─ 逐题解析（parse_single_block）
  │   ├─ 拆分题干与答案
  │   ├─ 识别选项结构
  │   └─ 判定题型
  │
  ├─ 文末答案区回填（_extract_answer_keys_from_text）
  │
  └─ 后处理
      ├─ PDF 判断题转换
      ├─ 填空题样式识别
      └─ 重新编号
```

### 4.2 题型识别逻辑

| 题型 | 识别条件 |
|------|----------|
| **单选题** | 有选项 + 1个正确答案 |
| **多选题** | 有选项 + 多个正确答案 |
| **判断题** | 仅有 A/B 选项 + 内容为"正确/错误" |
| **填空题** | 无选项 + 题干含下划线/空位占位符 |
| **简答题** | 无选项 + 非填空题 |

### 4.3 答案识别策略

#### 4.3.1 客观题答案

1. **文末答案区**：解析"参考答案"区域的题号-答案映射
2. **红色标记**（DOCX）：红色字体的选项为正确答案
3. **勾选符号**：选项前的 ☑、✅、√ 等符号
4. **样式标记**（PDF）：下划线、粗体、非黑色字体

#### 4.3.2 主观题答案

1. **显式答案行**：匹配"答案：xxx"格式
2. **样式提取**（DOCX/PDF）：强调样式文本视为答案片段
3. **填空自动挖空**：将答案片段替换为 `______`

### 4.4 关键函数说明

#### 文本提取

| 函数 | 作用 |
|------|------|
| `extract_text_by_filetype()` | 按扩展名分发到对应的文本提取函数 |
| `_read_text_file()` | 读取 TXT 文件，自动检测编码 |
| `_extract_docx_text_with_style()` | 用 python-docx 提取 DOCX 文本，识别字体样式 |
| `_extract_docx_text_fallback()` | 不依赖第三方库的 DOCX XML 解析 |
| `_extract_pdf_text()` | PDF 文本提取，支持样式识别 |
| `_extract_doc_text_windows()` | 使用 Windows COM 转换 .doc 文件 |

#### 答案解析

| 函数 | 作用 |
|------|------|
| `_extract_answer_keys_from_text()` | 从文末答案区提取题号->答案映射 |
| `_fill_answers_from_answer_keys()` | 将答案区映射回题目列表 |
| `_extract_choice_answer()` | 从答案文本提取选项字母 |
| `_normalize_answer_text()` | 标准化答案文本（去空白、转大写、全角转半角） |

#### 题块解析

| 函数 | 作用 |
|------|------|
| `parse_single_block()` | 解析单个题块，识别选项与题型 |
| `_split_content_and_answer()` | 将题块行拆分为题干与答案 |
| `_looks_like_question_start_line()` | 判断是否为新题开始 |
| `_looks_like_option_line()` | 判断是否为选项行 |
| `_has_option_structure()` | 跨行判断是否包含选项结构 |

#### DOCX 专用解析

| 函数 | 作用 |
|------|------|
| `_parse_docx_questions_with_red()` | 按红色选项识别正确答案 |
| `_parse_docx_numbered_choice_questions()` | 解析 Word 自动编号的选择题 |
| `_parse_docx_styled_blank_questions()` | 从样式中提取填空题答案 |
| `_merge_docx_blank_questions()` | 合并样式填空题到解析结果 |

#### PDF 专用解析

| 函数 | 作用 |
|------|------|
| `_extract_styled_segments_from_spans()` | 从 PDF span 提取样式片段 |
| `_extract_underlined_segments_from_pdf_line()` | 从 PDF 行提取下划线标注 |
| `_build_blank_question_from_line()` | 将样式答案片段转为填空题 |
| `_postprocess_pdf_to_judge()` | PDF 后处理：转判断题 |
| `_split_compound_placeholder_judge_questions()` | 拆分复合判断题 |

### 4.5 `build_parse_candidates()` 函数

构建题库解析候选结果，供预览时切换比对：

```python
返回值: [(方案名, 题目列表, 方案描述), ...]
```

- **自动识别（推荐）**：综合策略自动选择并合并
- **仅红色选项识别**：仅对 DOCX，按红色字体识别答案
- **仅样式填空识别**：仅对 DOCX，将强调样式视为填空答案
- **红色客观 + 样式填空**：组合方案
- **强制判断题模式**：仅对 PDF，将无选项题按判断题处理

---

## 5. 题目编辑模块 (`question.py`)

负责题目编辑对话框和答案格式化。

### 5.1 关键函数

| 函数 | 作用 |
|------|------|
| `_format_answer_text()` | 将答案值统一格式化为可展示字符串 |
| `_parse_manual_answer_for_question()` | 按题型校验并解析手动输入答案 |
| `_format_options_for_edit()` | 把选项字典转成多行可编辑文本 |
| `_parse_manual_options_text()` | 解析手动输入的选项文本 |
| `_show_question_edit_dialog()` | 弹出题目编辑对话框 |

### 5.2 编辑对话框功能

- 修改题目文本（多行编辑）
- 切换题型（单选/多选/判断/填空/简答）
- 编辑选项（客观题）
- 编辑答案（支持字母输入和文本输入）
- 实时校验输入格式

---

## 6. 持久化模块 (`question_bank.py`)

管理应用状态、题目身份键和做题记录的持久化。

### 6.1 存储位置

```
Windows: %APPDATA%/SuperReciteHelper/
Linux:   ~/.local/share/SuperReciteHelper/
```

### 6.2 存储文件

| 文件 | 内容 |
|------|------|
| `app_state.json` | 应用状态（如最近打开文件） |
| `error_record.json` | 做题记录（作答次数、错误次数） |
| `question_edits.json` | 手动编辑的题目覆盖数据 |

### 6.3 题目身份键系统

为每道题生成稳定的身份标识，跨会话/跨题号可复用：

```
_base_key:  由"原始题干+选项+题型"计算的 SHA1 哈希
_record_key: 用于答题记录统计的键（默认与 _base_key 一致）
```

**设计目的**：即使题号变化或手动编辑题目，也能正确关联历史记录。

### 6.4 关键函数

#### 应用状态

| 函数 | 作用 |
|------|------|
| `load_app_state()` | 读取应用级状态 |
| `save_app_state()` | 保存应用级状态 |

#### 手动编辑

| 函数 | 作用 |
|------|------|
| `load_manual_question_edits()` | 读取用户手动编辑过的题目覆盖数据 |
| `save_manual_question_edits()` | 持久化题目手动修改映射 |
| `apply_manual_question_edits()` | 将持久化编辑应用到当前题集 |
| `upsert_manual_question_edit()` | 写入某题的手动修改 |

#### 做题记录

| 函数 | 作用 |
|------|------|
| `load_records()` | 加载历史错误记录（含旧版本迁移） |
| `save_records()` | 保存错误记录 |
| `get_record()` | 获取某题的记录 |
| `update_record()` | 更新某题的记录 |

#### 身份键

| 函数 | 作用 |
|------|------|
| `_build_question_record_key()` | 生成稳定题目键 |
| `_ensure_question_identity_fields()` | 确保题目具备稳定键与原始内容快照 |
| `_record_key()` | 统一记录键入口 |

---

## 7. 抽题策略模块 (`session.py`)

实现加权随机抽题算法。

### 7.1 `weighted_random_pick()` 算法

根据以下因素为每道题计算权重：

| 因素 | 权重影响 |
|------|----------|
| 错误次数 ≥5 | +3.0 |
| 错误次数 ≥3 | +2.0 |
| 错误次数 ≥1 | +1.0 |
| 错误率 >50% | +3.0 |
| 错误率 >30% | +1.5 |
| 错误率 >0% | +0.5 |
| 从未做过 | +0.5 |

**核心思想**：错得越多、错误率越高的题，被抽中的概率越大。

### 7.2 `PracticeSession` 类

轻量刷题会话对象：

```python
class PracticeSession:
    questions      # 当前题集
    records        # 做题记录
    current_question  # 当前题目

    pick_next()    # 按加权策略抽取下一题
```

---

## 8. 图形界面模块 (`ui_main.py`)

使用 Tkinter 构建的刷题主界面。

### 8.1 `QuizApp` 类

主界面类，包含所有 UI 逻辑和交互处理。

#### 初始化流程

```
__init__()
  │
  ├─ 加载手动编辑并应用
  │
  ├─ 构建重复题分组
  │
  ├─ 设置自适应窗口大小
  │
  ├─ 定义字体
  │
  └─ build_ui() 构建界面控件
```

#### 界面布局

```
┌─────────────────────────────────────────────┐
│ 顶部信息栏：题库名称 | 统计信息              │
├─────────────────────────────────────────────┤
│ 主内容区（可滚动）                           │
│  ├─ 题号标签                                 │
│  ├─ 题型标签                                 │
│  ├─ 题目文本                                 │
│  ├─ 选项区（客观题按钮 / 主观题自评按钮）     │
│  ├─ 结果显示                                 │
│  └─ 历史记录                                 │
├─────────────────────────────────────────────┤
│ 底部按钮栏                                   │
│  ├─ 操作行：提交答案 | 下一题 | 键盘输入      │
│  └─ 工具行：考频统计 | 编辑当前题 | 管理修改 | 重置记录 │
└─────────────────────────────────────────────┘
```

### 8.2 核心交互流程

#### 客观题作答流程

```
下一题 → 显示题目与选项 → 用户选择选项 → 提交答案
  │                                         │
  │                                         ├─ 自动判分
  │                                         ├─ 高亮正确/错误选项
  │                                         └─ 更新记录
  │
  └─ 下一题（回车或点击按钮）
```

#### 主观题作答流程

```
下一题 → 显示题目 → 用户自行作答 → 显示正确答案
  │                                      │
  │                                      └─ 用户自评（答对/答错）
  │                                              │
  │                                              └─ 更新记录
  └─ 下一题
```

### 8.3 关键方法

#### 界面构建

| 方法 | 作用 |
|------|------|
| `build_ui()` | 构建主界面控件 |
| `show_welcome()` | 显示欢迎页 |
| `_refresh_layout()` | 动态更新换行宽度 |

#### 题目显示

| 方法 | 作答 |
|------|------|
| `display_question()` | 渲染当前题目、历史记录与作答控件 |
| `update_stats()` | 刷新顶部作答统计信息 |

#### 作答交互

| 方法 | 作用 |
|------|------|
| `toggle_option()` | 处理选项点击（单选覆盖/多选切换） |
| `submit_answer()` | 提交答案，客观题自动判分，主观题显示答案 |
| `submit_subjective_result()` | 记录主观题自评结果 |
| `next_question()` | 抽取下一题（含重复题规避） |

#### 键盘输入

| 方法 | 作用 |
|------|------|
| `_process_keyboard_enter()` | 统一处理回车行为 |
| `_select_objective_by_keyboard()` | 键盘输入映射为选项选中 |
| `_submit_subjective_by_keyboard()` | 键盘输入映射为主观题自评 |
| `_normalize_keyboard_text()` | 规范化键盘输入（兼容全角） |

#### 重复题检测

| 方法 | 作用 |
|------|------|
| `_question_signature()` | 生成题目归一签名 |
| `_build_duplicate_groups()` | 构建重复题分组 |
| `_is_recent_duplicate_pick()` | 判断是否近期重复 |

#### 编辑功能

| 方法 | 作用 |
|------|------|
| `edit_current_question()` | 编辑当前题 |
| `manage_manual_edits()` | 管理已保存的手动改题记录 |

#### 统计功能

| 方法 | 作用 |
|------|------|
| `show_frequency_stats()` | 展示题目考频统计 |

### 8.4 其他工具函数

| 函数 | 作用 |
|------|------|
| `_enable_windows_high_dpi()` | 启用 Windows 高 DPI 感知 |
| `_dedupe_existing_paths()` | 去重并过滤不存在的文件路径 |
| `_choose_files_incrementally()` | 支持多次"增加文件"选择 |
| `_choose_startup_files()` | 启动时文件选择逻辑 |
| `show_import_preview()` | 解析预览窗口 |

---

## 9. 数据结构

### 9.1 题目对象 (`question`)

```python
{
    'id': int,                    # 题号（会话内连续）
    'text': str,                  # 题干文本
    'options': {                  # 选项字典（客观题）
        'A': str,
        'B': str,
        ...
    },
    'answer': list | str,         # 答案（客观题为字母列表，主观题为文本）
    'type': str,                  # 题型：single/multi/judge/blank/short
    'source_no': int | None,      # 原文题号
    'section_hint': str | None,   # 分区提示（single/multi/judge）
    'source_file': str | None,    # 来源文件名（多文件合并时）

    # 身份键字段（自动生成）
    '_base_key': str,             # 稳定主键（SHA1 哈希）
    '_record_key': str,           # 记录统计键
    '_orig_text': str,            # 原始题干快照
    '_orig_answer': str | list,   # 原始答案快照
    '_orig_type': str,            # 原始题型快照
    '_orig_options': dict,        # 原始选项快照
}
```

### 9.2 做题记录 (`record`)

```python
{
    'q:sha1hash': {
        'attempts': int,    # 作答次数
        'errors': int       # 错误次数
    },
    ...
}
```

### 9.3 手动编辑 (`edit`)

```python
{
    'q:sha1hash': {
        'type': str,              # 修改后的题型
        'options': dict,          # 修改后的选项
        'text': str,              # 修改后的题干
        'answer': list | str,     # 修改后的答案
        'orig_type': str,         # 原始题型
        'orig_options': dict,     # 原始选项
        'orig_text': str,         # 原始题干
        'orig_answer': str,       # 原始答案
        'updated_at': str,        # 更新时间
        'preview': str,           # 题干预览（前80字符）
    },
    ...
}
```

---

## 10. 依赖项

### 10.1 Python 标准库

- `tkinter` - GUI 框架
- `re` - 正则表达式
- `json` - JSON 序列化
- `os` - 文件系统操作
- `hashlib` - SHA1 哈希
- `zipfile` - ZIP 文件处理（DOCX 解析回退）
- `xml.etree.ElementTree` - XML 解析
- `ctypes` - Windows API 调用
- `tempfile` - 临时文件

### 10.2 第三方库（可选）

| 库 | 用途 | 必需 |
|----|------|------|
| `python-docx` | DOCX 文本提取与样式识别 | 推荐 |
| `PyMuPDF (fitz)` | PDF 文本提取与样式识别 | 推荐 |
| `pypdf` / `PyPDF2` | PDF 文本提取（回退） | 可选 |
| `win32com` | .doc 文件转换（仅 Windows） | 可选 |

---

## 11. 运行方式

```bash
# 直接运行
python main.py

# 或在项目目录下
cd SuperReciteHelper
python main.py
```

程序启动后会弹出文件选择对话框，选择题库文件即可开始刷题。

---

## 12. 扩展与定制建议

### 12.1 添加新题型

1. 在 `parser.py` 的 `SECTION_HEADING_PATTERNS` 中添加分区标题
2. 在 `parse_single_block()` 中添加题型识别逻辑
3. 在 `ui_main.py` 的 `display_question()` 中添加显示逻辑
4. 在 `ui_main.py` 的 `submit_answer()` 中添加判分逻辑

### 12.2 修改抽题策略

修改 `session.py` 中的 `weighted_random_pick()` 函数的权重计算逻辑。

### 12.3 添加新的文件格式支持

1. 在 `parser.py` 中实现新的文本提取函数
2. 在 `extract_text_by_filetype()` 中添加分支
3. 在 `ui_main.py` 的 `_choose_files_incrementally()` 中添加文件类型过滤器

---

## 13. 已知限制与待改进项

参考项目中的 `To Do List.txt`：

- UI 信息栏长文件名显示不全
- 填空题判断逻辑待完善
- 抽题概率策略可优化（增加"从未抽到题目保底"）
- 可考虑添加笔记功能
