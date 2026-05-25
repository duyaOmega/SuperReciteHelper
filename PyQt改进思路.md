# PyQt 改进思路

本文档用于记录当前从 Tkinter 迁移到 PyQt6 的设计思路、已完成改动和后续拆分计划，避免把 README 写得过长。

## ui_pyqt.py
保留主窗口 `QuizWindow`。

负责：

```text
主界面布局
显示题目
下一题
提交答案
调用编辑/统计/管理弹窗
键盘快捷入口
```

## ui_pyqt_dialogs.py

放所有 PyQt 弹窗。

负责：

```text
题目编辑弹窗
管理修改弹窗
考频统计弹窗
题目详情弹窗
```

## ui_pyqt_utils.py


放和界面有关但不该塞进窗口类里的小逻辑。

负责：

```text
TYPE_LABELS
键盘输入规范化
重复题签名
重复题分组
```


## 最后结构大概是：

```text
main.py
parser.py
question.py
question_bank.py
session.py

ui_pyqt.py
ui_pyqt_dialogs.py
ui_pyqt_utils.py
```
