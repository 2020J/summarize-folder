# summarize-folder 错误记录

> 开始执行本 skill 时读取此文件。按严重度排列：🔴 阻断 → 🟡 警告 → ⚪ 提示

## 快速索引

| 关键词 | 严重度 | 一句话 |
|--------|--------|--------|
| `pip` `包名` `import` | 🔴 | `import docx` ≠ `pip install docx`，用错包名会装到错误的包 |
| `大文件` `context` `截断` | 🟡 | 总内容 > 50k 字时必须遵循 Step 2 的分级读取策略 |

---

## 🔴 阻断级

### 🔴 2026-07-24 · Python import 名和 pip 包名不一致导致装错包

- **关键词**: `pip` `包名` `import` `python-docx` `python-pptx`
- **场景**: Step 1 检查依赖时，用 `pip install` 按 import 名安装
- **错误**: 执行 `pip install docx` 或 `pip install pptx`——pip 会安装名为 `docx` / `pptx` 的包（如果 PyPI 上有的话），而不是需要的 `python-docx` / `python-pptx`。结果：import 仍然失败，或者更糟——装了一个同名但不同的包
- **正确**: 
  1. 用 `pip list | grep -iE "pdfplumber|python-docx|python-pptx|openpyxl|matplotlib"` 检查
  2. 对照 Step 1 中的 import→pip 映射表，确保安装时用 pip 包名
  3. 安装前用 `pip install --dry-run <包名>` 预览，确认不会覆盖已有版本
- **为何会犯**: Python 社区惯例是包名和 import 名相同（如 `numpy`、`pandas`），但这个项目依赖的两个关键包（`python-docx`、`python-pptx`）偏偏不一致。习惯性用 import 名去 pip install 就会踩坑

---

## 🟡 警告级

### 🟡 2026-07-24 · 大文件全量读入导致上下文截断

- **关键词**: `大文件` `context` `截断` `token`
- **场景**: extract.json 中多个文件总字数远超 context 容量，AI 在 Step 3a 阶段需要分析全部文件
- **错误**: 把所有文件 content 全量读入 context，导致模型上下文被截断，前半部分分析结果丢失，最终报告缺失关键章节
- **正确**: Step 2 已规定大文件策略：先扫描所有文件 `content_length`，总计 > 50k 字时按策略执行（摘要、首尾采样、XLSX 优先全量）。关键是在进入 Step 3a 前强制检查总字数
- **为何会犯**: 「数据越完整结果越好」的直觉在有限 context 窗口下是错的。SKILL.md 已经写了策略，但没有在 Step 3a 开头设一个硬性检查点（如「如果总字数 > 50k 且你还没做摘要，就在这里停下先做摘要」），容易跳过去

---

## ⚪ 提示级

<!-- 暂无，发现后在此添加 -->

---

## 已归档

<!-- 被淘汰或不再适用的条目移到这里 -->
