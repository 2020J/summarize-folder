# summarize-folder

一个 Claude Code Skill：读取本地文件夹中的 **PDF / DOCX / XLSX / PPTX / 文本 / 图片**，跨文件整合分析后生成含图表与数据表的专业报告（Word / Excel / Markdown）。

不是逐文件罗列摘要，而是**按主题聚合**——把散落在多个文件里的同一件事合并成一条连贯叙述，并标注每条信息的来源文件。

## 它解决什么问题

一堆文件（需求文档、预算表、方案 PPT、会议记录）放在同一个文件夹里，要产出一份汇总报告时：

- 人工通读要几小时，且容易漏掉文件之间的**矛盾**——Word 写合计 174.7 万但 Excel 里各项只加到 174.5 万；PPT 说「增长强劲」，Excel 却显示连续三个月环比下滑；同一个指标 Word 写 17.5 万、Excel 里是 12.5 万
- 直接丢给 AI 逐文件总结，得到的是一堆碎片，不是一份报告

这个 skill 在分析**之前**先做四维交叉核验，把矛盾显式标出来。

## 核心能力

### 四维交叉核验（分析前强制执行）

| 维度 | 检查什么 |
|------|----------|
| 数据一致性 | 同一数据点在不同文件中是否一致 |
| 计算校验 | 合计 = 各项之和？百分比 = 分子/分母？环比计算是否正确 |
| 时间线对齐 | 不同文件提到的时间节点是否吻合 |
| 逻辑一致性 | 结论是否与数据自洽 |

核验结果写入报告，发现的矛盾以 `⚠️ 核验发现：...` 显式高亮，并强制写入报告的「补充信息与细节」章节。

### 图片理解

图片**不做 OCR**。脚本只登记路径与体积，由 AI 用 Read 工具**直接看图**——因此能理解图表、截图、版式，而不只是识别文字。中文场景下效果远好于本地 OCR，且**零额外依赖**。

默认最多查看 20 张，超出时按路径排序取前 20，并在报告附录注明未纳入分析的数量。

### 三种输出格式

- **DOCX** — 封面 + 内嵌图表 + 原生表格 + 页码
- **XLSX** — 多 Sheet
- **Markdown** — 表格 + 图片引用

图表用 matplotlib 生成，中文字体自动检测（微软雅黑 / 黑体 / 宋体 / Noto Sans CJK）。

## 安装

克隆到 Claude Code 的 skills 目录：

```bash
git clone https://github.com/2020J/summarize-folder.git ~/.claude/skills/summarize-folder
```

或手动把整个目录复制到 `~/.claude/skills/` 下。

### 依赖

```bash
pip install pdfplumber python-docx python-pptx openpyxl matplotlib
```

> ⚠️ Python import 名和 pip 包名**不是一一对应的**：`import docx` 对应 `pip install python-docx`，`import pptx` 对应 `pip install python-pptx`。用 import 名去装会装到别的包。

## 用法

```
/summarize-folder <路径1> [<路径2> ...]     → 生成 docx
/summarize-folder <路径...> --format xlsx   → 生成 xlsx
/summarize-folder <路径...> --format md     → 生成 md
```

| 参数 | 必需 | 默认 | 说明 |
|------|------|------|------|
| `<路径...>` | 是 | - | 一个或多个路径，目录→扫描，文件→直接处理，可混合 |
| `--format` | 否 | `docx` | 输出格式：`docx` / `xlsx` / `md` |
| `--output` | 否 | `summary.<ext>` | 本地输出路径 |
| `--recursive` | 否 | `false` | 目录路径时是否递归子目录 |

### 示例

```
/summarize-folder D:/docs/project-a
/summarize-folder D:/docs/a.pdf b.xlsx c.pptx --format xlsx
/summarize-folder D:/docs/project --recursive
```

## 支持的文件类型

| 类别 | 扩展名 |
|------|--------|
| 文档 | `.pdf` `.docx` |
| 表格 | `.xlsx` `.xls` `.csv` |
| 演示 | `.pptx` |
| 图片 | `.png` `.jpg` `.jpeg` `.gif` `.bmp` `.webp` `.tif` `.tiff` |
| 文本 | `.txt` `.md` `.json` `.yaml` `.py` `.js` `.sql` `.log` 等 |

## 目录结构

```
summarize-folder/
  SKILL.md              ← 编排指令（Claude 读取）
  mistakes.md           ← 已知错误记录，执行前加载
  summarize_folder.py   ← 阶段1：文本提取
  render_report.py      ← 阶段3：报告渲染
```

## 工作流程

```
提取（extract.json）→ 四维交叉核验 → AI 分析（report.json）→ 渲染 → 自省
```

`report.json` 是中间产物，采用结构化 Schema（`meta` / `verification` / `overview` / `sections` / `appendix`），三种格式的渲染器都从它读取，因此同一份分析可以任意切换输出格式而不必重新分析。

## 已知限制

- **只处理独立图片文件**，PDF / DOCX / PPTX **内嵌**的图片不会被抽出
- 大文件按分级策略采样（总内容 > 50,000 字时每文件只取前 3000 字），避免 context 截断
- 加密文件（需密码的 PDF 等）会提取失败，在报告附录中列出原因，**不中断整体流程**

## License

[MIT](LICENSE)
