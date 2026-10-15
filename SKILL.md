---
name: summarize-folder
version: 2.0.0
description: "读取路径（文件夹或文件）中的 PDF、DOCX、XLSX、PPTX、文本、图片，AI 跨文件整合分析后生成含图表/数据表的专业报告（Word/Excel/Markdown）。"
metadata:
  requires:
    bins: ["python"]
---

# /summarize-folder

提取 → AI 分析（→ report.json） → 渲染 → 自省。

## 触发方式

```
/summarize-folder <路径1> [<路径2> ...]              → 生成 docx
/summarize-folder <路径...> --format xlsx            → 生成 xlsx
/summarize-folder <路径...> --format md              → 生成 md
```

路径可以是文件夹（扫描其中的支持文件）或直接指定文件，可混合。

## 参数

| 参数 | 必需 | 默认 | 说明 |
|------|------|------|------|
| `<路径...>` | 是 | - | 一个或多个路径，目录→扫描，文件→直接处理 |
| `--format` | 否 | `docx` | 输出格式：`docx` / `xlsx` / `md` |
| `--output` | 否 | `summary.<ext>` | 本地输出路径 |
| `--recursive` | 否 | `false` | 目录路径时是否递归子目录 |

## 脚本位置

所有 Python 脚本与本 SKILL.md 在同一目录。下文命令中的脚本路径均为相对本目录的文件名——执行前先用 `cd` 进入此目录：

```bash
cd <本 skill 所在目录>
```

（例如 Claude Code 中通常安装于 `~/.claude/skills/summarize-folder`，按你的实际路径替换。）

```
summarize-folder/
  SKILL.md              ← 本文件
  mistakes.md           ← 本 skill 的已知错误（Step 0 加载）
  summarize_folder.py   ← 阶段1：文本提取
  render_report.py      ← 阶段3：报告渲染
```

## 执行流程

### Step 0: 读取错误记录

开始执行前，读取本目录的 `mistakes.md`，扫描「快速索引」表，确认本次任务涉及的关键词是否已有已知错误。若匹配到 🔴 阻断级条目，必须在对应步骤显式遵循。

### Step 1: 检查依赖

> ⚠️ 注意：Python import 名和 pip 包名**不是一一对应的**——`import docx` 对应 `pip install python-docx`，`import pptx` 对应 `pip install python-pptx`。用错名字会导致装错包或重复安装。

用 pip 包名一次性检查所有依赖：

```bash
pip list 2>/dev/null | grep -iE "pdfplumber|python-docx|python-pptx|openpyxl|matplotlib"
```

对照下表判断每个包的状态：

| import 名 | pip 包名 | 用途 |
|-----------|----------|------|
| `pdfplumber` | `pdfplumber` | PDF 解析 |
| `docx` | `python-docx` | Word 文档 |
| `pptx` | `python-pptx` | PowerPoint |
| `openpyxl` | `openpyxl` | Excel |
| `matplotlib` | `matplotlib` | 图表渲染 |

对于缺失的包，先 `pip install --dry-run <包名>` 确认不会覆盖已有版本，再正式安装。已安装但版本过旧的用 `pip install --upgrade`。

### Step 2: 提取文本 → extract.json

```bash
python summarize_folder.py \
  <路径1> <路径2> ... \
  --output <temp_dir>/extract.json \
  [--recursive] [--cache <temp_dir>/extract_cache.json]
```

输出 `extract.json`，每个文件包含 `relative_path`、`type`、`status`、`content`、`content_length`、`metadata`。

> **大文件处理**：每个文件均包含 `content_length` 字段。读取 extract.json 后先扫描所有文件的 `content_length`：
> - 总计 < 50,000 字 → 全部读入 context
> - 总计 > 50,000 字 → 对每个文件只取前 3000 字用于 AI 分析，完整内容仍保留在 `extract.json` 的 `content` 字段供查证
> - 单个文件 > 20,000 字 → 读取前 2000 字 + 结尾 1000 字（首尾采样）
> - XLSX/CSV 优先全量读取（结构化数据密度高）

> **图片处理**：`type` 为 `image` 的条目**不做 OCR**，`content` 为空字符串，`metadata.size_bytes` 记录文件体积。这类条目要**用 Read 工具直接查看图片**（Read 支持 PNG/JPG 等格式，会以视觉方式呈现），而不是去读 `content`。
>
> - 图片同样占用 context，且比文字贵得多。**默认最多查看 20 张**；超出时按 `relative_path` 排序取前 20 张，并在报告附录注明「另有 N 张图片未纳入分析」
> - 图片是独立文件时才有 `type=image`；PDF/DOCX/PPTX **内嵌**的图片不会被抽出，脚本不会登记它们
> - 用户明确要求覆盖全部图片时，可分批查看，但要在报告中说明分析是基于全部还是部分图片

### Step 3a: 系统性交叉核验（必须先做）

读取 `extract.json` 后（遵循上述大文件策略），**先核验再分析**。对全部文件内容进行四维交叉检查：

#### 核验维度

**1. 数据一致性**
- 同一数据点在不同文件中是否一致？
- 例（真冲突）：DOCX 说 6 月销售额 17.5 万，CSV 里该月是 12.5 万 → 不一致
- ⚠️ **别把单位换算误判为冲突**：`17.5 万` 与 `175,000` 是同一个数，应判为**一致**；同理 `1.2 亿` = `120,000,000`。只有数值本身不同才算冲突
- 若不一致 → 记录差异，分析可能原因

**2. 计算校验**
- 对 XLSX/CSV 中的汇总数据进行验算
- 合计项 = 各项之和？百分比 = 分子/分母？环比增长 = (本期-上期)/上期？
- 若计算不吻合 → 标注"计算疑似有误"

**3. 时间线对齐**
- 不同文件中提到的时间节点是否吻合？
- 例：PPT 说 Q3 启动，DOCX 说 9 月完成 → 是否在同一时间框架内？
- 若有矛盾 → 列出时间线冲突

**4. 逻辑一致性**
- 不同文件的结论是否自洽？
- 例：DOCX 说市场增长强劲，XLSX 却显示连续下滑？
- 若逻辑矛盾 → 标注"结论与数据不一致"

#### 核验输出

将核验结果记录为 `verification` 字段（写入 report.json 顶层）。格式：

```json
"verification": {
  "status": "passed|warnings|conflicts",
  "consistency_checks": [
    {"check": "销售额数据", "result": "ok", "detail": "DOCX 与 CSV 数据一致"},
    {"check": "预算合计", "result": "mismatch", "detail": "XLSX 合计 174.7 万，各项之和为 174.5 万，差额 0.2 万", "sources": ["budget.xlsx"]}
  ],
  "timeline_check": [
    {"check": "Q3时间线", "result": "ok", "detail": "PPT与DOCX时间框架一致"}
  ],
  "logic_check": [
    {"check": "增长结论", "result": "conflict", "detail": "DOCX称持续增长，但CSV显示5月环比仅增1.9%，增长明显放缓", "sources": ["report.docx", "data.csv"]}
  ]
}
```

**status 判定规则**：
- `passed`：所有检查通过
- `warnings`：存在不一致但可解释（如四舍五入差异、时间口径不同）
- `conflicts`：存在无法解释的矛盾，必须写入报告"补充信息与细节"章节

### Step 3b: 生成 report.json

基于核验结果生成结构化报告。核验中发现的 **conflicts 和 warnings 必须出现在报告的"补充信息与细节"中**。

#### report.json Schema

```json
{
  "meta": {
    "title": "文件汇总 — <主题概括>",
    "source": "<源路径>",
    "generated_at": "<ISO时间>",
    "file_count": 5,
    "error_count": 0
  },
  "verification": { ... },
  "overview": {
    "summary": "<一段话概括所有文件的核心内容>",
    "key_findings": ["发现1", "发现2", "发现3"],
    "data_snapshot": { "指标名": "数值", ... }
  },
  "sections": [
    {
      "heading": "主题 A 标题",
      "blocks": [
        { "type": "paragraph", "text": "..." },
        { "type": "chart", "chart_type": "line|bar|pie", "title": "...",
          "x_values": [...], "series": [{"label":"...","values":[...]}] },
        { "type": "table", "headers": [...], "rows": [[...], ...] },
        { "type": "cross_ref", "text": "来源说明" }
      ]
    }
  ],
  "appendix": {
    "file_list": [{"path":"...","type":"...","status":"ok|error","summary":"..."}],
    "errors": ["错误信息（如有）"]
  }
}
```

#### 新增 Block 类型

| type | 何时用 | 渲染效果 |
|------|--------|---------|
| `verification_note` | 核验发现的不一致/冲突 | `⚠️ 核验发现：...` 黄色高亮 |

#### 分析原则

1. **先核验再分析**：Step 3a 的核验结果必须完成才能进入 3b
2. **全局理解**：通读全部内容，识别统一主题和逻辑线索
3. **主题聚合**：不同文件中讨论同一主题的内容合并到一起
4. **数据关联**：Excel 有数据、Word 有分析、PPT 有结论 → 整合为连贯叙述
5. **去重合并**：多个文件讲同一件事 → 合并，不重复
6. **交叉引用**：每条关键信息标注来源文件

#### 重要规则

- **chart 的数据直接填进 JSON**，不要写描述——把数值写完整
- **中文引号用「」**，避免 JSON 转义
- 从 XLSX/CSV 提取的数值表格，优先做 `table` + `chart` 组合
- section 内部的 blocks 顺序即报告阅读顺序
- **核验发现的冲突，必须写入"补充信息与细节"章节**

### Step 4: 渲染最终文档

```bash
python render_report.py \
  report.json \
  --format <md|docx|xlsx> \
  --output <输出路径> \
  --chart-dir <图表目录>
```

渲染器自动：
- 从 report.json 读取内容
- 用 matplotlib 生成 chart PNG（中文字体自动检测）
- 按格式输出：MD（表格+图片引用）/ DOCX（封面+内嵌图表+原生表格+页码）/ XLSX（多Sheet）

## 本地文件操作命令速查

```bash
# 提取
python summarize_folder.py <路径...> --output extract.json [--recursive] [--cache extract_cache.json]

# 渲染
python render_report.py report.json --format md  --output summary.md   --chart-dir ./charts
python render_report.py report.json --format docx --output summary.docx --chart-dir ./charts
python render_report.py report.json --format xlsx --output summary.xlsx --chart-dir ./charts
```

## 示例

```
/summarize-folder D:/docs/project-a
/summarize-folder D:/docs/a.pdf b.xlsx c.pptx --format xlsx
/summarize-folder D:/docs/project --recursive
/summarize-folder report.docx data.xlsx slides.pptx --format md
```

## Step 5: 自省与迭代

全部步骤完成后，花 30 秒做一次轻量回顾。目的不是写报告，而是**捕捉一个可以立即行动的小改进**。

### 5.1 快速扫描

在脑海里过一遍刚才的执行过程，问自己：

1. 有没有哪个步骤出现了偏差，需要在指令里说清楚？
2. 有没有遗漏的边界情况？（比如文件损坏、编码异常、空文件夹）
3. 用户是否纠正了某个判断或输出？
4. 有没有重复操作可以合并？

### 5.2 判断是否值得记录

- **会复现的** → 写入本目录的 `mistakes.md`
- **只是本次特殊情况** → 不记录

### 5.3 写入格式

按严重度将新条目插入 mistakes.md 对应的分区（`## 🔴 阻断级` / `## 🟡 警告级` / `## ⚪ 提示级`），同时更新顶部「快速索引」表。条目格式：

```markdown
### 🔴 YYYY-MM-DD · 一句话标题

- **关键词**: `tag1` `tag2`
- **场景**: <在什么情况下发生的>
- **错误**: <实际做了什么导致问题>
- **正确**: <应该怎么做>
- **为何会犯**: <为什么当时会选错做法>
```

严重度：🔴 阻断（会直接导致失败）· 🟡 警告（降低质量）· ⚪ 提示（效率技巧）

### 5.4 去重 & 归档

- 写入前先看已有条目，有类似的就更新而非新建
- 如果活跃区已有 15+ 条，顺便把最不可能复现的 2-3 条移到 `## 已归档`

