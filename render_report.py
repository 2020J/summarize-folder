#!/usr/bin/env python3
"""
render_report.py — render structured report.json to MD / DOCX / XLSX.

Input:  report.json (structured by Claude AI)
Output: summary.md | summary.docx | summary.xlsx
Side:   chart PNGs saved to --chart-dir
"""

import argparse
import io
import json
import os
import sys
from pathlib import Path

# ===========================================================================
# Chinese font detection
# ===========================================================================
def setup_matplotlib():
    """Detect a CJK-capable font and configure matplotlib."""
    import matplotlib

    matplotlib.use("Agg")  # headless
    import matplotlib.pyplot as plt
    from matplotlib.font_manager import FontProperties

    # Common Windows CJK fonts in preference order
    candidates = [
        "Microsoft YaHei",      # 微软雅黑
        "SimHei",               # 黑体
        "SimSun",               # 宋体
        "KaiTi",                # 楷体
        "FangSong",             # 仿宋
        "Noto Sans CJK SC",
    ]
    available = {f.name for f in FontProperties().get_font_family_names()} if hasattr(FontProperties(), 'get_font_family_names') else set()
    # Fallback: try to find via font_manager
    from matplotlib import font_manager
    sys_fonts = {f.name for f in font_manager.fontManager.ttflist}
    for name in candidates:
        if name in sys_fonts:
            plt.rcParams["font.family"] = name
            plt.rcParams["axes.unicode_minus"] = False
            return name
    return "sans-serif"


# ===========================================================================
# Chart generation
# ===========================================================================
def render_chart(block, index, chart_dir):
    """Generate a single chart PNG. Returns (filename, Path)."""
    import matplotlib.pyplot as plt
    import matplotlib.ticker as mticker

    chart_dir = Path(chart_dir)
    chart_dir.mkdir(parents=True, exist_ok=True)

    chart_type = block.get("chart_type", "bar")
    title = block.get("title", f"Chart {index}")
    x_values = block.get("x_values", [])
    series_list = block.get("series", [])
    width_px = block.get("width", 800)
    height_px = block.get("height", 450)

    dpi = 100
    fig_w = width_px / dpi
    fig_h = height_px / dpi
    fig, ax = plt.subplots(figsize=(fig_w, fig_h), dpi=dpi)

    x = range(len(x_values))
    colors = ["#2E86AB", "#A23B72", "#F18F01", "#C73E1D", "#6A994E"]

    for si, s in enumerate(series_list):
        label = s["label"]
        vals = s["values"]
        color = colors[si % len(colors)]

        if chart_type == "line":
            ax.plot(x, vals, marker="o", label=label, color=color, linewidth=2)
        elif chart_type == "bar":
            offset = (si - (len(series_list) - 1) / 2) * 0.8 / max(len(series_list), 2)
            bar_w = 0.8 / len(series_list)
            ax.bar([i + offset for i in x], vals, bar_w, label=label, color=color)
        elif chart_type == "pie":
            if si == 0:
                ax.pie(vals, labels=x_values, autopct="%1.1f%%", colors=colors[:len(vals)])
            break
        else:
            ax.bar(x, vals, label=label, color=color)

    if chart_type != "pie":
        ax.set_xticks(list(x))
        ax.set_xticklabels(x_values, rotation=30, ha="right", fontsize=9)
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
        ax.legend(fontsize=9)

    ax.set_title(title, fontsize=13, fontweight="bold")
    fig.tight_layout()

    fname = f"chart_{index:03d}.png"
    fig.savefig(chart_dir / fname, dpi=dpi)
    plt.close(fig)
    return fname


# ===========================================================================
# MD renderer
# ===========================================================================
def render_md(report, chart_dir, output_path):
    lines = []
    meta = report["meta"]

    lines.append(f"# {meta['title']}")
    lines.append(f"")
    lines.append(f"> 📂 源路径：{meta['source']}  |  🕐 {meta['generated_at'][:10]}  |  📄 {meta['file_count']} 个文件")
    lines.append(f"")

    # Overview
    ov = report.get("overview", {})
    lines.append(f"## 综合概述")
    lines.append(f"")
    lines.append(ov.get("summary", ""))
    lines.append(f"")

    if ov.get("key_findings"):
        lines.append("### 核心发现")
        lines.append("")
        for f in ov["key_findings"]:
            lines.append(f"- {f}")
        lines.append("")

    if ov.get("data_snapshot"):
        ds = ov["data_snapshot"]
        lines.append("### 数据总览")
        lines.append("")
        lines.append("| 指标 | 数值 |")
        lines.append("| --- | --- |")
        for k, v in ds.items():
            lines.append(f"| {k} | {v} |")
        lines.append("")

    # Sections
    chart_idx = 0
    for sec in report.get("sections", []):
        lines.append(f"## {sec['heading']}")
        lines.append("")
        for block in sec.get("blocks", []):
            bt = block["type"]
            if bt == "paragraph":
                lines.append(block["text"])
                lines.append("")
            elif bt == "table":
                headers = block.get("headers", [])
                rows = block.get("rows", [])
                if headers:
                    lines.append("| " + " | ".join(headers) + " |")
                    lines.append("| " + " | ".join(["---"] * len(headers)) + " |")
                for row in rows:
                    lines.append("| " + " | ".join(str(c) for c in row) + " |")
                lines.append("")
            elif bt == "chart":
                chart_idx += 1
                fname = render_chart(block, chart_idx, chart_dir)
                img_path = f"{chart_dir}/{fname}".replace("\\", "/")
                lines.append(f"![{block.get('title', 'Chart')}]({img_path})")
                lines.append(f"*{block.get('data_source', '')}*")
                lines.append("")
            elif bt == "cross_ref":
                lines.append(f"> 📎 {block['text']}")
                lines.append("")
            elif bt == "verification_note":
                lines.append(f"> ⚠️ **核验发现**：{block['text']}")
                lines.append("")

    # Verification summary (if present)
    ver = report.get("verification")
    if ver:
        lines.append("---")
        lines.append("## 交叉核验结果")
        lines.append(f"状态：{'✅ 全部通过' if ver.get('status')=='passed' else '⚠️ 存在不一致' if ver.get('status')=='warnings' else '❌ 存在冲突'}")
        lines.append("")
        for check in ver.get("consistency_checks", []):
            icon = "✅" if check["result"] == "ok" else "⚠️" if check["result"] == "mismatch" else "❌"
            lines.append(f"- {icon} {check['check']}：{check['detail']}")
        lines.append("")

    # Appendix
    lines.append("---")
    lines.append("## 附录：文件清单")
    lines.append("")
    lines.append("| 文件 | 类型 | 状态 | 说明 |")
    lines.append("|------|------|------|------|")
    for f in report.get("appendix", {}).get("file_list", []):
        lines.append(
            f"| {f.get('path', '')} | {f.get('type', '')} | "
            f"{f.get('status', '')} | {f.get('summary', '')} |"
        )

    errors = report.get("appendix", {}).get("errors", [])
    if errors:
        lines.append("")
        lines.append("## 附录：处理错误")
        lines.append("")
        for e in errors:
            lines.append(f"- ⚠️ {e}")

    output_path = Path(output_path)
    output_path.write_text("\n".join(lines), encoding="utf-8")
    return output_path


# ===========================================================================
# DOCX renderer
# ===========================================================================
def render_docx(report, chart_dir, output_path):
    from docx import Document
    from docx.shared import Inches, Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.section import WD_ORIENT

    doc = Document()

    # ── Default style ─────────────────────────────────
    style = doc.styles["Normal"]
    style.font.name = "Microsoft YaHei"
    style.font.size = Pt(10.5)
    style.paragraph_format.space_after = Pt(6)
    style.paragraph_format.line_spacing = 1.35
    # CJK fallback: set east-asian font
    from docx.oxml.ns import qn
    style.element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")

    meta = report["meta"]

    # ── Helper: format a table ─────────────────────────
    def fmt_table(table, col_widths=None):
        """Apply consistent formatting to a table."""
        from docx.oxml import OxmlElement
        for row_idx, row in enumerate(table.rows):
            for col_idx, cell in enumerate(row.cells):
                for p in cell.paragraphs:
                    p.paragraph_format.space_before = Pt(2)
                    p.paragraph_format.space_after = Pt(2)
                    for run in p.runs:
                        run.font.name = "Microsoft YaHei"
                        run.element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
                        if row_idx == 0:
                            run.font.size = Pt(9.5)
                            run.font.bold = True
                        else:
                            run.font.size = Pt(9)
                # Alignment: header center, data left
                if row_idx == 0:
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        # Column widths
        if col_widths:
            for ci, w in enumerate(col_widths):
                if ci < len(table.columns):
                    for row in table.rows:
                        row.cells[ci].width = Cm(w)

    def estimate_col_widths(headers, rows):
        """Estimate column widths (cm) from content."""
        widths = []
        for ci, h in enumerate(headers):
            max_len = len(str(h))
            for row in rows:
                if ci < len(row):
                    max_len = max(max_len, len(str(row[ci])))
            # Approximate: CJK char ~0.5cm, ASCII char ~0.3cm
            w = max_len * 0.32 + 1.0
            widths.append(min(w, 12))  # cap at 12cm
        return widths

    # ── Cover page ─────────────────────────────────────
    doc.add_paragraph("")
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_p.add_run(meta["title"])
    run.font.size = Pt(24)
    run.font.name = "Microsoft YaHei"
    run.element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
    run.bold = True

    info_p = doc.add_paragraph()
    info_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    info_run = info_p.add_run(f"\n📂 {meta['source']}\n🕐 {meta['generated_at'][:10]}\n📄 {meta['file_count']} 个文件")
    info_run.font.size = Pt(11)
    info_run.font.name = "Microsoft YaHei"
    info_run.element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")

    doc.add_page_break()

    # ── Overview ───────────────────────────────────────
    ov = report.get("overview", {})
    doc.add_heading("综合概述", level=1)
    if ov.get("summary"):
        doc.add_paragraph(ov["summary"])

    if ov.get("key_findings"):
        doc.add_heading("核心发现", level=2)
        for f in ov["key_findings"]:
            p = doc.add_paragraph(f, style="List Bullet")

    if ov.get("data_snapshot"):
        doc.add_heading("数据总览", level=2)
        ds = ov["data_snapshot"]
        t = doc.add_table(rows=len(ds)+1, cols=2, style="Light Grid Accent 1")
        t.cell(0, 0).text = "指标"
        t.cell(0, 1).text = "数值"
        for i, (k, v) in enumerate(ds.items(), 1):
            t.cell(i, 0).text = k
            t.cell(i, 1).text = str(v)
        fmt_table(t, col_widths=[4.5, 10])

    # ── Sections ───────────────────────────────────────
    chart_idx = 0
    chart_dir_p = Path(chart_dir)
    for sec in report.get("sections", []):
        doc.add_heading(sec["heading"], level=1)
        for block in sec.get("blocks", []):
            bt = block["type"]
            if bt == "paragraph":
                doc.add_paragraph(block["text"])
            elif bt == "table":
                headers = block.get("headers", [])
                rows = block.get("rows", [])
                num_cols = len(headers) or (len(rows[0]) if rows else 1)
                t = doc.add_table(rows=len(rows)+1, cols=num_cols, style="Light Grid Accent 1")
                for ci, h in enumerate(headers):
                    t.cell(0, ci).text = h
                for ri, row in enumerate(rows, 1):
                    for ci, val in enumerate(row):
                        t.cell(ri, ci).text = str(val)
                col_ws = estimate_col_widths(headers, rows)
                fmt_table(t, col_widths=col_ws)
                doc.add_paragraph("")
            elif bt == "chart":
                chart_idx += 1
                fname = render_chart(block, chart_idx, chart_dir)
                img_path = chart_dir_p / fname
                if img_path.exists():
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.add_run().add_picture(str(img_path), width=Inches(5.5))
                    ds = block.get("data_source", "")
                    if ds:
                        cap = doc.add_paragraph(ds)
                        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        if cap.runs:
                            cap.runs[0].font.size = Pt(9)
                doc.add_paragraph("")
            elif bt == "cross_ref":
                p = doc.add_paragraph()
                run = p.add_run(f"📎 {block['text']}")
                run.font.size = Pt(9)
                run.font.color.rgb = RGBColor(128, 128, 128)
            elif bt == "verification_note":
                p = doc.add_paragraph()
                run = p.add_run(f"⚠️ 核验发现：{block['text']}")
                run.font.size = Pt(10)
                run.font.color.rgb = RGBColor(200, 150, 0)

    # ── Verification summary ───────────────────────────
    ver = report.get("verification")
    if ver:
        doc.add_page_break()
        doc.add_heading("交叉核验结果", level=1)
        status = ver.get("status", "")
        status_text = "✅ 全部通过" if status == "passed" else "⚠️ 存在不一致" if status == "warnings" else "❌ 存在冲突"
        doc.add_paragraph(f"状态：{status_text}")
        for check in ver.get("consistency_checks", []):
            icon = "✅" if check["result"] == "ok" else "⚠️" if check["result"] == "mismatch" else "❌"
            doc.add_paragraph(f"{icon} {check['check']}：{check['detail']}", style="List Bullet")

    # ── Appendix ───────────────────────────────────────
    doc.add_page_break()
    doc.add_heading("附录：文件清单", level=1)
    fl = report.get("appendix", {}).get("file_list", [])
    t = doc.add_table(rows=len(fl)+1, cols=4, style="Light Grid Accent 1")
    for ci, h in enumerate(["文件", "类型", "状态", "说明"]):
        t.cell(0, ci).text = h
    for ri, f in enumerate(fl, 1):
        t.cell(ri, 0).text = f["path"]
        t.cell(ri, 1).text = f["type"]
        t.cell(ri, 2).text = f["status"]
        t.cell(ri, 3).text = f.get("summary", "")
    fmt_table(t, col_widths=[5, 2, 1.5, 8])

    errors = report.get("appendix", {}).get("errors", [])
    if errors:
        doc.add_heading("处理错误", level=2)
        for e in errors:
            doc.add_paragraph(f"⚠️ {e}", style="List Bullet")

    doc.save(str(output_path))
    return output_path


# ===========================================================================
# XLSX renderer
# ===========================================================================
def render_xlsx(report, chart_dir, output_path):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    meta = report["meta"]

    # Helper
    header_font = Font(bold=True, size=11)
    header_fill = PatternFill(start_color="2E86AB", end_color="2E86AB", fill_type="solid")
    header_font_w = Font(bold=True, size=11, color="FFFFFF")
    wrap = Alignment(wrap_text=True, vertical="top")
    thin = Side(style="thin")
    border = Border(top=thin, left=thin, right=thin, bottom=thin)

    def style_header(ws, row, num_cols):
        for c in range(1, num_cols+1):
            cell = ws.cell(row=row, column=c)
            cell.font = header_font_w
            cell.fill = header_fill
            cell.alignment = wrap
            cell.border = border

    # --- Sheet: 报告概要 ---
    ws0 = wb.active
    ws0.title = "报告概要"
    ws0.cell(1, 1, meta["title"]).font = Font(bold=True, size=16)
    ws0.cell(2, 1, f"源路径：{meta['source']}    日期：{meta['generated_at'][:10]}    文件数：{meta['file_count']}")
    ov = report.get("overview", {})
    ws0.cell(4, 1, "综合概述").font = header_font
    ws0.cell(5, 1, ov.get("summary", ""))

    if ov.get("key_findings"):
        row = 7
        ws0.cell(row, 1, "核心发现").font = header_font
        row += 1
        for f in ov["key_findings"]:
            ws0.cell(row, 1, f"• {f}")
            row += 1

    if ov.get("data_snapshot"):
        ds = ov["data_snapshot"]
        row = ws0.max_row + 2
        ws0.cell(row, 1, "数据总览").font = header_font
        row += 1
        ws0.cell(row, 1, "指标"); ws0.cell(row, 2, "数值")
        style_header(ws0, row, 2)
        row += 1
        for k, v in ds.items():
            ws0.cell(row, 1, k); ws0.cell(row, 2, str(v))
            row += 1
    ws0.column_dimensions["A"].width = 30
    ws0.column_dimensions["B"].width = 20

    # --- Per-section sheets ---
    chart_idx = 0
    for sec in report.get("sections", []):
        safe_name = sec["heading"][:31].replace("/", "-").replace("\\", "-")
        ws = wb.create_sheet(title=safe_name)
        row = 1
        ws.cell(row, 1, sec["heading"]).font = Font(bold=True, size=14)
        row += 2

        for block in sec.get("blocks", []):
            bt = block["type"]
            if bt == "paragraph":
                ws.cell(row, 1, block["text"])
                ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
                ws.cell(row, 1).alignment = wrap
                row += 1
            elif bt == "table":
                headers = block.get("headers", [])
                rows_data = block.get("rows", [])
                for ci, h in enumerate(headers, 1):
                    ws.cell(row, ci, h)
                if headers:
                    style_header(ws, row, len(headers))
                row += 1
                for rd in rows_data:
                    for ci, val in enumerate(rd, 1):
                        ws.cell(row, ci, str(val)).border = border
                    row += 1
                row += 1
            elif bt == "chart":
                chart_idx += 1
                fname = render_chart(block, chart_idx, chart_dir)
                img_path = Path(chart_dir) / fname
                if img_path.exists():
                    from openpyxl.drawing.image import Image as XLImage
                    img = XLImage(str(img_path))
                    img.width = min(img.width, 600)
                    img.height = min(img.height, 337)
                    ws.add_image(img, f"A{row}")
                    row += 20  # leave space for image
                ws.cell(row, 1, f"图表: {block.get('title', '')}  (来源: {block.get('data_source', '')})")
                row += 2
            elif bt == "cross_ref":
                ws.cell(row, 1, f"📎 {block['text']}").font = Font(size=9, color="808080")
                row += 1
            elif bt == "verification_note":
                ws.cell(row, 1, f"⚠️ 核验发现：{block['text']}").font = Font(size=10, color="C89600")
                row += 1

        ws.column_dimensions["A"].width = 25
        for c in range(2, 7):
            ws.column_dimensions[get_column_letter(c)].width = 20

    # --- Verification sheet ---
    ver = report.get("verification")
    if ver:
        ws_v = wb.create_sheet(title="交叉核验")
        ws_v.cell(1, 1, "交叉核验结果").font = Font(bold=True, size=14)
        status = ver.get("status", "")
        status_text = "全部通过" if status == "passed" else "存在不一致" if status == "warnings" else "存在冲突"
        ws_v.cell(2, 1, f"状态：{status_text}")
        ws_v.cell(4, 1, "检查项").font = header_font
        ws_v.cell(4, 2, "结果").font = header_font
        ws_v.cell(4, 3, "说明").font = header_font
        style_header(ws_v, 4, 3)
        r = 5
        for check in ver.get("consistency_checks", []):
            ws_v.cell(r, 1, check["check"])
            ws_v.cell(r, 2, check["result"])
            ws_v.cell(r, 3, check["detail"])
            r += 1
        ws_v.column_dimensions["A"].width = 25
        ws_v.column_dimensions["B"].width = 12
        ws_v.column_dimensions["C"].width = 60

    # --- Appendix sheet ---
    ws_app = wb.create_sheet(title="附录")
    ws_app.cell(1, 1, "文件清单").font = Font(bold=True, size=14)
    fl = report.get("appendix", {}).get("file_list", [])
    hdrs = ["文件", "类型", "状态", "说明"]
    for ci, h in enumerate(hdrs, 1):
        ws_app.cell(3, ci, h)
    style_header(ws_app, 3, 4)
    for ri, f in enumerate(fl, 4):
        ws_app.cell(ri, 1, f["path"]).border = border
        ws_app.cell(ri, 2, f["type"]).border = border
        ws_app.cell(ri, 3, f["status"]).border = border
        ws_app.cell(ri, 4, f.get("summary", "")).border = border
    for c in range(1, 5):
        ws_app.column_dimensions[get_column_letter(c)].width = 30

    errors = report.get("appendix", {}).get("errors", [])
    if errors:
        ws_app.cell(ws_app.max_row + 2, 1, "处理错误").font = header_font
        for ri, e in enumerate(errors, ws_app.max_row + 1):
            ws_app.cell(ri, 1, f"⚠️ {e}")

    wb.save(str(output_path))
    return output_path


# ===========================================================================
# Validation
# ===========================================================================
def validate_report(report):
    """Validate report.json structure. Returns list of error strings (empty = valid)."""
    errors = []

    if not isinstance(report, dict):
        return ["report.json 必须是一个 JSON 对象"]

    # meta
    meta = report.get("meta")
    if not isinstance(meta, dict):
        errors.append("缺少 meta 字段或格式错误")
    else:
        for field in ("title", "source", "generated_at", "file_count", "error_count"):
            if meta.get(field) is None:
                errors.append(f"meta.{field} 不能为空（必须是数字或字符串）")

    # sections
    sections = report.get("sections")
    if not isinstance(sections, list) or len(sections) == 0:
        errors.append("sections 必须是非空数组")
    else:
        valid_block_types = {"paragraph", "table", "chart", "cross_ref", "verification_note"}
        valid_chart_types = {"line", "bar", "pie"}

        for si, sec in enumerate(sections):
            if not isinstance(sec.get("heading"), str) or not sec["heading"].strip():
                errors.append(f"sections[{si}].heading 不能为空")
            blocks = sec.get("blocks")
            if not isinstance(blocks, list) or len(blocks) == 0:
                errors.append(f"sections[{si}] ({sec.get('heading', '?')}).blocks 必须是非空数组")
                continue
            for bi, block in enumerate(blocks):
                bt = block.get("type")
                if bt not in valid_block_types:
                    errors.append(
                        f"sections[{si}].blocks[{bi}]: type \"{bt}\" 无效，"
                        f"允许值: {', '.join(sorted(valid_block_types))}"
                    )
                    continue
                if bt == "paragraph" and not isinstance(block.get("text"), str):
                    errors.append(f"sections[{si}].blocks[{bi}] (paragraph): 缺少 text 字段")
                if bt == "table":
                    if not isinstance(block.get("headers"), list):
                        errors.append(f"sections[{si}].blocks[{bi}] (table): headers 必须是数组")
                    if not isinstance(block.get("rows"), list):
                        errors.append(f"sections[{si}].blocks[{bi}] (table): rows 必须是数组")
                if bt == "chart":
                    ct = block.get("chart_type", "")
                    if ct not in valid_chart_types:
                        errors.append(
                            f"sections[{si}].blocks[{bi}] (chart): chart_type \"{ct}\" 无效，"
                            f"允许值: {', '.join(sorted(valid_chart_types))}"
                        )
                    if not isinstance(block.get("x_values"), list):
                        errors.append(f"sections[{si}].blocks[{bi}] (chart): x_values 必须是数组")
                    series = block.get("series")
                    if not isinstance(series, list) or len(series) == 0:
                        errors.append(f"sections[{si}].blocks[{bi}] (chart): series 必须是非空数组")
                    else:
                        for sri, sr in enumerate(series):
                            if not isinstance(sr.get("label"), str):
                                errors.append(
                                    f"sections[{si}].blocks[{bi}] (chart): "
                                    f"series[{sri}].label 缺失"
                                )
                            if not isinstance(sr.get("values"), list):
                                errors.append(
                                    f"sections[{si}].blocks[{bi}] (chart): "
                                    f"series[{sri}].values 必须是数组"
                                )
                if bt == "cross_ref" and not isinstance(block.get("text"), str):
                    errors.append(f"sections[{si}].blocks[{bi}] (cross_ref): 缺少 text 字段")
                if bt == "verification_note" and not isinstance(block.get("text"), str):
                    errors.append(f"sections[{si}].blocks[{bi}] (verification_note): 缺少 text 字段")

    # appendix
    appendix = report.get("appendix")
    if appendix is not None and not isinstance(appendix, dict):
        errors.append("appendix 必须是 JSON 对象")
    elif isinstance(appendix, dict):
        file_list = appendix.get("file_list")
        if file_list is not None:
            if not isinstance(file_list, list):
                errors.append("appendix.file_list 必须是数组")
            else:
                for fi, f in enumerate(file_list):
                    if not isinstance(f, dict):
                        errors.append(f"appendix.file_list[{fi}] 必须是 JSON 对象")
                        continue
                    for field in ("path", "type", "status"):
                        value = f.get(field)
                        if not isinstance(value, str) or not value.strip():
                            errors.append(f"appendix.file_list[{fi}].{field} 不能为空")

    return errors


# ===========================================================================
# Main
# ===========================================================================
def main():
    parser = argparse.ArgumentParser(description="Render report.json to MD/DOCX/XLSX")
    parser.add_argument("report", help="Path to report.json")
    parser.add_argument("--format", required=True, choices=["md", "docx", "xlsx"])
    parser.add_argument("--output", required=True, help="Output file path")
    parser.add_argument("--chart-dir", default="./charts", help="Directory for chart PNGs")
    args = parser.parse_args()

    with open(args.report, "r", encoding="utf-8") as fh:
        report = json.load(fh)

    # Validate before rendering
    validation_errors = validate_report(report)
    if validation_errors:
        print("ERROR: report.json 校验失败：", file=sys.stderr)
        for e in validation_errors:
            print(f"  - {e}", file=sys.stderr)
        sys.exit(1)

    setup_matplotlib()

    renderers = {
        "md": render_md,
        "docx": render_docx,
        "xlsx": render_xlsx,
    }
    renderer = renderers[args.format]
    out = renderer(report, args.chart_dir, args.output)
    print(f"Report rendered: {out}")


if __name__ == "__main__":
    main()
