#!/usr/bin/env python3
"""
summarize_folder.py — extract text from all files in one or more paths.
Usage: python summarize_folder.py <path...> --output <JSON_PATH> [--recursive] [--cache <CACHE_JSON>]
  - Directory paths: scan for supported files.
  - File paths: process directly.
"""

import argparse
import csv as csv_mod
import hashlib
import json
import sys
from datetime import datetime, timezone
from pathlib import Path

# ---------------------------------------------------------------------------
# File-type registry
# ---------------------------------------------------------------------------
TEXT_EXTS = {
    ".txt", ".md", ".json", ".xml", ".html", ".htm",
    ".py", ".js", ".ts", ".jsx", ".tsx", ".css", ".scss", ".less",
    ".yaml", ".yml", ".toml", ".ini", ".cfg", ".conf", ".env",
    ".log", ".sh", ".bat", ".ps1", ".sql", ".r", ".rb", ".go",
    ".rs", ".java", ".kt", ".swift", ".c", ".cpp", ".h", ".hpp",
    ".tex", ".rst", ".org",
}

IMAGE_EXTS = {
    ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".tif", ".tiff",
}

SUPPORTED = {
    **{ext: "text" for ext in TEXT_EXTS},
    **{ext: "image" for ext in IMAGE_EXTS},
    ".csv": "csv",
    ".pdf": "pdf",
    ".docx": "docx",
    ".xlsx": "xlsx",
    ".xls": "xls",
    ".pptx": "pptx",
}


# ---------------------------------------------------------------------------
# Dependency check (extraction only — matplotlib is render-phase only)
# ---------------------------------------------------------------------------
def check_dependencies():
    missing = []
    for mod, pkg in [
        ("pdfplumber", "pdfplumber"),
        ("docx", "python-docx"),
        ("pptx", "python-pptx"),
        ("openpyxl", "openpyxl"),
    ]:
        try:
            __import__(mod)
        except ImportError:
            missing.append(pkg)

    if missing:
        print("ERROR: Missing Python packages: " + ", ".join(missing), file=sys.stderr)
        print("Run: pip install " + " ".join(missing), file=sys.stderr)
        sys.exit(1)


# ---------------------------------------------------------------------------
# File collection
# ---------------------------------------------------------------------------
def collect_files(paths, recursive):
    """Return sorted list of (Path, root_dir) tuples for supported files."""
    files = []
    for p in paths:
        p = Path(p).resolve()
        if p.is_dir():
            iterator = p.rglob("*") if recursive else p.glob("*")
            for fp in iterator:
                if fp.is_file() and fp.suffix.lower() in SUPPORTED:
                    files.append((fp, p))
        elif p.is_file():
            if p.suffix.lower() in SUPPORTED:
                files.append((p, p.parent))
            else:
                print(f"[SKIP] unsupported: {p}", file=sys.stderr)
        else:
            print(f"ERROR: path not found: {p}", file=sys.stderr)
    return sorted(set(files), key=lambda x: str(x[0]).lower())


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _file_sha256(path):
    """Compute SHA256 hex digest of a file."""
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def _read_text_raw(path):
    for enc in ("utf-8", "gbk", "latin-1"):
        try:
            with open(path, "r", encoding=enc) as fh:
                return fh.read()
        except (UnicodeDecodeError, UnicodeError):
            continue
    with open(path, "r", encoding="utf-8", errors="replace") as fh:
        return fh.read()


def _check_encrypted(exc, label):
    """If an exception indicates an encrypted/password-protected file,
    raise a friendlier ValueError. Otherwise re-raise the original."""
    msg = str(exc).lower()
    if any(kw in msg for kw in ("password", "encrypted", "encryption", "not a zip file")):
        raise ValueError(f"文件已加密，无法提取（需要密码）: {label}") from exc
    raise


# ---------------------------------------------------------------------------
# Handlers — each returns (content: str, metadata: dict)
# ---------------------------------------------------------------------------
def handle_text(path):
    return _read_text_raw(path), {"encoding": "utf-8"}


def handle_csv(path):
    rows = []
    with open(path, "r", encoding="utf-8", errors="replace") as fh:
        reader = csv_mod.reader(fh)
        for row in reader:
            rows.append("\t".join(row))
    content = "\n".join(rows) if rows else ""
    return content, {"rows": len(rows), "format": "csv"}


def handle_pdf(path):
    import pdfplumber

    try:
        pages_text = []
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    pages_text.append(t)
        return "\n\n".join(pages_text), {"pages": len(pdf.pages)}
    except Exception as e:
        _check_encrypted(e, str(path))
        raise


def handle_docx(path):
    """Extract DOCX content preserving paragraph/table document order."""
    import docx
    from docx.oxml.ns import qn

    doc = docx.Document(path)

    body = doc.element.body
    parts = []
    para_count = 0
    table_count = 0

    for child in body:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag == "p":
            texts = []
            for t_node in child.iter(qn("w:t")):
                if t_node.text:
                    texts.append(t_node.text)
            text = "".join(texts)
            if text.strip():
                parts.append(text)
                para_count += 1
        elif tag == "tbl":
            rows = []
            for tr in child.iter(qn("w:tr")):
                cells = []
                for tc in tr.iter(qn("w:tc")):
                    cell_texts = []
                    for p in tc.iter(qn("w:p")):
                        for t_node in p.iter(qn("w:t")):
                            if t_node.text:
                                cell_texts.append(t_node.text)
                    cells.append("".join(cell_texts))
                if cells:
                    rows.append(" | ".join(cells))
            if rows:
                parts.append("[TABLE]\n" + "\n".join(rows) + "\n[/TABLE]")
                table_count += 1

    content = "\n\n".join(parts)
    return content, {"paragraphs": para_count, "tables": table_count}


def handle_xlsx(path):
    import openpyxl

    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        _check_encrypted(e, str(path))
        raise

    sheets_data = {}
    sheet_count = 0
    skipped_empty = 0

    for name in wb.sheetnames:
        ws = wb[name]
        rows = []
        for row in ws.iter_rows(values_only=True):
            rows.append("\t".join(str(c) if c is not None else "" for c in row))
        if rows:
            sheets_data[name] = "\n".join(rows)
            sheet_count += 1
        else:
            skipped_empty += 1

    wb.close()

    content = "\n\n".join(
        f"--- Sheet: {name} ---\n{data}" for name, data in sheets_data.items()
    )

    meta = {"sheets": list(sheets_data.keys()), "data_sheets": sheet_count}
    if skipped_empty:
        meta["skipped_empty"] = skipped_empty
    return content, meta


def handle_pptx(path):
    """Extract PPTX text from shapes, tables, and group shapes recursively."""
    import pptx
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    def extract_shape_text(shape):
        texts = []
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                if para.text.strip():
                    texts.append(para.text)
        if shape.has_table:
            table = shape.table
            for row in table.rows:
                cells = [cell.text for cell in row.cells]
                texts.append(" | ".join(cells))
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for child in shape.shapes:
                texts.extend(extract_shape_text(child))
        return texts

    prs = pptx.Presentation(path)
    slides_text = []
    for i, slide in enumerate(prs.slides, 1):
        texts = []
        for shape in slide.shapes:
            texts.extend(extract_shape_text(shape))
        if texts:
            slides_text.append(f"--- Slide {i} ---\n" + "\n".join(texts))
    return "\n\n".join(slides_text), {"slides": len(prs.slides)}


def handle_image(fp):
    """图片不做 OCR —— 由 AI 用 Read 工具直接查看，脚本只登记路径与体积。"""
    return "", {"size_bytes": fp.stat().st_size}


# ---------------------------------------------------------------------------
# Dispatch
# ---------------------------------------------------------------------------
HANDLERS = {
    "text": handle_text,
    "csv": handle_csv,
    "pdf": handle_pdf,
    "docx": handle_docx,
    "xlsx": handle_xlsx,
    "xls": handle_xlsx,
    "pptx": handle_pptx,
    "image": handle_image,
}


def safe_process(fp, handler):
    ext = fp.suffix.lower()
    try:
        content, meta = handler(fp)
        return {
            "relative_path": str(fp),
            "type": SUPPORTED.get(ext, "unknown"),
            "status": "ok",
            "content": content,
            "content_length": len(content),
            "error": None,
            "metadata": meta,
        }
    except Exception as exc:
        return {
            "relative_path": str(fp),
            "type": SUPPORTED.get(ext, "unknown"),
            "status": "error",
            "content": "",
            "content_length": 0,
            "error": str(exc),
            "metadata": {},
        }


# ---------------------------------------------------------------------------
# Cache helpers
# ---------------------------------------------------------------------------
def load_cache(cache_path):
    if cache_path and Path(cache_path).exists():
        try:
            with open(cache_path, "r", encoding="utf-8") as fh:
                return json.load(fh)
        except (json.JSONDecodeError, OSError):
            return {}
    return {}


def save_cache(cache_path, cache_data):
    if cache_path:
        Path(cache_path).parent.mkdir(parents=True, exist_ok=True)
        with open(cache_path, "w", encoding="utf-8") as fh:
            json.dump(cache_data, fh, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description="Extract text from files in directories or specified files"
    )
    parser.add_argument(
        "paths", nargs="+", help="One or more paths (directories and/or files)"
    )
    parser.add_argument("--output", required=True, help="Output JSON file path")
    parser.add_argument(
        "--recursive", action="store_true", help="Recurse into subdirectories"
    )
    parser.add_argument(
        "--cache", default=None, help="Optional cache JSON path (file-hash to content)"
    )
    parser.add_argument(
        "--no-cache",
        action="store_true",
        help="Skip cache read (still writes cache on success)",
    )
    args = parser.parse_args()

    check_dependencies()

    files = collect_files(args.paths, args.recursive)

    if not files:
        result = {
            "sources": [str(Path(p).resolve()) for p in args.paths],
            "processed_at": datetime.now(timezone.utc).isoformat(),
            "total": 0,
            "ok": 0,
            "errors": 0,
            "files": [],
        }
        with open(args.output, "w", encoding="utf-8") as fh:
            json.dump(result, fh, ensure_ascii=False, indent=2)
        print("0 files found. Empty output written to", args.output)
        return

    cache = {} if args.no_cache else load_cache(args.cache)
    cache_updated = False

    total = len(files)
    results = []
    cache_hits = 0

    for idx, (fp, root) in enumerate(files):
        ext = fp.suffix.lower()
        cat = SUPPORTED.get(ext)
        handler = HANDLERS.get(cat)

        if handler is None:
            results.append({
                "relative_path": str(fp),
                "type": ext,
                "status": "skipped",
                "content": "",
                "content_length": 0,
                "error": "unsupported format",
                "metadata": {},
            })
            continue

        # Cache check
        file_hash = _file_sha256(fp)
        cache_key = str(fp.resolve())
        cached_entry = cache.get(cache_key)

        if cached_entry and cached_entry.get("hash") == file_hash:
            r = dict(cached_entry["result"])
            try:
                r["relative_path"] = str(fp.relative_to(root))
            except ValueError:
                r["relative_path"] = str(fp)
            results.append(r)
            cache_hits += 1
            continue

        # Progress feedback (only when actually extracting)
        print(f"[{idx+1}/{total}] extracting {fp.name}...", file=sys.stderr)

        r = safe_process(fp, handler)
        try:
            r["relative_path"] = str(fp.relative_to(root))
        except ValueError:
            r["relative_path"] = str(fp)
        results.append(r)

        # Store in cache
        if r["status"] == "ok":
            cache[cache_key] = {
                "hash": file_hash,
                "result": {
                    "relative_path": str(fp.resolve()),
                    "type": r["type"],
                    "status": r["status"],
                    "content": r["content"],
                    "content_length": r["content_length"],
                    "error": r["error"],
                    "metadata": r["metadata"],
                },
            }
            cache_updated = True

    ok_count = sum(1 for r in results if r["status"] == "ok")
    err_count = sum(1 for r in results if r["status"] == "error")

    result = {
        "sources": [str(Path(p).resolve()) for p in args.paths],
        "processed_at": datetime.now(timezone.utc).isoformat(),
        "total": len(results),
        "ok": ok_count,
        "errors": err_count,
        "files": results,
    }

    with open(args.output, "w", encoding="utf-8") as fh:
        json.dump(result, fh, ensure_ascii=False, indent=2)

    if cache_hits:
        print(
            f"Processed {len(results)} files: {ok_count} OK, {err_count} errors  "
            f"({cache_hits} from cache)"
        )
    else:
        print(f"Processed {len(results)} files: {ok_count} OK, {err_count} errors")

    if err_count:
        print("Errors:")
        for r in results:
            if r["status"] == "error":
                print(f"  - {r['relative_path']}: {r['error']}")

    if cache_updated and args.cache:
        save_cache(args.cache, cache)
        print(f"Cache updated: {args.cache}")

    print(f"Output written to: {args.output}")


if __name__ == "__main__":
    main()
