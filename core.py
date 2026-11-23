# -*- coding: utf-8 -*-
"""PDF 页面级文字水印去除 + PDF 转 Word（GUI 与 CLI 共用的核心逻辑）。

去除策略（按优先级自动选择）：
1. Acrobat 标准水印：每页内容流里形如
       /Artifact <</Subtype /Watermark ...>> BDC ... /Fm0 Do ... EMC
   的绘制指令，其中 Fm0 是带 /PieceInfo<</ADBE_CompoundType<</Private/Watermark>>>
   标记的 Form XObject。处理方式：仅删除指向该表单的 `/Fm0 Do` 指令，
   其余指令原样保留 —— 无损，不影响正文文字、图片与矢量元素。
   注意：不可整块删除 BDC..EMC，块内的平衡 Q 可能负责闭合页面外层的 q，
   整块删除会破坏图形状态平衡，导致文本提取异常。
2. 普通文字水印：对指定/检测到的水印文本做 redaction 擦除，
   images=PDF_REDACT_IMAGE_NONE 保证只擦文字、不碰图片。

限制：烧在图片像素内部的水印属于图片本身，无法无损去除，将原样保留。
"""
from __future__ import annotations

import argparse
import re
import sys
import time
from collections import defaultdict
from pathlib import Path

try:
    import pymupdf
except ImportError:  # 兼容 PyMuPDF<1.24 的旧包名
    import fitz as pymupdf  # type: ignore

from pdf2docx import Converter

__all__ = ["find_repeated_texts", "clean_pdf", "convert_single", "batch_convert"]


# ---------------------------------------------------------------- 水印检测

def find_repeated_texts(pdf_path, min_ratio=0.5, max_len=60, top=12):
    """统计“同一文本出现在 >= min_ratio 比例页面”的短文本行，作为水印候选。

    返回按出现页数降序的 [{"text", "pages", "ratio"}, ...]。
    """
    doc = pymupdf.open(pdf_path)
    try:
        n = doc.page_count
        pages_by_text = defaultdict(set)
        for pno in range(n):
            page = doc[pno]
            for block in page.get_text("dict").get("blocks", []):
                if block.get("type") != 0:
                    continue
                for line in block.get("lines", []):
                    text = re.sub(
                        r"\s+", "",
                        "".join(sp.get("text", "") for sp in line.get("spans", [])),
                    )
                    if text:
                        pages_by_text[text].add(pno)
        cands = [
            {"text": t, "pages": len(ps), "ratio": len(ps) / n}
            for t, ps in pages_by_text.items()
            if n and len(ps) / n >= min_ratio and len(t) <= max_len
        ]
        cands.sort(key=lambda c: (-c["pages"], -len(c["text"])))
        return cands[:top]
    finally:
        doc.close()


# ---------------------------------------------------------------- 水印去除

def _watermark_form_names(doc, page):
    """识别页面上带 Acrobat Private/Watermark 标记的水印 Form XObject 名称。"""
    names = []
    for xref, name, _invoker, _bbox in page.get_xobjects():
        try:
            obj = doc.xref_object(xref, compressed=True)
        except Exception:
            continue
        if "ADBE_CompoundType" in obj and "/Watermark" in obj:
            names.append(name.encode("latin-1"))
    return names


def _remove_form_do_ops(data: bytes, names) -> tuple:
    new = data
    removed = 0
    for n in names:
        new, k = re.subn(rb"/" + re.escape(n) + rb"\s+Do", b"", new)
        removed += k
    return new, removed


def _redact_text(doc, text: str) -> int:
    """回退方案：擦除普通文字水印（不碰图片）。"""
    removed = 0
    for page in doc:
        rects = page.search_for(text)
        if not rects:
            continue
        for r in rects:
            page.add_redact_annot(r)
        page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE)
        removed += len(rects)
    return removed


def clean_pdf(src, dst, watermark_text=None):
    """去除 PDF 水印并另存为 dst，返回 (移除数, 是否残留水印文本, 页数)。"""
    src, dst = Path(src), Path(dst)
    doc = pymupdf.open(src)
    try:
        removed = 0
        for page in doc:
            names = _watermark_form_names(doc, page)
            if not names:
                continue
            for xref in page.get_contents():
                data = doc.xref_stream(xref)
                new, k = _remove_form_do_ops(data, names)
                if k:
                    doc.update_stream(xref, new)
                    removed += k
        if removed == 0 and watermark_text:
            removed = _redact_text(doc, watermark_text)
        page_count = doc.page_count
        doc.save(dst, garbage=4, deflate=True)
    finally:
        doc.close()

    residue = False
    if watermark_text:
        with pymupdf.open(dst) as check:
            residue = any(watermark_text in p.get_text() for p in check)
    return removed, residue, page_count


# ---------------------------------------------------------------- 转换

def convert_single(pdf_path, out_docx, watermark_text=None) -> dict:
    """转换单个 PDF -> docx，返回结果 dict（ok/pages/removed/.../error）。"""
    src = Path(pdf_path)
    out_docx = Path(out_docx)
    clean = out_docx.with_name(out_docx.stem + "_clean.pdf")
    t0 = time.time()
    result = {"file": src.name, "ok": False}
    try:
        removed, residue, pages = clean_pdf(src, clean, watermark_text)
        result.update(pages=pages, removed=removed, watermark_gone=not residue)
        cv = Converter(str(clean))
        cv.convert(str(out_docx))
        cv.close()
        result["docx_kb"] = round(out_docx.stat().st_size / 1024, 1)
        result["ok"] = True
    except Exception as e:  # noqa: BLE001
        result["error"] = f"{type(e).__name__}: {e}"
    finally:
        clean.unlink(missing_ok=True)
    result["seconds"] = round(time.time() - t0, 1)
    return result


def batch_convert(pdfs, out_dir, watermark_text=None, progress=None) -> list:
    """批量转换。progress(done, total, current) 会在每个文件前后各回调一次：
    转换前 current 是文件名 str，转换后 current 是结果 dict。"""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    results = []
    total = len(pdfs)
    for i, pdf in enumerate(pdfs, 1):
        if progress:
            progress(i - 1, total, Path(pdf).name)
        r = convert_single(pdf, out_dir / (Path(pdf).stem + ".docx"), watermark_text)
        results.append(r)
        if progress:
            progress(i, total, r)
    return results


# ---------------------------------------------------------------- CLI

def main(argv=None) -> int:
    if sys.platform == "win32":
        for stream in (sys.stdout, sys.stderr):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass

    parser = argparse.ArgumentParser(
        description="PDF 去文字水印并批量转 Word（.docx）")
    parser.add_argument("input", nargs="?", default=None,
                        help="PDF 文件或包含 PDF 的目录")
    parser.add_argument("-o", "--out", default=None,
                        help="输出目录（默认为 input 同级的 word 目录）")
    parser.add_argument("-w", "--watermark", default=None,
                        help="水印文本；缺省时自动检测（也可仅移除 Acrobat 水印表单）")
    parser.add_argument("--detect", action="store_true",
                        help="仅检测并打印水印候选，不转换")
    parser.add_argument("--min-ratio", type=float, default=0.5,
                        help="检测阈值：文本出现页数占比（默认 0.5）")
    args = parser.parse_args(argv)

    if args.detect:
        if not args.input:
            parser.error("--detect 需要提供一个 PDF 文件路径")
        for c in find_repeated_texts(args.input, args.min_ratio):
            print(f"{c['ratio']:>5.0%}  {c['pages']:>4} 页  {c['text']}")
        return 0

    if not args.input:
        parser.print_help()
        return 1

    src = Path(args.input)
    if src.is_dir():
        pdfs = sorted(src.glob("*.pdf"))
        out_dir = Path(args.out) if args.out else src / "word"
    else:
        pdfs = [src]
        out_dir = Path(args.out) if args.out else src.parent / "word"
    if not pdfs:
        print("未找到 PDF 文件")
        return 1

    watermark = args.watermark
    if watermark is None:
        cands = [c for c in find_repeated_texts(pdfs[0], args.min_ratio)
                 if len(c["text"]) >= 4]
        if cands:
            watermark = cands[0]["text"]
            print(f"自动检测水印: {watermark} ({cands[0]['pages']} 页, {cands[0]['ratio']:.0%})")
        else:
            print("未检测到重复文本水印，将仅移除 Acrobat 水印表单")

    results = batch_convert(
        pdfs, out_dir, watermark,
        progress=lambda done, total, cur: (
            print(f"[{done + 1}/{total}] 转换中: {Path(cur).name}")
            if isinstance(cur, str) else None),
    )
    ok = sum(1 for r in results if r["ok"])
    for r in results:
        flag = "OK  " if r["ok"] else "FAIL"
        extra = "" if r["ok"] else f"  -> {r.get('error', '')}"
        print(f"[{flag}] {r['file']}  pages={r.get('pages')} "
              f"removed={r.get('removed')} watermark_gone={r.get('watermark_gone')}{extra}")
    print(f"完成: {ok}/{len(results)} -> {out_dir}")
    return 0 if ok == len(results) else 2


if __name__ == "__main__":
    raise SystemExit(main())
