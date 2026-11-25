# -*- coding: utf-8 -*-
"""PDF 逐页门禁（exam-word-skill，跨课程通用）：空白页 / 孤页 / 页数。

使用前需先把 DOCX 转成 PDF（LibreOffice headless）。本脚本只读 PDF。

检查项
------
1. 无空白页：每页有非空白文本（正文/页码都算）；
2. 无孤页：末页文本字数 >= min_last_page_chars，或页数==1；
3. 页数在 [min_pages, max_pages] 内（可选，0 表示不限）。

依赖 PyMuPDF（import pymupdf 或 fitz 均可）。退出码 0=通过 1=失败 2=缺依赖/读不了。

用法
----
    python check_pdf_pages.py paper.pdf
    python check_pdf_pages.py paper.pdf --min-last-page-chars 80
    python check_pdf_pages.py paper.pdf --min-pages 6 --max-pages 12
"""
import argparse
import sys
from pathlib import Path


def _open_pdf(path):
    try:
        import pymupdf
    except ImportError:
        try:
            import fitz as pymupdf
        except ImportError:
            return None, "缺依赖：需要 pymupdf（pip install pymupdf）"
    try:
        return pymupdf.open(str(path)), None
    except Exception as e:
        return None, f"无法读取 PDF：{e}"


def audit(path, *, min_last_page_chars=80, min_pages=0, max_pages=0):
    doc, err = _open_pdf(path)
    if err:
        return {"verdict": "ERROR", "message": err, "checks": []}
    checks = []
    n = len(doc)
    texts = [(p.get_text() or "").strip() for p in doc]

    blanks = [i + 1 for i, t in enumerate(texts) if not t]
    checks.append({"name": "无空白页", "ok": not blanks,
                   "detail": f"空白页={blanks}" if blanks else "0 个空白页"})

    last_chars = len(texts[-1]) if texts else 0
    lone = n > 1 and last_chars < min_last_page_chars
    checks.append({"name": "无孤页", "ok": not lone,
                   "detail": f"末页 {last_chars} 字 < {min_last_page_chars}" if lone
                             else f"末页 {last_chars} 字（或仅 1 页）"})

    if min_pages or max_pages:
        lo = min_pages or 0
        hi = max_pages or 10 ** 6
        ok = lo <= n <= hi
        checks.append({"name": "页数范围", "ok": ok,
                       "detail": f"{n} 页，允许 [{lo}, {hi}]"})

    verdict = "PASS" if all(c["ok"] for c in checks) else "FAIL"
    return {"verdict": verdict, "pages": n, "checks": checks,
            "file": Path(path).name}


def main(argv=None):
    ap = argparse.ArgumentParser(description="PDF 逐页门禁")
    ap.add_argument("pdf")
    ap.add_argument("--min-last-page-chars", type=int, default=80)
    ap.add_argument("--min-pages", type=int, default=0)
    ap.add_argument("--max-pages", type=int, default=0)
    args = ap.parse_args(argv)
    rep = audit(args.pdf, min_last_page_chars=args.min_last_page_chars,
                min_pages=args.min_pages, max_pages=args.max_pages)
    print("文件：" + Path(args.pdf).name)
    for c in rep["checks"]:
        print(f"  {'✔' if c['ok'] else '✘'} {c['name']}: {c['detail']}")
    print("结论：" + rep["verdict"])
    if rep["verdict"] == "ERROR":
        print(rep.get("message", ""))
        return 2
    return 0 if rep["verdict"] == "PASS" else 1


if __name__ == "__main__":
    sys.exit(main())
