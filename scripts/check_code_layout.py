# -*- coding: utf-8 -*-
"""代码块排版质量门禁（exam-word-skill，跨课程通用）。

为什么需要
----------
代码类试卷的排版问题（散落小字、无边界、字号不一、缩进错乱、黑框、行号）
靠"人眼逐页看"不可靠、不可迁移。本脚本把``references/exam-authoring-pitfalls.md``
§9.3 的视觉规范变成**可执行的量化门禁**：任何课程、任何语言的代码卷，只要用
``build_code_exam.py``（或等效渲染器）生成，都能一键验收。

检查项（全部可从 DOCX 客观读取，不依赖试卷内容）
------------------------------------------------
视觉边界由 ``--style`` 决定，默认 ``plain``（教师定稿：卷面干净、无底纹无竖线）：
1. ``no_shade`` / ``no_border``  plain 模式：代码块无底纹、无左侧竖线；
   ``has_shade`` / ``has_border``  shaded 模式：有浅灰底纹 + 左灰竖线（且色值在白名单）；
2. ``monospace``      代码 run 的 ascii/hAnsi/cs 字体属于等宽白名单；
3. ``size_ok``        代码字号在 [min-size, max-size] 区间；
4. ``indent_ok``      代码块左缩进 >= 正文左缩进（建立"区块内收"感）；
5. ``no_line_number`` 无 ``w:lnNumType``（行号）；
6. ``no_table_grid``  代码所在表格非 ``Table Grid``（黑框）；
7. ``ppr_order_ok``   pPr 子元素顺序合法（``w:pBdr``/``w:shd`` 在 ``w:spacing`` 前），
                      否则 Word 会提示"文档需要修复"。

可迁移设计
----------
* 字体白名单、底纹白名单、字号区间、代码特征词**全部由命令行参数给出**，
  不写死某门课/某种语言；默认覆盖 C/C++/Java/Python 常见等宽字体与特征。
* 判据只看 OOXML 客观属性，不比对具体题目文本，换课程零改动。
* 既可直接跑，也可 ``from check_code_layout import audit`` 被组卷脚本调用。

用法
----
    python check_code_layout.py paper.docx
    python check_code_layout.py paper.docx --json report.json
    python check_code_layout.py paper.docx --style shaded   # 校验"底纹+左竖线"样式
    python check_code_layout.py paper.docx --font Consolas --font Menlo \
        --min-size 9.5 --max-size 11 --shade F5F5F5 --body-indent 283

退出码：0 = 全部通过；1 = 存在未通过项；2 = 未识别到任何代码块（可能渲染器用错）。
"""
import argparse
import json
import sys
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

# 默认等宽字体白名单（跨语言/跨平台）
DEFAULT_MONO = {"Consolas", "Courier New", "Menlo", "Monaco", "Source Code Pro",
                "DejaVu Sans Mono", "Liberation Mono", "JetBrains Mono", "宋体"}
# 默认代码特征（命中任一即认为是代码行；可用 --hint 覆盖扩充）
DEFAULT_HINTS = ("#include", "#define", "printf", "scanf", "return ", ";", "{", "}",
                 "int ", "char ", "float ", "void ", "for (", "if (", "while (",
                 "class ", "public ", "def ", "import ", "cout", "cin", "println")


def _iter_body_paras(doc):
    """按文档顺序遍历全部正文段落（含表格单元格内）。"""
    for el in doc.element.body.iter(qn("w:p")):
        yield Paragraph(el, doc)


def _all_text(el):
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def _has_break(p):
    return len(p._p.findall(".//" + qn("w:br"))) >= 1


def _code_runs(p):
    return [r for r in p.runs if r.text]


def _para_text(p):
    """优先用 python-docx 的 .text（会把 w:br 转成 \\n，w:tab 转成 \\t）。"""
    return p.text if p.text else _all_text(p._p)


def is_code_paragraph(p, hints):
    """含软换行且命中代码特征 -> 认定为代码块段落。"""
    if not _has_break(p):
        return False
    text = _para_text(p)
    # 至少两行才算代码块
    if text.count("\n") < 1:
        return False
    return any(h in text for h in hints)


def _rfonts_of(run):
    rPr = run._element.find(qn("w:rPr"))
    if rPr is None:
        return {}
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        return {}
    return {k: rf.get(qn("w:" + k)) for k in ("ascii", "hAnsi", "cs", "eastAsia")}


def _shade_fill(p):
    pPr = p._p.find(qn("w:pPr"))
    if pPr is None:
        return None
    shd = pPr.find(qn("w:shd"))
    return None if shd is None else shd.get(qn("w:fill"))


def _left_border(p):
    pPr = p._p.find(qn("w:pPr"))
    if pPr is None:
        return None
    bdr = pPr.find(qn("w:pBdr"))
    if bdr is None:
        return None
    return bdr.find(qn("w:left"))


def _ppr_order_ok(p):
    """w:pBdr/w:shd 必须排在 w:spacing/w:ind 之前。"""
    pPr = p._p.find(qn("w:pPr"))
    if pPr is None:
        return True
    tags = [c.tag for c in pPr]
    def pos(tag):
        return tags.index(qn(tag)) if qn(tag) in tags else -1
    sp = min([x for x in (pos("w:spacing"), pos("w:ind")) if x >= 0] or [10**6])
    for early in ("w:pBdr", "w:shd"):
        i = pos(early)
        if i >= 0 and i > sp:
            return False
    return True


def _in_grid_table(p):
    """段落是否位于 Table Grid（黑框）表格内。"""
    el = p._p.getparent()
    while el is not None and el.tag != qn("w:tbl"):
        el = el.getparent()
    if el is None:
        return False
    tblPr = el.find(qn("w:tblPr"))
    if tblPr is None:
        return False
    style = tblPr.find(qn("w:tblStyle"))
    return style is not None and "Grid" in (style.get(qn("w:val")) or "")


def audit(docx_path, *, mono_fonts=None, shades=None, min_size=9.5, max_size=11.0,
          body_indent=0, hints=None, style="plain", require_tabs=False):
    """审计一个 DOCX 的代码块排版，返回结构化报告 dict。"""
    mono = set(mono_fonts) if mono_fonts else DEFAULT_MONO
    shade_set = set(shades) if shades else {s.upper() for s in DEFAULT_SHADES}
    hs = tuple(hints) if hints else DEFAULT_HINTS
    doc = Document(str(docx_path))
    paras = list(_iter_body_paras(doc))

    codes = [p for p in paras if is_code_paragraph(p, hs)]
    if not codes:
        return {"file": Path(docx_path).name, "code_paras": 0,
                "verdict": "NO_CODE", "checks": {}, "details": [],
                "message": "未识别到代码块，请确认渲染器或 --hint 参数"}

    def pt(run):
        return run.font.size.pt if run.font.size else None

    def left_twips(p):
        v = p.paragraph_format.left_indent
        return None if v is None else v.twips

    shade_key = "no_shade" if style == "plain" else "has_shade"
    border_key = "no_border" if style == "plain" else "has_border"
    check_keys = [shade_key, border_key, "monospace", "size_ok",
                  "indent_ok", "no_line_number", "no_table_grid", "ppr_order_ok"]
    if require_tabs:
        check_keys.append("indent_by_tabs")
    checks = {k: 0 for k in check_keys}
    details = []
    for i, p in enumerate(codes):
        runs = _code_runs(p)
        fill = _shade_fill(p)
        border = _left_border(p)
        fonts = [_rfonts_of(r) for r in runs]
        sizes = [pt(r) for r in runs]
        ln = p._p.find(".//" + qn("w:lnNumType"))
        li = left_twips(p)
        # tab 缩进判定：段落有固定 tab stops 且行首缩进用了 w:tab 元素
        pPr = p._p.find(qn("w:pPr"))
        tabs_el = pPr.find(qn("w:tabs")) if pPr is not None else None
        has_tab_stops = tabs_el is not None and len(tabs_el.findall(qn("w:tab"))) > 0
        has_tab_runs = len(p._p.findall(".//" + qn("w:tab"))) > 0
        row = {
            "idx": i,
            "lines": _all_text(p._p).count("\n") + 1,
            "shade": fill,
            "border": border is not None,
            "fonts": sorted({f.get("ascii") for f in fonts if f.get("ascii")}),
            "sizes": sorted({s for s in sizes if s is not None}),
            "left_twips": li,
        }
        c = {
            shade_key: ((not fill) or fill.upper() in ("AUTO", "FFFFFF"))
                       if style == "plain" else (bool(fill) and fill.upper() in shade_set),
            border_key: (border is None) if style == "plain" else (border is not None),
            "monospace": bool(fonts) and all(
                any(f.get(a) in mono for a in ("ascii", "hAnsi", "cs")) for f in fonts),
            "size_ok": bool(sizes) and all(
                s is not None and min_size - 0.01 <= s <= max_size + 0.01 for s in sizes),
            "indent_ok": (li is not None and li >= body_indent),
            "no_line_number": ln is None,
            "no_table_grid": not _in_grid_table(p),
            "ppr_order_ok": _ppr_order_ok(p),
        }
        if require_tabs:
            c["indent_by_tabs"] = has_tab_stops and has_tab_runs
        for k, v in c.items():
            checks[k] += int(bool(v))
        row["checks"] = c
        details.append(row)

    n = len(codes)
    passed = all(v == n for v in checks.values())
    return {
        "file": Path(docx_path).name,
        "code_paras": n,
        "checks": {k: {"pass": v, "total": n, "ok": v == n} for k, v in checks.items()},
        "details": details,
        "verdict": "PASS" if passed else "FAIL",
    }


DEFAULT_SHADES = ["F5F5F5", "F2F2F2", "FAFAFA", "F7F7F7", "EFEFEF", "F8F8F8"]


def format_report(rep):
    lines = [f"文件：{rep['file']}", f"代码块段落：{rep['code_paras']}", ""]
    if rep["verdict"] == "NO_CODE":
        lines.append("结论：无代码块（" + rep.get("message", "") + "）")
        return "\n".join(lines)
    for k, v in rep["checks"].items():
        flag = "✔" if v["ok"] else "✘"
        lines.append(f"  {flag} {k:<16} {v['pass']}/{v['total']}")
    bad = [d["idx"] for d in rep["details"]
           if not all(d["checks"].values())]
    if bad:
        lines.append("")
        lines.append("未通过的代码块序号：" + ", ".join(map(str, bad)))
    lines.append("")
    lines.append("结论：" + rep["verdict"])
    return "\n".join(lines)


def main(argv=None):
    ap = argparse.ArgumentParser(description="代码块排版质量门禁")
    ap.add_argument("docx", help="待检查的 DOCX")
    ap.add_argument("--json", dest="json_out", help="把报告写成 JSON")
    ap.add_argument("--font", dest="fonts", action="append", help="等宽字体白名单（可多次）")
    ap.add_argument("--shade", dest="shades", action="append", help="底纹色白名单（可多次）")
    ap.add_argument("--min-size", type=float, default=9.5)
    ap.add_argument("--max-size", type=float, default=11.0)
    ap.add_argument("--body-indent", type=int, default=0, help="正文左缩进 twips")
    ap.add_argument("--hint", dest="hints", action="append", help="代码特征词（可多次）")
    ap.add_argument("--style", choices=["plain", "shaded"], default="plain",
                    help="plain=无底纹无竖线（默认）；shaded=浅灰底纹+左灰竖线")
    ap.add_argument("--require-tabs", action="store_true",
                    help="v6：强制代码缩进用 tab stops（字体回退也不错位）；老卷勿开")
    args = ap.parse_args(argv)

    rep = audit(args.docx, mono_fonts=args.fonts, shades=args.shades,
                min_size=args.min_size, max_size=args.max_size,
                body_indent=args.body_indent, hints=args.hints, style=args.style,
                require_tabs=args.require_tabs)
    print(format_report(rep))
    if args.json_out:
        Path(args.json_out).write_text(json.dumps(rep, ensure_ascii=False, indent=2),
                                       encoding="utf-8")
    if rep["verdict"] == "NO_CODE":
        return 2
    return 0 if rep["verdict"] == "PASS" else 1


if __name__ == "__main__":
    sys.exit(main())
