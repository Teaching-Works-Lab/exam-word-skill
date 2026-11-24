# -*- coding: utf-8 -*-
"""代码类试卷专用渲染器（exam-word skill）。

何时用
------
试卷里含有「程序阅读 / 查错与测试 / 代码填空」等**多行代码块**时，用本脚本，
不要用 ``build_from_spec.py``。原因见
``references/exam-authoring-pitfalls.md`` §9.5：

* 逐行段落 + ``keep_with_next`` 长链在 LibreOffice 会失效，代码被劈页；
* ``keep_with_next`` 对表格无效，代码块会与其后的选项分家；
* 「整题推页」会在页尾留下大片空白。

本脚本的 v4 版面结构
--------------------
每道大题一张无边框表格，**一道题一行**，行设 ``w:cantSplit``：

* 行内不跨页 -> 题干、代码、选项、小问永远在同一页；
* 行与行之间允许跨页 -> 不会因为某题放不下就在页尾留大片空白。

表格 ``cellMar=0``、``tblInd=0``，单元格内段落缩进与正文一致：
题干 283（悬挂）、代码/小问 425、选项 567。文档不能以表格结尾，
末尾补一个 1pt 极小空段。

spec 契约
---------
结构沿用 ``build_from_spec.py``：``{"metadata": {...}, "score_table": [...],
"pages": [{"blocks": [...]}]}``，并在 blocks 里额外支持：

* ``{"type": "code", "lines": ["...", "..."]}`` —— 代码块，逐行软换行；
* ``{"type": "option-row", "items": ["A. ...", "B. ..."]}`` —— 两/四选项一行，
  按 tab 位排版，避免长选项折行。

其余 ``heading`` / ``question`` / ``text`` / ``spacer`` 语义不变。
``build_from_spec.py`` 的 ``_validate_spec`` 不认识 ``code``，所以本脚本
自行渲染，但**复用它的模板部件函数**（``_clear_body`` / ``_add_header`` /
``_populate_score_table`` / ``_restore_parts``），保证 7 个保留部件
byte-identical、``verify_docx.py`` 仍可通过。

用法
----
    python scripts/build_code_exam.py spec.json out.docx [template.docx]

未传 template 时使用 ``assets/reference.docx``。
生成后仍须执行 ``verify_docx.py`` + LibreOffice 渲染 PDF 逐页复核（§5 门禁）。
"""
import sys
from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, Twips

SKILL = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SKILL / "scripts"))
import build_from_spec as bfs  # noqa: E402

DEFAULT_TEMPLATE = SKILL / "assets/reference.docx"

# 版面常量（twips / pt）
Q_LEFT = 283          # 题干正文缩进（悬挂：题号顶格）
BODY_LEFT = 425       # 小问 / 代码基准缩进
OPT_LEFT = 567        # 选项再内收
CODE_SIZE = 9.5       # 代码字号
CODE_LINE = 12.5      # 代码固定行距
TEXT_LINE = 20        # 正文行距
OPT_SIZE = 11         # 选项字号
UNIT_GAP = 6          # 题与题之间的间距


def _add_keep_lines(p):
    pPr = p._p.get_or_add_pPr()
    if pPr.find(qn("w:keepLines")) is None:
        pPr.append(OxmlElement("w:keepLines"))


def new_para(doc, text="", *, left=0, first=0, before=0, after=0, line=TEXT_LINE,
             bold=False, size=None, font=None, keep=False):
    p = doc.add_paragraph(style="Normal")
    pf = p.paragraph_format
    pf.left_indent = Twips(left)
    pf.first_line_indent = Twips(first)
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = Pt(line)
    pf.keep_with_next = keep
    _add_keep_lines(p)
    if text:
        run = p.add_run(text)
        run.bold = bold
        if size:
            run.font.size = Pt(size)
        if font:
            run.font.name = font
    return p


def _cell_para(cell, is_first, text="", *, left=0, first=0, before=0, after=0,
               line=TEXT_LINE, bold=False):
    p = cell.paragraphs[0] if is_first else cell.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Twips(left)
    pf.first_line_indent = Twips(first)
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = Pt(line)
    _add_keep_lines(p)
    if text:
        run = p.add_run(text)
        run.bold = bold
    return p


def _set_code_font(run, size):
    """代码字体：ascii/hAnsi/cs 用 Consolas，eastAsia 用宋体。"""
    run.font.size = Pt(size)
    run.font.name = "Consolas"
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs"):
        rFonts.set(qn(attr), "Consolas")
    rFonts.set(qn("w:eastAsia"), "宋体")


def _body_width(doc):
    sec = doc.sections[0]
    return sec.page_width.twips - sec.left_margin.twips - sec.right_margin.twips


def _make_question_table(doc, n_rows):
    """一道大题一张表：一题一行，行不跨页。"""
    tbl = doc.add_table(rows=n_rows, cols=1)
    tbl.style = "Normal Table"   # 不要 Table Grid，否则代码块出现黑框
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    tbl.autofit = False
    tblPr = tbl._tbl.tblPr

    ind = OxmlElement("w:tblInd")
    ind.set(qn("w:w"), "0")
    ind.set(qn("w:type"), "dxa")
    tblPr.append(ind)

    mar = OxmlElement("w:tblCellMar")
    for side in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), "0")
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.append(mar)

    width = _body_width(doc)
    tbl.columns[0].width = Twips(width)
    for row in tbl.rows:
        trPr = row._tr.get_or_add_trPr()
        trPr.append(OxmlElement("w:cantSplit"))
        row.cells[0].width = Twips(width)
    return tbl


def _render_unit(cell, blocks):
    """把一道题（题干 + 代码 + 选项 + 小问）写进一个单元格。"""
    first = True
    for idx, block in enumerate(blocks):
        kind = block["type"]
        top = idx == 0
        before = UNIT_GAP if top else 0
        if kind == "question":
            _cell_para(cell, first, block["text"], left=Q_LEFT, first=-Q_LEFT,
                       before=before, after=0)
        elif kind == "text":
            _cell_para(cell, first, block["text"],
                       left=block.get("left_twips", BODY_LEFT),
                       before=(block.get("before_pt") if top else 0) or before,
                       after=block.get("after_pt", 0),
                       bold=block.get("bold", False))
        elif kind == "code":
            p = _cell_para(cell, first, "", left=BODY_LEFT, before=3,
                           after=4, line=CODE_LINE)
            lines = block["lines"]
            for i, line in enumerate(lines):
                run = p.add_run(line if line else " ")
                _set_code_font(run, CODE_SIZE)
                if i < len(lines) - 1:
                    run.add_break()
        elif kind == "option-row":
            items = block["items"]
            stops = (1.75, 3.50, 5.25) if len(items) >= 4 else ((3.40,) if len(items) == 2 else ())
            p = _cell_para(cell, first, "", left=OPT_LEFT, before=before, after=0, line=18)
            for pos in stops:
                p.paragraph_format.tab_stops.add_tab_stop(Inches(pos), WD_TAB_ALIGNMENT.LEFT)
            for i, item in enumerate(items):
                if i:
                    p.add_run("\t")
                run = p.add_run(item)
                run.font.size = Pt(OPT_SIZE)
        elif kind == "spacer":
            _cell_para(cell, first, "", left=BODY_LEFT, before=0,
                       after=block.get("after_pt", 0), line=TEXT_LINE)
        first = False


def render_section(doc, blocks):
    """一个大题：先说明段落，再一题一行建表。"""
    idx = 0
    while idx < len(blocks) and blocks[idx]["type"] != "question":
        b = blocks[idx]
        if b["type"] == "text":
            new_para(doc, b["text"], left=b.get("left_twips", BODY_LEFT),
                     before=b.get("before_pt", 0), after=b.get("after_pt", 0),
                     bold=b.get("bold", False), keep=True)
        idx += 1

    units = []
    for b in blocks[idx:]:
        if b["type"] == "question":
            units.append([b])
        elif units:
            units[-1].append(b)

    if not units:
        return
    tbl = _make_question_table(doc, len(units))
    for r, unit in enumerate(units):
        _render_unit(tbl.cell(r, 0), unit)


def render_document(doc, blocks):
    i = 0
    while i < len(blocks):
        b = blocks[i]
        if b["type"] == "heading":
            new_para(doc, b["text"], before=b.get("before_pt", 4),
                     after=b.get("after_pt", 2), bold=True, keep=True)
            j = i + 1
            while j < len(blocks) and blocks[j]["type"] != "heading":
                j += 1
            render_section(doc, blocks[i + 1:j])
            i = j
        else:
            render_section(doc, [b])
            i += 1
    # 文档不能以表格结尾，补一个极小空段
    tail = doc.add_paragraph(style="Normal")
    tail.paragraph_format.space_before = Pt(0)
    tail.paragraph_format.space_after = Pt(0)
    tail.paragraph_format.line_spacing = Pt(1)
    for run in tail.runs:
        run.font.size = Pt(1)


def build(spec_path: Path, output_path: Path, template_path: Path = DEFAULT_TEMPLATE):
    import json
    spec = json.loads(spec_path.read_text(encoding="utf-8-sig"))
    document = Document(str(template_path))
    score_tbl = deepcopy(document.tables[0]._tbl) if document.tables else None
    bfs._clear_body(document)
    document.core_properties.author = ""
    document.core_properties.last_modified_by = ""
    from datetime import datetime
    now = datetime.now().replace(microsecond=0)
    document.core_properties.created = now
    document.core_properties.modified = now

    update_fields = document.settings._element.find(qn("w:updateFields"))
    if update_fields is None:
        update_fields = OxmlElement("w:updateFields")
        document.settings._element.append(update_fields)
    update_fields.set(qn("w:val"), "true")

    bfs._add_header(document, spec["metadata"])
    if score_tbl is not None:
        document._body._element.insert(len(document._body._element) - 1, score_tbl)
        bfs._populate_score_table(document.tables[0], spec.get("score_table"))

    for page in spec["pages"]:
        render_document(document, page["blocks"])

    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(output_path))
    bfs._restore_parts(template_path, output_path)


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        raise SystemExit("usage: build_code_exam.py spec.json out.docx [template.docx]")
    template = Path(sys.argv[3]) if len(sys.argv) > 3 else DEFAULT_TEMPLATE
    build(Path(sys.argv[1]), Path(sys.argv[2]), template)
    print("built", sys.argv[2])
