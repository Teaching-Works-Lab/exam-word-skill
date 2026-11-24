# -*- coding: utf-8 -*-
"""答案卷 / 评分细则专用渲染器（exam-word skill）。

何时用
------
生成「参考答案与评分细则」「答案速查」「A/B 卷答案对照」等**教师用**文档。
它与试卷卷面的区别（见 references/exam-authoring-pitfalls.md §10）：

* 卷头替换成教师用版：无「序号/学号/姓名/专业及班级」栏、无卷首得分表；
  显式标注「参考答案与评分细则 · 教师用」「请勿下发给学生」。
* 内容允许自然跨页（不像学生卷要留答题空白），但代码块与表格行各自
  ``cantSplit``，且表头跨页重复（``w:tblHeader``）。
* 表格用细灰边框 + 表头底纹；宽表（> 5 列，如双栏答案速查表）自动拆成
  左右两张，避免挤压折行。
* 代码块仍是 §9.3 规范：无边框、无行号、Consolas + 宋体、左缩进 425 twips。

spec 契约
---------
``{"metadata": {...}, "pages": [{"blocks": [...]}]}``，metadata 与试卷同源，
额外字段 ``doc_type: "answer"``。支持的 block：

* ``heading`` / ``subheading``（level 3 / 4）
* ``text``（可带 ``bold``）/ ``bullet``（可带 ``mark``）/ ``quote``
* ``table``（``rows``，``header`` 为 true 时首行加粗并重复）
* ``code``（``lines``）/ ``spacer``

复用 ``build_from_spec.py`` 的 ``_clear_body`` / ``_restore_parts``，
保证 7 个保留部件 byte-identical、``verify_docx.py`` 可通过。

用法
----
    python scripts/build_answer_docx.py spec.json out.docx [template.docx]
"""
import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor, Twips

SKILL = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SKILL / "scripts"))
import build_from_spec as bfs  # noqa: E402

DEFAULT_TEMPLATE = SKILL / "assets/reference.docx"

BODY_LEFT = 425
CODE_SIZE = 8.5
CODE_LINE = 11.0
BODY_SIZE = 10.0
BODY_LINE = 14.0
TABLE_SIZE = 8.5
TABLE_LINE = 11.5
GREY = RGBColor(0x59, 0x59, 0x59)


def _rfonts(run, ascii_font, east_font):
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), ascii_font)
    rFonts.set(qn("w:hAnsi"), ascii_font)
    rFonts.set(qn("w:cs"), ascii_font)
    rFonts.set(qn("w:eastAsia"), east_font)


def _para(doc, text="", *, size=BODY_SIZE, line=BODY_LINE, left=0, first=0,
          before=0, after=0, bold=False, color=None, keep=False,
          align=None, style="Normal"):
    p = doc.add_paragraph(style=style)
    pf = p.paragraph_format
    pf.left_indent = Twips(left)
    pf.first_line_indent = Twips(first)
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = Pt(line)
    pf.keep_with_next = keep
    pf.keep_together = True
    if align is not None:
        p.alignment = align
    if text:
        run = p.add_run(text)
        run.bold = bold
        run.font.size = Pt(size)
        if color is not None:
            run.font.color.rgb = color
        _rfonts(run, "Times New Roman", "宋体")
    return p


def _add_answer_header(doc, meta):
    """教师用卷头：沿用模板样式，但去掉学号栏与得分表。

    不用 tab 定位：学校名/学院名较长时 tab 位会超出正文宽度，
    导致「命题教师」姓名被挤到页外。改用全角空格串联、显式字号。
    """
    _para(doc, meta["school"], size=16, line=22, before=0, after=3,
          bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _para(doc,
          f"{meta['academic_term']}　命题单位：{meta['unit']}　命题教师：{meta['teacher']}",
          size=10.5, line=18, after=1)

    course = doc.add_paragraph(style="Normal")
    pf = course.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(1)
    pf.line_spacing = Pt(18)
    pf.keep_together = True
    r = course.add_run(f"【{meta['course']}】课程")
    r.bold = True
    r.font.size = Pt(10.5)
    _rfonts(r, "Times New Roman", "宋体")
    r = course.add_run(f"　（{meta.get('paper', 'A 卷')}）")
    r.font.size = Pt(10.5)
    _rfonts(r, "Times New Roman", "宋体")

    _para(doc, "参考答案与评分细则 · 教师用（请勿下发给学生）",
          size=11, line=17, before=3, after=2, bold=True,
          align=WD_ALIGN_PARAGRAPH.CENTER)
    _para(doc, meta["summary"], size=10, line=15, before=0, after=6,
          color=GREY)


def _table_borders(tbl):
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:color"), "909090")
        borders.append(el)
    tbl._tbl.tblPr.append(borders)


def _cell_shading(cell, fill="EDEDED"):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), fill)
    cell._tc.get_or_add_tcPr().append(shd)


def _render_table(doc, rows, header=True):
    if not rows:
        return
    cols = max(len(r) for r in rows)
    # 双栏答案速查表（8/6 列）拆成左右两张，避免挤压
    if cols > 5:
        half = cols // 2
        left = [r[:half] for r in rows]
        right = [r[half:half * 2] for r in rows
                 if any(str(c).strip() for c in r[half:half * 2])]
        _render_table(doc, left, header)
        _render_table(doc, right, header)
        return

    tbl = doc.add_table(rows=len(rows), cols=cols)
    tbl.style = "Normal Table"
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    tbl.autofit = True
    _table_borders(tbl)
    for i, row in enumerate(rows):
        trPr = tbl.rows[i]._tr.get_or_add_trPr()
        trPr.append(OxmlElement("w:cantSplit"))
        if header and i == 0:
            trPr.append(OxmlElement("w:tblHeader"))
        for j in range(cols):
            cell = tbl.cell(i, j)
            value = str(row[j]) if j < len(row) else ""
            cell.text = ""
            p = cell.paragraphs[0]
            pf = p.paragraph_format
            pf.space_before = Pt(1)
            pf.space_after = Pt(1)
            pf.line_spacing = Pt(TABLE_LINE)
            pf.keep_together = True
            run = p.add_run(value)
            run.font.size = Pt(TABLE_SIZE)
            run.bold = bool(header and i == 0)
            _rfonts(run, "Times New Roman", "宋体")
            if header and i == 0:
                _cell_shading(cell)
    # 小表整表不拆页：除末行外所有段落 keep_with_next
    if len(rows) <= 8:
        for i in range(len(rows) - 1):
            for cell in tbl.rows[i].cells:
                for p in cell.paragraphs:
                    p.paragraph_format.keep_with_next = True
    _para(doc, "", size=1, line=4, before=0, after=0)


def _render_code(doc, lines):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = "Normal Table"          # 无边框、无行号
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
    trPr = tbl.rows[0]._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:cantSplit"))

    cell = tbl.cell(0, 0)
    p = cell.paragraphs[0]
    pf = p.paragraph_format
    pf.left_indent = Twips(BODY_LEFT)
    pf.space_before = Pt(3)
    pf.space_after = Pt(4)
    pf.line_spacing = Pt(CODE_LINE)
    pf.keep_together = True
    for i, line in enumerate(lines):
        run = p.add_run(line if line else " ")
        run.font.size = Pt(CODE_SIZE)
        _rfonts(run, "Consolas", "宋体")
        if i < len(lines) - 1:
            run.add_break()
    _para(doc, "", size=1, line=4, before=0, after=0)


def _render_block(doc, block, next_type=None):
    kind = block["type"]
    if kind == "heading":
        _para(doc, block["text"], size=11.5, line=16, before=8, after=3,
              bold=True, keep=True)
    elif kind == "subheading":
        level = block.get("level", 3)
        _para(doc, block["text"], size=10.5 if level == 3 else 10,
              line=14, before=6 if level == 3 else 4, after=2,
              bold=True, keep=True)
    elif kind == "text":
        text = block["text"]
        # 标签行（评分：/ 给分提示：/ 评分要点：等）整行强调，便于阅卷定位
        emphasize = block.get("bold") or bool(
            __import__("re").match(r"^(评分|给分提示|评分要点|优秀示例|说明|指出|改进版|修改后|测试用例|分级样例)", text)
        )
        # 「评分要点：」这类标签必须与紧随其后的表格 / 代码同页
        keep = bool(next_type in ("table", "code") and text.rstrip().endswith(("：", ":")))
        _para(doc, text, bold=emphasize, keep=keep)
    elif kind == "bullet":
        _para(doc, (block.get("mark", "• ") + block["text"]),
              left=283, first=-283, after=0)
    elif kind == "quote":
        _para(doc, block["text"], size=9.5, line=13, left=BODY_LEFT,
              before=2, after=3, color=GREY)
    elif kind == "table":
        _render_table(doc, block["rows"], block.get("header", True))
    elif kind == "code":
        _render_code(doc, block["lines"])
    elif kind == "spacer":
        _para(doc, "", size=1, line=2, after=block.get("after_pt", 4))


def build(spec_path: Path, output_path: Path, template_path: Path = DEFAULT_TEMPLATE):
    import json
    from datetime import datetime

    spec = json.loads(spec_path.read_text(encoding="utf-8-sig"))
    document = Document(str(template_path))
    bfs._clear_body(document)

    document.core_properties.author = ""
    document.core_properties.last_modified_by = ""
    now = datetime.now().replace(microsecond=0)
    document.core_properties.created = now
    document.core_properties.modified = now

    update_fields = document.settings._element.find(qn("w:updateFields"))
    if update_fields is None:
        update_fields = OxmlElement("w:updateFields")
        document.settings._element.append(update_fields)
    update_fields.set(qn("w:val"), "true")

    _add_answer_header(document, spec["metadata"])
    for page in spec["pages"]:
        blocks = page["blocks"]
        for idx, block in enumerate(blocks):
            nxt = blocks[idx + 1]["type"] if idx + 1 < len(blocks) else None
            _render_block(document, block, nxt)

    # 文档不能以表格结尾
    _para(document, "", size=1, line=1, before=0, after=0)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(output_path))
    bfs._restore_parts(template_path, output_path)


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        raise SystemExit("usage: build_answer_docx.py spec.json out.docx [template.docx]")
    template = Path(sys.argv[3]) if len(sys.argv) > 3 else DEFAULT_TEMPLATE
    build(Path(sys.argv[1]), Path(sys.argv[2]), template)
    print("built", sys.argv[2])
