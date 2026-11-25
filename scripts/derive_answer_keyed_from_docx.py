# -*- coding: utf-8 -*-
"""卷面标注答案卷·通用派生器（exam-word-skill §11 / §12.4）。

从**已有的学生卷 DOCX** 直接派生，不要求 paper-spec.json / objective.json /
spacer 结构。兼容两种学生卷版式：
  A 型：一题一表，题干在段落（学习通 C&C++ A 卷）
  B 型：一大题一表、一行一题，题干在表格单元格（学习通 C&C++ B 卷）

三改：①卷头「（X 卷）」→「（X 卷）答案」；②客观题答案填入题干原有空括号
（段落或表格单元格内）；③主观题答案块插在该题**全部内容之后**（代码/小问
之后），标签绑定首要点。保存后恢复模板 7 个保留部件，保证 verify_docx 保真。

用法
----
    python scripts/derive_answer_keyed_from_docx.py <student.docx> <answers.json> <out.docx>

answers.json（内容来自答案 Markdown，同源；题数由数组长度决定，无硬编码）：
{
  "paper": "B 卷",
  "objective": {"一": ["A", ...], "二": ["ABD", ...], "三": ["×", ...]},
  "rubric": {
    "四": [[[2, "要点文本"], [3, "要点文本"]], ...],
    "五": [[...]],
    "六": [[...]]
  }
}
"""
import json
import re
import shutil
import sys
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt
from docx.text.paragraph import Paragraph

SKILL = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SKILL / "scripts"))
import build_from_spec as bfs  # noqa: E402

TEMPLATE = SKILL / "assets/reference.docx"
LABEL = "参考答案与评分要点："
CHAPTER_RE = re.compile(r"^([一二三四五六七八九十])、")
QUESTION_RE = re.compile(r"^\d+\s*[、.]")
BRACKET_RE = re.compile(r"([（(])\s*[\u3000 ]*\s*([)）])")


def para_text(el):
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def set_para_text(p, text):
    runs = p.runs
    if runs:
        runs[0].text = text
        for r in runs[1:]:
            r.text = ""
    else:
        p.add_run(text)


def add_answer_para(parent, text, after_el, pt=18, left=283, first=-283, bold=False):
    """在 after_el 之后插入答案段落；parent 为任意 Paragraph，提供 _parent。"""
    new_p = OxmlElement("w:p")
    after_el.addnext(new_p)
    para = Paragraph(new_p, parent._parent)
    pf = para.paragraph_format
    pf.left_indent = Pt(left / 20)
    pf.first_line_indent = Pt(first / 20)
    pf.line_spacing = Pt(pt)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.keep_together = True
    run = para.add_run(text)
    run.font.size = Pt(10.5)
    if bold:
        run.bold = True
    rPr = run._r.get_or_add_rPr()
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rPr.append(rf)
    rf.set(qn("w:ascii"), "Times New Roman")
    rf.set(qn("w:eastAsia"), "宋体")
    return para


def fill_bracket_in_para(p, ans, where):
    full = "".join(r.text for r in p.runs)
    matches = list(BRACKET_RE.finditer(full))
    if not matches:
        raise ValueError(f"题干无空括号：{full[:50]} ({where})")
    last = matches[-1]
    set_para_text(p, full[:last.start()] + last.group(1) + ans + last.group(2)
                  + full[last.end():])


def derive(student, answers_path, out):
    answers = json.loads(Path(answers_path).read_text(encoding="utf-8-sig"))
    obj_map, rubric = answers["objective"], answers["rubric"]
    paper = answers.get("paper", "A 卷")

    doc = Document(str(student))
    body = doc.element.body
    tbl_by_el = {tb._tbl: tb for tb in doc.tables}
    first_para = doc.paragraphs[0]

    # ================= 客观题：逐章按题号填括号（段落题 / 表格行题）=================
    section, counters, filled = None, {k: 0 for k in obj_map}, 0

    def feed_obj(text, container_paras, where):
        nonlocal filled
        if not (section in obj_map and QUESTION_RE.match(text)):
            return
        i = counters[section]
        if i >= len(obj_map[section]):
            raise ValueError(f"第{section}章题数超出 answers.json：{text[:40]}")
        ans = obj_map[section][i]
        counters[section] += 1
        for p in reversed(container_paras):
            if BRACKET_RE.search("".join(r.text for r in p.runs)):
                fill_bracket_in_para(p, ans, where)
                filled += 1
                return
        raise ValueError(f"第{section}章第{i+1}题找不到空括号：{text[:40]}")

    for el in body:
        tag = el.tag.split("}")[-1]
        if tag == "p":
            t = para_text(el).strip()
            cm = CHAPTER_RE.match(t)
            if cm:
                section = cm.group(1)
                continue
            if re.match(r"^[四五]、|^六、", t):
                section = None
                continue
            p = Paragraph(el, doc._body)
            feed_obj(t, [p], "段落题")
        elif tag == "tbl" and section in obj_map:
            tb = tbl_by_el.get(el)
            if tb is None:
                continue
            for ri, row in enumerate(tb.rows):
                cell = row.cells[0]
                txt = cell.text.strip()
                first_line = txt.splitlines()[0] if txt else ""
                if QUESTION_RE.match(first_line):
                    feed_obj(first_line, list(cell.paragraphs), f"表行 r{ri}")

    expected_obj = sum(len(v) for v in obj_map.values())
    if filled != expected_obj:
        raise ValueError(f"客观题填答 {filled} / 答案数据 {expected_obj}")
    for ch, n in counters.items():
        if n != len(obj_map[ch]):
            raise ValueError(f"第{ch}章卷面 {n} 题 / 答案数据 {len(obj_map[ch])} 题")

    # ================= 主观题定位（段落题 / 表格行题）=================
    # element 序列（只含 p / tbl）
    items = []
    for el in body:
        tag = el.tag.split("}")[-1]
        if tag in ("p", "tbl"):
            items.append((tag, el))
    idx_of = {id(el): i for i, (_, el) in enumerate(items)}

    chapter_now = None
    entries = []  # (kind, anchor, ch)
    # 段落题：(kind='p', el, ch)；表格题：(kind='cell', 题目所在单元格, ch)
    for tag, el in items:
        t = para_text(el).strip()
        if tag == "p":
            cm = CHAPTER_RE.match(t)
            if cm:
                chapter_now = cm.group(1)
                continue
            if (chapter_now in rubric and QUESTION_RE.match(t)
                    and re.search(r"（\d+\s*分）", t)):
                entries.append(("p", el, chapter_now))
        elif tag == "tbl" and chapter_now in rubric:
            tb = tbl_by_el.get(el)
            if tb is None:
                continue
            for row in tb.rows:
                cell = row.cells[0]
                txt = cell.text.strip()
                first_line = txt.splitlines()[0] if txt else ""
                if QUESTION_RE.match(first_line) and re.search(r"（\d+\s*分）", first_line):
                    entries.append(("cell", cell, chapter_now))

    expected_subj = sum(len(v) for v in rubric.values())
    if len(entries) != expected_subj:
        raise ValueError(f"主观题卷面 {len(entries)} / 答案数据 {expected_subj}")

    idx_in_ch = {ch: 0 for ch in rubric}
    plan = []  # (anchor_para, after_el, ch, n, kind)
    for i, (kind, obj, ch) in enumerate(entries):
        n = idx_in_ch[ch]
        idx_in_ch[ch] += 1
        if kind == "cell":
            # 表格题：答案插在该题单元格末尾（紧跟本题，不与邻题错位）
            last_para = obj.paragraphs[-1]
            plan.append((last_para, last_para._p, ch, n, "cell"))
            continue
        # 段落题：插在本题最后一个元素之后
        el = obj
        nxt_el = entries[i + 1][1] if i + 1 < len(entries) and entries[i + 1][0] == "p" else None
        cur, last = el.getnext(), el
        while cur is not None and cur is not nxt_el:
            if cur.tag in (qn("w:p"), qn("w:tbl")):
                if CHAPTER_RE.match(para_text(cur).strip()):
                    break
                last = cur
            cur = cur.getnext()
        plan.append((Paragraph(el, doc._body), last, ch, n, "p"))

    # 从后往前插入答案块
    inserted = 0
    for anchor, after_el, ch, n, kind in reversed(plan):
        points = rubric[ch][n]
        # 单元格内答案缩进对齐题干段（283），不加首行悬挂
        left = 283 if kind == "cell" else 0
        first = 0 if kind == "cell" else 0
        a = add_answer_para(anchor, LABEL, after_el, left=left, first=first, bold=True)
        a.paragraph_format.keep_with_next = True
        inserted += 1
        for mark, text in points:
            a = add_answer_para(a, f"（{mark} 分）{text}", a._p, left=left, first=first)
            inserted += 1

    for p in doc.paragraphs[:6]:
        if f"（{paper}）" in p.text and "答案" not in p.text:
            set_para_text(p, p.text.replace(f"（{paper}）", f"（{paper}）答案", 1))
            break

    tmp = Path("/tmp/_derive_keyed_raw.docx")
    doc.save(tmp)
    preserved = bfs._restore_parts(TEMPLATE, tmp)
    shutil.copy(tmp, out)
    return {
        "output": str(out), "objective_filled": filled,
        "subjective_blocks": len(entries), "paragraphs_inserted": inserted,
        "preserved_parts": len(preserved),
    }


def main():
    if len(sys.argv) != 4:
        print(__doc__)
        return 2
    result = derive(*[Path(x) for x in sys.argv[1:4]])
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
