# -*- coding: utf-8 -*-
"""卷面标注答案卷·通用派生器（exam-word-skill §11 补充路径）。

与 build_answer_keyed_paper.py 的区别：本脚本从**已有的学生卷 DOCX** 直接派生，
不要求 paper-spec.json / objective.json / spacer 结构。适合学生卷已排版定稿、
只有答案 Markdown 的情形（如学习通课程导出卷）。

流程：学生卷 DOCX + 答案 JSON → 卷面标注答案卷 DOCX
  ① 卷头「（X 卷）」→「（X 卷）答案」
  ② 客观题答案填入题干原有空括号（括号形态沿用原样）
  ③ 主观题答案块插在**该题全部内容之后**（代码/小问之后），标签绑定首要点
  保存后恢复模板 7 个保留部件，保证 verify_docx 保真通过。

用法
----
    python scripts/derive_answer_keyed_from_docx.py <student.docx> <answers.json> <out.docx>

answers.json 结构（内容来自答案 Markdown，同源）：
{
  "paper": "A 卷",
  "objective": {
    "一": ["B", "B", "..."],           // 按章，按卷面题号顺序
    "二": ["ABC", "..."],
    "三": ["√", "×", "..."]
  },
  "rubric": {
    "四": [[[2, "要点文本"], [3, "要点文本"]], ...],   // 每题 [[分值, 文本], ...]
    "五": [[...]],
    "六": [[...]]
  }
}

注意：不硬编码题数、章节名、分值分布；章节由 answers.json 的键决定，
题数由实际数组长度决定。题数与卷面不符时直接报错，不静默跳过。
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
    """整段写回第一个 run，清空其余（保留段落格式）。"""
    runs = p.runs
    if runs:
        runs[0].text = text
        for r in runs[1:]:
            r.text = ""
    else:
        p.add_run(text)


def add_answer_para(parent_para, text, after_el=None, pt=18, left=283, first=-283, bold=False):
    """在 after_el（或 parent_para）之后插入答案段落，返回新段落。"""
    new_p = OxmlElement("w:p")
    (after_el if after_el is not None else parent_para._p).addnext(new_p)
    from docx.text.paragraph import Paragraph
    para = Paragraph(new_p, parent_para._parent)
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


def derive(student, answers_path, out):
    answers = json.loads(Path(answers_path).read_text(encoding="utf-8-sig"))
    obj_map = answers["objective"]
    rubric = answers["rubric"]
    paper = answers.get("paper", "A 卷")

    doc = Document(str(student))
    paras = list(doc.paragraphs)

    # ① 客观题：按章计数，把答案填入题干最后一个空括号
    section, counters, filled = None, {k: 0 for k in obj_map}, 0
    for p in paras:
        s = p.text.strip()
        m0 = CHAPTER_RE.match(s)
        if m0:
            section = m0.group(1)
            continue
        if section in obj_map and QUESTION_RE.match(s):
            i = counters[section]
            if i >= len(obj_map[section]):
                raise ValueError(f"第{section}章题数超出 answers.json：题干 {s[:40]}")
            ans = obj_map[section][i]
            counters[section] += 1
            full = "".join(r.text for r in p.runs)
            matches = list(BRACKET_RE.finditer(full))
            if not matches:
                raise ValueError(f"题干无空括号：{full[:50]}")
            last = matches[-1]
            newtext = (full[:last.start()] + last.group(1) + ans + last.group(2)
                       + full[last.end():])
            set_para_text(p, newtext)
            filled += 1
    expected_obj = sum(len(v) for v in obj_map.values())
    if filled != expected_obj:
        raise ValueError(f"客观题填答 {filled} / 答案数据 {expected_obj}")
    for ch, n in counters.items():
        if n != len(obj_map[ch]):
            raise ValueError(f"第{ch}章卷面 {n} 题 / 答案数据 {len(obj_map[ch])} 题")

    # ② 主观题定位：章节属于 rubric 键、带（N 分）的题号段
    subject_starts = []
    for p in paras:
        s = p.text.strip()
        if not (QUESTION_RE.match(s) and p.paragraph_format.left_indent is not None):
            continue
        if not re.search(r"（\d+\s*分）", s):
            continue
        el = p._p.getprevious()
        ch = None
        while el is not None:
            mm = CHAPTER_RE.match(para_text(el).strip())
            if mm:
                ch = mm.group(1)
                break
            el = el.getprevious()
        if ch in rubric:
            subject_starts.append((p, ch))
    expected_subj = sum(len(v) for v in rubric.values())
    if len(subject_starts) != expected_subj:
        raise ValueError(f"主观题卷面 {len(subject_starts)} / 答案数据 {expected_subj}")

    idx_in_ch = {ch: 0 for ch in rubric}
    plan = []
    for i, (p, ch) in enumerate(subject_starts):
        n = idx_in_ch[ch]
        idx_in_ch[ch] += 1
        # 插入点 = 本题最后一个元素；遇下一章标题停止（章标题带 keep_next）
        nxt = subject_starts[i + 1][0]._p if i + 1 < len(subject_starts) else None
        cur, last = p._p.getnext(), p._p
        while cur is not None and cur is not nxt:
            if cur.tag in (qn("w:p"), qn("w:tbl")):
                t = para_text(cur).strip()
                if CHAPTER_RE.match(t):
                    break
                last = cur
            cur = cur.getnext()
        plan.append((p, last, ch, n))

    # ③ 从后往前插入答案块（标签绑定首要点、全块不拆）
    inserted = 0
    for anchor, last_el, ch, n in reversed(plan):
        points = rubric[ch][n]
        a = add_answer_para(anchor, LABEL, after_el=last_el, left=0, first=0, bold=True)
        a.paragraph_format.keep_with_next = True
        inserted += 1
        for mark, text in points:
            a = add_answer_para(a, f"（{mark} 分）{text}")
            inserted += 1

    # ④ 卷头加「答案」
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
        "subjective_blocks": len(subject_starts), "paragraphs_inserted": inserted,
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
