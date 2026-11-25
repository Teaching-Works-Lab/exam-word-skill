# -*- coding: utf-8 -*-
"""卷面标注答案卷生成器（exam-word-skill §11）。

把「学生卷」派生成「答案卷」：卷面全量重印，客观题答案填进题干原有括号，
主观题答题空白换成统一标签 + 逐条带分值的参考答案要点。

用法
----
    python scripts/build_answer_keyed_paper.py <paper-spec.json> <objective.json> <answer.md> <out.docx>
        [--spec-out keyed-spec.json] [--marks-json marks.json]
        [--template reference.docx] [--label 答案]

输入
----
paper-spec.json   学生卷结构化 spec（metadata / score_table / pages）
objective.json    客观题同源 JSON（list，含 type / stem / answer）
answer.md         参考答案与评分标准 Markdown（答案速查表 + 主观题参考答案 / 评分要点）

门禁（构建即断言，失败直接报错）
----
* 客观题：速查表答案 == objective.json 答案；题干去标点后包含对应 stem；
  18 个空括号全部填上，括号形态沿用学生卷原样；判断题统一 √ / ×。
* 主观题：spacer 数 == 主观题数；每条参考答案要点挂整数分值，
  格式统一 `（N 分）`，且每题分值合计 == 该题满分。
* 派生 spec 合并为单页自然流动，不预留硬分页（§11.3-7）。
* 卷头卷别后插「答案」，学号姓名栏与空得分表原样保留（§11.3-1）。

分值解析
----
优先从「评分要点」自动提取三级策略（（N）X 分 / ①X 分 / 分号分段取首个 X 分），
过滤「扣 X 分」「不超过 X 分」等干扰；解析结果与要点条数、满分不符时报错，
可用 --marks-json 显式覆盖：{"简答1": [3, 4, 3], "综合2": [4, 4, 4, 3]}。
"""
import argparse
import json
import re
import sys
import tempfile
from pathlib import Path

SKILL = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SKILL / "scripts"))
import build_from_spec as bfs  # noqa: E402

DEFAULT_TEMPLATE = SKILL / "assets/reference.docx"
TYPES = ("单选题", "多选题", "判断题")
RUBRIC_LABEL = "参考答案与评分要点："
# 卷面题干沿用学生卷 20pt；答案块统一 18pt（密度更高，避免末页孤页）。
# 两者都由生成器显式设定，不是从答案册粘贴带来的「混血」（§11.3-6）。
PAPER_LINE_PT = 20
ANSWER_LINE_PT = 18
TRUE_FALSE = {
    "正确": "√", "错误": "×", "√": "√", "×": "×",
    "对": "√", "错": "×", "T": "√", "F": "×",
}
BRACKET_RE = re.compile(r"([（(])\s*([)）])")
SUBJ_HEAD = re.compile(r"^[四五]、")
# 主观题「主问题」：章标题下的题号行（全角/半角题号均可），区别于小问（1）（2）
MAINQ_RE = re.compile(r"^\d+\s*[．.、]")
NORM_RE = re.compile(r"[\s，。、（）()：:；;「」『』【】\[\],.?？!！]")


def norm(s):
    return NORM_RE.sub("", s or "")


def parse_answer_md(path):
    """解析答案册：客观题速查表 + 主观题参考答案 / 评分要点。"""
    txt = Path(path).read_text(encoding="utf-8")
    quick = {}
    for typ in TYPES:
        m = re.search(rf"^### {typ}[^\n]*$(.*?)(?=^### |^## |\Z)", txt, re.S | re.M)
        rows = []
        for line in (m.group(1).splitlines() if m else []):
            line = line.strip()
            if not line.startswith("|") or re.match(r"^\|\s*-", line):
                continue
            cells = [c.strip() for c in line.strip("|").split("|")]
            if cells and cells[0] == "题号":
                continue
            for i in range(0, len(cells) - 1, 2):
                num, ans = cells[i], cells[i + 1]
                if num.isdigit() and ans:
                    rows.append((int(num), ans))
        quick[typ] = rows

    subj = []
    marks = list(re.finditer(
        r"^(### (四、简答题|五、综合应用题).*|#### 第 (\d+) 题（(U\d)，(\d+) 分）.*)$",
        txt, re.M))
    kind = None
    for i, m in enumerate(marks):
        if m.group(2):
            kind = "简答" if m.group(2).startswith("四") else "综合"
            continue
        end = marks[i + 1].start() if i + 1 < len(marks) else len(txt)
        body = txt[m.end():end]
        am = re.search(r"\*\*参考答案\*\*(.*?)\*\*评分要点\*\*", body, re.S)
        rm = re.search(r"\*\*评分要点\*\*(.*)", body, re.S)
        if not am or not rm:
            raise ValueError(f"答案册缺少参考答案/评分要点：第 {m.group(3)} 题")
        subj.append({
            "kind": kind, "n": int(m.group(3)), "unit": m.group(4),
            "score": int(m.group(5)),
            "answer": [l.strip() for l in am.group(1).splitlines() if l.strip()],
            "rubric": [l.strip() for l in rm.group(1).splitlines() if l.strip()],
        })
    return quick, subj


def parse_marks(rubric, n_points, full, where):
    """从评分要点提取各要点分值；失败抛错，要求显式 --marks-json。"""
    joined = " ".join(rubric)
    numbered = re.findall(r"（(\d+)）\s*(\d+)\s*分", joined)
    circled = re.findall(r"[①②③④⑤⑥⑦⑧⑨⑩]\s*(\d+)\s*分", joined)
    if numbered and len(numbered) == n_points:
        vals = [int(v) for _, v in numbered]
    elif circled and len(circled) == n_points:
        vals = [int(v) for v in circled]
    else:
        vals = []
        for seg in re.split(r"[；;。]", joined):
            if re.search(r"扣|不超过|总分|不给满分", seg):
                continue
            m = re.search(r"(\d+)\s*分", seg)
            if m:
                vals.append(int(m.group(1)))
    if len(vals) != n_points or sum(vals) != full:
        raise ValueError(
            f"{where}: 分值解析失败（要点 {n_points} 条 / 满分 {full}，"
            f"解析到 {vals}，合计 {sum(vals)}）；请用 --marks-json 显式提供")
    return vals


def format_point(text, mark):
    """把分值挂到要点名之后，全卷统一 `（N 分）`。

    先归一化正文里已有的紧凑分值写法（`2分`/`(2分)`/`（2 分）` 之外的变体），
    避免同一份答案卷里出现 `（2分）` 与 `（2 分）` 混排（§11.5）。
    """
    text = re.sub(r"[（(]\s*(\d+)\s*分\s*[）)]", "", text).strip()
    label = f"（{mark} 分）"
    m = re.match(r"^(【[^】]+】)", text)
    if m:
        return text[:m.end()] + label + text[m.end():].lstrip("：: ")
    m = re.match(r"^([（(]\d+[)）]|[①②③④⑤⑥⑦⑧⑨⑩])", text)
    if m:
        return text[:m.end()] + label + text[m.end():]
    return label + text


def ensure_subjective_spacers(blocks, expected):
    """主观题答题空白不足时，在每道主问题末尾自动补 spacer。

    返回 ``(new_blocks, added)``。主问题 = 主观题章（四/五/六）内的题号
    question 行；每道主问题的「答题空白」= 其后到下一主问题/章标题之间的
    spacer。已有 spacer 的题不重复补。这样学生卷 spec 漏写答题空白时，
    不必手工派生 spec（§11 易用性增强）。
    """
    have = sum(1 for b in blocks if b["type"] == "spacer")
    if have >= expected:
        return list(blocks), 0
    marks = [i for i, b in enumerate(blocks)
             if b["type"] == "heading" and SUBJ_HEAD.match(b.get("text", ""))]
    if not marks:
        return list(blocks), 0
    start = min(marks)
    main, in_subj = [], False
    for i in range(start, len(blocks)):
        b = blocks[i]
        if b["type"] == "heading":
            in_subj = bool(SUBJ_HEAD.match(b.get("text", "")))
            continue
        if in_subj and b["type"] == "question" and MAINQ_RE.match(b.get("text", "")):
            main.append(i)
    insert_at = []
    for j, ms in enumerate(main):
        end = main[j + 1] if j + 1 < len(main) else len(blocks)
        for k in range(ms + 1, end):
            if blocks[k]["type"] == "heading":
                end = k
                break
        if not any(blocks[k]["type"] == "spacer" for k in range(ms, end)):
            insert_at.append(end)
    out = list(blocks)
    for pos in sorted(set(insert_at), reverse=True):
        out.insert(pos, {"type": "spacer", "text": "", "after_pt": 8})
    return out, len(insert_at)


def fill_bracket(stem, ans, where):
    matches = list(BRACKET_RE.finditer(stem))
    if not matches:
        raise ValueError(f"{where}: 题干末尾找不到空括号：{stem[:60]}")
    last = matches[-1]
    return stem[:last.start()] + last.group(1) + ans + last.group(2) + stem[last.end():]


def check_objective(objective, quick):
    for typ in TYPES:
        group = [x for x in objective if x["type"] == typ]
        rows = quick[typ]
        if len(group) != len(rows):
            raise ValueError(f"{typ}: objective {len(group)} 题 / 速查表 {len(rows)} 题不一致")
        for (num, ans), item in zip(rows, group):
            a = TRUE_FALSE.get(item["answer"], item["answer"])
            q = TRUE_FALSE.get(ans, ans)
            if a != q:
                raise ValueError(f"{typ} 第 {num} 题答案不一致：objective={a} / 速查表={q}")
    return sum(len(quick[t]) for t in TYPES)


def build_keyed_spec(paper_path, objective_path, answer_md_path,
                     marks_override=None, label="答案"):
    paper = json.loads(Path(paper_path).read_text(encoding="utf-8-sig"))
    objective = json.loads(Path(objective_path).read_text(encoding="utf-8"))
    quick, subj = parse_answer_md(answer_md_path)
    marks_override = marks_override or {}

    total_obj = check_objective(objective, quick)
    meta = dict(paper["metadata"])
    meta["answer_label"] = label
    meta["doc_type"] = "answer-keyed"

    blocks = [dict(b) for page in paper["pages"] for b in page["blocks"]]

    # 0) 答题空白自动补齐：spacer 数少于主观题数时，在每道主问题末尾补
    blocks, spacers_added = ensure_subjective_spacers(blocks, len(subj))

    # 1) 客观题：按 objective 顺序把答案填进题干原有空括号
    obj_iter = iter(objective)
    in_objective, filled = True, 0
    for b in blocks:
        if b["type"] == "heading":
            if SUBJ_HEAD.match(b.get("text", "")):
                in_objective = False
            continue
        if b["type"] == "question" and in_objective:
            item = next(obj_iter)
            stem = b["text"]
            if norm(item["stem"]) not in norm(stem):
                raise ValueError(f"客观题题干与 objective 错位：{stem[:40]} / {item['stem'][:40]}")
            ans = TRUE_FALSE.get(item["answer"], item["answer"])
            b["text"] = fill_bracket(stem, ans, f"客观题第 {filled + 1} 题")
            filled += 1
    if filled != total_obj:
        raise ValueError(f"客观题填答数 {filled} != {total_obj}")

    # 2) 主观题：每个 spacer 换成一题的答案块
    answer_blocks = []
    for s in subj:
        key = f"{s['kind']}{s['n']}"
        marks = (marks_override.get(key)
                 if key in marks_override
                 else parse_marks(s["rubric"], len(s["answer"]), s["score"], key))
        if len(marks) != len(s["answer"]) or sum(marks) != s["score"]:
            raise ValueError(f"{key}: 分值 {marks} 与要点数/满分不符")
        points = [format_point(t, m) for t, m in zip(s["answer"], marks)]
        answer_blocks.append([
            {"type": "text", "text": RUBRIC_LABEL, "bold": True,
             "keep_next": True, "left_twips": 0, "first_twips": 0,
             "line_pt": ANSWER_LINE_PT},
        ] + [
            {"type": "text", "text": p, "left_twips": 283, "first_twips": -283,
             "line_pt": ANSWER_LINE_PT, "keep_together": True}
            for p in points
        ])

    out, si = [], 0
    for b in blocks:
        if b["type"] == "spacer":
            if si >= len(answer_blocks):
                raise ValueError("spacer 数多于主观题数")
            out.extend(answer_blocks[si])
            si += 1
        else:
            out.append(b)
    if si != len(answer_blocks) or len(subj) != si:
        raise ValueError(f"主观题 {len(subj)} 道 / spacer {si} 个，无法一一对应")
    if any(b["type"] == "spacer" for b in out):
        raise ValueError("仍有残留答题空白")

    spec = {"metadata": meta, "score_table": paper.get("score_table"),
            "pages": [{"blocks": out}]}
    stats = {
        "objective_filled": filled,
        "subjective": [{"key": f"{s['kind']}{s['n']}", "points": len(s["answer"]),
                        "marks": (marks_override.get(f"{s['kind']}{s['n']}")
                                  or parse_marks(s["rubric"], len(s["answer"]), s["score"],
                                                 f"{s['kind']}{s['n']}")),
                        "full": s["score"]} for s in subj],
        "blocks": len(out),
        "spacers_added": spacers_added,
        "hard_page_breaks": 0,
    }
    return spec, stats


def main():
    ap = argparse.ArgumentParser(description="Build an answer-keyed exam paper (§11)")
    ap.add_argument("paper_spec")
    ap.add_argument("objective_json")
    ap.add_argument("answer_md")
    ap.add_argument("out_docx")
    ap.add_argument("--spec-out")
    ap.add_argument("--marks-json")
    ap.add_argument("--template", default=str(DEFAULT_TEMPLATE))
    ap.add_argument("--label", default="答案")
    args = ap.parse_args()

    override = None
    if args.marks_json:
        override = json.loads(Path(args.marks_json).read_text(encoding="utf-8"))
    spec, stats = build_keyed_spec(args.paper_spec, args.objective_json,
                                   args.answer_md, override, args.label)
    if args.spec_out:
        Path(args.spec_out).parent.mkdir(parents=True, exist_ok=True)
        Path(args.spec_out).write_text(
            json.dumps(spec, ensure_ascii=False, indent=2), encoding="utf-8")
        spec_path = Path(args.spec_out)
    else:
        tmp = tempfile.NamedTemporaryFile("w", suffix=".json", delete=False, encoding="utf-8")
        json.dump(spec, tmp, ensure_ascii=False)
        tmp.close()
        spec_path = Path(tmp.name)

    result = bfs.build_document(Path(args.template), spec_path, Path(args.out_docx))
    print(json.dumps({"output": result["output"], "stats": stats}, ensure_ascii=False, indent=2))
    if not args.spec_out:
        spec_path.unlink(missing_ok=True)


if __name__ == "__main__":
    main()
