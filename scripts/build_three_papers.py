# -*- coding: utf-8 -*-
"""三卷同出：一份源数据同时生成学生卷 / 卷面标注答案卷 / 学习通导入版。

为什么必须同出（2026-11-25 实战复盘）
----
同一套卷子在教学中有三种形态，格式互不兼容，人工分别维护必然漂移：
  1. 学生卷（正式 Word 排版，印刷下发）—— exam-spec.json → build_from_spec.py
  2. 卷面标注答案卷（学生卷原卷面填空+答案要点，教师阅卷）—— 同一 spec 派生
  3. 学习通导入版（每行 `序号.【题型】题干` + `答案：` 行，供 Word 智能导入）——
      objective.json + 答案 md 派生，编号是全局连续，与前两者分题型编号不同

只改其中一份 docx 是本次真实事故：第 14 题换题只改了学习通导入版，学生卷与
标注答案卷仍是旧题，三份内容不一致。根因是「改 docx 而不是改源 + 三份各自生成」。

用法
----
    python scripts/build_three_papers.py \
        --spec exam-spec-A.json \
        --objective objective-A.json \
        --answers answer-A.md \
        --outdir out/ --stem 智能制造装备-期末笔试A卷 [--template assets/reference.docx]

输出（均在 outdir）
----
    <stem>.docx                       学生卷
    <stem>-标注答案卷.docx            卷面标注答案卷
    <stem>-学习通导入.docx            学习通导入版

硬约束
----
* spec / objective / answers 是唯一内容源；脚本只读它们，绝不读旧 docx。
  换题、改答案一律先改源，再重跑本脚本，三份一起重建（§15 改动传播）。
* 三份共享同一 objective.json，题数/答案天然一致；脚本构建后交叉断言：
  学生卷与标注卷题干逐字一致（标注卷仅多答案）、客观题答案与 objective 一致。
* 每份都跑 verify_docx.py（模板保真）与 check_final_gate.py（卷面干净）。
"""
import argparse
import json
import re
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
SKILL = HERE.parent


def run(cmd):
    print("  $", " ".join(str(c) for c in cmd), flush=True)
    r = subprocess.run([str(c) for c in cmd], capture_output=True, text=True)
    if r.stdout:
        print(r.stdout.strip()[:800], flush=True)
    if r.returncode != 0:
        print(r.stderr.strip()[:1200], file=sys.stderr)
        raise SystemExit(f"命令失败({r.returncode}): {cmd[0]}")


def norm(s):
    return re.sub(r"\s+", "", str(s or ""))


def parse_subjective(md):
    """从答案 md 解析主观题 [(参考答案, 评分要点)]，按出现顺序。"""
    sec = md.split("## 二、主观题参考答案与评分要点", 1)
    if len(sec) < 2:
        return []
    sec = sec[1]
    parts = re.split(r"^####\s*第\s*\d+\s*题[^\n]*$", sec, flags=re.M)
    out = []
    for body in parts[1:]:
        m = re.search(r"\*\*参考答案\*\*\s*\n(.*?)(?=\n\*\*评分要点\*\*|\Z)", body, re.S)
        ans = m.group(1).strip() if m else ""
        m2 = re.search(r"\*\*评分要点\*\*\s*\n(.*?)(?=\n#{3,4}\s|\Z)", body, re.S)
        rub = m2.group(1).strip() if m2 else ""
        out.append((ans, rub))
    return out


def subjective_stems(spec):
    """按出现顺序取学生卷 spec 的主观题题干（大题起始按 `N.` / `N、`）。"""
    ordered, cur, in_subj = [], None, False
    for page in spec.get("pages", []):
        for b in page.get("blocks", []):
            t = b.get("text", "")
            if b.get("type") == "heading":
                # 主观题大题标题关键词。除「简答/综合/分析」外还必须认「应用题」
                # ——漏识会让学习通导入版少题、三卷同出断言失败（2026-11-26 实战）。
                # 程序类卷的「程序阅读/查错/编程」同属主观题，一并覆盖。
                in_subj = any(k in t for k in (
                    "简答", "应用", "综合", "分析", "论述", "设计",
                    "程序", "编程", "阅读", "查错", "改错", "证明", "计算题"))
                continue
            if not in_subj or b.get("type") != "question" or not t.strip():
                continue
            if re.match(r"^\s*\d+\s*[.．、]", t):
                cur = [re.sub(r"^\s*\d+\s*[.．、]\s*", "", t)]
                ordered.append(cur)
            elif cur is not None:
                cur.append(t)
    return ["\n".join(x) for x in ordered]


def norm_answer(a):
    a = str(a).strip()
    if a in ("√", "对", "T", "True", "正确"):
        return "正确"
    if a in ("×", "错", "F", "False", "错误"):
        return "错误"
    return a


def build_chaoxing_import(objective, spec, answers, out):
    """学习通 Word 智能导入版：客观题 + 主观题，全局连续编号。"""
    from docx import Document
    from docx.shared import Pt

    subj = parse_subjective(answers)
    stems = subjective_stems(spec)
    assert len(subj) == len(stems), f"主观题数不符: 答案{len(subj)} vs 题干{len(stems)}"

    doc = Document()
    st = doc.styles["Normal"]
    st.font.name, st.font.size = "宋体", Pt(10.5)

    for i, q in enumerate(objective, 1):
        typ = q["type"]
        doc.add_paragraph(f"{i}.【{typ}】{q['stem']}")
        for opt in q.get("options", []):
            doc.add_paragraph(opt)
        doc.add_paragraph(f"答案：{norm_answer(q['answer'])}")
        if q.get("topic"):
            doc.add_paragraph(f"知识点：{q['topic']}")
        if q.get("diff"):
            doc.add_paragraph(f"难易程度：{q['diff']}")

    for i, ((ans, rub), stem) in enumerate(zip(subj, stems), 1):
        kind = "简答题" if "简答" in stem[:40] or i <= 4 else "综合应用题"
        doc.add_paragraph(f"{i}.【{kind}】{stem}")
        doc.add_paragraph(f"答案：{ans}")
        if rub:
            doc.add_paragraph(f"答案解析：{rub}")
    doc.save(out)
    return len(objective), len(subj)


def stem_core(stem):
    """题干主干：去掉末尾答案括号（学生卷是空括号、标注卷已填答案），用于跨卷比对。

    标注卷括号里可能是单选字母(B)、多选组合(BCD)、判断记号(√/×)，此前只匹配
    空括号导致多选/判断题干比对失败（2026-11-26 实战）。
    """
    s = str(stem or "").strip()
    # 把括号内的答案归一为空：学生卷（　）/ 标注卷（B）（BCD）（√）统一成（　），
    # 保留括号骨架，否则删空括号后与卷面文本对不上（2026-11-26 实战）。
    s = re.sub(r"[（(][^）)]*[）)]", "（　）", s)
    s = re.sub(r"^\s*\d+\s*[.．、]\s*", "", s)  # 题号前缀：objective stem 无题号，卷面题干有
    return norm(s)


def cross_check(student, keyed, objective):
    """三卷一致性断言：学生卷与标注卷题干主干一致；标注卷已把答案填进括号。"""
    from docx import Document

    s = [p.text.strip() for p in Document(student).paragraphs if p.text.strip()]
    k = [p.text.strip() for p in Document(keyed).paragraphs if p.text.strip()]
    # 卷面文本的括号里可能是已填答案（B/BCD/√），与 stem_core 的归一规则保持一致。
    blank_brackets = lambda t: norm(re.sub(r"[（(][^）)]*[）)]", "（　）", t))
    s_all, k_all = blank_brackets("".join(s)), blank_brackets("".join(k))
    cores = [stem_core(q["stem"]) for q in objective if q["stem"].strip()]
    missing_s = [x for x in cores if x and x not in s_all]
    missing_k = [x for x in cores if x and x not in k_all]
    if missing_s or missing_k:
        raise SystemExit(
            f"一致性断言失败: 学生卷缺{len(missing_s)}题, 标注卷缺{len(missing_k)}题; "
            f"样例={ (missing_s or missing_k)[:1] }"
        )
    # 标注卷客观题答案必须已填：空 bracket 不应成对出现在题干行
    print(f"  一致性断言 PASS: {len(cores)} 道客观题题干主干在学生卷与标注卷均一致", flush=True)


def main():
    ap = argparse.ArgumentParser(description="三卷同出：学生卷 + 标注答案卷 + 学习通导入版")
    ap.add_argument("--spec", required=True, help="学生卷 spec.json")
    ap.add_argument("--objective", required=True, help="客观题同源 JSON（list）")
    ap.add_argument("--answers", required=True, help="参考答案与评分标准 md")
    ap.add_argument("--outdir", required=True)
    ap.add_argument("--stem", required=True, help="输出文件名主干")
    ap.add_argument("--template", default=str(SKILL / "assets" / "reference.docx"))
    ap.add_argument("--keyed-spec", default=None,
                    help="标注答案卷 spec.json；不给则由 build_answer_keyed_paper.py 派生")
    a = ap.parse_args()

    out = Path(a.outdir).resolve()
    out.mkdir(parents=True, exist_ok=True)
    spec = json.loads(Path(a.spec).read_text(encoding="utf-8-sig"))
    obj = json.loads(Path(a.objective).read_text(encoding="utf-8-sig"))
    if isinstance(obj, dict):
        obj = obj.get("questions", obj.get("items", []))
    md = Path(a.answers).read_text(encoding="utf-8")

    student = out / f"{a.stem}.docx"
    keyed = out / f"{a.stem}-标注答案卷.docx"
    cx = out / f"{a.stem}-学习通导入.docx"

    print("[1/3] 学生卷", flush=True)
    run([sys.executable, HERE / "build_from_spec.py", a.template, a.spec, student])

    print("[2/3] 卷面标注答案卷", flush=True)
    if a.keyed_spec:
        run([sys.executable, HERE / "build_from_spec.py", a.template, a.keyed_spec, keyed])
    else:
        derived = out / f"{a.stem}.keyed-spec.json"
        run([sys.executable, HERE / "build_answer_keyed_paper.py", a.spec, a.objective,
             a.answers, keyed, "--spec-out", derived, "--template", a.template])

    print("[3/3] 学习通导入版", flush=True)
    no, ns = build_chaoxing_import(obj, spec, md, cx)
    print(f"  客观题={no} 主观题={ns}", flush=True)

    print("[校验] 一致性 + 模板保真 + 卷面干净", flush=True)
    cross_check(student, keyed, obj)
    run([sys.executable, HERE / "verify_docx.py", student, a.template])
    run([sys.executable, HERE / "verify_docx.py", keyed, a.template])
    run([sys.executable, HERE / "check_final_gate.py", student, keyed])

    print("\n三卷同出完成：", flush=True)
    for p in (student, keyed, cx):
        print(f"  {p}  ({p.stat().st_size} B)", flush=True)


if __name__ == "__main__":
    main()
