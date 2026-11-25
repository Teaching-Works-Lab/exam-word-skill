#!/usr/bin/env python3
"""卷内内容质量门禁：查前后泄题、重复考点、单元覆盖缺口。

现有 verify_docx.py 只查模板保真，verify_answer_keyed_paper.py 只查答案卷
结构；本脚本补上“题目内容本身是否有问题”的通用门禁，适用于任意课程试卷。

输入：
  paper-spec.json    学生卷结构化 spec
  objective.json     [{"type","stem","answer"}, ...]
  answer.md          参考答案与评分标准（与 build_answer_keyed_paper.py 同格式）
  --blueprint bp.json 可选：{"units":[{"id","name","keywords":[...]}]}

检查：
  1. 卷内重复：客观题之间、主观题之间的题干高度相似或完全相同。
  2. 前后泄题：前面客观题的“题干 + 正确选项”与后面主观题的
     “题干 + 参考答案 + 评分要点”高度重合，等于先把答案说了。
  3. 覆盖缺口：blueprint 中每个单元的关键词是否在卷面/答案中出现。
  4. 题干清晰性：主观题是否用未定义的口号式代称或指代要求考生说明/比较。

输出 JSON / Markdown，含每条问题的证据与“怎么办”建议。
"""

from __future__ import annotations

import argparse
import difflib
import html
import json
import re
import sys
import unicodedata
from pathlib import Path

# 阈值：经验值，宁可多报“需复核”也不漏掉真泄题；文档里有说明。
DUP_STEM_RATIO = 0.72
LEAK_RATIO = 0.60
LEAK_CONTAIN = 0.50
LEAK_TERM_COVER = 0.42
MIN_CLAIM_LEN = 8
NGRAM = 5
MAX_TOKEN_LEN = 5

PUNCT_RE = re.compile(r"[\s，。、（）()：:；;「」『』【】\[\],.?？!！“”\"'’‘—\-_/\\]+")
OPT_PREFIX_RE = re.compile(r"^[A-Ea-e][\.．、]\s*")
TERM_RE = re.compile(r"[A-Za-z][A-Za-z0-9+#._-]*|[\u4e00-\u9fff]{2,}")

STOP_TERMS = {
    "人工智能", "机器学习", "深度学习", "大语言", "模型", "数据", "算法", "输出", "输入",
    "训练", "学习", "问题", "应用", "技术", "系统", "方法", "结果", "能力", "环境",
    "内容", "可以", "主要", "用于", "下列", "说法", "正确", "错误", "进行", "通过",
    "需要", "使用", "具有", "相关", "生成", "智能", "一个", "这种", "其他", "以及",
    "以下", "什么", "下列", "哪项", "属于", "不属", "用于判断", "下列哪", "下列哪个",
}

SECTION_CN = {"一": "单选", "二": "多选", "三": "判断", "四": "简答", "五": "综合应用"}
HEAD_RE = re.compile(r"^([一二三四五六])、")

# 题干清晰性：抓住“自造口号式代称 + 要求解释/比较 + 未给定义”的命题缺陷。
QUOTED_RE = re.compile(r"[“\"]([^”\"]{2,24})[”\"]")
SLOGAN_VERB_RE = re.compile(r"(改造|借用|借助|利用|复用|提升|优化|强化|改善|扩充|处理)")
GENERIC_NOUN_RE = re.compile(r"(样本|能力|特征|模型|数据|知识|方法|效果|资源|内容)")
DEFINE_RE = re.compile(r"(即|是指|所谓|定义为|分别指|具体指|含义是|意思是|例如|比如|包括)")
EXPLAIN_RE = re.compile(r"(说明|比较|分析|评价|简述|阐述|解释|指出|给出|为什么)")
TWO_WAYS_RE = re.compile(r"两(?:个|条|种)(?:思路|做法|方法|方式|途径)")
VAGUE_REF_RE = re.compile(r"(上述|前述|该做法|这种做法|此做法|这两(?:个|条|种)(?:思路|做法|方法|方式))")


def norm(s: str) -> str:
    s = unicodedata.normalize("NFKC", html.unescape(s or ""))
    return PUNCT_RE.sub("", s)


def ngrams(s: str, n: int = NGRAM) -> set[str]:
    s = norm(s)
    if len(s) < n:
        return {s} if s else set()
    return {s[i:i + n] for i in range(len(s) - n + 1)}


def ratio(a: str, b: str) -> float:
    a, b = norm(a), norm(b)
    if not a or not b:
        return 0.0
    return difflib.SequenceMatcher(None, a, b).ratio()


def containment(short: str, long: str) -> float:
    """短文本的信息有多少被长文本覆盖；用于“正确选项被主观题答案照搬”。"""
    sa, la = ngrams(short), ngrams(long)
    if not sa or not la:
        return 0.0
    return len(sa & la) / len(sa)


def terms(s: str) -> set[str]:
    """抽出中文 2~4 字片段与英文术语，作为“考点指纹”。"""
    s = norm(s)
    out: set[str] = set()
    for m in TERM_RE.finditer(s):
        raw = m.group(0)
        if re.match(r"[A-Za-z]", raw):
            out.add(raw.lower())
            continue
        for n in (2, 3, 4):
            for i in range(len(raw) - n + 1):
                t = raw[i:i + n]
                if t not in STOP_TERMS:
                    out.add(t)
    return out


def shared_chunks(a: str, b: str, minimum: int = 3) -> list[str]:
    """两段文本连续相同的片段；这是“换句式重复”最稳的证据。"""
    aa, bb = norm(a), norm(b)
    out: list[str] = []
    for block in difflib.SequenceMatcher(None, aa, bb, autojunk=False).get_matching_blocks():
        chunk = aa[block.a:block.a + block.size]
        if block.size >= minimum and chunk not in STOP_TERMS:
            out.append(chunk)
    return sorted(out, key=len, reverse=True)


def concept_coverage(short: str, long: str, minimum: int = 3) -> tuple[float, list[str]]:
    """短文本里有多少“连续相同片段”出现在长文本中；片段越长权重越高。"""
    chunks = shared_chunks(short, long, minimum)
    if not chunks:
        return 0.0, []
    denom = max(1, len(norm(short)))
    score = sum(len(c) ** 1.5 for c in chunks) / denom
    return min(1.0, score), chunks


def flatten_blocks(spec: dict) -> list[dict]:
    out = []
    for page in spec.get("pages", []):
        out.extend(page.get("blocks", []))
    return out


def questions_from_spec(spec: dict) -> list[dict]:
    """按卷面顺序抽出题目，附带所属大题、选项、小问文本。"""
    qs: list[dict] = []
    cur_section = None
    cur: dict | None = None

    def flush() -> None:
        nonlocal cur
        if cur is not None:
            qs.append(cur)
            cur = None

    for b in flatten_blocks(spec):
        typ = b.get("type", "")
        text = b.get("text", "")
        if typ == "heading":
            flush()
            m = HEAD_RE.match(text)
            cur_section = m.group(1) if m else None
            continue
        if typ == "question":
            flush()
            cur = {"section": cur_section, "stem": text, "options": [], "detail": []}
        elif typ == "option-row" and cur is not None:
            cur["options"].extend(b.get("items", []))
        elif typ == "option" and cur is not None:
            cur["options"].append(text)
        elif typ == "text" and cur is not None:
            cur["detail"].append(text)
    flush()
    return qs


def assign_labels(qs: list[dict]) -> None:
    counters: dict[str, int] = {}
    for i, q in enumerate(qs, 1):
        sec = q.get("section") or "?"
        counters[sec] = counters.get(sec, 0) + 1
        q["label"] = f"{sec}、{counters[sec]}"
        q["index"] = i


def parse_answer_md(path: Path) -> tuple[dict[str, list[tuple[int, str]]], list[dict]]:
    """复用答案册格式：客观题速查表 + 主观题参考答案/评分要点。"""
    txt = path.read_text(encoding="utf-8")
    quick: dict[str, list[tuple[int, str]]] = {}
    for typ in ("单选题", "多选题", "判断题"):
        m = re.search(rf"^### {typ}[^\n]*$(.*?)(?=^### |^## |\Z)", txt, re.S | re.M)
        rows: list[tuple[int, str]] = []
        for line in (m.group(1).splitlines() if m else []):
            line = line.strip()
            if not line.startswith("|") or re.match(r"^\|\s*-", line):
                continue
            cells = [c.strip() for c in line.strip("|").split("|")]
            if cells and cells[0] == "题号":
                continue
            for i in range(0, len(cells) - 1, 2):
                if cells[i].isdigit() and cells[i + 1]:
                    rows.append((int(cells[i]), cells[i + 1]))
        quick[typ] = rows

    marks = list(re.finditer(
        r"^(### (四、简答题|五、综合应用题).*|#### 第 (\d+) 题（(U\d+)，(\d+) 分）.*)$",
        txt, re.M))
    subj: list[dict] = []
    kind = None
    for i, m in enumerate(marks):
        if m.group(2):
            kind = "简答" if m.group(2).startswith("四") else "综合应用"
            continue
        end = marks[i + 1].start() if i + 1 < len(marks) else len(txt)
        body = txt[m.end():end]
        am = re.search(r"\*\*参考答案\*\*(.*?)\*\*评分要点\*\*", body, re.S)
        rm = re.search(r"\*\*评分要点\*\*(.*)", body, re.S)
        if not am or not rm:
            raise ValueError(f"答案册缺少参考答案/评分要点：第 {m.group(3)} 题")
        subj.append({
            "kind": kind,
            "n": int(m.group(3)),
            "unit": m.group(4),
            "score": int(m.group(5)),
            "answer": [l.strip() for l in am.group(1).splitlines() if l.strip()],
            "rubric": [l.strip() for l in rm.group(1).splitlines() if l.strip()],
        })
    return quick, subj


def correct_option_text(q: dict, ans: str) -> str:
    opts = q.get("options", [])
    if not opts:
        return ""
    pieces = []
    for ch in ans.upper():
        idx = ord(ch) - 65
        if 0 <= idx < len(opts):
            pieces.append(OPT_PREFIX_RE.sub("", opts[idx]).strip())
    return "；".join(pieces)


def objective_claim(q: dict, ans: str) -> str:
    """把客观题转成一句“如果前面答对就等于说了什么”的话。"""
    if q.get("options"):
        return f"{q['stem']}{correct_option_text(q, ans)}"
    mark = "正确" if ans in ("√", "正确", "对") else "错误"
    return f"{q['stem']}{mark}"


def finding(kind: str, severity: str, title: str, a: str, b: str,
            score: float, evidence: str, suggestion: str) -> dict:
    return {
        "kind": kind,
        "severity": severity,
        "title": title,
        "a": a,
        "b": b,
        "score": round(score, 3),
        "evidence": evidence,
        "suggestion": suggestion,
    }


def check_duplicates(obj_claims: list[dict], subj: list[dict]) -> list[dict]:
    out = []
    for i in range(len(obj_claims)):
        for j in range(i + 1, len(obj_claims)):
            a, b = obj_claims[i], obj_claims[j]
            r = ratio(a["stem"], b["stem"])
            cc, chunks = concept_coverage(a["stem"], b["stem"], minimum=4)
            if norm(a["stem"]) and norm(a["stem"]) == norm(b["stem"]):
                out.append(finding("卷内重复", "error", "客观题题干完全相同",
                                   a["label"], b["label"], 1.0, a["stem"],
                                   "删除其中一题，换成未覆盖单元的新题。"))
            elif r >= DUP_STEM_RATIO or (cc >= 0.28 and chunks):
                out.append(finding("卷内重复", "warning", "客观题题干考点高度重合",
                                   a["label"], b["label"], max(r, cc),
                                   f"{a['stem']}｜{b['stem']}｜重合片段：{'、'.join(chunks[:4])}",
                                   "核对是否换汤不换药；若考点相同，合并为一题或改成不同考点。"))
    for i in range(len(subj)):
        for j in range(i + 1, len(subj)):
            a, b = subj[i], subj[j]
            r = ratio(" ".join(a["answer"]), " ".join(b["answer"]))
            if r >= DUP_STEM_RATIO:
                out.append(finding("卷内重复", "warning", "主观题参考答案高度相似",
                                   f"主观第{a['n']}题", f"主观第{b['n']}题", r,
                                   f"{a['answer'][0]}｜{b['answer'][0]}",
                                   "两题考查同一套要点；改写其中一题的情境或命题角度。"))
    return out


def check_leakage(obj_claims: list[dict], subj: list[dict]) -> list[dict]:
    out = []
    for obj in obj_claims:
        claim = obj["claim"]
        if len(norm(claim)) < MIN_CLAIM_LEN:
            continue
        for s in subj:
            later = " ".join([s.get("stem", "")] + s["answer"] + s["rubric"])
            r = ratio(claim, later)
            c = containment(claim, later)
            cc, chunks = concept_coverage(claim, later, minimum=4)
            if r >= LEAK_RATIO or c >= LEAK_CONTAIN or (cc >= 0.25 and chunks):
                score = max(r, c, cc)
                sev = "error" if r >= 0.72 or c >= 0.70 or cc >= 0.40 else "warning"
                out.append(finding(
                    "前后泄题", sev, "前面客观题把后面主观题答案先说了",
                    obj["label"], f"主观第{s['n']}题", score,
                    f"前题：{claim}｜后题答案：{' '.join(s['answer'])[:120]}｜重合片段：{'、'.join(chunks[:4])}",
                    "删除或改写前题的正确选项/结论性表述；或把后题换成不同考点/不同情境，避免背一套答案通吃。",
                ))
    return out


def check_coverage(blueprint: dict | None, all_text: str) -> dict:
    if not blueprint:
        return {"status": "skipped", "reason": "未提供 --blueprint，无法自动判定覆盖完整性"}
    units = blueprint.get("units", [])
    if not units:
        return {"status": "skipped", "reason": "blueprint 中没有 units"}
    rows = []
    for u in units:
        hits = [k for k in u.get("keywords", []) if norm(k) and norm(k) in norm(all_text)]
        rows.append({
            "id": u.get("id", ""),
            "name": u.get("name", ""),
            "keywords": u.get("keywords", []),
            "hits": hits,
            "covered": bool(hits),
        })
    missing = [r for r in rows if not r["covered"]]
    return {
        "status": "checked",
        "units": rows,
        "missing": [r["id"] for r in missing],
        "findings": [finding(
            "覆盖缺口", "error", "单元未被卷面覆盖",
            r["id"], r["name"], 0.0, "、".join(r["keywords"]),
            f"在重复/泄题题位补一道 {r['name']} 题，或替换该单元已有的低质量题。",
        ) for r in missing],
    }


def is_slogan_label(text: str) -> bool:
    """宽泛动词 + 宽泛名词构成的口号式代称，例如“借用已有能力”。"""
    return bool(SLOGAN_VERB_RE.search(text) and GENERIC_NOUN_RE.search(text))


def check_clarity(qs: list[dict]) -> list[dict]:
    """主观题必须自包含：不能让考生猜题干里自造的口号或指代。"""
    out: list[dict] = []
    for q in qs:
        if q.get("section") not in ("四", "五"):
            continue
        stem = q.get("stem", "")
        slogans = [x for x in QUOTED_RE.findall(stem) if is_slogan_label(x)]
        if slogans and EXPLAIN_RE.search(stem) and not DEFINE_RE.search(stem):
            two_ways = bool(TWO_WAYS_RE.search(stem))
            out.append(finding(
                "题干不清", "error" if two_ways else "warning",
                "用未定义的口号式代称要求说明/比较",
                q.get("label", "?"), "—", 0.9 if two_ways else 0.6,
                f"题干：{stem}｜未定义代称：{'、'.join(slogans)}",
                "把引号里的代称改成标准术语，或在题干中先给出定义/例子；同时明确要求回答的维度。",
            ))
        vague = VAGUE_REF_RE.findall(stem)
        if vague and not DEFINE_RE.search(stem):
            out.append(finding(
                "题干不清", "warning", "指代词缺少题干内先行词",
                q.get("label", "?"), "—", 0.5,
                f"题干：{stem}｜指代：{'、'.join(sorted(set(vague)))}",
                "把“上述/该做法/这两条”等换成明确对象，或在同题内先说明它指什么。",
            ))
    return out


# ---- 来源亲和性校验：防止拿错课程的 objective.json（§13.8）----
# objective_claim() 只取 objective 的 answer 字母，题干/选项却来自 spec。
# 若 objective 是另一门课的题库（题数恰好相同），错字母会索引到错误选项，
# 产出"看似有证据"的虚假泄题报告。objective 标准格式自带 stem 字段却一直闲置，
# 这里用它做同卷校验，错配则 fail-fast。
SOURCE_ALIGN_RATIO = 0.45   # 单题干判定"对齐"的相似度
SOURCE_ALIGN_PASS = 0.80    # 至少 80% 对齐才算同卷
SOURCE_ALIGN_FAIL = 0.50    # 低于 50% 判定错配，直接中止


def check_source_alignment(qs: list[dict], objective: list[dict]) -> dict:
    """校验 objective.json 与 spec 是否来自同一份卷。

    返回 {"verdict": pass|warn|fail, ...}。verdict=fail 时必须中止，
    不得继续产出泄题/重复等结论（那些结论建立在错配之上，不可信）。
    """
    n = min(len(qs), len(objective))
    if n == 0:
        return {"verdict": "warn", "ratio": 0.0, "aligned": 0, "total": 0,
                "samples": [], "note": "无题可校验"}
    scores: list[float] = []
    samples: list[dict] = []
    has_stem = 0
    for i in range(n):
        o_stem = str(objective[i].get("stem", "") or "").strip()
        s_stem = str(qs[i].get("stem", "") or "").strip()
        if o_stem:
            has_stem += 1
        r = ratio(o_stem, s_stem) if (o_stem and s_stem) else 0.0
        scores.append(r)
        if i < 2:
            samples.append({"index": i + 1, "spec_stem": s_stem[:60],
                            "objective_stem": o_stem[:60], "ratio": round(r, 3)})
    if has_stem == 0:
        return {"verdict": "warn", "ratio": 0.0, "aligned": 0, "total": n,
                "samples": samples,
                "note": "objective.json 全部缺 stem，无法校验来源；请用带 stem 的标准格式"}
    avg = sum(scores) / len(scores)
    aligned = sum(1 for r in scores if r >= SOURCE_ALIGN_RATIO)
    frac = aligned / len(scores)
    verdict = ("pass" if frac >= SOURCE_ALIGN_PASS else
               "fail" if frac < SOURCE_ALIGN_FAIL else "warn")
    note = ("objective 与 spec 题干对齐，可继续" if verdict == "pass" else
            "疑似跨课程/跨卷错配：objective 题干与 spec 大范围不符，结论不可信" if verdict == "fail" else
            "部分题干不一致，请人工确认 objective 与 spec 是否同卷")
    return {"verdict": verdict, "ratio": round(avg, 3), "aligned": aligned,
            "total": len(scores), "frac": round(frac, 3), "samples": samples,
            "note": note}


def markdown_report(report: dict) -> str:
    lines = ["# 试卷内容质量门禁报告", ""]
    lines.append(f"- 结论：**{'通过' if report['passed'] else '未通过'}**")
    lines.append(f"- 卷内重复：{len(report['duplicates'])} 条；前后泄题：{len(report['leakage'])} 条")
    lines.append(f"- 题干清晰性：{len(report['clarity'])} 条")
    lines.append(f"- 覆盖检查：{report['coverage']['status']}")
    lines += ["", "## 需要处理的问题", ""]
    items = (report["duplicates"] + report["leakage"] + report["clarity"]
             + report["coverage"].get("findings", []))
    if not items:
        lines.append("无。")
    else:
        lines += ["| 类型 | 等级 | 位置 | 问题 | 证据 | 怎么办 |",
                  "|---|---|---|---|---|---|"]
        for f in items:
            ev = f["evidence"].replace("|", "\\|")
            lines.append(
                f"| {f['kind']} | {f['severity']} | {f['a']} ↔ {f['b']} | "
                f"{f['title']}（{f['score']:.2f}） | {ev[:100]} | {f['suggestion']} |")
    if report["coverage"]["status"] == "checked":
        lines += ["", "## 单元覆盖", "", "| 单元 | 名称 | 命中关键词 | 覆盖 |", "|---|---|---|---|"]
        for r in report["coverage"]["units"]:
            lines.append(f"| {r['id']} | {r['name']} | {'、'.join(r['hits']) or '—'} | "
                         f"{'是' if r['covered'] else '否'} |")
    return "\n".join(lines) + "\n"


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("spec", type=Path, help="学生卷 spec.json")
    ap.add_argument("objective", type=Path, help="客观题答案 JSON")
    ap.add_argument("answer_md", type=Path, help="参考答案与评分标准 Markdown")
    ap.add_argument("--blueprint", type=Path, help="单元/关键词 blueprint.json（可选）")
    ap.add_argument("--json-out", type=Path, help="JSON 报告输出路径")
    ap.add_argument("--md-out", type=Path, help="Markdown 报告输出路径")
    ap.add_argument("--gate-mode", choices=["full", "alignment"], default="full",
                    help="full=内容 error 也阻塞(默认); alignment=只把来源对齐当硬门禁,内容 error 降级为 warn 不阻塞(供 pipeline 用,见 pitfalls §13.7/§15)")
    args = ap.parse_args(argv)

    spec = json.loads(args.spec.read_text(encoding="utf-8"))
    objective = json.loads(args.objective.read_text(encoding="utf-8"))
    quick, subj = parse_answer_md(args.answer_md)
    qs = questions_from_spec(spec)
    assign_labels(qs)

    obj_qs = qs[:len(objective)]
    if len(obj_qs) != len(objective):
        raise ValueError(f"spec 题目数 {len(qs)} 与客观题答案数 {len(objective)} 不匹配")

    # §13.8 来源亲和性：题数相同≠同一份卷。先确认 objective 的题干与 spec 对得上，
    # 否则 answer 字母索引到的是另一门课的选项，后面所有结论都是虚假的。
    align = check_source_alignment(qs, objective)
    if align["verdict"] == "fail":
        detail = json.dumps({"passed": False, "source_alignment": align,
                             "error": "objective.json 与题干 spec 疑似不是同一份卷，已中止门禁；"
                                      "请核对来源（课程/试卷/A-B 卷）后重跑，勿采信任何泄题结论"},
                            ensure_ascii=False, indent=2)
        print(detail)
        if args.json_out:
            args.json_out.write_text(detail + "\n", encoding="utf-8")
        return 2

    obj_claims = []
    for i, item in enumerate(objective):
        q = obj_qs[i]
        ans = str(item.get("answer", "")).strip()
        obj_claims.append({
            "label": q["label"],
            "stem": q["stem"],
            "claim": objective_claim(q, ans),
        })

    subj_qs = [q for q in qs if q.get("section") in ("四", "五")]
    for i, s in enumerate(subj):
        s["stem"] = subj_qs[i]["stem"] if i < len(subj_qs) else ""

    duplicates = check_duplicates(obj_claims, subj)
    leakage = check_leakage(obj_claims, subj)
    clarity = check_clarity(qs)
    all_text = "\n".join(
        [b.get("text", "") for b in flatten_blocks(spec)]
        + [" ".join(b.get("items", [])) for b in flatten_blocks(spec)]
        + [s["stem"] + " " + " ".join(s["answer"] + s["rubric"]) for s in subj])
    blueprint = json.loads(args.blueprint.read_text(encoding="utf-8")) if args.blueprint else None
    coverage = check_coverage(blueprint, all_text)

    errors = [f for f in duplicates + leakage + clarity + coverage.get("findings", [])
              if f["severity"] == "error"]
    report = {
        "passed": not errors,
        "source_alignment": align,
        "duplicates": duplicates,
        "leakage": leakage,
        "clarity": clarity,
        "coverage": coverage,
        "thresholds": {
            "dup_stem_ratio": DUP_STEM_RATIO,
            "leak_ratio": LEAK_RATIO,
            "leak_containment": LEAK_CONTAIN,
            "ngram": NGRAM,
        },
    }
    payload = json.dumps(report, ensure_ascii=False, indent=2)
    print(payload)
    if args.json_out:
        args.json_out.write_text(payload + "\n", encoding="utf-8")
    if args.md_out:
        args.md_out.write_text(markdown_report(report), encoding="utf-8")
    if args.gate_mode == "alignment":
        # 硬门禁只认来源对齐（跨课程/跨卷错配才是致命的）。内容类 error 多为
        # 「同素材既出客观题又出主观题」等已知误报（§13.7），产出报告供人工
        # 复核即可，不阻塞流水线；否则每次命题设计都需要先消除这些 warning。
        print("[gate-mode=alignment] 内容类 error 已降级为 warn，详见报告；来源对齐为硬门禁。")
        return 0
    return 0 if report["passed"] else 1


if __name__ == "__main__":
    sys.exit(main())
