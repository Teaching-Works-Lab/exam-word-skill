# -*- coding: utf-8 -*-
"""客观题答案分布均衡门禁（exam-word-skill，跨课程通用）。

为什么需要
----------
命题大忌：答案集中在单一选项。真实踩坑（2026-11-24 C&C++Ⅰ 期末 A 卷）：
单选题 10 题里 B 出现 6 次、A 仅 1 次、D 为 0；多选题 5 题答案**全是 ABC**
（干扰项一律放在 D）。学生不看题也能猜中，属命题事故。教师明确要求：
「答案全部是一个选项」「判断题全部对/全部错」都必须避免，且与本卷现状无关——
这是**任何试卷都要守的通用规则**。

判据（全部可量化，参数化，换课程零改动）
----------------------------------------
单选题 (single)：
  * 单一字母出现次数 <= single_max_ratio * n（默认 0.4，即不超过 40%）
  * 最长连续同一答案 <= max_run（默认 2）
多选题 (multiple)：
  * 完全相同的答案组合出现次数 <= combo_max_repeat（默认 2）
  * 干扰项（每题唯一错误选项）所在字母的分散度：不能全部堆在同一字母
    （多选题每题 3 项正确，正确项字母频次天然偏高，故不用字母占比判据，
     改用「干扰项位置」判据 —— 这才是 ABC×5 的真正病根）
判断题 (judge)：
  * 「对」占比落在 [judge_min_true_ratio, judge_max_true_ratio]（默认 0.3~0.7）
    -> 排除全对、全错、9:1 等极端

可迁移设计
----------
* 只读答案数据（answers.json 的 objective 段，或 objective.json），不依赖题干文本；
* 章节名支持「一/二/三」与「single/multiple/judge」两组别名，可经命令行映射；
* 所有阈值参数化；模型无关、课程无关、语言无关；
* 可独立运行，也可 ``from check_answer_distribution import audit`` 被组卷流程调用。

用法
----
    python check_answer_distribution.py answers.json
    python check_answer_distribution.py answers.json --json report.json
    python check_answer_distribution.py answers.json --single-max-ratio 0.35 --max-run 2

answers.json 结构::
    {"paper": "A 卷",
     "objective": {"一": ["B","B",...], "二": ["ABC",...], "三": ["√",...]}}

退出码：0 = 全部达标；1 = 存在分布事故；2 = 未识别到答案数据。
"""
import argparse
import json
import sys
from collections import Counter
from pathlib import Path

# 章节别名 -> 题型。顺序即检查顺序。
TYPE_ALIASES = {
    "single": {"一", "1", "single", "单选", "单选题", "选择题"},
    "multiple": {"二", "2", "multiple", "多选", "多选题"},
    "judge": {"三", "3", "judge", "判断", "判断题", "对错"},
}


def classify_key(key):
    for typ, aliases in TYPE_ALIASES.items():
        if str(key).strip() in aliases:
            return typ
    return None


def longest_run(seq):
    best = cur = 0
    prev = object()
    for x in seq:
        cur = cur + 1 if x == prev else 1
        best = max(best, cur)
        prev = x
    return best


def normalize(a):
    """统一答案记号：对/错 -> √/×，去空格。"""
    s = str(a).strip().upper().replace(" ", "")
    return {"TRUE": "√", "T": "√", "对": "√", "正确": "√", "YES": "√", "Y": "√",
            "FALSE": "×", "F": "×", "错": "×", "错误": "×", "NO": "×", "N": "×"}.get(s, s)


def audit(answers, *, single_max_ratio=0.4, max_run=2,
          combo_max_repeat=2, distractor_max_ratio=0.6,
          judge_min_true_ratio=0.3, judge_max_true_ratio=0.7):
    """审计答案分布，返回结构化报告 dict。"""
    obj = answers.get("objective", answers)
    grouped = {}
    for key, seq in obj.items():
        typ = classify_key(key)
        if typ:
            grouped.setdefault(typ, []).extend(normalize(a) for a in seq)

    if not grouped:
        return {"verdict": "NO_DATA", "checks": [],
                "message": "未识别到客观题答案，检查 answers.json 的 objective 段或章节名"}

    checks = []
    def add(typ, name, ok, detail):
        checks.append({"type": typ, "check": name, "ok": bool(ok), "detail": detail})

    # —— 单选题 ——
    if "single" in grouped:
        seq = grouped["single"]
        n = len(seq)
        cnt = Counter(seq)
        top_letter, top_n = (cnt.most_common(1)[0] if cnt else ("-", 0))
        add("single", "单一选项不过载", top_n <= single_max_ratio * n + 1e-9,
            f"最多={top_letter}×{top_n}/{n}，上限={single_max_ratio:.0%}({single_max_ratio*n:.1f})；分布={dict(sorted(cnt.items()))}")
        run = longest_run(seq)
        add("single", "无连续同答案", run <= max_run,
            f"最长连续={run}，上限={max_run}；序列={' '.join(seq)}")

    # —— 多选题 ——
    if "multiple" in grouped:
        seq = grouped["multiple"]
        n = len(seq)
        combos = Counter("".join(sorted(x)) for x in seq)
        top_combo, top_n = (combos.most_common(1)[0] if combos else ("-", 0))
        add("multiple", "答案组合不重复", top_n <= combo_max_repeat,
            f"最多组合={top_combo}×{top_n}/{n}，上限={combo_max_repeat}；组合={dict(combos)}")
        # 干扰项 = 每题唯一不在答案里的选项字母（4 选 3 的多选题）
        ALL = set("ABCD")
        distractors = ["".join(sorted(ALL - set(x))) for x in seq]
        dcnt = Counter(distractors)
        dl, dn = (dcnt.most_common(1)[0] if dcnt else ("-", 0))
        add("multiple", "干扰项位置分散", dn <= distractor_max_ratio * n + 1e-9,
            f"干扰项最多={dl}×{dn}/{n}，上限={distractor_max_ratio:.0%}({distractor_max_ratio*n:.1f})；干扰项分布={dict(dcnt)}")

    # —— 判断题 ——
    if "judge" in grouped:
        seq = grouped["judge"]
        n = len(seq)
        true_n = sum(1 for x in seq if x == "√")
        ratio = true_n / n if n else 0
        ok = judge_min_true_ratio - 1e-9 <= ratio <= judge_max_true_ratio + 1e-9
        add("judge", "对错比例均衡", ok,
            f"√={true_n}/{n}={ratio:.0%}，允许[{judge_min_true_ratio:.0%},{judge_max_true_ratio:.0%}]；序列={' '.join(seq)}")

    verdict = "PASS" if all(c["ok"] for c in checks) else "FAIL"
    return {"verdict": verdict, "checks": checks,
            "counts": {k: len(v) for k, v in grouped.items()}}


def format_report(rep):
    lines = []
    for c in rep["checks"]:
        flag = "✔" if c["ok"] else "✘"
        lines.append(f"  {flag} [{c['type']}] {c['check']}: {c['detail']}")
    lines.append("")
    lines.append("结论：" + rep.get("verdict", "-"))
    if rep.get("message"):
        lines.append(rep["message"])
    return "\n".join(lines)


def main(argv=None):
    ap = argparse.ArgumentParser(description="客观题答案分布均衡门禁")
    ap.add_argument("answers", help="answers.json（含 objective 段）")
    ap.add_argument("--json", dest="json_out", help="报告写成 JSON")
    ap.add_argument("--single-max-ratio", type=float, default=0.4)
    ap.add_argument("--max-run", type=int, default=2)
    ap.add_argument("--combo-max-repeat", type=int, default=2)
    ap.add_argument("--distractor-max-ratio", type=float, default=0.6)
    ap.add_argument("--judge-min-true-ratio", type=float, default=0.3)
    ap.add_argument("--judge-max-true-ratio", type=float, default=0.7)
    args = ap.parse_args(argv)

    data = json.loads(Path(args.answers).read_text(encoding="utf-8-sig"))
    rep = audit(data, single_max_ratio=args.single_max_ratio, max_run=args.max_run,
                combo_max_repeat=args.combo_max_repeat,
                distractor_max_ratio=args.distractor_max_ratio,
                judge_min_true_ratio=args.judge_min_true_ratio,
                judge_max_true_ratio=args.judge_max_true_ratio)
    print("文件：" + Path(args.answers).name)
    print(format_report(rep))
    if args.json_out:
        Path(args.json_out).write_text(json.dumps(rep, ensure_ascii=False, indent=2), encoding="utf-8")
    if rep["verdict"] == "NO_DATA":
        return 2
    return 0 if rep["verdict"] == "PASS" else 1


if __name__ == "__main__":
    sys.exit(main())
