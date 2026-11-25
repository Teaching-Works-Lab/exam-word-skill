# -*- coding: utf-8 -*-
"""§11 卷面标注答案卷回归测试：路由 → 生成 → 门禁 → 模板保真。

    python3 -m unittest discover -s tests -v
"""
import json
import sys
import tempfile
import unittest
from pathlib import Path

SKILL = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SKILL / "scripts"))

import build_answer_keyed_paper as bkp  # noqa: E402
import build_from_spec as bfs  # noqa: E402
import route_request as rr  # noqa: E402
import verify_answer_keyed_paper as vak  # noqa: E402
import verify_docx  # noqa: E402

TEMPLATE = SKILL / "assets" / "reference.docx"
WORKSPACE = Path("/Users/aa/WorkBuddy/2026-11-24-06-42-58")


class RouterTests(unittest.TestCase):
    def test_answer_alone_is_keyed_paper(self):
        self.assertEqual(rr.route_request("出一下答案", []), "answer-keyed-paper")
        self.assertEqual(rr.route_request("把参考答案给我", []), "answer-keyed-paper")
        self.assertEqual(rr.route_request("标准答案", []), "answer-keyed-paper")
        self.assertEqual(rr.route_request("answer key", []), "answer-keyed-paper")

    def test_booklet_is_explicit_opt_in(self):
        self.assertEqual(rr.route_request("教师用答案册", []), "answer-booklet")
        self.assertEqual(rr.route_request("参考答案与评分细则，教师用", []), "answer-booklet")
        self.assertEqual(rr.route_request("A/B 卷答案对照表", []), "answer-booklet")

    def test_file_operations_win(self):
        self.assertEqual(rr.route_request("对比两份答案", ["a.docx", "b.docx"]), "compare")
        self.assertEqual(rr.route_request("把答案改成标准格式", ["a.docx"]), "normalize")
        self.assertEqual(rr.route_request("转成可编辑", ["a.pdf"]), "pdf-to-word")

    def test_legacy_modes_unchanged(self):
        self.assertEqual(rr.route_request("对比", ["a.docx", "b.docx"]), "compare")
        self.assertEqual(rr.route_request("修正并对比", ["a.docx", "b.docx"]), "compare-and-fix")
        self.assertEqual(rr.route_request("检查", ["a.docx"]), "inspect")
        self.assertEqual(rr.route_request("随便看看", ["a.docx"]), "inspect")

    def test_non_answer_without_files_raises(self):
        with self.assertRaises(ValueError):
            rr.route_request("看看这份卷子", [])


class MarksTests(unittest.TestCase):
    def test_numbered_rubric(self):
        rubric = ["（1）4 分：a 1、b 1。", "（2）4 分：c。",
                  "（3）4 分：d。", "（4）3 分：e。",
                  "只罗列设备清单者总分不超过 8 分。"]
        self.assertEqual(bkp.parse_marks(rubric, 4, 15, "综合1"), [4, 4, 4, 3])

    def test_semicolon_rubric_ignores_deductions(self):
        rubric = ["状态信号 3 分（≥5 项）；定位方法 4 分（须体现分层）；"
                  "复测闭环 3 分。只罗列名称扣 2 分。"]
        self.assertEqual(bkp.parse_marks(rubric, 3, 10, "简答1"), [3, 4, 3])

    def test_circled_rubric(self):
        self.assertEqual(bkp.parse_marks(["①2 分 ②2 分 ③2 分 ④2 分 ⑤2 分"], 5, 10, "简答2"),
                         [2, 2, 2, 2, 2])

    def test_unparseable_rubric_raises(self):
        with self.assertRaises(ValueError):
            bkp.parse_marks(["随便写几句，没有分值"], 3, 10, "简答1")


class FormatTests(unittest.TestCase):
    def test_bracket_full_and_half_width(self):
        self.assertEqual(bkp.fill_bracket("题干（ ）", "B", "q1"), "题干（B）")
        self.assertEqual(bkp.fill_bracket("题干?( )", "ABC", "q2"), "题干?(ABC)")
        self.assertEqual(bkp.fill_bracket("A（说明）题干（ ）", "√", "q3"),
                         "A（说明）题干（√）")
        with self.assertRaises(ValueError):
            bkp.fill_bracket("没有括号", "B", "q4")

    def test_mark_placement(self):
        self.assertEqual(bkp.format_point("【状态信号】主轴负载电流…", 3),
                         "【状态信号】（3 分）主轴负载电流…")
        self.assertEqual(bkp.format_point("（1）先把零件归并…", 4),
                         "（1）（4 分）先把零件归并…")
        self.assertTrue(bkp.format_point("普通文本", 2).startswith("（2 分）"))

    def test_objective_cross_check(self):
        objective = [{"type": "单选题", "stem": "题干（ ）", "answer": "B"},
                     {"type": "判断题", "stem": "判断题干( )", "answer": "正确"}]
        quick = {"单选题": [(1, "B")], "多选题": [], "判断题": [(1, "正确")]}
        self.assertEqual(bkp.check_objective(objective, quick), 2)
        with self.assertRaises(ValueError):
            bkp.check_objective(objective, {**quick, "单选题": [(1, "C")]})


MINI_SPEC = {
    "metadata": {
        "school": "测试学院", "academic_term": "2026－2027学年第一学期",
        "unit": "机械学院", "teacher": "命题教师", "course": "智能制造装备",
        "paper": "B 卷", "exam_mode": "闭卷",
        "summary": "本试卷共有五道大题（满分：100 分；考试时间：120 分钟）",
    },
    "score_table": [
        ["题号", "一", "二", "三", "四", "五", "", "总分"],
        ["分数", "", "", "", "", "", "", ""],
        ["得分", "", "", "", "", "", "", ""],
    ],
    "pages": [
        {"blocks": [
            {"type": "heading", "text": "一、单项选择题（共 1 题，每题 2 分，共 2 分）"},
            {"type": "question", "text": "1、题干（ ）"},
            {"type": "option", "text": "A. 选项一"},
            {"type": "option", "text": "B. 选项二"},
            {"type": "heading", "text": "三、判断题（共 1 题，每题 1 分，共 1 分）"},
            {"type": "question", "text": "1、判断题干( )"},
        ]},
        {"blocks": [
            {"type": "heading", "text": "四、简答题（共 1 题，每题 5 分，共 5 分）"},
            {"type": "question", "text": "1. 主观题干"},
            {"type": "question", "text": "（1）小问一；",
             "left_twips": 425, "first_twips": -142},
            {"type": "spacer", "after_pt": 190},
        ]},
    ],
}
MINI_OBJECTIVE = [
    {"type": "单选题", "stem": "题干（ ）", "answer": "B"},
    {"type": "判断题", "stem": "判断题干( )", "answer": "错误"},
]
MINI_MD = """# 测试卷 · 参考答案与评分标准

## 一、客观题答案速查

### 单选题

| 题号 | 答案 |
|---|---|
| 1 | B |

### 多选题

| 题号 | 答案 |
|---|---|

### 判断题

| 题号 | 答案 |
|---|---|
| 1 | 错误 |

## 二、主观题参考答案与评分要点

### 四、简答题（共 1 题，每题 5 分，共 5 分）

#### 第 1 题（U2，5 分）

**参考答案**

【要点一】内容一。
（2）内容二。

**评分要点**

（1）3 分：a 1、b 1。
（2）2 分：c 1、d 1。
"""


class BuildTests(unittest.TestCase):
    def test_end_to_end_mini_spec(self):
        with tempfile.TemporaryDirectory() as tmp:
            t = Path(tmp)
            (t / "spec.json").write_text(json.dumps(MINI_SPEC, ensure_ascii=False), encoding="utf-8")
            (t / "objective.json").write_text(
                json.dumps(MINI_OBJECTIVE, ensure_ascii=False), encoding="utf-8")
            (t / "answer.md").write_text(MINI_MD, encoding="utf-8")

            spec, stats = bkp.build_keyed_spec(
                t / "spec.json", t / "objective.json", t / "answer.md")
            self.assertEqual(spec["metadata"]["answer_label"], "答案")
            self.assertEqual(spec["metadata"]["doc_type"], "answer-keyed")
            self.assertEqual(len(spec["pages"]), 1, "派生 spec 应为单页自然流动")
            self.assertEqual(stats["objective_filled"], 2)
            self.assertEqual(stats["hard_page_breaks"], 0)
            self.assertEqual([s["marks"] for s in stats["subjective"]], [[3, 2]])
            blocks = spec["pages"][0]["blocks"]
            self.assertFalse(any(b["type"] == "spacer" for b in blocks))
            self.assertEqual(sum(1 for b in blocks
                                 if b.get("text") == bkp.RUBRIC_LABEL), 1)
            self.assertIn("（B）", blocks[1]["text"])
            self.assertIn("(×)", blocks[5]["text"])

            (t / "keyed.json").write_text(
                json.dumps(spec, ensure_ascii=False), encoding="utf-8")
            out = t / "keyed.docx"
            bfs.build_document(TEMPLATE, t / "keyed.json", out)
            self.assertTrue(out.exists() and out.stat().st_size > 0)
            report = verify_docx.verify_document(out, TEMPLATE)
            self.assertTrue(report["valid"], report["errors"])

    @unittest.skipUnless(
        (WORKSPACE / ".workbuddy/exam-spec-B.json").exists(),
        "workspace B-paper data not present")
    def test_full_b_paper_gate(self):
        """端到端：真实 B 卷数据重建一份，跑 11 项独立门禁。"""
        data = WORKSPACE / ".workbuddy"
        with tempfile.TemporaryDirectory() as tmp:
            t = Path(tmp)
            spec, stats = bkp.build_keyed_spec(
                data / "exam-spec-B.json",
                data / "tmp-paper-B-objective.json",
                WORKSPACE / "智能制造装备-期末笔试B卷-参考答案与评分标准.md")
            (t / "keyed.json").write_text(
                json.dumps(spec, ensure_ascii=False), encoding="utf-8")
            out = t / "b-keyed.docx"
            bfs.build_document(TEMPLATE, t / "keyed.json", out)
            self.assertEqual(stats["objective_filled"], 18)
            self.assertEqual([s["marks"] for s in stats["subjective"]],
                             [[3, 4, 3], [2, 3, 2, 3], [3, 3, 2, 2],
                              [3, 3, 2, 2], [4, 4, 4, 3], [4, 4, 4, 3]])
            report = vak.run(out)
            self.assertTrue(report["valid"], report["errors"])
            # 门禁项随结构性修复扩展（11 → 15+），断言关键项在且全绿，不写死数量
            for key in ("answer_label", "student_header", "score_table_blank",
                        "objective_answers", "rubric_labels", "mark_sum",
                        "no_answer_space", "no_hard_page_break", "no_ai_marks"):
                self.assertIn(key, report["checks"])
                self.assertTrue(report["checks"][key]["ok"], key)
            fidelity = verify_docx.verify_document(out, TEMPLATE)
            self.assertTrue(fidelity["valid"], fidelity["errors"])


FULLWIDTH_SPEC = {
    "metadata": {
        "school": "测试学院", "academic_term": "2026－2027学年第一学期",
        "unit": "机械学院", "teacher": "命题教师", "course": "人工智能技术及应用",
        "paper": "A 卷", "exam_mode": "闭卷",
        "summary": "本试卷共有四道大题（满分：100 分）",
    },
    "score_table": [
        ["题号", "一", "三", "四", "五", "", "总分"],
        ["分数", "", "", "", "", "", ""],
        ["得分", "", "", "", "", "", ""],
    ],
    "pages": [
        {"blocks": [
            {"type": "heading", "text": "一、单项选择题（共 1 小题，每小题 1 分，共 1 分）"},
            {"type": "question", "text": "1．图灵测试主要用于判断（　）。"},
            {"type": "option", "text": "A. 运算速度"},
            {"type": "option", "text": "B. 智能行为"},
            {"type": "heading", "text": "三、判断题（共 1 小题，每小题 1 分，共 1 分）"},
            {"type": "question", "text": "2．模型输出无需人工核实。（　）"},
        ]},
        {"blocks": [
            {"type": "heading", "text": "四、简答题（共 1 小题，每小题 6 分，共 6 分）"},
            {"type": "question", "text": "1．简述三大流派。"},
            {"type": "spacer", "text": "", "after_pt": 5},
            # 故意漏掉应用题的答题空白，验证生成器自动补齐
            {"type": "heading", "text": "五、应用题（共 1 小题，每小题 10 分，共 10 分）"},
            {"type": "question", "text": "1．写一条提示词。"},
        ]},
    ],
}
FULLWIDTH_OBJECTIVE = [
    {"type": "单选题", "stem": "1．图灵测试主要用于判断（　）。", "answer": "B"},
    {"type": "判断题", "stem": "2．模型输出无需人工核实。（　）", "answer": "错误"},
]
FULLWIDTH_MD = """# 测试卷 · 参考答案与评分标准

## 一、客观题答案速查

### 单选题

| 题号 | 答案 |
|---|---|
| 1 | B |

### 多选题

| 题号 | 答案 |
|---|---|

### 判断题

| 题号 | 答案 |
|---|---|
| 1 | × |

## 二、主观题参考答案与评分要点

### 四、简答题（共 1 题，每题 6 分）

#### 第 1 题（U1，6 分）

**参考答案**

【要点一】符号主义与连接主义。
要点二（2分）行为主义。

**评分要点**

（1）3 分：两种流派。
（2）3 分：行为主义。

### 五、综合应用题（共 1 题，每题 10 分）

#### 第 1 题（U1，10 分）

**参考答案**

【角色设定】明确教师角色。
【校验要求】标注待核。

**评分要点**

（1）4 分：角色设定。
（2）6 分：校验要求。
"""


class PaperTypeAdaptationTests(unittest.TestCase):
    """异卷型（全角题号 / 25+9 / 每题分值不同）不得触发门禁误报。"""

    def test_fullwidth_numbering_and_missing_spacer(self):
        with tempfile.TemporaryDirectory() as tmp:
            t = Path(tmp)
            (t / "spec.json").write_text(
                json.dumps(FULLWIDTH_SPEC, ensure_ascii=False), encoding="utf-8")
            (t / "objective.json").write_text(
                json.dumps(FULLWIDTH_OBJECTIVE, ensure_ascii=False), encoding="utf-8")
            (t / "answer.md").write_text(FULLWIDTH_MD, encoding="utf-8")

            spec, stats = bkp.build_keyed_spec(
                t / "spec.json", t / "objective.json", t / "answer.md")
            self.assertEqual(stats["objective_filled"], 2)
            self.assertEqual(stats["spacers_added"], 1)
            blocks = spec["pages"][0]["blocks"]
            self.assertFalse(any(b["type"] == "spacer" for b in blocks))
            self.assertFalse(any("（2分）" in b.get("text", "") for b in blocks))
            self.assertIn("（B）", blocks[1]["text"])

            (t / "keyed.json").write_text(
                json.dumps(spec, ensure_ascii=False), encoding="utf-8")
            out = t / "keyed.docx"
            bfs.build_document(TEMPLATE, t / "keyed.json", out)
            report = vak.run(out, t / "keyed.json")
            self.assertTrue(report["valid"], report["errors"])
            for key in ("answer_label", "objective_answers", "rubric_labels",
                        "mark_format", "mark_sum", "no_answer_space"):
                self.assertIn(key, report["checks"])
                self.assertTrue(report["checks"][key]["ok"], key)
            self.assertEqual(report["objective_count"], 2)
            self.assertEqual(report["subjective_count"], 2)
            self.assertTrue(verify_docx.verify_document(out, TEMPLATE)["valid"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
