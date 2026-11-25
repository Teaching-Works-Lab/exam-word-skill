# -*- coding: utf-8 -*-
"""卷内内容质量门禁回归测试。

    python3 -m unittest discover -s tests -v
"""
import contextlib
import io
import json
import sys
import tempfile
import unittest
from pathlib import Path

SKILL = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SKILL / "scripts"))

import check_paper_content as cpc  # noqa: E402


def spec_with(obj_questions, subj_questions):
    blocks = [
        {"type": "heading", "text": "一、单项选择题（共 2 题）"},
        {"type": "question", "text": obj_questions[0]},
        {"type": "option", "text": "A. 甲"},
        {"type": "option", "text": "B. 连接主义"},
        {"type": "heading", "text": "二、简答题（共 1 题）"},
    ]
    for s in subj_questions:
        blocks.append({"type": "question", "text": s})
        blocks.append({"type": "spacer", "text": ""})
    return {"pages": [{"blocks": blocks}]}


ANSWER_MD = """# 测试卷 · 参考答案

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

## 二、主观题参考答案与评分要点

### 四、简答题（共 1 题）

#### 第 1 题（U1，6 分）

**参考答案**

连接主义认为智能源于神经网络，靠数据自动学习连接权重，是人工智能三大流派之一。

**评分要点**

（1）3 分：连接主义与神经网络、连接权重。
（2）3 分：属于人工智能流派。
"""

BLUEPRINT = {"units": [
    {"id": "U1", "name": "人工智能基础", "keywords": ["人工智能", "三大流派"]},
    {"id": "U2", "name": "深度学习", "keywords": ["卷积", "池化"]},
]}


class ContentGateTests(unittest.TestCase):
    def test_detects_front_to_back_leak(self):
        qs = spec_with(
            ["1．人工智能三大流派中，主张模拟人脑神经网络结构、通过数据自动学习权重的是（　）。"],
            ["1．简述人工智能三大流派。"])
        with tempfile.TemporaryDirectory() as tmp:
            t = Path(tmp)
            (t / "spec.json").write_text(json.dumps(qs, ensure_ascii=False), encoding="utf-8")
            (t / "obj.json").write_text(json.dumps(
                [{"type": "单选题", "stem": "人工智能三大流派", "answer": "B"}]), encoding="utf-8")
            (t / "answer.md").write_text(ANSWER_MD, encoding="utf-8")
            with contextlib.redirect_stdout(io.StringIO()):
                cpc.main([str(t / "spec.json"), str(t / "obj.json"), str(t / "answer.md"),
                          "--json-out", str(t / "report.json")])
            report = json.loads((t / "report.json").read_text(encoding="utf-8"))
        self.assertTrue(report["leakage"])
        self.assertEqual(report["leakage"][0]["a"], "一、1")

    def test_detects_duplicate_stems(self):
        a = {"label": "一、1", "stem": "1．图灵测试主要用于判断（　）。", "claim": "x"}
        b = {"label": "三、1", "stem": "1．图灵测试通过对话判断机器智能行为。（　）", "claim": "x"}
        out = cpc.check_duplicates([a, b], [])
        self.assertTrue(out)
        self.assertIn("卷内重复", [f["kind"] for f in out])

    def test_detects_coverage_gap(self):
        all_text = "图灵测试 三大流派 人工智能"
        out = cpc.check_coverage(BLUEPRINT, all_text)
        self.assertEqual(out["missing"], ["U2"])
        self.assertFalse(out["findings"][0]["suggestion"] == "")

    def test_clean_paper_passes(self):
        a = {"label": "一、1", "stem": "1．下列哪一项是输出设备？（　）", "claim": "显示器"}
        b = {"label": "三、1", "stem": "1．二进制 1010 转十进制是多少？（　）", "claim": "x"}
        subj = [{"n": 1, "stem": "简述计算机网络的作用", "answer": ["实现资源共享与通信"]}]
        self.assertEqual(cpc.check_duplicates([a, b], subj), [])
        self.assertEqual(cpc.check_leakage([a, b], subj), [])


if __name__ == "__main__":
    unittest.main(verbosity=2)
