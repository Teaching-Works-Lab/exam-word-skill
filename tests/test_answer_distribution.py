# -*- coding: utf-8 -*-
"""§14 答案分布均衡门禁的回归测试（跨课程通用判据）。"""
from conftest import load_script

gate = load_script("check_answer_distribution")


def test_concentrated_single_and_abc_multi_fails():
    """A 卷旧分布：单选 B×6 + 多选全 ABC -> FAIL。"""
    answers = {"objective": {
        "一": ["B", "B", "B", "A", "C", "B", "C", "B", "B", "C"],
        "二": ["ABC", "ABC", "ABC", "ABC", "ABC"],
        "三": ["√", "×", "√", "√", "√", "×", "√", "√", "√", "×"],
    }}
    rep = gate.audit(answers)
    assert rep["verdict"] == "FAIL"
    failed = {c["check"] for c in rep["checks"] if not c["ok"]}
    assert "单一选项不过载" in failed
    assert "无连续同答案" in failed
    assert "答案组合不重复" in failed
    assert "干扰项位置分散" in failed


def test_balanced_distribution_passes():
    """B 卷分布：单选均衡、多选 4 种组合、判断 4:6 -> PASS。"""
    answers = {"objective": {
        "一": ["A", "A", "C", "D", "B", "C", "D", "A", "D", "B"],
        "二": ["ABD", "ACD", "BCD", "ABD", "ACD"],
        "三": ["×", "√", "×", "×", "√", "√", "×", "×", "×", "√"],
    }}
    rep = gate.audit(answers)
    assert rep["verdict"] == "PASS", rep["checks"]


def test_fixed_a_after_rebalance_passes():
    """A 卷修复后：选项位置调换、内容不变 -> PASS。"""
    answers = {"objective": {
        "一": ["C", "A", "D", "B", "C", "A", "D", "B", "C", "A"],
        "二": ["BCD", "ACD", "ABD", "ABC", "ACD"],
        "三": ["√", "×", "√", "√", "√", "×", "√", "√", "√", "×"],
    }}
    rep = gate.audit(answers)
    assert rep["verdict"] == "PASS", rep["checks"]


def test_all_true_judge_fails():
    """判断题全对 -> FAIL。"""
    answers = {"objective": {
        "三": ["√", "√", "√", "√", "√", "√", "√", "√", "√", "√"],
    }}
    rep = gate.audit(answers)
    assert rep["verdict"] == "FAIL"
    judge = [c for c in rep["checks"] if c["type"] == "judge"][0]
    assert judge["ok"] is False


def test_alias_and_normalization():
    """支持 single/multiple/judge 别名与 对/错、T/F 记号。"""
    answers = {"objective": {
        "single": ["A", "B", "C", "D"],
        "multiple": ["ABD", "ACD"],
        "judge": ["对", "错", "T", "F"],
    }}
    rep = gate.audit(answers)
    assert set(rep["counts"]) == {"single", "multiple", "judge"}
    assert rep["verdict"] == "PASS", rep["checks"]


def test_no_data_returns_no_data():
    rep = gate.audit({"objective": {"主观": ["要点"]}})
    assert rep["verdict"] == "NO_DATA"
