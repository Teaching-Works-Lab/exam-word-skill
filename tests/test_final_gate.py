# -*- coding: utf-8 -*-
"""check_final_gate.py 与 clean_markdown 回归测试（不依赖 pytest）。"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))

from derive_answer_keyed_from_docx import clean_markdown  # noqa: E402
import check_final_gate as gate  # noqa: E402


def test_clean_backtick():
    assert clean_markdown("输出：`sum=10`") == "输出：sum=10"


def test_clean_multiple_backticks():
    assert clean_markdown("`x << 1` 左移，`x >> 1` 右移") == "x << 1 左移，x >> 1 右移"


def test_clean_bold_and_bullet():
    assert clean_markdown("- （2 分）**输出：9**") == "（2 分）输出：9"


def test_clean_html_escape():
    assert clean_markdown("若 a &lt; b 则 a &amp; 1") == "若 a < b 则 a & 1"


def test_clean_preserves_c_code_text():
    # 自然语言里出现的 C 表达式不应被破坏
    s = "s[1] = s[0] 是结构体整体赋值"
    assert clean_markdown(s) == s


def test_gate_detects_backtick():
    """门禁正则确实命中反引号（单元级，不造 docx）。"""
    assert gate.BACKTICK.search("输出：`sum=10`")
    assert not gate.BACKTICK.search("输出：sum=10")


def test_gate_detects_html_and_ai():
    assert gate.HTML_ESC.search("a &lt; b")
    assert gate.AI_STAMP.search("作为一个AI，我为你生成了试卷")
    assert not gate.HTML_ESC.search("a < b")


def test_gate_checks_list_complete():
    assert set(gate.CHECKS) == {"no_backtick", "no_html_escape",
                                "no_md_bold", "no_bullet", "no_ai_stamp"}


if __name__ == "__main__":
    fns = [n for n in dir() if n.startswith("test_") and callable(globals()[n])]
    ok = 0
    for n in fns:
        try:
            globals()[n]()
            print("PASS", n)
            ok += 1
        except AssertionError as e:
            print("FAIL", n, e)
    print(f"{ok}/{len(fns)} passed")
    sys.exit(0 if ok == len(fns) else 1)
