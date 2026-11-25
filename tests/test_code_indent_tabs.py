# -*- coding: utf-8 -*-
"""代码缩进 tab 化回归测试（v6，防回退到行首空格）。不依赖 pytest。"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))

import build_code_exam as bce  # noqa: E402


def test_level1_indent():
    assert bce._leading_spaces_to_tabs("    int score;") == "\tint score;"


def test_level2_indent():
    assert bce._leading_spaces_to_tabs("        s += i;") == "\t\ts += i;"


def test_top_level_no_tab():
    assert bce._leading_spaces_to_tabs("int main(void) {") == "int main(void) {"


def test_inline_spaces_preserved():
    # 行内注释对齐空格不能被转 tab
    s = "        sum += i;               /* 注释 */"
    out = bce._leading_spaces_to_tabs(s)
    assert out.startswith("\t\tsum += i;")
    assert "               /* 注释 */" in out


def test_remainder_kept():
    # 非 4 倍数余量保留为空格（容错）
    assert bce._leading_spaces_to_tabs("  return 0;") == "  return 0;"


def test_constants():
    assert bce.CODE_LEFT == 480
    assert bce.CODE_TAB_STEP == 440
    assert bce.CODE_INDENT_UNIT == 4


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
