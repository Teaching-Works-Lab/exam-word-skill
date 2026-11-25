#!/usr/bin/env python3
"""检查试卷 Markdown 中「运行下列代码」代码块是否为完整可编译 C 程序。

用法：python scripts/check_code_completeness.py paper.md [paper2.md ...]

豁免（不算片段）：题干明确问「关于下面这个函数」的函数定义题；
「查错与测试」题中故意给出的病态程序。
"""
import re
import sys
from pathlib import Path

RUN_PAT = re.compile(r"运行下列代码|执行以下代码|输出结果是|输出是")
EXEMPT = re.compile(r"关于下面这个函数|关于函数|查错|有\s*\d+\s*处\s*bug|找出.*错误")


def check(path):
    s = Path(path).read_text(encoding="utf-8")
    bad = []
    total = 0
    for m in re.finditer(r"```c\n(.*?)\n```", s, re.S):
        code = m.group(1)
        pre = s[max(0, m.start() - 260):m.start()]
        q = pre.strip().splitlines()
        stem = q[-1] if q else ""
        if not RUN_PAT.search(stem) or EXEMPT.search(pre[-400:]):
            continue
        total += 1
        missing = []
        if "#include" not in code:
            missing.append("#include")
        if not re.search(r"\bmain\s*\(", code):
            missing.append("main()")
        if "return 0" not in code:
            missing.append("return 0")
        if code.count("{") != code.count("}"):
            missing.append("braces")
        if missing:
            bad.append((stem[:48], missing))
    return total, bad


if __name__ == "__main__":
    rc = 0
    for arg in sys.argv[1:]:
        total, bad = check(arg)
        print(f"{Path(arg).name}: 运行类代码块 {total} 个，不完整 {len(bad)} 个")
        for stem, missing in bad:
            print(f"  FAIL {stem} -> 缺 {', '.join(missing)}")
        rc |= 1 if bad else 0
    sys.exit(rc)
