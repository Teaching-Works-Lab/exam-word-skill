# -*- coding: utf-8 -*-
"""最终交付强制门禁（§15.6 长约束 / 最后一道保险）。

定位：**独立于生成时清理**。生成脚本负责"产出时 clean"，本门禁负责
"交付前 verify"——直接读成品 DOCX 的卷面文本，任何 Markdown/HTML 残留或
AI 痕迹都让整条流水线 fail-fast。这样即使换了生成路径、历史文件、人工
粘贴，只要卷面带残留就过不了关。

检查项（全部从成品 DOCX 客观读取，不依赖生成过程）
------------------------------------------------
1. ``no_backtick``      卷面无反引号 `（Markdown 行内代码标记误印）
2. ``no_html_escape``   无 &lt; &gt; &amp; &quot; &#39; 等 HTML 转义残留
3. ``no_md_bold``       无 ** / __ 强调标记残留
4. ``no_bullet``        自然语言段落无行首 "- " / "* " 列表标记
5. ``no_ai_stamp``      无"作为AI/以下是/希望这些对你/Ok, here"等 AI 痕迹
空段落/空白页由 check_pdf_pages.py 单独把关（表格布局需要空段，此处不查）。

豁免：C 代码段落（含 ``#include`` / ``int main`` / ``w:br`` 软换行的段）
不参与 1-5 检查——代码里的 ``*``、``/* */``、``#`` 都是合法 C 语法。
代码排版由 check_code_layout.py 单独把关，职责不重叠。

用法
----
    python check_final_gate.py paper.docx
    python check_final_gate.py paper-answer.docx --json report.json
    # 批量
    for f in *.docx; do python check_final_gate.py "$f" || exit 1; done

退出码：0=PASS；1=发现残留（FAIL）；2=文件不可读/无内容。
"""
import argparse
import json
import re
import sys
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

BACKTICK = re.compile(r"`")
HTML_ESC = re.compile(r"&(?:lt|gt|amp|quot|nbsp|apos|#\d+);", re.I)
MD_BOLD = re.compile(r"\*\*|__")
BULLET = re.compile(r"^\s*[-+*]\s+\S")
AI_STAMP = re.compile(
    r"(作为(?:一个)?AI|我是AI|以下(?:是|为)[^。\n]{0,12}(?:试卷|答案|内容)|"
    r"希望(?:这|以上)[^。\n]{0,10}(?:帮助|有用)|祝你|Here(?:'s| is)|Ok[, ]|"
    r"注[：:]?\s*本(?:试卷|答案)[^。\n]{0,20}(?:由|经)AI)", re.I)

CHECKS = ("no_backtick", "no_html_escape", "no_md_bold",
          "no_bullet", "no_ai_stamp")


def is_code_para(p):
    """代码段豁免：含软换行或典型 C 语法标记。"""
    if any(True for _ in p.iter(qn("w:br"))):
        return True
    t = "".join(x.text or "" for x in p.iter(qn("w:t")))
    return any(k in t for k in ("#include", "int main", "#define", "printf(", "scanf("))


def audit(docx_path):
    doc = Document(str(docx_path))
    findings = {k: [] for k in CHECKS}
    for idx, p in enumerate(doc.element.body.iter(qn("w:p"))):
        text = "".join(t.text or "" for t in p.iter(qn("w:t")))
        if not text.strip():
            continue
        if is_code_para(p):
            continue
        for name, pat in (("no_backtick", BACKTICK), ("no_html_escape", HTML_ESC),
                          ("no_md_bold", MD_BOLD), ("no_ai_stamp", AI_STAMP)):
            if pat.search(text):
                findings[name].append({"para": idx, "text": text[:90]})
        if BULLET.match(text):
            findings["no_bullet"].append({"para": idx, "text": text[:90]})

    checks = {k: {"pass": len(v) == 0, "count": len(v), "samples": v[:6]}
              for k, v in findings.items()}
    passed = all(c["pass"] for c in checks.values())
    return {"file": str(docx_path), "passed": passed,
            "verdict": "PASS" if passed else "FAIL", "checks": checks}


def main(argv=None):
    ap = argparse.ArgumentParser(description="最终交付强制门禁：卷面 Markdown/HTML/AI 残留扫描")
    ap.add_argument("docx", nargs="+", help="待检查的 DOCX（可多个）")
    ap.add_argument("--json", dest="json_out", help="JSON 报告输出路径")
    args = ap.parse_args(argv)

    reports = []
    for f in args.docx:
        try:
            reports.append(audit(f))
        except Exception as e:  # noqa: BLE001
            reports.append({"file": f, "passed": False, "verdict": "ERROR",
                            "error": f"{type(e).__name__}: {e}", "checks": {}})

    all_pass = all(r["passed"] for r in reports)
    print("=" * 64)
    print("最终交付强制门禁（§15.6 长约束）")
    print("=" * 64)
    for r in reports:
        print(f"\n[{r['verdict']}] {Path(r['file']).name}")
        if r["verdict"] == "ERROR":
            print("   ", r["error"])
            continue
        for name, c in r["checks"].items():
            mark = "✔" if c["pass"] else "✘"
            print(f"  {mark} {name:18s} {c['count']} 处")
            for s in c["samples"][:3]:
                print(f"        p{s['para']}: {s['text']}")
    print("\n结论：" + ("PASS — 全部卷面干净" if all_pass else "FAIL — 存在残留，禁止交付"))

    if args.json_out:
        Path(args.json_out).write_text(
            json.dumps(reports, ensure_ascii=False, indent=2), encoding="utf-8")
    return 0 if all_pass else 1


if __name__ == "__main__":
    sys.exit(main())
