#!/usr/bin/env python3
"""去掉试卷/答案文档里「子题干」的加粗，保留标题/卷头/答案标签/表头的加粗。

背景（2026-11-24 复盘）：
    应用题/查错题常带 (1)(2)(3)(4) 子问题。生成器若把整段子题干置为
    bold=True，会出现「题干加粗、题目主体不加粗」的不一致。教师明确要求：
    所有题目的题干都不加粗。

判定规则（精准，避免误伤）：
    段落文本（strip 后）匹配「^( 或 （) 数字 ( 或 ）)」的，视为子题干，
    将其所有 run 的 bold 置为 False。其余一律保留原样。

保真契约：
    python-docx 重存会改写 styles/numbering/footer 等部件。本脚本在修改前
    读取「文档自身」的这些部件，保存后原样写回，因此即便文档与 reference.docx
    不一致（例如人工另存过的答案册），也不会被模板覆盖。

用法：
    python unbold_substems.py FILE.docx [FILE.docx ...]
    python unbold_substems.py --drop-junk FILE.docx   # 附带清理卷头残留垃圾 run
"""
import re
import sys
from pathlib import Path
from zipfile import ZipFile

from docx import Document

PRESERVED_PARTS = {
    "word/styles.xml",
    "word/numbering.xml",
    "word/theme/theme1.xml",
    "word/footer1.xml",
    "word/fontTable.xml",
    "word/footnotes.xml",
    "word/endnotes.xml",
}

SUBSTEM_RE = re.compile(r"^\s*[\(（]\s*\d+\s*[\)）]")


def iter_paragraphs(doc):
    """遍历正文段落 + 所有表格单元格内的段落。"""
    for p in doc.paragraphs:
        yield p
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p


def unbold(path: Path, drop_junk: bool = False) -> dict:
    with ZipFile(path) as z:
        own_parts = {
            name: z.read(name)
            for name in PRESERVED_PARTS
            if name in z.namelist()
        }

    doc = Document(path)
    substems = 0
    runs_changed = 0
    junk_removed = 0

    for p in iter_paragraphs(doc):
        if SUBSTEM_RE.match(p.text.strip()):
            substems += 1
            for run in p.runs:
                if run.bold:
                    run.bold = False
                    runs_changed += 1

    if drop_junk and doc.paragraphs:
        first = doc.paragraphs[0]
        for run in list(first.runs):
            if run.text.strip() in {"这个", "这个 "}:
                run._element.getparent().remove(run._element)
                junk_removed += 1

    doc.save(path)

    tmp = path.with_suffix(".unbold.tmp.docx")
    with ZipFile(path) as zin, ZipFile(tmp, "w") as zout:
        for info in zin.infolist():
            zout.writestr(
                info,
                own_parts.get(info.filename, zin.read(info.filename)),
            )
    tmp.replace(path)

    return {
        "file": str(path),
        "substems": substems,
        "runs_unbolded": runs_changed,
        "junk_removed": junk_removed,
        "preserved_parts": len(own_parts),
    }


def main():
    args = sys.argv[1:]
    drop_junk = False
    if args and args[0] == "--drop-junk":
        drop_junk = True
        args = args[1:]
    if not args:
        print(__doc__)
        sys.exit(1)
    import json

    for arg in args:
        report = unbold(Path(arg), drop_junk=drop_junk)
        print(json.dumps(report, ensure_ascii=False))


if __name__ == "__main__":
    main()
