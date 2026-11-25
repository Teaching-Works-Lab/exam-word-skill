import argparse
import json
from pathlib import Path


COMPARE_TERMS = ("对比", "比较", "差异", "diff", "compare")
INSPECT_TERMS = ("检查", "审查", "核对", "inspect", "audit")
MODIFY_TERMS = (
    "修正",
    "修改",
    "改格式",
    "改成标准格式",
    "改为标准格式",
    "调整格式",
    "规范格式",
    "重新排版",
    "套用",
    "套版",
    "统一",
    "标准化",
    "整理",
    "fix",
    "normalize",
    "apply",
)
CONVERT_TERMS = ("转换", "转成", "转为", "可编辑", "convert")

# 「说答案」的默认出口：卷面标注答案卷（§11）。
ANSWER_TERMS = (
    "答案", "标准答案", "参考答案", "评分标准", "出答案", "给答案",
    "answer", "answer key", "answer-keyed",
)
# 教师答案册是显式例外，不能靠默认猜出来。
BOOKLET_TERMS = (
    "教师用", "教师卷", "答案册", "评分细则", "答案速查",
    "对照表", "rubric", "booklet",
)


def _contains(text: str, terms: tuple[str, ...]) -> bool:
    lowered = text.casefold()
    return any(term.casefold() in lowered for term in terms)


def route_request(request: str, files: list[str]) -> str:
    suffixes = [Path(file).suffix.casefold() for file in files]
    docx_count = suffixes.count(".docx")
    pdf_count = suffixes.count(".pdf")

    wants_compare = _contains(request, COMPARE_TERMS)
    wants_modify = _contains(request, MODIFY_TERMS)
    wants_convert = _contains(request, CONVERT_TERMS)
    wants_answer = _contains(request, ANSWER_TERMS)
    wants_booklet = _contains(request, BOOKLET_TERMS)

    # 文件操作类意图优先，避免「对比两份答案」「把答案改成标准格式」被误判成创作。
    if pdf_count:
        if wants_convert or wants_modify or not docx_count:
            return "pdf-to-word"
        return "inspect"
    if docx_count >= 2:
        return "compare-and-fix" if wants_modify else "compare"

    # 答案创作：默认卷面标注答案卷；只有显式教师用才走答案册。
    if wants_answer and not wants_modify and not wants_compare:
        return "answer-booklet" if wants_booklet else "answer-keyed-paper"

    if wants_modify:
        return "normalize"
    if wants_compare or _contains(request, INSPECT_TERMS):
        return "inspect"
    if not docx_count and not pdf_count:
        raise ValueError("At least one DOCX or PDF input is required for non-answer requests")
    return "inspect"


def main() -> None:
    parser = argparse.ArgumentParser(description="Route an exam Word/PDF request")
    parser.add_argument("request")
    parser.add_argument("files", nargs="*")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()
    mode = route_request(args.request, args.files)
    if args.json:
        print(json.dumps({"mode": mode}, ensure_ascii=False))
    else:
        print(mode)


if __name__ == "__main__":
    main()
