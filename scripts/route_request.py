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


def _contains(text: str, terms: tuple[str, ...]) -> bool:
    lowered = text.casefold()
    return any(term.casefold() in lowered for term in terms)


def route_request(request: str, files: list[str]) -> str:
    suffixes = [Path(file).suffix.casefold() for file in files]
    docx_count = suffixes.count(".docx")
    pdf_count = suffixes.count(".pdf")
    if not docx_count and not pdf_count:
        raise ValueError("At least one DOCX or PDF input is required")

    wants_compare = _contains(request, COMPARE_TERMS)
    wants_modify = _contains(request, MODIFY_TERMS)
    wants_convert = _contains(request, CONVERT_TERMS)

    if pdf_count:
        if wants_convert or wants_modify or not docx_count:
            return "pdf-to-word"
        return "inspect"
    if docx_count >= 2:
        return "compare-and-fix" if wants_modify else "compare"
    if wants_modify:
        return "normalize"
    if wants_compare or _contains(request, INSPECT_TERMS):
        return "inspect"
    return "inspect"


def main() -> None:
    parser = argparse.ArgumentParser(description="Route an exam Word/PDF request")
    parser.add_argument("request")
    parser.add_argument("files", nargs="+")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()
    mode = route_request(args.request, args.files)
    if args.json:
        print(json.dumps({"mode": mode}, ensure_ascii=False))
    else:
        print(mode)


if __name__ == "__main__":
    main()
