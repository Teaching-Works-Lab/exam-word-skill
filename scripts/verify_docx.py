import argparse
import hashlib
import json
from pathlib import Path
from zipfile import BadZipFile, ZipFile

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
REQUIRED_PARTS = {"[Content_Types].xml", "word/document.xml", "word/settings.xml", "word/styles.xml"}


def _digest(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _length(value):
    return int(value) if value is not None else None


def _page_setup(path: Path):
    document = Document(path)
    attributes = (
        "page_width",
        "page_height",
        "left_margin",
        "right_margin",
        "top_margin",
        "bottom_margin",
        "header_distance",
        "footer_distance",
    )
    return [
        {attribute: _length(getattr(section, attribute)) for attribute in attributes}
        for section in document.sections
    ]


def verify_document(output_path, template_path, rendered_pdf=None, expected_pages=None):
    output_path = Path(output_path)
    template_path = Path(template_path)
    report = {
        "output": str(output_path),
        "template": str(template_path),
        "valid": False,
        "errors": [],
        "warnings": [],
        "preserved_parts_match": False,
    }
    if not output_path.exists():
        report["errors"].append("output_missing")
        return report
    if not template_path.exists():
        report["errors"].append("template_missing")
        return report
    if output_path.stat().st_size == 0:
        report["errors"].append("output_empty")
        return report

    try:
        with ZipFile(output_path) as output_archive, ZipFile(template_path) as template_archive:
            bad_member = output_archive.testzip()
            if bad_member:
                report["errors"].append(f"corrupt_member:{bad_member}")
            output_names = set(output_archive.namelist())
            missing_required = sorted(REQUIRED_PARTS - output_names)
            report["errors"].extend(f"missing_part:{part}" for part in missing_required)
            mismatches = []
            for part in sorted(PRESERVED_PARTS):
                if part not in template_archive.namelist():
                    continue
                if part not in output_names:
                    mismatches.append(f"missing:{part}")
                elif _digest(output_archive.read(part)) != _digest(template_archive.read(part)):
                    mismatches.append(f"changed:{part}")
            report["preserved_part_mismatches"] = mismatches
            report["preserved_parts_match"] = not mismatches
            if mismatches:
                report["errors"].append("preserved_parts_mismatch")
            footer_text = output_archive.read("word/footer1.xml") if "word/footer1.xml" in output_names else b""
            if b"PAGE" not in footer_text or b"NUMPAGES" not in footer_text:
                report["errors"].append("page_fields_missing")
    except BadZipFile:
        report["errors"].append("invalid_docx_package")
        return report

    try:
        if _page_setup(output_path) != _page_setup(template_path):
            report["errors"].append("page_setup_mismatch")
    except Exception as error:
        report["errors"].append(f"document_open_failed:{type(error).__name__}")

    if expected_pages is not None:
        if rendered_pdf is None:
            report["errors"].append("rendered_pdf_required")
        else:
            try:
                from pypdf import PdfReader

                actual_pages = len(PdfReader(str(rendered_pdf)).pages)
                report["rendered_pages"] = actual_pages
                if actual_pages != expected_pages:
                    report["errors"].append(f"page_count:{actual_pages}!={expected_pages}")
            except Exception as error:
                report["errors"].append(f"pdf_check_failed:{type(error).__name__}")

    report["valid"] = not report["errors"]
    return report


def main() -> None:
    parser = argparse.ArgumentParser(description="Verify an exam DOCX against its template")
    parser.add_argument("output")
    parser.add_argument("template")
    parser.add_argument("--rendered-pdf")
    parser.add_argument("--expected-pages", type=int)
    args = parser.parse_args()
    report = verify_document(args.output, args.template, args.rendered_pdf, args.expected_pages)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    raise SystemExit(0 if report["valid"] else 1)


if __name__ == "__main__":
    main()
