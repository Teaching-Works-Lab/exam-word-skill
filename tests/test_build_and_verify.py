import hashlib
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from zipfile import ZipFile

from docx import Document

from conftest import ROOT, load_script


build_from_spec = load_script("build_from_spec")
verify_docx = load_script("verify_docx")


PRESERVED_PARTS = {
    "word/styles.xml",
    "word/numbering.xml",
    "word/theme/theme1.xml",
    "word/footer1.xml",
    "word/fontTable.xml",
    "word/footnotes.xml",
    "word/endnotes.xml",
}


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def part_digests(path: Path):
    with ZipFile(path) as archive:
        return {
            part: hashlib.sha256(archive.read(part)).hexdigest()
            for part in PRESERVED_PARTS
            if part in archive.namelist()
        }


class BuildAndVerifyTests(unittest.TestCase):
    def test_builds_from_reference_without_mutating_template(self):
        with TemporaryDirectory() as directory:
            temp_path = Path(directory)
            template = ROOT / "assets" / "reference.docx"
            before = digest(template)
            spec_path = temp_path / "试卷结构.json"
            output = temp_path / "标准试卷.docx"
            spec_path.write_text(
                json.dumps(
                    {
                        "metadata": {
                            "school": "沧州交通学院",
                            "academic_term": "2025－2026学年第二学期",
                            "unit": "计算机与信息技术学院",
                            "teacher": "测试教师",
                            "course": "测试课程",
                            "paper": "A 卷",
                            "exam_mode": "闭卷",
                            "summary": "本试卷共有一道大题（满分：100 分）",
                        },
                        "pages": [
                            {
                                "blocks": [
                                    {"type": "heading", "text": "一、测试题（满分100分，共1题）"},
                                    {"type": "question", "text": "1、这是测试题目。（100分）"},
                                ]
                            }
                        ],
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )

            result = build_from_spec.build_document(template, spec_path, output)

            self.assertTrue(output.exists())
            self.assertEqual(digest(template), before)
            self.assertEqual(part_digests(output), part_digests(template))
            self.assertEqual(result["output"], str(output))
            self.assertIn("这是测试题目", "\n".join(p.text for p in Document(output).paragraphs))

            verification = verify_docx.verify_document(output, template)
            self.assertTrue(verification["valid"])
            self.assertTrue(verification["preserved_parts_match"])


    def test_verifier_rejects_missing_docx(self):
        with TemporaryDirectory() as directory:
            template = ROOT / "assets" / "reference.docx"
            result = verify_docx.verify_document(Path(directory) / "不存在.docx", template)
            self.assertFalse(result["valid"])
            self.assertIn("output_missing", result["errors"])

    def test_score_table_clears_unspecified_template_cells(self):
        with TemporaryDirectory() as directory:
            temp_path = Path(directory)
            template = ROOT / "assets" / "reference.docx"
            spec_path = temp_path / "spec.json"
            output = temp_path / "output.docx"
            spec_path.write_text(
                json.dumps(
                    {
                        "metadata": {
                            "school": "沧州交通学院",
                            "academic_term": "2025－2026学年第二学期",
                            "unit": "计算机与信息技术学院",
                            "teacher": "测试教师",
                            "course": "测试课程",
                            "summary": "本试卷共有两道大题（满分：100 分）",
                        },
                        "score_table": [
                            ["题号", "一", "二", "总分"],
                            ["分数", "40", "60", "100"],
                            ["得分", "", "", ""],
                        ],
                        "pages": [{"blocks": [{"type": "text", "text": "测试"}]}],
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )

            build_from_spec.build_document(template, spec_path, output)
            table = Document(output).tables[0]

            self.assertEqual([cell.text for cell in table.rows[0].cells][4:], ["", "", "", ""])

    def test_paragraph_flow_controls_are_applied(self):
        """Regression: keep_next / keep_together / indents must survive into the DOCX."""
        with TemporaryDirectory() as directory:
            temp_path = Path(directory)
            template = ROOT / "assets" / "reference.docx"
            spec_path = temp_path / "spec.json"
            output = temp_path / "output.docx"
            spec_path.write_text(
                json.dumps(
                    {
                        "metadata": {
                            "school": "测试学校",
                            "academic_term": "2025－2026学年第二学期",
                            "unit": "测试学院",
                            "teacher": "测试教师",
                            "course": "测试课程",
                            "summary": "本试卷共有一道大题（满分：100 分）",
                        },
                        "pages": [
                            {
                                "blocks": [
                                    {"type": "heading", "text": "一、测试题"},
                                    {
                                        "type": "question",
                                        "text": "1、题干",
                                        "keep_next": True,
                                        "keep_together": True,
                                        "left_twips": 425,
                                        "first_twips": -142,
                                    },
                                ]
                            }
                        ],
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            build_from_spec.build_document(template, spec_path, output)
            paragraph = [p for p in Document(output).paragraphs if p.text.strip() == "1、题干"][0]
            fmt = paragraph.paragraph_format
            self.assertTrue(fmt.keep_with_next)
            self.assertTrue(fmt.keep_together)
            self.assertEqual(fmt.left_indent.twips, 425)
            self.assertEqual(fmt.first_line_indent.twips, -142)

    def test_score_table_rows_are_blank_for_new_papers(self):
        """Regression: 分数/得分 rows are grading space, not printed scores."""
        with TemporaryDirectory() as directory:
            temp_path = Path(directory)
            template = ROOT / "assets" / "reference.docx"
            spec_path = temp_path / "spec.json"
            output = temp_path / "output.docx"
            spec_path.write_text(
                json.dumps(
                    {
                        "metadata": {
                            "school": "测试学校",
                            "academic_term": "2025－2026学年第二学期",
                            "unit": "测试学院",
                            "teacher": "测试教师",
                            "course": "测试课程",
                            "summary": "本试卷共有一道大题（满分：100 分）",
                        },
                        "score_table": [
                            ["题号", "一", "二", "三", "四", "五", "", "总分"],
                            ["分数", "", "", "", "", "", "", ""],
                            ["得分", "", "", "", "", "", "", ""],
                        ],
                        "pages": [{"blocks": [{"type": "text", "text": "测试"}]}],
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            build_from_spec.build_document(template, spec_path, output)
            table = Document(output).tables[0]
            for row in table.rows[1:]:
                self.assertEqual([c.text for c in row.cells][1:], [""] * 7)

    def test_skill_points_to_authoring_pitfalls(self):
        """Regression: the authoring checklist stays wired into the Skill entry."""
        skill = (ROOT / "SKILL.md").read_text(encoding="utf-8")
        self.assertIn("references/exam-authoring-pitfalls.md", skill)
        self.assertTrue((ROOT / "references" / "exam-authoring-pitfalls.md").exists())


if __name__ == "__main__":
    unittest.main()
