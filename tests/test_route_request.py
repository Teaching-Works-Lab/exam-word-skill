import unittest

from conftest import load_script


route_request = load_script("route_request").route_request


class RouteRequestTests(unittest.TestCase):
    def test_routes_supported_exam_requests(self):
        cases = [
            ("对比这两个 Word 的内容和格式差异", ["标准.docx", "试卷.docx"], "compare"),
            ("检查这份 Word 是否符合标准模板", ["试卷.docx"], "inspect"),
            ("检查这份试卷 Word 是否符合标准格式", ["试卷.docx"], "inspect"),
            ("按标准模板修正并统一这份 Word", ["试卷.docx"], "normalize"),
            ("帮我改格式，不改题目内容", ["试卷.docx"], "normalize"),
            ("把这份计算机学院程序设计试卷改成标准格式", ["程序设计试卷.docx"], "normalize"),
            ("把高等数学试卷 PDF 转成可编辑 Word", ["高等数学试卷.pdf"], "pdf-to-word"),
            ("把这个 PDF 转成可编辑的标准试卷 Word", ["试卷.pdf"], "pdf-to-word"),
            ("对比这两个 Word 并按模板修正", ["标准.docx", "试卷.docx"], "compare-and-fix"),
            ("帮我处理一下这个文件", ["试卷.docx"], "inspect"),
        ]
        for request, files, expected in cases:
            with self.subTest(request=request):
                self.assertEqual(route_request(request, files), expected)


    def test_rejects_unsupported_inputs(self):
        with self.assertRaisesRegex(ValueError, "DOCX or PDF"):
            route_request("处理文件", ["notes.txt"])


if __name__ == "__main__":
    unittest.main()
