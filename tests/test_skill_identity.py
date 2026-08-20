import re
import unittest
from pathlib import Path


SKILL_ROOT = Path(__file__).resolve().parents[1]


class SkillIdentityTests(unittest.TestCase):
    def test_uses_general_exam_word_identity(self):
        skill_text = (SKILL_ROOT / "SKILL.md").read_text(encoding="utf-8")
        agent_text = (SKILL_ROOT / "agents" / "openai.yaml").read_text(encoding="utf-8")

        name = re.search(r"^name:\s*(.+)$", skill_text, re.MULTILINE)
        self.assertIsNotNone(name)
        self.assertEqual(name.group(1).strip(), "exam-word")
        self.assertIn("$exam-word", agent_text)
        self.assertNotIn("mechanical", skill_text.casefold())
        self.assertNotIn("机械学院", skill_text)
        self.assertNotIn("机械学院", agent_text)


if __name__ == "__main__":
    unittest.main()
