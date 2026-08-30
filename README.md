# Exam Word Skill

一个面向高校试卷的 Codex Skill，用于把不同学院、不同课程的 Word/PDF 试卷统一到标准模板，并稳定完成格式检查、Word 对比、格式修正和 PDF 转可编辑 Word。

显式触发名称：`$exam-word`

它属于 [Teaching Works Lab 课程教学 Skill 体系](https://github.com/Teaching-Works-Lab)，负责试卷 Word 的对比与规范化；课程考核设计由 `course-teaching-workflows` 负责，课程考核资料归档由 `course-assessment-archive-skill` 负责。

## 能做什么

- **单个 Word 改格式**：保留题目内容，按标准模板生成新的修正版 DOCX。
- **单个 Word 格式检查**：只报告格式问题，不修改原文件。
- **两个 Word 对比**：分别检查内容差异与格式差异。
- **对比后修正**：输出差异报告，并生成新的修正版 DOCX。
- **PDF 转试卷 Word**：提取试卷内容，再按模板生成可编辑 DOCX。

适用于程序设计、机械、数学、英语等课程。学院或课程名称不决定是否触发；只要任务明确涉及试卷及受支持的操作即可。

## 快速使用

把试卷文件附加到 Codex，然后直接输入：

```text
$exam-word 帮我把这份程序设计试卷按标准模板改格式，不改题目内容，另存新文件。
```

也可以使用自然语言自动触发：

```text
检查这份试卷 Word 是否符合标准格式。
对比这两个试卷 Word 的内容和格式差异。
对比后，把第二份试卷按第一份的格式修正。
把这份试卷 PDF 转成可编辑的标准 Word。
```

## 工作模式

| 模式 | 典型输入 | 默认结果 |
|---|---|---|
| `inspect` | 一个 Word，要求检查或意图不明确 | 只读检查报告 |
| `normalize` | 一个 Word，要求改格式、套模板或重新排版 | 新的修正版 DOCX |
| `compare` | 两个 Word，要求对比或查找差异 | JSON 与 Markdown 差异报告 |
| `compare-and-fix` | 两个 Word，要求对比并修正 | 差异报告和新的修正版 DOCX |
| `pdf-to-word` | PDF，要求转换、可编辑或套模板 | 新的可编辑 DOCX |

路由由 `scripts/route_request.py` 根据请求文本和附件类型确定。意图不明确时默认进入只读检查，不擅自修改文件。

## 核心规则

1. **原文件不覆盖**：所有修改结果另存为新文件。
2. **内容与格式分离**：题目、代码、公式、图片、分值和课程信息来自源文件；模板只控制页面、字体、表格、页眉页脚和分页。
3. **模板优先级明确**：用户提供的指定模板优先；否则使用 `assets/reference.docx`。
4. **结果必须验证**：生成后进行 DOCX 结构检查，并渲染检查全部页面。
5. **隐私最小化**：仓库和 Skill 不保存原始试卷、学生信息、生成结果或 OCR/渲染缓存。

扫描件、复杂公式、浮动图片和特殊分页可能需要 OCR 或人工复核；遇到内容不清、分值冲突或模板不唯一时，Skill 会停止猜测并说明问题。

## 安装

推荐先添加组织 Marketplace，再选择安装本 Plugin：

```text
codex plugin marketplace add Teaching-Works-Lab/.github
codex plugin add exam-word-skill@teaching-works-lab
```

也可以继续按独立 Skill 方式安装：

在 PowerShell 中克隆到 Codex Skills 目录：

```powershell
git clone https://github.com/Teaching-Works-Lab/exam-word-skill.git "$env:USERPROFILE\.codex\skills\exam-word"
```

重新启动 Codex 后，即可使用 `$exam-word`。主要脚本依赖 Python 3.12 和 `python-docx`；Windows 上建议使用 Microsoft Word 完成最终高保真渲染检查。

## 目录说明

```text
exam-word/
├── SKILL.md                     # 触发条件、路由和执行规则
├── agents/openai.yaml           # Codex 展示名称与默认提示
├── assets/reference.docx        # 内置标准试卷模板
├── references/                  # 模式与模板契约
├── scripts/                     # 路由、对比、构建和验证脚本
└── tests/                       # 回归测试与脱敏测试夹具
```

## 开发验证

```powershell
py -3.12 -m unittest discover -s tests -p "test_*.py" -v
py -3.12 scripts/route_request.py "帮我把这份试卷改成标准格式" "试卷.docx"
```

第二条命令应输出 `normalize`。更详细的执行约束见 [`SKILL.md`](SKILL.md)，模式输入输出见 [`references/modes.md`](references/modes.md)。
