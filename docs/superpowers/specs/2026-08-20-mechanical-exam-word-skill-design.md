# 机械学院试卷 Word Skill 设计

## 目标

创建一个可自动发现、也可通过 `$mechanical-exam-word` 显式调用的个人 Skill，稳定处理机械学院标准试卷的 Word 对比、格式检查、模板套用、PDF 转标准 Word 和结果验证。

## 范围

Skill 支持五种模式：

1. `compare`：比较两个 Word 的内容与格式，默认只生成差异报告。
2. `inspect`：将一个 Word 与内置标准模板比较，默认只报告问题。
3. `normalize`：保持题目内容，按内置标准模板生成新的修正版 Word。
4. `pdf-to-word`：从 PDF 提取试卷结构与文本，再按内置模板生成可编辑 Word。
5. `compare-and-fix`：先比较，再生成差异报告与修正版 Word。

不处理无关的普通公文、论文和简历，不覆盖用户原文件，不把扫描页整页作为图片伪装成可编辑 Word。

## 路由

- 用户显式写 `$mechanical-exam-word` 时始终加载 Skill。
- 输入包含 PDF 时，只在用户要求转换或套版时进入 `pdf-to-word`。
- 输入包含两个 Word 且请求含“对比、比较、差异”时进入 `compare`。
- 输入包含一个 Word 且请求含“检查、审查、核对”时进入 `inspect`。
- 请求含“修正、套用、统一、标准化”时进入 `normalize` 或 `compare-and-fix`。
- 意图不清时先只检查，不修改。

## 架构

- `SKILL.md`：触发条件、路由、执行约束和清理规则。
- `assets/reference.docx`：标准模板唯一事实来源。
- `references/template-contract.md`：模板参数和必须保留的 OOXML 部件。
- `references/modes.md`：各模式输入、输出和停止条件。
- `scripts/route_request.py`：根据文件类型和意图给出确定性路由建议。
- `scripts/compare_docx.py`：比较 Word 的页面、样式、表格、页眉页脚、编号与正文结构。
- `scripts/build_from_spec.py`：把结构化试卷 JSON 写入模板副本，保留模板关键部件。
- `scripts/verify_docx.py`：检查 OOXML、页数约束和模板关键部件。
- `tests/`：路由、对比、构建和清理边界回归测试；测试夹具由代码生成，不保存学生数据或原始试卷。

## 数据流

输入文件先被分类并复制到一次性任务目录。Word 对比直接产生 JSON 与 Markdown 报告；Word 标准化先提取结构化内容，再由模板构建器输出新文件；PDF 模式先调用 PDF/OCR 工作流得到结构化试卷 JSON，再调用同一模板构建器。所有修改模式最后必须经过 Word 渲染和结构验证。

## 一致性策略

- 模板文件和模板契约纳入版本控制。
- 所有文档生成都从同一模板副本开始。
- 固定恢复 `styles.xml`、`numbering.xml`、主题、页脚、字体表、脚注和尾注等模板关键部件。
- 使用结构化 JSON 作为 PDF/Word 内容提取与 Word 构建之间的稳定接口。
- 只有最终输出和简短运行清单属于交付物；PNG、渲染 PDF、候选 Word 和 OCR 缓存属于临时文件。

## 清理与隐私

仓库只保留模板、脚本、规则、测试及脱敏夹具。不上传原始试卷 PDF、用户生成的最终 Word、QA 渲染页面、临时依赖目录或学生个人信息。每次运行使用独立任务目录，成功后删除中间产物，失败时保留任务目录并在报告中给出路径供排查。

## 验证

- 路由测试覆盖五种模式和模糊请求的只读回退。
- Word 对比测试能区分内容差异与格式差异。
- 构建测试验证输出为可打开的 DOCX，关键模板部件哈希一致。
- 最终在本机 Word 中渲染代表性输出，确认页码、表格和分页无异常。
- 发布前验证仓库为 Private，且提交清单不含原始试卷和临时文件。
