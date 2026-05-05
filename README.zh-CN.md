# DOCX Template Translator Skill

一个用于将 **LaTeX、PDF、Markdown 或粗转 DOCX** 转译成指定 Word 模板格式的 Codex skill。

它面向毕业论文、学位论文、单位报告、标准文档等场景，尤其适合那些 **pandoc 默认 DOCX 输出不够用**、必须严格套用学校或机构 Word 模板的任务。

English: [README.md](README.md)

## 为什么不是直接用 pandoc？

pandoc 很适合做第一轮转换，尤其适合 Markdown 和 LaTeX 类源文件。`--reference-doc`
也能复用 Word 样式，但它只知道“样式定义”，不知道模板里的页面和段落语义。

严格模板里经常会遇到这些问题：

- 封面、英文封面、声明页没有正确填写；
- 签名页、授权页、页眉页脚不符合模板；
- 模板样式名具有迷惑性，例如 `Body Text` 实际用于英文封面；
- 正文段落误套成封面样式；
- 图表标题丢失编号；
- 表格没有变成三线表；
- 参考文献引用不是上标，也不能跳转；
- 超链接以蓝色下划线显示，不适合打印版论文；
- 目录、页码、交叉引用字段没有更新；
- 没有自动导出 PDF 进行视觉检查。

这个 skill 的定位不是替代 pandoc，而是把 pandoc、Word 和 Python 组合起来，让 AI
根据用户上传的 Word 模板和源文件，自动写一版 **项目专用的后处理脚本**。

## 现有项目调研

- [pandoc](https://pandoc.org/MANUAL.html) 支持 `--reference-doc`，可以复用
  DOCX 样式。官方文档说明：reference DOCX 的正文内容会被忽略，主要使用其中的样式和文档属性。
- [Quarto DOCX](https://quarto.org/docs/reference/formats/docx.html) 也支持
  `reference-doc`、目录、章节编号、引用和参考文献链接。
- [pdf2docx](https://github.com/ArtifexSoftware/pdf2docx) 可以把原生 PDF 转成
  DOCX；项目文档说明其底层使用 PyMuPDF 解析 PDF、规则化分析版面，并用 `python-docx` 生成 Word。
- [python-docx-template](https://docxtpl.readthedocs.io/) 适合在预制 Word 模板里做 Jinja 风格变量替换。
- [MDDoc](https://www.mddoc.app/) 是商业化 Markdown 转 Word 工具，可以将 Markdown 元素映射到上传的 Word 模板。
- [wmvanvliet/pandoc-tutorial](https://github.com/wmvanvliet/pandoc-tutorial)
  展示了复杂 LaTeX 转 DOCX 时的现实情况：pandoc 能完成大部分转换，但最后的模板适配通常需要自定义处理。
- [openclaw/pdf-to-docx skill](https://playbooks.com/skills/openclaw/skills/pdf-to-docx)
  是一个基于 pdf2docx 的 PDF 转 DOCX skill。

本项目的差异点是：它不是一个固定转换器，而是一个 **开放的 AI 模板重建工作流**。它让 AI 先检查用户上传的 Word 模板，再为该模板生成一版专用 Python 后处理脚本，从而适配学校、单位、期刊等强模板场景。

## 核心思路

1. 把 LaTeX/PDF/Markdown 当作 **内容源**。
2. 把 `.docx` 模板当作 **排版源**。
3. 先用 pandoc、Word 导入或其他工具生成粗略 body DOCX。
4. 检查模板样式、段落、编号、表格、超链接和 OOXML。
5. 让 AI 基于模板检查结果编写或修改 Python 重建脚本。
6. 用 `python-docx` 和底层 OOXML 操作生成最终 DOCX。
7. 用 Word COM 更新目录、页码等字段，并导出 PDF。
8. 把 PDF 渲染成预览拼图，快速检查封面、目录、图表、公式、参考文献。

## 支持输入

- LaTeX 项目：`main.tex`、章节、图片、BibTeX；
- Markdown 文件；
- 原生数字 PDF；
- 已有粗转 Word 文件。

如果有 LaTeX 或 Markdown 源文件，优先使用源文件。PDF 的语义信息较弱，不适合作为首选输入。

## Skill 内容

skill 位于：

```text
skills/docx-template-translator/
```

包含：

- `SKILL.md`：AI 工作流程；
- `scripts/inspect_docx_template.py`：检查 Word 模板结构；
- `scripts/adaptive_docx_pipeline.py`：可改造的 DOCX 重建脚本起点；
- `scripts/finalize_word_docx.py`：通过 Word 更新目录/字段并导出 PDF；
- `scripts/render_pdf_preview.py`：将 PDF 渲染为预览拼图；
- `references/pandoc-limitations.md`：说明和 pandoc 默认功能的区别；
- `references/zhengzhou-case-study.md`：郑州大学毕业论文模板转换案例。

## 安装

将 skill 复制到 Codex skills 目录：

```bash
mkdir -p ~/.codex/skills
cp -r skills/docx-template-translator ~/.codex/skills/
```

然后可以这样调用：

```text
Use $docx-template-translator to convert my LaTeX thesis into Word using this .docx template.
```

## Python 依赖

推荐安装：

```bash
pip install python-docx pywin32 pymupdf pillow
```

可选：

```bash
pip install pdf2docx
```

LaTeX 或 Markdown 转 DOCX 时建议另外安装 pandoc。

最终更新目录和导出 PDF 时，推荐在 Windows 上使用本机 Microsoft Word，因为脚本通过
Word COM 调用 Word 更新字段和导出 PDF。

## 典型流程

1. 用户提供源文件和 Word 模板。
2. Codex 检查模板：

```bash
python scripts/inspect_docx_template.py template.docx --out template_report.json
```

3. Codex 生成粗略 body DOCX：

```bash
pandoc main.tex --citeproc --reference-doc template.docx -o body.docx
```

4. Codex 根据模板和 body 输出，改造 `scripts/adaptive_docx_pipeline.py`。
5. 最终处理：

```bash
python scripts/finalize_word_docx.py final.docx --pdf
python scripts/render_pdf_preview.py final.pdf --pages 1-8
```

## 和 pandoc 默认功能的本质区别

pandoc 是“格式转换器”，这个 skill 是“模板适配工作流”。

pandoc 默认能力通常止步于：把内容转换成 DOCX，并应用参考文档里的样式定义。

这个 skill 的目标是进一步处理：

- 根据模板语义重建封面和声明页；
- 识别模板真正的正文、标题、目录、参考文献样式；
- 自动补图表编号；
- 改三线表；
- 修引用上标和跳转；
- 修超链接颜色；
- 保留公式和图片；
- 用 Word 更新目录和页码；
- 自动导出 PDF 做视觉质检。

## 项目状态

这是一个工作流 skill，不是万能一键转换器。它的核心价值是：每个严格 Word 模板都有自己的局部规则，因此让 AI 先检查模板，再写一版专用 Python 脚本，往往比写一个“通用转换器”更可靠。

## 社区

[LINUX DO — 中文开发者社区](https://linux.do/)

本项目感谢并致谢 LINUX DO 作为中文开发者社区在开源分享与技术讨论方面提供的社区土壤与交流氛围。除非该社区另行明确声明，本致谢不代表 LINUX DO 对本项目的官方背书或认可。

## 许可证

本项目采用 [Apache License 2.0](LICENSE) 开源许可。

Apache-2.0 允许使用、修改、分发、私有使用和商业使用，但需要遵守许可证条款。

项目的商业化方向不是限制核心 skill 的使用，而是围绕服务、定制模板适配、托管流程、企业支持、模板包和自愿赞助来实现可持续发展。详见 [COMMERCIAL.md](COMMERCIAL.md)。
