# DOCX Template Translator Skill

一个用于将 **LaTeX、PDF、Markdown 或粗转 DOCX** 转译成指定 Word 模板格式的 Codex skill。

它面向毕业论文、学位论文、单位报告、标准文档等场景，尤其适合那些 **pandoc 默认 DOCX 输出不够用**、必须严格套用学校或机构 Word 模板的任务。

English: [README.en.md](README.en.md)

## 一分钟看效果

仓库自带一个最小可跑示例 [`examples/minimal_markdown/`](examples/minimal_markdown/)，在仓库根目录跑：

```bash
python examples/minimal_markdown/run_example.py
```

脚本会现场生成一份 sample 模板，依次跑 inspect / adaptive / finalize，并在
Windows + Word 环境下输出 PDF 预览拼图：

![预览拼图](examples/minimal_markdown/expected/preview.png)

> 上图是用一份**真实的郑州大学学位论文模板**跑出来的 4 页拼图（封面 / 英文
> 封面 / 英文扉页 / 原创性声明）。它通过 `--template` 参数指定模板生成：
> `python examples/minimal_markdown/run_example.py --template <你的本地模板.docx>`。
> 不传 `--template` 时 `run_example.py` 会用 `build_template.py` 现场生成的
> 轻量 sample 模板，输出更简化但完全可重现。封面里的 "李四 / 王五 / 10459"
> 是学校模板自带的占位示例值，不是任何真实学生信息。

inspect / adaptive / preview 三步在 macOS / Linux / Windows 都能跑；用 Word COM 更新字段并导出 PDF 的 finalize 步骤只在 Windows + Microsoft Word 下可用。

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

将 skill 复制到 Codex skills 目录。

macOS / Linux：

```bash
mkdir -p ~/.codex/skills
cp -r skills/docx-template-translator ~/.codex/skills/
```

Windows（PowerShell）：

```powershell
New-Item -ItemType Directory -Force -Path "$HOME\.codex\skills" | Out-Null
Copy-Item -Recurse -Force skills\docx-template-translator "$HOME\.codex\skills\"
```

然后可以这样调用：

```text
Use $docx-template-translator to convert my LaTeX thesis into Word using this .docx template.
```

如果需要在 Claude Code、Cursor 或其他 AI agent 里使用同一份 skill 文件夹，参考下文 [兼容性](#兼容性) 一节。

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
   起步脚本支持 JSON 配置；如果是中文学位论文模板，可以直接复用内置 preset：

```bash
python skills/docx-template-translator/scripts/adaptive_docx_pipeline.py \
    --template template.docx \
    --body-docx body.docx \
    --out final.docx \
    --config skills/docx-template-translator/presets/zhengzhou_thesis.json \
    --three-line-tables
```

非中文学位论文模板请去掉 `--three-line-tables`，并使用自己的 config（或保留默认行为，默认不会改写正文字体）。

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

## 已知限制

- **finalize 阶段仅支持 Windows。** `finalize_word_docx.py` 通过 pywin32 调用 Microsoft Word COM 来更新字段/目录并导出 PDF，因此只在装有真实 Word 的 Windows 机器上能跑完整流程。模板检查、adaptive pipeline 和 PDF 预览拼图在 macOS / Linux 上不依赖 Word，可以直接运行。
- **不支持扫描版 / 纯图片 PDF。** PDF 输入需要原生数字文本层（pdf2docx / Word 导入才能抽出结构），先 OCR 再转，或直接用 LaTeX / Markdown 源文件。
- **模板含宏：宏会被禁用。** finalize 阶段会先设置 `Word.AutomationSecurity = msoAutomationSecurityForceDisable`，再打开文档，因此模板里的 AutoMacros / VBA 不会执行。详见 [SECURITY.md](SECURITY.md)。
- **starter pipeline 默认行为是保守的。** 三线表、正文字体覆盖等中文学位论文专属处理改成了 opt-in，需要通过 config（参考 `presets/zhengzhou_thesis.json`）或 CLI flag 显式打开。默认行为不会偷偷改写你的正文字体。
- **样式名识别覆盖英文和中文模板。** 其他本地化 Word 模板（日文 / 韩文等）可能需要在 config 里补充 `unnumbered_heading_styles` 和 `body_candidate_styles`。

## 安全说明

- finalize 阶段会用本机 Microsoft Word 打开用户提供的 `.docx`。脚本默认禁用宏，请勿在不可信来源的输入上放宽该设置。
- skill 工作流允许 AI agent 基于 `adaptive_docx_pipeline.py` 生成 / 改写一份项目专用 Python 脚本。**运行 AI 改写后的脚本前请人工 review diff**。本项目自带的脚本不进行任何网络 I/O，也只往用户显式指定的输出路径写文件。
- 完整威胁模型详见 [SECURITY.md](SECURITY.md)。

## 兼容性

skill 元数据（带 `name` + `description` frontmatter 的 `SKILL.md`）虽然为 Codex 设计，但格式与 Anthropic Claude Skills 规范有重叠，因此同一份 skill 文件夹可以稍作包装后被其他 AI agent 复用。

| Agent / IDE     | 状态        | 说明 |
| --------------- | ----------- | ---- |
| **Codex**       | 原生支持     | 直接放到 `~/.codex/skills/`，用 `Use $docx-template-translator …` 触发。 |
| **Claude Code** | 包装后可用   | `SKILL.md` 的 frontmatter 与 Claude Skills 规范一致，放到 Claude Skills 目录（例如 `~/.claude/skills/<name>/`）或包装成 Claude Code plugin 即可。Python 脚本无需改动。 |
| **Cursor**      | 手动 / Rules | Cursor 没有原生 "skill" 概念，可以把 SKILL.md 内容粘到 `.cursor/rules/*.mdc` 里，让 agent 直接调用 `scripts/*.py`。 |
| **OpenClaw**    | 可适配       | 结构与 OpenClaw skill 约定接近，但仓库里没有 OpenClaw 专属 manifest，发布到该平台前请补元数据。 |

脚本本身是纯 CPython，不依赖任何 agent runtime —— 任何能跑 shell 命令的 AI agent 都可以驱动这套工作流。

## 许可证

本项目采用 [Apache License 2.0](LICENSE) 开源许可。

Apache-2.0 允许使用、修改、分发、私有使用和商业使用，但需要遵守许可证条款。

## 社区

[LINUX DO — 中文开发者社区](https://linux.do/)

本项目认可并感谢 LINUX DO 社区在中文开发者开源交流、项目分享和技术讨论中的价值。除非社区另有明确说明，此处仅为社区致谢和链接，不代表官方背书。
