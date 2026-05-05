# Minimal Markdown 示例 — 一份最小可跑的 thesis-style 文档

## 摘要

这是一份用于回归测试和读者快速体验的最小输入。它故意覆盖了模板适配里几个
最常见的元素：标题层级、带编号的图表标题、需要变三线表的表格、列表、行内
代码以及一段带超链接的正文。

## 1 引言

DOCX Template Translator Skill 的核心命题是：每个严格的 Word 模板都有自己的
局部规则，与其指望一个"通用转换器"覆盖一切，不如让 AI 先 inspect 模板再
生成项目专用的后处理脚本。

本示例不追求论文质量，只追求**端到端可跑**。

## 2 一段需要排版的正文

下面这段正文用来验证 `body_style` 的 remap 是否生效。在中文学位论文 preset
下，正文应当被映射到 `论文正文` / 类似样式，且字号统一为 12pt、中文字体宋体、
英文字体 Times New Roman。如果你跑的是默认（非 preset）配置，则不会发生
字体覆盖。

This sentence is intentionally English to verify that the Latin font mapping
(Times New Roman) works alongside the East Asian font mapping (宋体) without
breaking either side.

## 3 列表与代码

无序列表：

- 第一项：测试列表是否进入正确的样式
- 第二项：测试中文混排
- 第三项：`inline code` 是否保留等宽字体

有序列表：

1. inspect 模板
2. 生成 rough body
3. 应用 adaptive pipeline
4. finalize + preview

## 4 表格（应被识别并可选地变成三线表）

| 步骤 | 工具 | 跨平台 |
| --- | --- | --- |
| 检查模板 | inspect_docx_template.py | ✓ |
| 生成 rough body | pandoc | ✓ |
| 适配模板 | adaptive_docx_pipeline.py | ✓ |
| finalize | Word COM (pywin32) | Windows only |
| 预览拼图 | render_pdf_preview.py | ✓ |

## 5 一个带 caption 的"图"占位

下面用一段 markdown 标题模拟图片标题，方便验证 caption 正则识别：

> 图 1.1  这是一个测试用的图标题

(实际 docx 里 pandoc 会把这段渲染成 blockquote，但模板适配的关注点在它的
caption 正则，不在排版本身。)

## 6 引用

更多模板转换背景，见
[pandoc 官方手册](https://pandoc.org/MANUAL.html) 与
[本仓库 README](../../README.md)。

## 7 结论

如果这份 markdown 能在你的机器上一路跑到 `final.docx` 并且打开后内容完整，
说明 `inspect_docx_template.py` 和 `adaptive_docx_pipeline.py` 在你的环境里
工作正常。如果你拿到了 Windows + Word + pywin32 的完整环境，还会额外得到
`final.pdf` 和 `final.preview.png`。
