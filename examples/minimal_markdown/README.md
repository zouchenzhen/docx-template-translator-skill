# Minimal Markdown 示例

这是一个**最小可跑**的端到端示例，用于演示这套 skill 的核心工作流。

它故意不依赖 Microsoft Word，因此可以在 macOS / Linux / Windows 上都跑通到
`final.docx`，验证 `inspect_docx_template.py` + `adaptive_docx_pipeline.py`
两步是否正常工作。Word COM 相关的 `finalize_word_docx.py` 因为只能在 Windows
上跑，本示例不强制依赖。

## 文件清单

```
examples/minimal_markdown/
├── README.md                  ← 你正在看的这份
├── sample.md                  ← 输入 Markdown（带标题/正文/列表/表格/图）
├── build_template.py          ← 用 python-docx 现场生成一个 sample_template.docx
├── run_example.py             ← 跨平台一键脚本：build template → pandoc → adaptive → preview
└── expected/
    ├── preview.png            ← 期望生成的 PDF 预览拼图（Windows + Word 才能产出）
    └── README.md              ← 说明 expected/ 下文件如何生成
```

仓库里**不**直接 commit `sample_template.docx`，而是用 `build_template.py`
现场生成。这样既避免在 git 历史里塞二进制 blob，也方便读者看清"模板里到底
预定义了哪些样式"。

## 运行需要

最小依赖：

```bash
pip install python-docx pillow pymupdf
```

如果想从 Markdown 构造 rough body，再装一个 pandoc。

如果要走完 finalize（更新字段 + 导出 PDF + 渲染预览），需要：

- Windows
- 已安装的 Microsoft Word
- `pip install pywin32`

## 一键运行

在仓库根目录下：

```bash
python examples/minimal_markdown/run_example.py
```

脚本会按顺序执行：

1. `build_template.py` 生成 `examples/minimal_markdown/sample_template.docx`；
2. 调 `inspect_docx_template.py`，输出 `sample_template.template-report.json`；
3. 如果检测到 `pandoc`：用 `pandoc sample.md --reference-doc sample_template.docx -o body.docx` 生成粗 body；
   否则用 python-docx 自己拼一个最小 body（保证示例在没装 pandoc 的环境下也能跑通）；
4. 调 `adaptive_docx_pipeline.py` 生成 `final.docx`；
5. 如果当前是 Windows + 装了 Word + 装了 pywin32：调 `finalize_word_docx.py --pdf` 并用 `render_pdf_preview.py` 输出 `preview.png`；
   否则跳过这一步并打印提示。

所有产物落在 `examples/minimal_markdown/`，方便和 `expected/` 对比。

## 期望产物

跑完之后你应该看到：

```
examples/minimal_markdown/
├── sample_template.docx                ← 由 build_template.py 生成
├── sample_template.template-report.json ← 由 inspect 步骤生成
├── body.docx                           ← 中间产物
├── final.docx                          ← 适配模板后的成品
├── final.pdf       (Windows + Word 才有)
└── final.preview.png  (Windows + Word 才有)
```

`expected/preview.png` 是仓库提供的参考截图，便于直接看到这个 skill 的产出
长什么样，不用真的本地装 Word 才能体验。

## 故障排除

| 现象 | 可能原因 / 处理 |
| --- | --- |
| `python-docx` 找不到 `Document` | 使用了 `docx` 这个不相关的同名包，请 `pip uninstall docx && pip install python-docx`。 |
| `pandoc` 报找不到命令 | 没装 pandoc。脚本会自动 fallback 到 python-docx 拼 body，仍能跑完到 `final.docx`。 |
| Windows 跑 finalize 时 `pywintypes.com_error` | 多半是 gen_py 缓存损坏。删 `%LOCALAPPDATA%\Temp\gen_py\` 后重试，本仓库脚本已加 fallback，但旧缓存仍可能干扰。 |
| 三线表样式没出现 | 默认行为是不强制三线表。本示例显式传了 `--three-line-tables`；如果你跑自定义命令，记得加上或在 config 里开 `enable_three_line_tables`。 |
