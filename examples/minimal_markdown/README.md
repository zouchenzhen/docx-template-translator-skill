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

### 用真实学位论文模板跑

不传参数时 `run_example.py` 会用 `build_template.py` 现场生成的轻量
sample 模板（小、可重现，不依赖任何学校资产）。如果你想看这套流水线在
**真实学位论文模板**上的输出（也就是仓库里 `expected/preview.png` 那张
图的来源），传一个本地的 `.docx` 模板路径即可：

```bash
python examples/minimal_markdown/run_example.py \
    --template /path/to/your-real-thesis-template.docx \
    --preview-pages "20"
```

脚本会把这份模板复制到 `sample_template.docx`，然后照常跑 inspect →
pandoc → adaptive → finalize → preview。`--preview-pages` 控制最后那张
PDF 拼图渲染哪些页：默认 `"1-4"` 在简化 demo 模板上就能看到核心结构，
但跑真实学位论文模板时前几页全是封面/声明/英文扉页，看不出 skill 对正文
内容做了什么；改用 `"20"` 之类的页码，可以直接渲染示例 `sample.md`
生成出来的正文页。当前 `expected/preview.png` 进一步从第 20 页裁剪出主体
区域，集中展示真实插图、图题、正文段落、显示公式和三线表。

模板本身不会被 commit（已在 `.gitignore` 里），只产出会写到
`examples/minimal_markdown/`。`expected/preview.png` 就是用这种方式针对
一份郑州大学学位论文模板跑出来的；具体每页是什么、为什么挑这几页，详见
[`expected/README.md`](expected/README.md)。

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
