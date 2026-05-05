# Contributing Guide

感谢对 docx-template-translator-skill 的兴趣！这份指南说明如何提 issue、提
PR、跑示例、以及哪些方向最欢迎贡献。

## 项目定位（提 PR 前请先阅读）

本仓库是一个 **AI 模板适配工作流 skill**，不是一个"通用一键转换器"。

- ✅ 欢迎：让 inspect / adaptive / finalize 流程在更多模板下更稳的修改；让
  脚本在更多平台 / Word 版本 / 本地化下能跑；新增 preset；新增 example；
  补文档与 case study；修 bug；提升安全性。
- ❌ 不太欢迎：把 `adaptive_docx_pipeline.py` 改成"内置一切规则的大转换器"。
  这个脚本是 **starter**，专门用于让 AI agent 复制并改造，不是要长成
  All-in-one。

如果你的需求是"我有自己学校的模板，想让脚本一键支持"，**正确做法是新增一份
preset**（参考 `skills/docx-template-translator/presets/zhengzhou_thesis.json`），
而不是把规则塞进 `adaptive_docx_pipeline.py` 默认行为。

## 提 Issue

提 issue 时请尽量包含：

1. 操作系统 + Python 版本 + Microsoft Word 版本（如果走 finalize）。
2. 脚本调用的完整命令。
3. 期望行为 vs 实际行为。
4. 如果是模板适配问题，附 `inspect_docx_template.py` 输出的 JSON 报告
   （删掉敏感个人信息），不要直接附整份 docx。

如果是安全相关问题，请走 [SECURITY.md](SECURITY.md) 流程，而不是公开 issue。

## 本地开发环境

最小依赖：

```bash
pip install python-docx pillow pymupdf
# Windows 走 finalize 还需要：
pip install pywin32
# Markdown / LaTeX 输入需要 pandoc，从 https://pandoc.org/installing.html 安装
```

跑端到端示例验证你的改动没破坏主流程：

```bash
python examples/minimal_markdown/run_example.py
```

非 Windows 环境下该脚本会自动跳过 finalize 步骤，但 inspect / adaptive /
preview 必须仍能跑通。

## 提 PR 前的自检

提 PR 前请确认：

- [ ] 修改的脚本能通过 `python -m py_compile <file>` 语法检查。
- [ ] `examples/minimal_markdown/run_example.py` 在你的环境下还能跑完。
- [ ] 没有把 `*.docx` / `*.pdf` / `*.preview.png` 等运行产物 commit 进 git
      （`.gitignore` 已覆盖 example 目录下的常见产物）。
- [ ] 不要在 PR 里附带任何含个人姓名、学号、学校签章等隐私信息的真实文档。
- [ ] commit message 写明 **why**，而不只是 **what**。

## Coding Style

- Python 用 4 空格缩进，类型注解视情况而定，不强求 100% 标注。
- 脚本对外的 CLI 通过 `argparse`，不要引入新的 CLI 框架。
- 直接对 `docx.oxml` 操作 OOXML 是允许的，但请加注释说明操作的是哪段 XML
  以及为什么必须用 raw XML 而不是 python-docx 高层 API。
- 不要默默改写用户的正文字体 / 表格样式。任何 opinionated 行为必须是
  opt-in（CLI flag 或 config key），并在 README "Known Limitations" 同步说明。
- 不要在脚本里发起网络请求。

## 新增 Preset

把目标模板的所有 opinionated 规则放进一个 JSON：

```json
{
  "_comment": "<人话描述这个 preset 适用于哪个模板 / 哪种使用场景>",
  "body_style": "...",
  "body_candidate_styles": ["..."],
  "unnumbered_h1": ["..."],
  ...
}
```

文件名建议形如 `<institution>_<doc_type>.json`，放在
`skills/docx-template-translator/presets/`。

提 PR 时附一段简短的"我用这个 preset 跑过哪份模板、效果如何"。**请不要把
真实学校模板 / 论文 commit 进仓库**，附到 PR 描述或 issue 评论即可。

## 安全相关贡献

请优先关注：

- finalize 阶段对 Word COM 的安全使用（保留 `AutomationSecurity = 3`）；
- 所有路径参数都应基于用户显式输入，不应猜测 / glob 用户磁盘；
- 任何"自动写脚本到 ~/.codex/skills"之类的便利功能要走显式确认。

## 版权与协议

本项目使用 [Apache License 2.0](LICENSE)。提交 PR 即表示你同意你的贡献以
同样的协议被接受、分发和再许可。
