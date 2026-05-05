# Security Policy

This skill runs on a developer's local machine and is invoked by an AI agent.
It accepts user-supplied `.docx`, `.tex`, `.md`, and `.pdf` files, and it can
drive Microsoft Word through COM. A few of these paths deserve attention.

## Threat Model

- Inputs (`.docx` template, `.docx` body, `.tex`, `.md`, `.pdf`) come from the
  user or are produced by upstream tools (pandoc, pdf2docx). They are treated
  as **untrusted** by this project's own scripts.
- The AI-generated postprocessing pipeline (a copy of
  `scripts/adaptive_docx_pipeline.py`) is treated as **partially trusted**:
  the human user is expected to review the diff before running it.
- The `finalize_word_docx.py` step launches the **real local Microsoft Word**
  via COM, which is the most security-sensitive step in the workflow.

## Word Macros / AutoMacros

`scripts/finalize_word_docx.py` opens the final `.docx` with Word through
`win32com.client`. Any `.docx` (template or final document) could in principle
contain VBA macros or AutoMacros that run on open.

The script mitigates this by setting:

```python
word.AutomationSecurity = 3   # msoAutomationSecurityForceDisable
```

before opening the document, so macros and AutoMacros are disabled for that
session.

**Do not weaken this** unless you fully trust the input file. If you receive a
`.docx` template from an unknown source, prefer to inspect it manually in Word
with macros disabled first, and only then run the finalize step.

## AI-Generated Postprocessing Scripts

The skill workflow encourages an AI agent to generate or patch a project-
specific Python pipeline based on `scripts/adaptive_docx_pipeline.py`. As with
any AI-generated code, **review the diff before running it**, especially:

- file deletion / overwriting paths,
- `subprocess` / shell calls,
- network access,
- arbitrary `eval` / `exec`.

The reference scripts shipped with this skill do not perform any network I/O
and do not delete user files; the only filesystem writes are to the explicit
output paths the user passes on the command line.

## Dependencies

The recommended Python dependencies (`python-docx`, `pywin32`, `pymupdf`,
`pillow`, optionally `pdf2docx`) are mainstream and actively maintained. The
project does not vendor any binary blobs. `pandoc` and Microsoft Word, if
used, must be installed separately by the user.

## Reporting

If you find a security issue, please open a private security advisory on
GitHub or contact the maintainer through the repository instead of filing a
public issue.
