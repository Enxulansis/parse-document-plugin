---
name: parse_document
description: >
  Parse PDF, DOCX, PPTX, XLSX, and image files (PNG, JPG, JPEG) into Markdown
  using the minerU toolchain. Trigger whenever the user wants to read, view,
  open, inspect, or extract content from any document file — even casual
  phrasing ("看看这个pdf", "读一下那个docx", "打开看看", "里面写了什么", "帮我查看",
  "提取文字"). If a file is binary (PDF, DOCX, PPTX, XLSX, images), Claude cannot
  read it directly, so this skill MUST be invoked first. Also triggers for: batch
  processing ("把所有pdf都解析"), legacy Office formats that won't open
  (.doc/.ppt/.xls "打不开"), academic/research workflows ("整理参考文献",
  "翻译这篇论文", "快速查看这篇论文", "用draw.io重绘这篇论文的流程图",
  "总结这篇论文的方法", "提取论文里的数据"), image text extraction, and any
  scenario where parsed document content is needed as input for another skill
  (e.g., generate_image) or downstream task (translate, summarize, analyze,
  redraw). This skill is often the REQUIRED FIRST STEP whenever the user wants
  to do anything with a binary document they cannot directly read. Do NOT trigger
  for plain-text formats (.md, .txt, .py, .json) that Claude can read directly.
---

## Document parsing with minerU

Parse binary documents into Markdown for Claude to read and analyze.

### Environment check (run first)

Before first use, verify the minerU environment is available:

```bash
conda run -n mineru mineru --version
```

If this fails, guide the user to install minerU: https://github.com/opendatalab/MinerU

Also check for optional Office format conversion support:

```bash
conda run -n mineru python -c "import win32com.client; print('Office COM: available')"
```

If this fails, legacy formats (.doc/.ppt/.xls) will require manual conversion. Tell the user they can install pywin32: `conda run -n mineru pip install pywin32`

### Key paths

This skill ships a helper module at `scripts/mineru_helper.py` relative to the plugin directory. Find it by locating the SKILL.md file and navigating to `../scripts/mineru_helper.py` relative to it.

Alternatively, the helper is installed alongside the plugin — find it with:

```python
import os
skill_dir = os.path.dirname(os.path.abspath(__file__))  # not exact, illustrative
helper_path = os.path.join(skill_dir, "..", "scripts", "mineru_helper.py")
```

The helper auto-detects mineru location via `conda run -n mineru python -c "import shutil; print(shutil.which('mineru'))"`. If conda is not available, ask the user for the full path to their mineru installation.

### Format support

| Supported (parse directly) | Auto-convert (if Office+pywin32) | Not supported |
|----------------------------|----------------------------------|---------------|
| `.pdf` `.docx` `.pptx` `.xlsx` | `.doc` → `.docx` | Other formats |
| `.png` `.jpg` `.jpeg` | `.ppt` → `.pptx` | |
| | `.xls` → `.xlsx` | |

### Output directory rules

- **Single file**: `<project_dir>/<stem>_<ext>/`
- **Batch (2+ files)**: `<project_dir>/output/<stem>_<ext>/`

### Workflow

bash cannot pass Chinese characters reliably via command-line arguments. Instead, embed paths inside a Python runner script written with the Write tool.

#### Step 1: Determine what to parse

Identify the target files from the user's request. Use `os.listdir` + extension filtering for files with special Unicode characters.

#### Step 2: Write the runner script

Write a runner to a temp location (e.g., `/tmp/run_parse.py` on Linux, or a Write-tool-accessible path):

```python
import sys, os
# Add the plugin's scripts directory to path
sys.path.insert(0, r"<PATH_TO_PLUGIN_SCRIPTS>")
import mineru_helper

results = mineru_helper.run(
    project_dir=r"<PROJECT_DIR>",
    filenames=["<file1>", ...],  # or None to auto-detect all
)
```

Replace `<PATH_TO_PLUGIN_SCRIPTS>` with the actual path to the plugin's `scripts/` directory. Replace `<PROJECT_DIR>` with the absolute project path.

#### Step 3: Execute

```bash
conda run -n mineru python "<PATH_TO_RUNNER>"
```

If conda is not available, use the full path to the mineru environment's python.

#### Step 4: Read output

The runner prints parsed content. Read it from the command output and respond to the user's original question.

### Edge cases

- **Long filenames (>80 chars)**: The helper auto-copies to a short temp name to avoid Windows MAX_PATH errors.
- **Special Unicode characters**: Use `os.listdir` + extension filtering if hardcoded filenames fail.
- **No GPU available**: minerU falls back to CPU mode (slower but functional).
- **Office not installed**: Legacy formats (.doc/.ppt/.xls) require manual conversion.

### Log interpretation

When parsing Office formats, `WARNING: No valid PDF or image files to process` and `Skipping visualization...` are normal — the Office file was processed successfully.
