# parse-document

Claude Code plugin for parsing PDF, DOCX, PPTX, XLSX, and image files into Markdown using [minerU](https://github.com/opendatalab/MinerU).

## Features

- Parse PDF, DOCX, PPTX, XLSX, PNG, JPG, JPEG → Markdown
- Auto-convert legacy formats: .doc → .docx, .ppt → .pptx, .xls → .xlsx
- Batch processing with smart output directory management
- Automatic cache detection (skip already-parsed files)
- Academic workflow support (translate, summarize, extract, redraw from papers)
- Handles Chinese paths and long filenames gracefully

## Prerequisites

1. **minerU** installed with GPU support. Follow the [official guide](https://github.com/opendatalab/MinerU).
2. **Conda** environment named `mineru` with mineru installed.
3. *(Optional)* **pywin32** for automatic legacy Office format conversion:
   ```bash
   conda run -n mineru pip install pywin32
   ```
   Requires Microsoft Office installed (Windows only).

## Installation

### Option 1: Manual install (current)

```bash
# 1. Clone the repo
git clone https://github.com/Enxulansis/parse-document-plugin.git

# 2. Add as local marketplace
claude plugin marketplace add ./parse-document-plugin --scope user

# 3. Install the plugin
claude plugin install parse-document@parse-document-plugin --scope user
```

### Option 2: One-click install (coming soon)

After marketplace registration, install with a single command:

```bash
claude plugin install parse-document@claude-plugins-official
```

*This plugin is pending submission to the official Claude Code plugin marketplace.*

## Usage

Just talk to Claude naturally — the skill triggers automatically:

- "看看这个PDF里面写了什么"
- "帮我把这些docx都转成markdown"
- "翻译这篇论文"
- "用draw.io重绘这篇论文的流程图"

Or invoke explicitly:

```
/parse-document:parse_document
```

## Environment Check

On first use, the skill auto-detects your environment and reports what's available:

```
minerU: detected
Office COM: available
Conda: available
```

If something is missing, the skill provides installation guidance.

## Supported Formats

| Parse directly | Auto-convert (Office + pywin32) |
|---------------|--------------------------------|
| .pdf .docx .pptx .xlsx | .doc → .docx |
| .png .jpg .jpeg | .ppt → .pptx |
| | .xls → .xlsx |

## Output Structure

- Single file: `<stem>_<ext>/` (e.g., `report_docx/`, `photo_png/`)
- Batch (2+ files): `output/<stem>_<ext>/`

## License

MIT
