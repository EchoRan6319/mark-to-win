<div align="right">

English | [简体中文](README_CN.md)

</div>

# Mark to Win

Convert AI-generated Markdown into Word, PDF, or Excel files.

## Features

- Word (`.docx`), PDF (`.pdf`), and Excel (`.xlsx`) output
- Pandoc-first workflow with native fallbacks
- Markdown-aware handling for headings, lists, tables, quotes, and code blocks
- Optional CJK font support for PDF output

## Quick Start

```bash
python scripts/mark_to_win.py example.md --to docx --output example.docx
python scripts/mark_to_win.py example.md --to pdf --output example.pdf
python scripts/mark_to_win.py example.md --to xlsx --output example.xlsx
```

## Installation

### Claude Code

```bash
git clone https://github.com/EchoRan6319/mark-to-win.git ~/.claude/skills/mark-to-win
```

### Manual

Copy `SKILL.md` and `scripts/` into your Claude skills directory:

```bash
mkdir -p ~/.claude/skills/mark-to-win
cp SKILL.md ~/.claude/skills/mark-to-win/
cp -r scripts ~/.claude/skills/mark-to-win/
```

## Usage

```bash
python scripts/mark_to_win.py input.md --to docx --output output.docx --engine auto
python scripts/mark_to_win.py input.md --to pdf --output output.pdf --engine html --keep-html
python scripts/mark_to_win.py input.md --to docx,pdf,xlsx --output-dir exports
```

## Optional Dependencies

- `python-docx` for Word output
- `openpyxl` for unstable Excel output
- `reportlab` for native PDF fallback

## Notes

- Use `--engine auto` for most cases.
- Use `--engine pandoc` when you need higher-fidelity DOCX/PDF output.
- Use `--engine html` for a GitHub-style PDF reader layout.

## License

MIT
