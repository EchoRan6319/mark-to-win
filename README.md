<div align="right">

English | [简体中文](README_CN.md)

</div>

# Mark to Win

Mark to Win is a universal Skill for converting AI-generated Markdown to Word, PDF, or Excel formats.

## Features

- 📄 Multiple output formats: Word (.docx), PDF, Excel (.xlsx) ⚠️ *unstable*
- 🔧 Simple command-line interface
- 📊 Preserves Markdown formatting and structure

## Files

- `SKILL.md`: Skill instructions
- `scripts/mark_to_win.py`: Markdown conversion helper

## Installation

### Claude Code

```bash
# Clone to Claude skills directory
git clone https://github.com/EchoRan6319/mark-to-win.git ~/.claude/skills/mark-to-win
```

### Manual Installation

1. Download or clone this repository
2. Copy `SKILL.md` and `scripts/` to your Claude skills directory:
   ```bash
   mkdir -p ~/.claude/skills/mark-to-win
   cp SKILL.md ~/.claude/skills/mark-to-win/
   cp -r scripts ~/.claude/skills/mark-to-win/
   ```

### Verify Installation

After installation, the skill will be automatically available in Claude Code. You can verify by checking if `mark-to-win` appears in your skills list.

## Quick Start

```bash
python scripts/mark_to_win.py example.md --to docx --output example.docx
python scripts/mark_to_win.py example.md --to xlsx --output example.xlsx
python scripts/mark_to_win.py example.md --to pdf --output example.pdf
```

## Optional Dependencies

The script uses optional dependencies:

- `python-docx` - Word document support
- `openpyxl` - Excel file support (⚠️ unstable)
- `reportlab` - PDF generation support

## PDF Optimization

For best PDF fidelity, install and use Pandoc from the Skill workflow when available.

## Usage Examples

```bash
# Convert to Word document
python scripts/mark_to_win.py README.md --to docx --output readme.docx

# Convert to Excel spreadsheet (⚠️ unstable - may lose formatting)
python scripts/mark_to_win.py data.md --to xlsx --output data.xlsx

# Convert to PDF file
python scripts/mark_to_win.py report.md --to pdf --output report.pdf
```

## License

This project is licensed under the MIT License.
