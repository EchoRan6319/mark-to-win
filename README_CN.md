<div align="right">

[English](README.md) | 简体中文

</div>

# Mark to Win

将 AI 生成的 Markdown 转换为 Word、PDF 或 Excel 文件。

## 功能

- 支持 Word（`.docx`）、PDF（`.pdf`）和 Excel（`.xlsx`）输出（不稳定）
- 默认优先使用 Pandoc，必要时自动回退到本地实现
- 适配 Markdown 的标题、列表、表格、引用和代码块
- PDF 支持可选中文字体

## 快速开始

```bash
python scripts/mark_to_win.py example.md --to docx --output example.docx
python scripts/mark_to_win.py example.md --to pdf --output example.pdf
python scripts/mark_to_win.py example.md --to xlsx --output example.xlsx
```

## 安装

### Claude Code

```bash
git clone https://github.com/EchoRan6319/mark-to-win.git ~/.claude/skills/mark-to-win
```

### 手动安装

将 `SKILL.md` 和 `scripts/` 复制到 Claude skills 目录：

```bash
mkdir -p ~/.claude/skills/mark-to-win
cp SKILL.md ~/.claude/skills/mark-to-win/
cp -r scripts ~/.claude/skills/mark-to-win/
```

## 使用

```bash
python scripts/mark_to_win.py input.md --to docx --output output.docx --engine auto
python scripts/mark_to_win.py input.md --to pdf --output output.pdf --engine html --keep-html
python scripts/mark_to_win.py input.md --to docx,pdf,xlsx --output-dir exports
```

## 可选依赖

- `python-docx`：Word 输出
- `openpyxl`：不稳定的 Excel 输出
- `reportlab`：本地 PDF 回退

## 说明

- 大多数场景直接使用 `--engine auto`
- 需要更高保真 DOCX/PDF 时使用 `--engine pandoc`
- 需要 GitHub 风格 PDF 阅读效果时使用 `--engine html`

## 许可证

MIT
