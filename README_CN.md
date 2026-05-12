<div align="right">

[English](README.md) | 简体中文

</div>

# 马克吐文

马克吐文是一个通用 Skill，用来把 AI 生成的 Markdown 转成 Word、PDF 或 Excel。

## 功能特性

- 📄 支持多种输出格式：Word (.docx)、PDF、Excel (.xlsx) ⚠️ *不稳定*
- 🔧 简单的命令行接口
- 📊 保留 Markdown 格式和结构

## 文件说明

- `SKILL.md`: Skill 指令文件
- `scripts/mark_to_win.py`: Markdown 转换辅助脚本

## 安装

### Claude Code

```bash
# 克隆到 Claude skills 目录
git clone https://github.com/EchoRan6319/mark-to-win.git ~/.claude/skills/mark-to-win
```

### 手动安装

1. 下载或克隆本仓库
2. 将 `SKILL.md` 和 `scripts/` 复制到 Claude skills 目录：
   ```bash
   mkdir -p ~/.claude/skills/mark-to-win
   cp SKILL.md ~/.claude/skills/mark-to-win/
   cp -r scripts ~/.claude/skills/mark-to-win/
   ```

### 验证安装

安装完成后，Skill 会自动在 Claude Code 中可用。可以通过检查 `mark-to-win` 是否出现在 skills 列表中来验证。

## 快速开始

```bash
python scripts/mark_to_win.py example.md --to docx --output example.docx
python scripts/mark_to_win.py example.md --to xlsx --output example.xlsx
python scripts/mark_to_win.py example.md --to pdf --output example.pdf
```

## 可选依赖

脚本使用以下可选依赖：

- `python-docx` - Word 文档支持
- `openpyxl` - Excel 文件支持 (⚠️ 不稳定)
- `reportlab` - PDF 生成支持

## PDF 优化

为了获得最佳的 PDF 转换质量，建议在 Skill 工作流中安装并使用 Pandoc。

## 使用示例

```bash
# 转换为 Word 文档
python scripts/mark_to_win.py README.md --to docx --output readme.docx

# 转换为 Excel 表格 (⚠️ 不稳定 - 可能丢失格式)
python scripts/mark_to_win.py data.md --to xlsx --output data.xlsx

# 转换为 PDF 文件
python scripts/mark_to_win.py report.md --to pdf --output report.pdf
```

## 许可证

本项目采用 MIT 许可证。
