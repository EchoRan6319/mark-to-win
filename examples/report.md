# AI 项目周报

## 本周进展

- 完成 Markdown 到 Word 的基础转换。
- 支持 Markdown 表格导出为 Excel 工作表。
- PDF 输出优先建议使用 Pandoc，高兼容场景可使用脚本 fallback。

## 风险清单

| 风险 | 影响 | 建议 |
| --- | --- | --- |
| PDF 样式不一致 | 中 | 优先使用 Pandoc 或浏览器打印管线 |
| Excel 输入不是表格 | 低 | 退化为结构化文本 Sheet |

## 结论

马克吐文适合作为 AI 办公文档导出的通用 Skill 起点。
