# SlideForge

**一条命令，把 Markdown 变成能编辑的 PowerPoint。生成的是原生文本框，不是图片 — 每一页都能在 PowerPoint / WPS 里继续改。**

**Markdown → editable .pptx in one command. Native text boxes, not screenshots.**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
![Python](https://img.shields.io/badge/Python-3.9%2B-blue)

## 为什么 / Why

“文档转 PPT”工具很多，但写东西的人大多在 Markdown 里起草。SlideForge 就补这一块：读纯 Markdown，出真正的 PPTX。

- **可编辑**：标题、要点都是原生文本框，不是把页面截成图片贴上去
- **零学习成本**：会写 Markdown 就会做幻灯片，`#` 是封面/章节，`##` 是一页，`-` 是要点，`>` 是金句页
- **本地运行**：不上传文档，没有账号，没有水印
- **能进流水线**：命令行工具，可以被脚本和 AI Agent 调用，先出草稿再人工精修

## 快速开始 / Quick start

```bash
git clone https://github.com/Yangdingcheng/slideforge.git
cd slideforge
pip install -r requirements.txt
python3 slideforge.py examples/deck.md --theme dark --brand "@我的账号"
# 生成 examples/deck.pptx，16:9，共 6 页
```

## Markdown 语法 / Syntax

| 写法 | 生成 |
|---|---|
| `# 标题`（文件第一个） | 封面页（下一行普通文字是副标题） |
| `# 标题`（后面的） | 章节分隔页 |
| `## 标题` | 内容页 |
| `- 要点` | 一级要点 |
| `  - 要点`（两个空格） | 二级要点 |
| `1. 步骤` | 编号要点 |
| `> 金句` | 金句页 |

## 主题 / Themes

`dark` 黑红 · `light` 浅色 · `ocean` 深蓝。页脚自动带品牌和页码（`当前/总页数`）。

## 适合谁

- 经常要把文章、笔记、周报快速变成汇报 PPT 的人
- 用 Obsidian / 语雀 / Markdown 写作，想一键出演示稿的人
- 给 AI Agent 打底稿：Agent 写 Markdown，SlideForge 出 PPT，人再精修

## Roadmap

- [ ] 表格和代码块渲染
- [ ] 图片页（`![](path)` 生成图文页）
- [ ] 自定义主题 JSON
- [ ] 演讲者备注（段落注释写入备注区）

## License

MIT
