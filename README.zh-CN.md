# Bookcraft 出书与封面 Skill

Bookcraft 是一个面向 Codex 的出书与商业封面包装 Skill，用于规划、写作、修订、维护、出版结构整理、封面方向设计，以及将 Markdown 书稿导出为 Word/DOCX。

## 作者与联系方式

作者：陈歆怡，海聚海外 CEO。

英文介绍：Xinyi Chen, founder of HEDGE Global.

联系方式：`chenxinyi_g`

## 关于陈歆怡

陈歆怡（Xinyi Chen），在上海的浙江人。

海聚海外（HEDGE Global）创始人兼 CEO，长期深耕高端国际教育、企业出海和科技投资。以前喜欢随心旅行，现在是 Vibe coding 爱好者。

想要交流 Vibe coding 的朋友，欢迎加微信：`chenxinyi_g`

![微信二维码](assets/wechat-qr.png)

## 能做什么

- 从书名、主题、一句话想法或已有草稿开始规划一本书。
- 生成 mdBook 风格书稿项目，包括 `book.toml`、`src/SUMMARY.md`、维护文件和章节文件。
- 用 `RESEARCH_INDEX.md` 管理资料、引用、待核查内容和参考文献状态。
- 按章节写作、审阅、扩写、修订，并维护 `CONTINUITY.md`、`STYLE_GUIDE.md`、`GLOSSARY.md`。
- 整理公开出版图书结构：扉页、版权页、自序、前言、正文、参考文献和可选后置页。
- 输出商业封面包装方向：一句话销售钩子、视觉符号、标题层级、货架测试、缩略图测试和图像生成提示词。
- 通过 `scripts/md2word.py` 导出可编辑 Word/DOCX 成品。

## 安装

```bash
mkdir -p ~/.codex/skills
git clone https://github.com/hedgeacademy/bookcraft ~/.codex/skills/bookcraft
```

开启新 Codex 对话后可以这样使用：

```text
Use $bookcraft to plan a book about [topic].
```

中文使用也可以直接说：

```text
用 bookcraft 帮我规划一本关于 AI 工作流的商业书。
```

## 常用提示词

```text
Use $bookcraft to create a book project from this outline.
```

```text
Use $bookcraft to review this chapter for structure, evidence, and continuity.
```

```text
Use $bookcraft to build a commercial cover brief for this manuscript.
```

```text
Use $bookcraft to export this manuscript to Word/DOCX.
```

## 封面方法

封面模块强调商业识别，而不是单纯“好看”。一个有效封面应在几秒内回答：

- 这本书讲什么？
- 这本书写给谁？
- 它承诺给读者什么结果或情绪？
- 为什么读者应该选择这本，而不是旁边同类书？

完整封面方法见：

- `references/09-cover-design.md`
- `references/09-cover-design.zh-CN.md`

## 导出要求

DOCX 导出需要 Python、`python-docx`、`pandoc`、完整的 `manuscript.md`、有效的 `book.toml`，并且 `# 参考文献` 下至少有一条真实参考文献。

```bash
python3 -m pip install python-docx
brew install pandoc
python3 scripts/md2word.py --input manuscript.md
```

## 许可证

MIT
