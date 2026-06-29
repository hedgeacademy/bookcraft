# Phase 3：mdBook 项目结构

书稿采用 mdBook 兼容的 Markdown 项目结构。创建项目时必须先创建目录和 `book.toml`，再创建基础文件，并写入初始大纲、风格指南、连续性记录和 `SUMMARY.md`。如果用户已有项目，不要覆盖已有文件，先读取并增量更新。

## 推荐目录结构

```text
book-project/
├── book.toml
├── README.md
├── OUTLINE.md
├── CONTINUITY.md
├── STYLE_GUIDE.md
├── GLOSSARY.md
├── RESEARCH_INDEX.md
├── CHANGELOG.md
├── manuscript.md
├── operation-log.md
├── src/
│   ├── SUMMARY.md
│   ├── front-matter/
│   │   ├── recommendation-preface-01.md（可选，序言一）
│   │   ├── recommendation-preface-02.md（可选，序言二）
│   │   ├── recommendation-preface-03.md（可选，序言三）
│   │   ├── author-preface.md
│   │   ├── preface.md
│   │   └── toc-placeholder.md
│   ├── chapters/
│   │   ├── chapter-01.md
│   │   ├── chapter-02.md
│   │   └── chapter-03.md
│   └── back-matter/
│       └── references.md
├── assets/
│   ├── images/
│   └── tables/
├── exports/
├── config/
│   └── word.json
└── scripts/
    └── md2word.py
```

## 文件职责

- `book.toml`：mdBook 配置文件。`[book]` 放 mdBook 标准字段；`[bookcraft.*]` 放出版和版权扩展字段，驱动 Word 版权页渲染。
- `src/SUMMARY.md`：维护全书目录结构。
- `OUTLINE.md`：记录完整大纲。
- `CONTINUITY.md`：记录长篇写作中的连续性信息。
- `STYLE_GUIDE.md`：记录全书写作风格、语气、段落节奏和表达规则。
- `GLOSSARY.md`：统一核心术语。
- `RESEARCH_INDEX.md`：管理研究资料和引用。
- `CHANGELOG.md`：记录重要修改。
- `manuscript.md`：组装后的完整书稿，用于 Word 转换。
- 扉页和版权页不创建独立 Markdown 正文文件，也不加入 `src/SUMMARY.md` 或 `manuscript.md`。Word 导出时统一由 `scripts/md2word.py` 根据 `book.toml` 生成。
- `operation-log.md`：记录假设、决策、创建文件、转换命令和验证结果。
- `exports/`：保存最终导出的 Word/DOCX 文件。
- `scripts/md2word.py`：Markdown 到 Word 的转换脚本，复制自本项目的 `scripts/md2word.py`。

## book.toml

创建项目时必须提前创建 `book.toml`，即使部分字段暂缺，也要先写入占位值。优先使用 Skill 自带的 `templates/book.toml`。

mdBook 规范中 `[book]` 常用字段：

- `title`：书名
- `authors`：作者列表，数组形式。创建项目时先从用户对话中获取作者姓名；获取成功时写入该姓名，未提供时使用 `["蜡笔小歆"]`。
- `description`：图书简介
- `src`：源文件目录，通常为 `"src"`
- `language`：语言，例如 `"zh-CN"`
- `text-direction`：文本方向，例如 `"ltr"`

不要在 `[book]` 中放 `publisher` 或单数 `author`。出版社、副书名、ISBN、CIP、定价、版权声明等出版元数据统一放入 `book.toml` 的 `[bookcraft.*]` 扩展段。

## mdBook 标准字段与 Bookcraft 扩展字段

| 字段/表 | mdBook 标准 | 当前模板是否包含 | 用途 | 处理方式 |
|---|---|---|---|---|
| `[book].title` | 是 | 是 | mdBook 书名；Word 扉页书名 | mdBook 直接读取 |
| `[book].authors` | 是 | 是 | mdBook 作者列表；Word 扉页作者 | mdBook 直接读取；脚本会合并为作者字符串 |
| `[book].description` | 是 | 是 | mdBook 简介；可同步为内容摘要 | mdBook 直接读取 |
| `[book].src` | 是 | 是 | mdBook 源目录 | mdBook 直接读取 |
| `[book].language` | 是 | 是 | 语言 | mdBook 直接读取 |
| `[book].text-direction` | 是 | 是 | 文本方向 | mdBook 直接读取 |
| `[bookcraft.book].subtitle` | 否 | 是 | 副书名 | mdBook 忽略；Word 脚本读取 |
| `[bookcraft.description]` | 否 | 是 | 内容摘要、详细简介、关键词、目标读者、分类 | mdBook 忽略；Word 版权页读取 |
| `[bookcraft.publisher]` | 否 | 是 | 出版方名称、地址、网址 | mdBook 忽略；Word 版权页读取 |
| `[bookcraft.edition]` | 否 | 是 | 版次、印次、开本、ISBN、定价 | mdBook 忽略；Word 版权页读取 |
| `[bookcraft.copyright_notice]` | 否 | 是 | 版权声明、免责声明 | mdBook 忽略；Word 版权页读取 |
| `[bookcraft.cip]` | 否 | 是 | CIP 数据 | mdBook 忽略；Word 版权页读取 |
| `[bookcraft.export]` | 否 | 是 | DOCX 输出路径、Word 配置、脚本路径 | mdBook 忽略；Agent 和导出脚本参考 |
| `[bookcraft.workflow.auto_writing]` | 否 | 是 | 自动写书的无人干预、完整写作、自动研究、自动修订、DOCX 导出和重试开关 | mdBook 忽略；Agent 执行自动写书时读取 |
| `[bookcraft.layout.*]` | 否 | 是 | 经典中文图书排版参数 | mdBook 忽略；Word 导出模块读取 |
| `[bookcraft.content_rules.*]` | 否 | 是 | 中文图书内容规范参数，包括大纲数量、字数和引文规划 | mdBook 忽略；规划、写作和审阅模块读取 |
| `[bookcraft.normalization]` | 否 | 是 | 术语替换、规范拼写和禁用拼写 | mdBook 忽略；写作和审阅模块读取 |

配置原则：

- `[book]` 只放 mdBook 标准字段，确保 `mdbook build` 不受影响。
- `[bookcraft.*]` 是本项目的自定义命名空间，mdBook 不使用这些字段，但 TOML 解析器会保留它们。
- `scripts/md2word.py` 只读取 `book.toml` 的 `[bookcraft.*]` 渲染版权页。
- 如果未来需要新增出版字段，优先放入 `[bookcraft.*]`，不要污染 `[book]`。

## book.toml 扩展字段

`book.toml` 同时包含自定义 `[bookcraft.*]` 扩展段。mdBook 会读取自己认识的标准字段，忽略这些额外配置；`scripts/md2word.py` 会读取 `[bookcraft.*]` 渲染紧凑的一页版权页。

核心字段：

- `[bookcraft.book]`：`subtitle`
- `[bookcraft.description]`：`summary`、`text`、`keywords`、`target_readers`、`category`
- `[bookcraft.publisher]`：`name`、`address`、`website`
- `[bookcraft.edition]`：`version`、`printing`、`format`、`isbn`、`price`
- `[bookcraft.copyright_notice]`：`text`、`disclaimer`
- `[bookcraft.cip]`：`text`
- `[bookcraft.export]`：`docx`、`word_config`、`script`
- `[bookcraft.workflow.auto_writing]`：`end_to_end`、`ask_for_confirmation`、`complete_all_chapters`、`auto_research`、`auto_review_and_revise`、`export_docx`、`retry_recoverable_errors`、`finish_only_after_docx_validation`
- `[bookcraft.layout.page]`：纸张、页边距、页眉页脚、奇偶页/首页设置
- `[bookcraft.layout.grid]`：每行字符数、每页行数
- `[bookcraft.layout.body]`：正文字体、字号、缩进、行距、对齐方式
- `[bookcraft.layout.headings]`：目录深度、标题层级、章标题编号、标题间过渡文字
- `[bookcraft.layout.code]`：代码字体、字号、行距、缩进、注释要求
- `[bookcraft.layout.figures_tables]`：图表题注、嵌入版式、图表位置和引用要求
- `[bookcraft.content_rules.*]`：书名、目录、大纲数量/字数/引文、内容简介、前言、语言、界面术语、操作步骤、图、物理量、参考文献和随书资源要求
- `[bookcraft.normalization]`：项目指定的词语替换、正确拼写和禁用拼写

创建项目时用 `[book]` 的 `title`、`authors`、`description` 初始化 `[bookcraft.*]` 中的书名相关展示信息；副书名、出版社、ISBN、CIP、定价等信息以 `[bookcraft.*]` 为准。

作者字段按以下优先级确定：用户对话中提供的姓名 > `蜡笔小歆`。不得使用机构名称替代作者姓名，除非用户明确要求机构署名。

副书名不得超过 15 字，必须是一句简短、完整的话，不得用冒号、破折号或竖线拆成两部分。

## 目标目录结构输出要求

创建或更新项目后，必须输出目标目录结构：

- 目录树要反映真实文件系统中的创建结果，而不是只输出推荐模板。
- 如果是新项目，说明项目根目录、已创建文件和后续需要补写的章节。
- 如果是已有项目，先保留原有文件，再列出新增或更新的文件。
- `exports/` 是导出结果目录；项目刚创建时可以为空，不要伪造 `book.docx`。
- `scripts/md2word.py` 必须复制到项目内，便于离线导出和后续修改。
- 如果旧项目已有 `src/front-matter/title-page.md` 或 `copyright-summary.md`，可以保留作为历史资料，但组装 `manuscript.md` 时必须排除，不能加入 Word 正文。
- 推荐序在成品中统一命名为"序言一""序言二""序言三"，最多三篇，默认不创建任何推荐序文件；只有存在真实推荐内容时才按顺序创建相应文件。
- 自序只有一篇且必须创建，文件固定为 `src/front-matter/author-preface.md`。
- 后置页默认只创建必选的 `src/back-matter/references.md`。结语、后记、致谢、附录、术语表、作者简介等均为可选项，没有用户要求或真实内容时不创建、不加入 `SUMMARY.md`、不加入 `manuscript.md`。
- `references.md` 必须包含至少一条经过核实的参考文献，不能只写"待补充""暂无"等占位文字。资料不足时先完成资料整理或停止导出，不得虚构文献。
- 普通模式只有在用户要求导出、且实际生成文件后，才在 `exports/` 中写入 Word/DOCX 成品；自动写书模式已默认包含导出任务，必须自动生成 Word/DOCX 成品。

章节文件数量应与当前大纲匹配。普通规划模式如果只有初步大纲，可以先创建前三章骨架；自动写书模式必须一次创建完整大纲对应的全部章节文件，并持续写完，不得留下"后续章节待扩展"。
