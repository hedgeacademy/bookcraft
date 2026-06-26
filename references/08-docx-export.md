# Phase 8：Word/DOCX 导出

最终基于 mdBook 兼容的 Markdown 源稿导出 Word/DOCX 成品。

## 导出前检查

- Markdown 源稿结构完整
- `book.toml` 已创建，且 `[book]` 包含 mdBook 标准字段
- `book.toml` 已包含 `[bookcraft.*]` 扩展段，用于版权页、CIP、ISBN、内容摘要、出版发行、版次印次、定价和版权声明
- `SUMMARY.md` 目录正确
- `manuscript.md` 已按 Word 成品顺序组装
- `scripts/md2word.py` 已复制到项目内
- 所有章节文件存在
- 图片、表格、脚注、引用和附录路径正确
- 标题层级符合出版结构
- 前置页、正文和后置页顺序正确
- 参考文献格式统一
- `manuscript.md` 必须包含"参考文献"一级标题和至少一条真实有效文献；缺失、为空或只有占位文字时禁止导出
- 结语、后记、致谢、附录、术语表、作者简介等可选后置页没有真实内容时不得加入 `manuscript.md`
- Word/DOCX 导出样式符合项目配置

导出的 Word 文件应便于后续编辑、校对和出版社排版。

## 排版默认值

参考旧版 `scripts/md2word.py` 的经典中文图书排版：

- 页面：16 开，18.4cm × 26cm
- 页边距：上 2.5cm，下 1.8cm，左 2cm，右 2cm
- 页眉/页脚距离：页眉 1.7cm，页脚 1.3cm
- 正文：宋体，五号 10.5pt，首行缩进两字符，两端对齐，1.5 倍行距
- 一级标题：黑体，三号 16pt，加粗，居中，章前自动分页
- 二级标题：黑体，四号 14pt，加粗
- 三级标题：黑体，小四 12pt，加粗
- 章标题编号：使用半角数字，例如 `第3章`，不要使用 `第三章`
- 自动生成目录，三级深度
- 扉页单独一页
- 版权页紧接扉页之后，由 `book.toml` 的 `[bookcraft.*]` 渲染，压缩在一页内
- 页眉：正文章节显示标题 + 细线；扉页、版权页、序言一至序言三、自序、前言和目录无页眉
- 页脚：扉页、版权页、序言一至序言三、自序和前言不显示页码；目录从 `I` 开始使用大写罗马数字；正文第一章从 `001` 重新开始，始终显示三位数字
- 图题/表题：居中，图号或表号与名称之间空两个空格；表题在表上方

不要询问用户排版偏好，默认使用以上参数。

## 排版参数来源

排版规范已参数化，主要落在两个位置。本模块负责在导出时读取和执行这些排版参数。

- `book.toml` 的 `[bookcraft.layout.page]`：纸张、页边距、页眉页脚距离、奇偶页不同、首页不同。
- `book.toml` 的 `[bookcraft.layout.grid]`：每行 40 字、每页 40 行。
- `book.toml` 的 `[bookcraft.layout.body]`：正文字体、字号、首行缩进、行距、两端对齐、正文中不手动设置字体字号。
- `book.toml` 的 `[bookcraft.layout.headings]`：最多三级标题、目录深度、章标题编号样式、标题间过渡文字要求。
- `book.toml` 的 `[bookcraft.layout.code]`：代码字体 Courier New、8pt、1.5 倍行距、缩进和注释要求。
- `book.toml` 的 `[bookcraft.layout.figures_tables]`：图表题注居中、题注上下空行、图号/表号与名称之间空两个空格、表题在表上方、图片嵌入版式。
- `config/word.json`：导出脚本实际使用的 Word 排版参数副本，创建项目时应从 `book.toml` 同步生成。

若 `book.toml` 与 `config/word.json` 不一致，以 `book.toml` 的 `[bookcraft.layout.*]` 为准，并同步更新 `config/word.json`。

## manuscript.md 组装

导出前先组装完整书稿顺序：

1. 序言一至序言三：可选推荐序，默认没有；按 `recommendation-preface-01.md` 至 `recommendation-preface-03.md` 的存在情况顺序加入，最多三篇。
2. 自序：`src/front-matter/author-preface.md`
3. 前言：`src/front-matter/preface.md`
4. 正文章节：`src/chapters/chapter-XX.md`
5. 后置内容：必须追加含有效内容的 `src/back-matter/references.md`；其他可选后置页只在存在真实内容时追加

扉页由命令参数和 `book.toml` 生成，版权页由 `book.toml` 的 `[bookcraft.*]` 生成。`manuscript.md`、`src/SUMMARY.md` 均不得包含 `title-page.md`、`copyright-summary.md`、扉页正文或版权页正文。

旧项目即使保留 `src/front-matter/title-page.md` 和 `copyright-summary.md`，也只能作为历史资料，组装 Word 输入时必须排除。导出脚本还会清理常见的重复扉页、版权页标题，但不能以此替代正确组装。

## 导出执行要求

- 优先使用项目内 `scripts/md2word.py` 生成 `.docx`。该脚本来自本项目的 `scripts/md2word.py`，基于 pandoc + python-docx。
- `scripts/md2word.py` 会检查 pandoc 和 python-docx，生成 `reference.docx` 模板，调用 pandoc 转换 Markdown，再用 python-docx 后处理。
- Word 输出文件名默认使用 `书名-作者-YYYYMMDD-HHmm.docx`，例如 `马斯克传-作者-20260626-1530.docx`，保存在 `exports/` 目录，用于后续编辑、批注、校对和出版社排版。
- 脚本会优先在当前工作目录、书稿目录及其父目录查找 `book.toml`；读取 `[book]` 中的书名和作者，以及 `[bookcraft.*]` 中的副书名、出版和导出配置。
- 导出完成后必须检查文件是否存在，并在回复中给出绝对路径。
- 如果导出失败，不要声称已生成成品。普通模式说明失败原因；自动写书模式必须先自行诊断、安装或切换可用依赖、修复书稿和配置并重试，只有遇到无法自行恢复的硬故障时才报告中断。

推荐命令：

```bash
python scripts/md2word.py \
  --input manuscript.md \
  --title "书名" \
  --subtitle "副标题" \
  --author "作者名"
```

## 脚本后处理能力

`scripts/md2word.py` 会自动完成：

- 生成经典中文排版 `reference.docx`
- 转换 Markdown 到 Word
- 插入扉页和版权页
- 清理旧项目误拼入正文的重复扉页和版权页
- 读取 `book.toml` 的 `[bookcraft.*]` 渲染版权页
- 将版权页压缩为一页小字号版式
- 将目录放置在前言之后、第一章之前
- 处理自序、前言和可选推荐序等前辅文的页眉页脚
- 将自序文末的作者落款和"月日于地点"两行右对齐，并在落款前保留一个空行
- 正文章节页眉显示章节标题和细线
- 目录页脚居中显示大写罗马数字页码，从 `I` 开始
- 正文页脚居中显示三位数字页码，从 `001` 开始
- 调整首行缩进、居中段落、引用和列表段落
- 为表格添加题注和灰色表头
- 自动移除"章前引文"标签，仅保留引文正文
- 自动纠正节、小节编号，使其跟随所属章号
- 自动移除章末"第 X 章完""本章约 X 字""本章字数"等生成提示

## 验证输出

导出后至少验证：

- `exports/书名-作者-YYYYMMDD-HHmm.docx` 存在
- 文件大小合理，通常应大于 5KB
- `manuscript.md` 包含预期顶级标题
- `book.toml` 中 `[bookcraft.*]` 的内容摘要、ISBN/CIP、版权声明已进入 Word 版权页
- Word 成品包含"参考文献"及至少一条有效文献；若缺失则导出必须失败
- Word 成品没有空白或占位的可选后置页
- 验证结果写入 `operation-log.md`

如果用户要求"按出版社格式写书"，优先保证图书结构完整、章节层级清楚、版权页和前后置页齐全、正文版式整洁、引用和参考文献规范，并导出可编辑的 Word/DOCX 文档，而不是只生成普通 Markdown 文档。
