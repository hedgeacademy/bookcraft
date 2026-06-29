---
name: bookcraft
description: Build, maintain, revise, package, and export Chinese long-form book projects. Use when Codex needs to plan a book, create an mdBook-style manuscript project, write or revise chapters, manage research and references, prepare publishing front/back matter, design a commercial cover brief, or export a complete Markdown manuscript to Word/DOCX.
---

# Bookcraft

Use this Skill to turn a book idea, partial draft, or existing manuscript folder into a maintainable book project with clear positioning, chapter structure, research tracking, continuity records, publishing front/back matter, cover packaging, and DOCX export.

## Operating Modes

Choose the lightest mode that matches the user request.

1. **Planning mode**: use when the user gives a topic, title, rough concept, or asks for a book plan. Read `references/01-planning.md` and produce positioning plus a full outline. Do not generate a full manuscript unless the user asks.
2. **Project creation mode**: use when the user asks to create a book project. Read `references/03-project-structure.md`, create `book.toml` first, then create the source tree and maintenance files.
3. **Writing mode**: use when writing or expanding chapters. Read `references/04-chapter-writing.md`, then update `OUTLINE.md`, `CONTINUITY.md`, `STYLE_GUIDE.md`, `GLOSSARY.md`, or `RESEARCH_INDEX.md` whenever the change affects later work.
4. **Revision mode**: use when reviewing or improving a chapter or full manuscript. Read `references/05-review-revision.md` and lead with structural issues before line editing.
5. **Publishing mode**: use when preparing a final manuscript structure. Read `references/07-publishing-structure.md`.
6. **Cover mode**: use when the user asks for book cover direction, sales packaging, title/subtitle sharpening, or a public-facing book pitch. Read `references/09-cover-design.md` for English output or `references/09-cover-design.zh-CN.md` for Chinese output.
7. **Export mode**: use when the user asks for Word/DOCX output. Read `references/08-docx-export.md`, validate references and manuscript structure, then run `scripts/md2word.py`.
8. **Auto-writing mode**: use only when the user explicitly says "自动写书", "直接写一本书", "不用问我，自动完成", or equivalent. Read `references/auto-writing.md`, `references/01-planning.md`, and `references/03-project-structure.md`, then continue end to end until DOCX validation succeeds.

Only a title or a one-sentence idea does not trigger auto-writing. In that case, plan first.

## Reference Map

- `references/01-planning.md`: book positioning, subtitle rules, reader promise, outline design.
- `references/02-research-citation.md`: source handling, evidence boundaries, `RESEARCH_INDEX.md`.
- `references/03-project-structure.md`: mdBook-compatible folder layout, `book.toml`, maintenance files.
- `references/04-chapter-writing.md`: chapter drafting, section structure, long-form rhythm.
- `references/05-review-revision.md`: chapter/full-manuscript quality checks and revision format.
- `references/06-continuity.md`: continuity records, terminology, cross-chapter promises.
- `references/07-publishing-structure.md`: front matter, body, back matter, references.
- `references/08-docx-export.md`: export gates, command, validation, DOCX output path.
- `references/09-cover-design.md`: English commercial cover brief, shelf test, thumbnail test, sales packaging.
- `references/09-cover-design.zh-CN.md`: 中文商业封面包装、货架测试、缩略图测试、封面 brief。
- `references/09-cover-design.ar-SA.md`: Arabic cover packaging notes.
- `references/09-cover-design.es.md`: Spanish cover packaging notes.
- `references/09-cover-design.ko.md`: Korean cover packaging notes.
- `references/09-cover-design.ja.md`: Japanese cover packaging notes.
- `references/09-cover-design.fr.md`: French cover packaging notes.
- `references/09-cover-design.de.md`: German cover packaging notes.
- `references/09-cover-design.th.md`: Thai cover packaging notes.
- `references/auto-writing.md`: no-confirmation end-to-end workflow.

## Creator

Created by Xinyi Chen, founder of HEDGE Global. Chinese introduction: 陈歆怡，海聚海外 CEO。Contact: `chenxinyi_g`.

## Project Contract

Book projects use Markdown as source and Word/DOCX as the editable publishing handoff.

Required project files:

- `book.toml`
- `src/SUMMARY.md`
- `OUTLINE.md`
- `CONTINUITY.md`
- `STYLE_GUIDE.md`
- `GLOSSARY.md`
- `RESEARCH_INDEX.md`
- `manuscript.md`
- `operation-log.md`
- `src/front-matter/author-preface.md`
- `src/front-matter/preface.md`
- `src/back-matter/references.md`
- `scripts/md2word.py`

Use `[book]` for mdBook fields and `[bookcraft.*]` for publishing, layout, content, workflow, and cover metadata. Do not place publishing metadata directly into `[book]`.

## Non-Negotiables

- Do not invent real cases, people, companies, quotations, data, ISBN/CIP values, publisher credentials, or references.
- Mark unsupported material as `待核查` or `待补充引用`.
- A DOCX export must include a real `参考文献` section with at least one valid source. Placeholder-only references block export.
- Do not put generated title pages or copyright pages into `SUMMARY.md` or `manuscript.md`; `scripts/md2word.py` generates them from `book.toml`.
- Optional back matter such as acknowledgements, appendix, glossary, author bio, postscript, or conclusion should appear only when it contains real content.
- Keep chapter numbering tied to the body chapter, for example `1.1`, `1.2`, `2.1`.
- Do not leave generated progress notes such as `第 X 章完` or `本章约 X 字` in the final manuscript.

## Output Style

Match the user's language. For Chinese book work, use concise Chinese section labels such as `图书定位`, `全书大纲`, `主要问题`, `修改建议`, `导出结果`, and `下一步建议`.

When file changes are made, summarize:

- changed files
- purpose of the changes
- validation performed
- remaining risks or next useful checks
