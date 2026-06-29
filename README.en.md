# Bookcraft

Bookcraft is a Codex Skill for book writing and commercial cover packaging. It helps an AI agent plan, write, revise, package, and export a maintainable Chinese long-form book project.

## Author and Contact

Created by Xinyi Chen, founder of HEDGE Global.

Chinese introduction: 陈歆怡，海聚海外 CEO。

Contact: `chenxinyi_g`

## What It Covers

- Book positioning, reader promise, title, subtitle, and full outline.
- mdBook-style project creation with `book.toml`, `src/SUMMARY.md`, maintenance files, and chapter files.
- Research and citation tracking through `RESEARCH_INDEX.md`.
- Chapter writing, chapter review, continuity maintenance, style guide, and glossary management.
- Publishing structure: title page, copyright page, preface, body chapters, references, and optional back matter.
- Commercial cover packaging: sales hook, visual symbol, hierarchy, shelf test, thumbnail test, and prompt-ready creative briefs.
- Word/DOCX export through `scripts/md2word.py`.

## Install

```bash
mkdir -p ~/.codex/skills
git clone https://github.com/hedgeacademy/bookcraft ~/.codex/skills/bookcraft
```

Start a new Codex thread and use:

```text
Use $bookcraft to plan a book about [topic].
```

## Common Prompts

```text
Use $bookcraft to plan a business book about AI workflows for small companies.
```

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

## Cover Method

The cover workflow is designed around commercial clarity. A strong direction should answer, within a few seconds:

- What is the book about?
- Who is it for?
- What promise does it make?
- Why should this book be chosen instead of adjacent books?

For the full cover framework, see:

- `references/09-cover-design.md`
- `references/09-cover-design.zh-CN.md`

## Export Requirements

DOCX export requires Python, `python-docx`, `pandoc`, a complete `manuscript.md`, a valid `book.toml`, and at least one real reference under `# 参考文献`.

```bash
python3 -m pip install python-docx
brew install pandoc
python3 scripts/md2word.py --input manuscript.md
```

## License

MIT
