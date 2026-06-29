# Bookcraft

Bookcraft is a Codex Skill for building Chinese long-form book projects: positioning, outlines, research tracking, chapter writing, continuity, revision, publishing structure, commercial cover briefs, and Word/DOCX export.

It is designed for authors, editors, consultants, educators, and builders who want an AI-assisted writing workflow that produces maintainable source files instead of one-off long answers.

## Language Versions

- [English](README.en.md)
- [中文](README.zh-CN.md)

## Author and Contact

Bookcraft is created by Xinyi Chen, founder of HEDGE Global. Chinese introduction: 陈歆怡，海聚海外 CEO。

Contact: `chenxinyi_g`

## What It Does

- Plans a book from a title, topic, rough idea, or existing draft.
- Creates an mdBook-compatible manuscript project.
- Writes and revises chapters while maintaining continuity files.
- Tracks research sources, citations, unresolved claims, and reference status.
- Prepares publishing front matter and back matter.
- Builds a commercial cover brief using a shelf-first sales packaging method.
- Exports `manuscript.md` to editable Word/DOCX through `scripts/md2word.py`.

## Skill Structure

```text
bookcraft/
├── SKILL.md
├── AGENTS.md
├── README.md
├── README.en.md
├── README.zh-CN.md
├── references/
│   ├── 01-planning.md
│   ├── 02-research-citation.md
│   ├── 03-project-structure.md
│   ├── 04-chapter-writing.md
│   ├── 05-review-revision.md
│   ├── 06-continuity.md
│   ├── 07-publishing-structure.md
│   ├── 08-docx-export.md
│   ├── 09-cover-design.md
│   ├── 09-cover-design.zh-CN.md
│   └── auto-writing.md
├── templates/
│   ├── book.toml
│   ├── summary.md
│   └── word.json
└── scripts/
    └── md2word.py
```

## Install

Copy this repository into your Codex skills directory:

```bash
mkdir -p ~/.codex/skills
cp -R bookcraft ~/.codex/skills/bookcraft
```

Then start a new Codex thread and invoke it with:

```text
Use $bookcraft to plan a book about [topic].
```

If your Codex setup supports implicit Skill invocation, requests about writing, revising, exporting, or packaging a book should trigger the Skill automatically.

## Example Prompts

```text
Use $bookcraft to plan a business book about AI workflows for small companies.
```

```text
Use $bookcraft to create a book project for this outline and prepare the first three chapter files.
```

```text
Use $bookcraft to review chapter 2 for structure, evidence, and continuity.
```

```text
Use $bookcraft to build a commercial cover brief for this manuscript.
```

```text
Use $bookcraft to export this manuscript to Word/DOCX.
```

## Auto-Writing Mode

Bookcraft only enters full auto-writing mode when the user explicitly asks for it, for example:

```text
自动写一本书：《书名》
不用问我，自动完成并导出 Word。
```

In auto-writing mode, the Skill creates the project, writes all planned chapters, checks continuity, validates references, assembles `manuscript.md`, and runs DOCX export. If the user only provides a title or idea, Bookcraft plans first instead of silently generating a full book.

## DOCX Export Requirements

The exporter requires:

- Python 3.11 or newer.
- `python-docx`.
- `pandoc`.
- A complete `manuscript.md`.
- A valid `book.toml`.
- A `# 参考文献` section with at least one real source.

Install dependencies:

```bash
python3 -m pip install python-docx
```

On macOS:

```bash
brew install pandoc
```

On Debian or Ubuntu:

```bash
sudo apt-get install pandoc
```

Run:

```bash
python3 scripts/md2word.py --input manuscript.md
```

## Configuration

Bookcraft uses:

- `[book]` for mdBook standard fields.
- `[bookcraft.*]` for publishing metadata, layout settings, content rules, workflow behavior, and cover packaging.

Generated book projects should keep `book.toml` as the source of truth for title, author, subtitle, publishing metadata, export settings, and layout defaults.

## Cover Method

The cover module is documented in `references/09-cover-design.md`. It turns title, subtitle, audience, and manuscript promise into:

- a one-sentence sales hook
- a visual symbol
- hierarchy rules for title/subtitle/author
- shelf and thumbnail checks
- 3 cover directions
- prompt-ready creative briefs

The method is written as a neutral public framework for commercial book packaging.

## License

MIT
