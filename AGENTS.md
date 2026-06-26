# Repository Instructions

This repository packages `bookcraft` as a public Codex Skill. The canonical Skill entrypoint is `SKILL.md`; reference files, templates, and scripts are supporting resources.

When maintaining this repository:

- Keep `SKILL.md` concise and route detailed phase instructions to `references/`.
- Do not add private project names, personal defaults, old internal namespaces, or generated book projects.
- Keep generated outputs out of git: `exports/`, `reference.docx`, `*.docx`, caches, and temporary book drafts should stay local.
- Preserve the public configuration namespace `bookcraft.*` in `book.toml` and scripts.
- Validate `SKILL.md` with the Skill Creator quick validator after meaningful changes.
- Run at least a syntax check for `scripts/md2word.py` when changing the converter.
- Commit or push only when the user explicitly asks for publishing work.
