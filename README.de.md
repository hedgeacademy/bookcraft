# Bookcraft

Bookcraft ist eine Codex Skill für die Planung, Erstellung, Überarbeitung, Verpackung und den Export längerer Buchprojekte, einschließlich kommerzieller Cover-Briefs und Word/DOCX-Ausgabe.

## Autorin und Kontakt

Erstellt von Xinyi Chen, founder of HEDGE Global.

Chinesische Vorstellung: 陈歆怡，海聚海外 CEO。

Kontakt: `chenxinyi_g`

## Was die Skill abdeckt

- Buchpositionierung, Leserprechen, Titel, Untertitel und vollständige Gliederung.
- Erstellung eines mdBook-ähnlichen Projekts mit `book.toml`, `src/SUMMARY.md`, Verwaltungsdateien und Kapiteln.
- Verwaltung von Recherche, Quellen, Zitaten, Prüfpunkten und Literatur.
- Kapitelentwurf, Überarbeitung, Kontinuität, Stilhandbuch und Glossar.
- Veröffentlichungsstruktur: Vorspann, Kapitel, Referenzen und optionale Anhänge.
- Kommerzielles Cover-Packaging: Verkaufsversprechen, visuelles Symbol, Hierarchie, Regaltest und Thumbnail-Test.
- Export von `manuscript.md` nach Word/DOCX über `scripts/md2word.py`.

## Installation

```bash
mkdir -p ~/.codex/skills
git clone https://github.com/hedgeacademy/bookcraft ~/.codex/skills/bookcraft
```

Starten Sie einen neuen Codex-Thread und verwenden Sie:

```text
Use $bookcraft to plan a book about [topic].
```

## Häufige Prompts

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

## Cover-Methode

Ein Cover soll nicht nur schön sein, sondern kommerziell sofort verständlich. Eine starke Richtung beantwortet in wenigen Sekunden: Worum geht es, für wen ist es, welches Versprechen gibt es und warum sollte man dieses Buch statt ähnlicher Bücher wählen?

Siehe:

- `references/09-cover-design.de.md`
- `references/09-cover-design.md`

## Exportanforderungen

Der DOCX-Export benötigt Python, `python-docx`, `pandoc`, ein vollständiges `manuscript.md`, ein gültiges `book.toml` und einen echten Referenzabschnitt.

## Lizenz

MIT
