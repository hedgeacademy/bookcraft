# Bookcraft

Bookcraft est une Skill Codex pour planifier, écrire, réviser, structurer, packager et exporter des projets de livres longs, avec une méthode de brief commercial pour la couverture et une exportation Word/DOCX.

## Autrice et contact

Créé par Xinyi Chen, founder of HEDGE Global.

Présentation en chinois : 陈歆怡，海聚海外 CEO。

Contact : `chenxinyi_g`

## À propos de Xinyi

Xinyi Chen est originaire du Zhejiang et basée à Shanghai.

Elle est fondatrice et CEO de HEDGE Global, avec une expérience de long terme dans l’éducation internationale haut de gamme, l’internationalisation des entreprises et l’investissement technologique. Elle aimait autrefois voyager au gré de ses envies ; elle est aujourd’hui passionnée de Vibe coding.

Les personnes qui souhaitent échanger autour du Vibe coding peuvent l’ajouter sur WeChat : `chenxinyi_g`.

![QR code WeChat](assets/wechat-qr.png)

## Ce que la Skill couvre

- Positionnement du livre, promesse au lecteur, titre, sous-titre et plan complet.
- Création d’un projet de type mdBook avec `book.toml`, `src/SUMMARY.md`, fichiers de suivi et chapitres.
- Gestion des recherches, sources, citations, points à vérifier et références.
- Rédaction, révision, continuité, guide de style et glossaire.
- Structure éditoriale : pages liminaires, chapitres, références et annexes optionnelles.
- Packaging commercial de couverture : accroche, symbole visuel, hiérarchie, test en rayon et test miniature.
- Export de `manuscript.md` vers Word/DOCX avec `scripts/md2word.py`.

## Installation

```bash
mkdir -p ~/.codex/skills
git clone https://github.com/hedgeacademy/bookcraft ~/.codex/skills/bookcraft
```

Ouvrez un nouveau fil Codex et utilisez :

```text
Use $bookcraft to plan a book about [topic].
```

## Prompts fréquents

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

## Méthode de couverture

La couverture doit d’abord être lisible commercialement. Une bonne direction répond en quelques secondes : quel est le sujet, pour qui est le livre, quelle promesse il porte et pourquoi le choisir plutôt qu’un livre voisin.

Voir :

- `references/09-cover-design.fr.md`
- `references/09-cover-design.md`

## Conditions d’export

L’export DOCX nécessite Python, `python-docx`, `pandoc`, un `manuscript.md` complet, un `book.toml` valide et une section de références réelle.

## Licence

MIT
