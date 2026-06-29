# Bookcraft

Bookcraft は、長編書籍プロジェクトの企画、執筆、レビュー、出版構成、商業的な表紙ブリーフ作成、Word/DOCX 書き出しを支援する Codex Skill です。

## 作者と連絡先

Created by Xinyi Chen, founder of HEDGE Global.

中国語での紹介: 陈歆怡，海聚海外 CEO。

連絡先: `chenxinyi_g`

## できること

- 本のポジショニング、読者への約束、タイトル、サブタイトル、全体構成の設計。
- `book.toml`、`src/SUMMARY.md`、管理ファイル、章ファイルを含む mdBook 形式のプロジェクト作成。
- 調査資料、引用、要確認事項、参考文献の管理。
- 章ごとの執筆、加筆、レビュー、修正、継続性の維持。
- 出版構成の整理: 前付、本文、参考文献、任意の後付。
- 商業的な表紙パッケージ: 販売フック、視覚シンボル、タイトル階層、棚テスト、サムネイルテスト。
- `scripts/md2word.py` による `manuscript.md` から Word/DOCX への書き出し。

## インストール

```bash
mkdir -p ~/.codex/skills
git clone https://github.com/hedgeacademy/bookcraft ~/.codex/skills/bookcraft
```

新しい Codex スレッドで次のように使います。

```text
Use $bookcraft to plan a book about [topic].
```

## よく使うプロンプト

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

## 表紙メソッド

表紙は単なる装飾ではなく、商業的にすぐ理解されるシグナルです。優れた表紙は数秒で、何の本か、誰のための本か、何を約束するか、なぜ選ぶべきかを伝えます。

参照:

- `references/09-cover-design.ja.md`
- `references/09-cover-design.md`

## 書き出し要件

DOCX 書き出しには Python、`python-docx`、`pandoc`、完成した `manuscript.md`、有効な `book.toml`、実在する参考文献セクションが必要です。

## ライセンス

MIT
