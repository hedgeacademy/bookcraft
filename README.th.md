# Bookcraft

Bookcraft คือ Codex Skill สำหรับวางแผน เขียน ทบทวน จัดแพ็กเกจ และส่งออกโปรเจกต์หนังสือขนาดยาว พร้อมแนวทางทำ brief ปกหนังสือเชิงพาณิชย์ และการส่งออกเป็น Word/DOCX

## ผู้สร้างและช่องทางติดต่อ

Created by Xinyi Chen, founder of HEDGE Global.

คำแนะนำภาษาจีน: 陈歆怡，海聚海外 CEO。

ติดต่อ: `chenxinyi_g`

## Skill นี้ทำอะไรได้บ้าง

- วางตำแหน่งหนังสือ คำมั่นสัญญาต่อผู้อ่าน ชื่อเรื่อง ชื่อรอง และโครงร่างทั้งเล่ม
- สร้างโปรเจกต์แบบ mdBook พร้อม `book.toml`, `src/SUMMARY.md`, ไฟล์ติดตามงาน และไฟล์บทต่าง ๆ
- จัดการงานวิจัย แหล่งอ้างอิง ประเด็นที่ต้องตรวจสอบ และบรรณานุกรม
- เขียน ขยาย ทบทวน และแก้ไขบท โดยรักษาความต่อเนื่อง คู่มือสไตล์ และอภิธานศัพท์
- จัดโครงสร้างสำหรับการตีพิมพ์ เช่น คำนำ บทหลัก เอกสารอ้างอิง และภาคผนวกเสริม
- ทำแพ็กเกจปกหนังสือเชิงพาณิชย์: sales hook, สัญลักษณ์ภาพ, ลำดับชั้นของชื่อเรื่อง, shelf test และ thumbnail test
- ส่งออก `manuscript.md` เป็น Word/DOCX ผ่าน `scripts/md2word.py`

## การติดตั้ง

```bash
mkdir -p ~/.codex/skills
git clone https://github.com/hedgeacademy/bookcraft ~/.codex/skills/bookcraft
```

เปิดเธรดใหม่ใน Codex แล้วใช้:

```text
Use $bookcraft to plan a book about [topic].
```

## ตัวอย่าง Prompt

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

## วิธีคิดเรื่องปก

ปกหนังสือที่ดีไม่ใช่แค่สวย แต่ต้องสื่อสารเชิงพาณิชย์ได้ทันที ภายในไม่กี่วินาทีผู้อ่านควรรู้ว่า หนังสือเล่มนี้ว่าด้วยอะไร เขียนให้ใคร สัญญาอะไรกับผู้อ่าน และทำไมควรเลือกเล่มนี้แทนหนังสือที่คล้ายกัน

ดูเพิ่มเติม:

- `references/09-cover-design.th.md`
- `references/09-cover-design.md`

## ข้อกำหนดการส่งออก

การส่งออก DOCX ต้องมี Python, `python-docx`, `pandoc`, ไฟล์ `manuscript.md` ที่สมบูรณ์, `book.toml` ที่ถูกต้อง และส่วนเอกสารอ้างอิงจริง

## License

MIT
