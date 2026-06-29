# بوك كرافت Bookcraft

Bookcraft مهارة مخصصة لـ Codex تساعد على تخطيط الكتب الطويلة وكتابتها ومراجعتها وتجهيزها للنشر، مع إعداد موجز تجاري لتصميم الغلاف وتصدير المخطوط إلى Word/DOCX.

## المؤلفة وبيانات التواصل

المؤلفة: Xinyi Chen، مؤسسة HEDGE Global.

التعريف بالصينية: 陈歆怡，海聚海外 CEO。

التواصل: `chenxinyi_g`

## ماذا تقدم المهارة؟

- تحويل فكرة أو عنوان أو مسودة إلى خطة كتاب واضحة.
- إنشاء مشروع مخطوط بنمط mdBook يحتوي على `book.toml` وملفات الفصول وملفات المتابعة.
- إدارة البحث والمراجع والاقتباسات ونقاط التحقق.
- كتابة الفصول ومراجعتها مع الحفاظ على الاتساق والأسلوب والمصطلحات.
- تجهيز بنية النشر: المقدمة، الفصول، المراجع، والملحقات الاختيارية.
- إعداد اتجاهات غلاف تجارية: وعد القارئ، الرمز البصري، ترتيب العنوان، اختبار الرف، واختبار الصورة المصغرة.
- تصدير `manuscript.md` إلى ملف Word/DOCX قابل للتحرير.

## التثبيت

```bash
mkdir -p ~/.codex/skills
git clone https://github.com/hedgeacademy/bookcraft ~/.codex/skills/bookcraft
```

ابدأ محادثة جديدة في Codex واستخدم:

```text
Use $bookcraft to plan a book about [topic].
```

## أمثلة للاستخدام

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

## منهج الغلاف

يعتمد منهج الغلاف على الوضوح التجاري: يجب أن يوضح الغلاف بسرعة موضوع الكتاب، والجمهور المستهدف، والوعد الأساسي، وسبب اختيار هذا الكتاب بدلا من الكتب المشابهة.

انظر:

- `references/09-cover-design.ar-SA.md`
- `references/09-cover-design.md`

## متطلبات التصدير

يتطلب تصدير DOCX وجود Python و`python-docx` و`pandoc` وملف `manuscript.md` مكتمل وملف `book.toml` صالح وقسم مراجع حقيقي.

## الترخيص

MIT
