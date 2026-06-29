# Bookcraft

Bookcraft es una Skill de Codex para planificar, escribir, revisar, empaquetar y exportar proyectos de libros largos, con metodología de brief comercial para portadas y exportación a Word/DOCX.

## Autora y contacto

Creado por Xinyi Chen, founder of HEDGE Global.

Presentación en chino: 陈歆怡，海聚海外 CEO。

Contacto: `chenxinyi_g`

## Sobre Xinyi

Xinyi Chen es originaria de Zhejiang y vive en Shanghái.

Es fundadora y CEO de HEDGE Global, con una trayectoria centrada en educación internacional de alto nivel, expansión global de empresas e inversión tecnológica. Antes disfrutaba viajar de forma espontánea; ahora es entusiasta del Vibe coding.

Si quieres conversar sobre Vibe coding, puedes agregarla en WeChat: `chenxinyi_g`.

![Código QR de WeChat](assets/wechat-qr.png)

## Qué incluye

- Posicionamiento del libro, promesa al lector, título, subtítulo y esquema completo.
- Creación de un proyecto tipo mdBook con `book.toml`, `src/SUMMARY.md`, archivos de mantenimiento y capítulos.
- Gestión de investigación, fuentes, citas, puntos pendientes y bibliografía.
- Escritura y revisión de capítulos con continuidad, guía de estilo y glosario.
- Estructura editorial: portada interior, página legal, prefacio, capítulos, referencias y anexos opcionales.
- Packaging comercial de portada: gancho de venta, símbolo visual, jerarquía tipográfica, prueba de estantería y prueba de miniatura.
- Exportación de `manuscript.md` a Word/DOCX mediante `scripts/md2word.py`.

## Instalación

```bash
mkdir -p ~/.codex/skills
git clone https://github.com/hedgeacademy/bookcraft ~/.codex/skills/bookcraft
```

Abre un nuevo hilo de Codex y usa:

```text
Use $bookcraft to plan a book about [topic].
```

## Prompts habituales

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

## Método de portada

El método de portada prioriza la claridad comercial. Una buena dirección debe responder en pocos segundos: de qué trata el libro, para quién es, qué promete y por qué elegirlo frente a libros similares.

Consulta:

- `references/09-cover-design.es.md`
- `references/09-cover-design.md`

## Requisitos de exportación

La exportación DOCX requiere Python, `python-docx`, `pandoc`, un `manuscript.md` completo, un `book.toml` válido y una sección real de referencias.

## Licencia

MIT
