# docx-service

Generates Word (`.docx`) documents from a Java backend with
[docx4j](https://www.docx4java.org/). Java 25, Maven, no web framework.

Headings, body text, tables and SVG images with PNG fallbacks — appended in
order, with control over page size, margins, fonts, colours, alignment and how
content behaves at a page boundary.

**📖 [DOCUMENTATION.md](DOCUMENTATION.md) is the full developer guide** — API
reference, recipes, the unit traps, and the things that produce a wrong document
rather than an obvious error. Start there if you are using this as a library.

## Quick start

```bash
JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn test
```

```bash
JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn test-compile exec:java
```

The second writes `target/sample-<timestamp>.docx`.

A second, bigger sample — an incident postmortem long enough to run several pages,
with a paragraph deliberately written to split across a page boundary — is
`PostmortemSampleMain`. Run it via its own named exec execution:

```bash
JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn test-compile exec:java@postmortem
```

This writes `target/postmortem-<timestamp>.docx`.

## Using it

```java
var builder = WordDocument.builder()
        .pageSetup(PageSetup.a4())
        .heading("Quarterly Report", TextStyle.defaults());

for (Section section : sections) {          // 5, 6, 10 — decided at runtime
    builder.heading(section.title(), TextStyle.builder().sizePt(11).bold(true).build());
    for (String paragraph : section.paragraphs()) {
        builder.paragraph(paragraph);
    }
}

builder.svgImage(svgBytes, pngFallbackBytes);

byte[] docx = builder.build().toByteArray();
```

For large documents prefer `writeTo(out)` over `toByteArray()`: it avoids
buffering a second full copy of the file. `writeTo` does not close the stream
you give it, so a controller keeps ownership of the response.

## Layout

| Package | Contents |
| --- | --- |
| `com.example.docx` | `WordDocument` (facade), `Units`, `DocumentGenerationException` |
| `…​.page` | `PageSetup` — page size and margins |
| `…​.style` | `TextStyle` — run formatting; `ParagraphStyle` — spacing, alignment and keep-with-next; `Alignment` |
| `…​.content` | `Paragraphs` — stateless paragraph factory, used for headings and body alike |
| `…​.style` | also `TableStyle`, `TableBorderStyle`, `BorderLine`, `Edge` |
| `…​.part` | `ImageParts` — image parts and the SVG blip extension |
| `…​.sample` | Runnable `main` (test sources, so it stays out of the jar) |

`content` and `style` never touch `WordprocessingMLPackage`, so they are pure
functions of their arguments and need no fixtures. The facade owns the
package; `page` only produces a `w:sectPr` fragment for it.

## Tables

```java
builder.table(
        List.of("Region", "Revenue"),
        List.of(List.of("EMEA", "1 240")),
        TableStyle.builder()
                .headerBorder(TableBorderStyle.bottomOnly("#1F4E79", 1.0))
                .bodyBorder(TableBorderStyle.bottomOnly("#BFBFBF", 0.5))
                .build());
```

Borders are chosen per cell edge, separately for the header and the body:

| Wanted | Set |
| --- | --- |
| Rule under the header | header → `BOTTOM` |
| Rules between body rows | body → `BOTTOM` |
| Full grid | body → all four |

`TableBorderStyle.none()`, or a style with no edges, draws nothing.

## Things that bite

- **Twips, not pixels.** All page geometry is in twips (1/1440 inch). The
  default margin of 851 twips is ~1.5 cm.
- **851 ≠ `marginsCm(1.5)`.** The cm path computes 850 (1.5 / 2.54 × 1440 =
  850.39). The 0.02 mm difference is invisible, but the numbers are not equal.
- **A4 is a literal `11906 × 16838`.** Deriving 297 mm gives 16839, which is
  not what Word writes.
- **Half-points.** `w:sz` is twice the point size: 20 pt emits `40`.
- **Bare hex.** OOXML rejects `#` in `w:color`; the builder strips it.
- **`xml:space="preserve"`** is set on every `w:t`, or leading and trailing
  spaces vanish silently.
- **Latest docx4j is 17.0.2, not 11.5.x.** `search.maven.org` reports 11.5.3
  and is stale; check `maven-metadata.xml` on repo1 instead. In 17.x the `w:*`
  classes come from `docx4j-generated-objects`, not `docx4j-openxml-objects`.
- **Word wants SVG 1.1, not SVG 2.** Word's renderer is strict where browsers
  are lenient, so an SVG that looks right in Chrome can render wrong or not at
  all. Two common offenders: `height="auto"` (SVG 2 sizing — rejected outright),
  and 8-digit hex colours like `#444cf71a` (CSS Color 4 — silently falls back to
  **black**). Use explicit `width`/`height`, and `fill="#444CF7"` with a separate
  `fill-opacity`.
- **SVGs carry a PNG twin.** `svgImage` embeds both: the PNG is what the blip
  points at, and the SVG rides along as an extension. Word desktop draws the
  vector, Word Online and older Word draw the PNG. You supply both — the library
  does not rasterise, and does not check that they match.
- **Content is appended, in order.** `heading`, `paragraph` and `svgImage` each add
  one item; none of them replaces a previous call. Build a document by looping.
- **Headings keep with the next paragraph.** `ParagraphStyle.heading()` sets
  `w:keepNext`, so a heading never strands at the foot of a page with its body
  overleaf. Body paragraphs deliberately do not set it.
- **Three different page-flow controls, three different jobs.** `keepWithNext`
  glues a paragraph to the next one (headings use it). `keepLines` keeps one
  paragraph's lines together, moving the whole paragraph rather than splitting
  it. `widowControl` allows the split but forbids a lone stranded line. For body
  text `widowControl` is right and `keepLines` is usually wrong — forcing whole
  paragraphs over leaves large gaps at page ends.
- **None of them is a guarantee.** A paragraph taller than the text area is split
  by Word regardless of `keepLines`; the alternative would be losing text.
- **Page breaks are a paragraph property, not content.**
  `ParagraphStyle.pageBreakBefore(true)` makes that paragraph start a new page.
  There is no separate break object to insert, so nothing drifts out of position
  when the content above it changes.
- **Justified is `both` in OOXML.** `Alignment.JUSTIFY` emits `w:jc w:val="both"`,
  not `"justify"`. `Alignment.LEFT` emits no `w:jc` at all, since left is Word's
  own default.
- **Points, not pixels.** An 11 pt heading is `sizePt(11)`, emitting `w:sz` 22.
  11 px would be 8.25 pt and noticeably smaller.
- **An empty `w:tc` makes the document unopenable.** Not misrendered — unopenable.
  Every cell gets a paragraph, including blank and null ones.
- **`w:tblGrid` is mandatory**, with a fixed layout. Without both, Word auto-fits
  to content and the same table renders differently in different clients.
- **Border width is in eighths of a point** — a third unit alongside twips and
  half-points. 1 pt is `w:sz="8"`; a 0.5 pt hairline is `4`.
- **A spacer paragraph follows every table.** Two adjacent tables merge into one
  in Word, and a body ending in a table is irregular.
- **Table borders are per cell edge, not `w:tblBorders`.** Table-level borders
  cannot give the header a border different from the body's. One consequence:
  adjacent cells each own their edges, so `LEFT` + `RIGHT` on the body puts two
  borders between columns and the thicker one wins.
- **A linked paragraph defers to `build()`.** `paragraph(text, Hyperlink)` and its two
  siblings need the package for the link's relationship, so — unlike every other
  `paragraph(...)` overload — they validate at `build()` time, not at the call site.

## Notes

- `WordDocument` is not thread safe. Build one per document.
- The first document in a JVM pays a one-off ~1s JAXB context initialisation.
  Warm it at startup with `Context.getWmlObjectFactory()` if latency matters.
- The heading uses direct formatting only and carries no `Heading1` style id,
  so it does not appear in Word's Navigation pane. That is a deliberate
  phase-one trade-off.
- Images are drawn at half the usable page width (page width less both margins),
  with height from the PNG's aspect ratio.
