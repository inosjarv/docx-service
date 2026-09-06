# docx-service — Developer Guide

Generate Word `.docx` files from a Java backend. Java 25, Maven, [docx4j](https://www.docx4java.org/) 17.0.2, no web framework.

This is the full reference. [README.md](README.md) is the short orientation.

---

## Contents

1. [Installing](#installing)
2. [Quick start](#quick-start)
3. [How the model works](#how-the-model-works)
4. [Units — read this first](#units--read-this-first)
5. [API reference](#api-reference)
6. [Recipes](#recipes)
7. [Things that will bite you](#things-that-will-bite-you)
8. [Error handling](#error-handling)
9. [Threading and performance](#threading-and-performance)
10. [Extending the library](#extending-the-library)
11. [What this does not do](#what-this-does-not-do)

---

## Installing

```xml
<dependency>
    <groupId>com.example</groupId>
    <artifactId>docx-service</artifactId>
    <version>1.0.0-SNAPSHOT</version>
</dependency>
```

**Requires Java 25.** The jar is compiled with `--release 25`.

`docx4j-JAXB-ReferenceImpl:17.0.2` arrives transitively, bringing the Glassfish JAXB
runtime with it. Inside an application server that already provides JAXB, exclude it and
substitute `docx4j-JAXB-MOXy`:

```xml
<dependency>
    <groupId>com.example</groupId>
    <artifactId>docx-service</artifactId>
    <version>1.0.0-SNAPSHOT</version>
    <exclusions>
        <exclusion>
            <groupId>org.docx4j</groupId>
            <artifactId>docx4j-JAXB-ReferenceImpl</artifactId>
        </exclusion>
    </exclusions>
</dependency>
```

The jar contains library code only — no logging binding, no demo assets, no sample class.
You supply your own SLF4J binding.

### Building from source

```bash
JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn clean test
```

```bash
JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn test-compile exec:java
```

The second writes a timestamped `target/sample-<millis>.docx` you can open. Note
`test-compile`, not `compile`: the sample lives in test sources so it stays out of the
published jar.

A second, bigger sample — `PostmortemSampleMain` — runs via its own named exec
execution instead: `mvn test-compile exec:java@postmortem`, writing a timestamped
`target/postmortem-<millis>.docx`. See the README for details.

---

## Quick start

```java
import com.example.docx.WordDocument;
import com.example.docx.page.PageSetup;
import com.example.docx.style.TextStyle;

byte[] docx = WordDocument.builder()
        .pageSetup(PageSetup.a4())
        .heading("Quarterly Report")
        .paragraph("Revenue grew twelve per cent quarter on quarter.")
        .build()
        .toByteArray();
```

From a Spring controller:

```java
@GetMapping(value = "/reports/{id}",
        produces = "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
ResponseEntity<StreamingResponseBody> download(@PathVariable String id) {
    WordDocument document = reports.render(id);
    return ResponseEntity.ok()
            .header(HttpHeaders.CONTENT_DISPOSITION,
                    ContentDisposition.attachment().filename("report.docx").build().toString())
            .body(document::writeTo);
}
```

`writeTo` does **not** close the stream you give it, so the container keeps ownership of
the response.

---

## How the model works

### Content is appended, in order

`heading`, `paragraph`, `svgImage` and `table` each **add** one item. None replaces a
previous call. A document is built by looping:

```java
var builder = WordDocument.builder().pageSetup(PageSetup.a4());

for (Section section : sections) {          // 5, 6, 10 — decided at runtime
    builder.heading(section.title());
    for (String text : section.paragraphs()) {
        builder.paragraph(text);
    }
}

byte[] docx = builder.build().toByteArray();
```

### Two style axes

Formatting splits along the same line OOXML does:

| | Controls | Applies to |
| --- | --- | --- |
| `TextStyle` | font, size, bold, italic, colour | the **run** — the characters |
| `ParagraphStyle` | spacing, alignment, page-flow keeps | the **paragraph** — the block |

Headings and body paragraphs share both types; they differ only in the values passed.
That is why there is no separate `HeadingStyle`.

### A builder is single-use

Text paragraphs and tables are built eagerly when you call `heading`/`paragraph`/`table`,
and the stored entry returns that same object every time. Calling `build()` twice on one
builder would hand both documents the same underlying nodes. **Build one document per
builder** — construction is cheap, and that is the natural lifetime in a request handler.

### Validation timing

Text and tables validate at the **call site**, because they are built there. A blank
heading fails on the line that added it, not at `build()` — which matters when a loop
adds thirty of them.

Images validate at `build()`, because an image needs the package, which does not exist
until then. A linked paragraph — `paragraph(text, Hyperlink)` and its two siblings —
follows the same rule as images and for the same reason: the link's URL is a
relationship, which needs the package too. Everything about that call, including the
leading text's blank-check, is deferred to `build()`, not validated at the call site
the way a plain `paragraph(...)` call is.

---

## Units — read this first

OOXML uses **three different units**, and mixing them up produces documents that look
subtly wrong rather than failing. This is the single most common source of error.

| Unit | Where | Conversion |
| --- | --- | --- |
| **Twip** (1/1440 inch) | page size, margins, table widths, paragraph spacing | `Units.cmToTwips`, `Units.inchesToTwips` |
| **Half-point** | font size | `Units.pointsToHalfPoints` — 11 pt → `22` |
| **Eighth of a point** | border width | `Units.pointsToEighths` — 1.0 pt → `8` |
| EMU (1/914400 inch) | image extents | `Units.twipsToEmu` — 1 twip → `635` exactly |

You rarely call these directly: `sizePt(11)` and `widthPt(1.0)` take points and convert
internally. They are public so you never have to invent a conversion.

**There are no pixels anywhere in OOXML.** If you have a pixel measurement from a web
design, convert it before use. A value that looks like a sensible pixel count is almost
always a wrong twip count — 851 twips is 1.5 cm, but 851 pixels would be 8.86 inches,
wider than an A4 page.

### Handy values

| | Twips |
| --- | --- |
| A4 portrait | `11906 × 16838` |
| Default margin | `851` (≈1.5 cm) |
| A4 usable width at default margins | `10204` |
| 1 cm | 567 |
| 1 inch | 1440 |

`marginsCm(1.5)` computes **850**, not 851. Both are ~1.5 cm and the 0.02 mm difference
is invisible, but they are different numbers — do not treat them as interchangeable.

---

## API reference

### `WordDocument`

| Method | Returns | Notes |
| --- | --- | --- |
| `WordDocument.builder()` | `Builder` | |
| `toByteArray()` | `byte[]` | Buffers the whole file |
| `writeTo(OutputStream)` | `void` | Does **not** close the stream |
| `toHtml()` | `String` | Self-contained XHTML; buffers the whole thing |
| `writeHtmlTo(OutputStream)` | `void` | Does **not** close the stream |

Prefer `writeTo`/`writeHtmlTo` for large documents: they avoid buffering a second full
copy.

`toHtml()`/`writeHtmlTo` render via docx4j's own XHTML converter — the `<style>` block
comes entirely from docx4j's conversion of the document's paragraph and run formatting,
not from anything authored in this library. Images embed as base64 `data:` URIs, so the
output is one self-contained document with nothing alongside it to ship separately.

### `WordDocument.Builder`

| Method | Default styles used |
| --- | --- |
| `pageSetup(PageSetup)` | — |
| `heading(String)` | `TextStyle.defaults()`, `ParagraphStyle.heading()` |
| `heading(String, TextStyle)` | given, `ParagraphStyle.heading()` |
| `heading(String, TextStyle, ParagraphStyle)` | both given |
| `paragraph(String)` | `TextStyle.body()`, `ParagraphStyle.body()` |
| `paragraph(String, TextStyle)` | given, `ParagraphStyle.body()` |
| `paragraph(String, TextStyle, ParagraphStyle)` | both given |
| `paragraph(String, Hyperlink)` | `TextStyle.body()`, `ParagraphStyle.body()` — defers to `build()`, see below |
| `paragraph(RichText)` | `ParagraphStyle.body()` |
| `paragraph(RichText, ParagraphStyle)` | given |
| `svgImage(byte[] svg, byte[] pngFallback)` | drawn at **half** the usable page width |
| `table(headers, rows, TableStyle)` | spans the **full** usable page width |
| `table(headers, rows, TableStyle, int widthTwips)` | explicit width |
| `build()` | Requires at least one content item |

**`table(headers, rows, style)` defers the page setup to `build()`**, the same as
`svgImage`. It may be called before or after `pageSetup(...)` with the same result. Pass
an explicit width with the four-argument overload to pin the width regardless of page
setup.

### `PageSetup`

```java
PageSetup.a4()                                          // A4, 851-twip margins
PageSetup.builder().a4().marginsTwips(851).build()
PageSetup.builder().a4().marginsCm(2.0).build()
PageSetup.builder().a4().left(1440).right(1440).build()
PageSetup.builder().pageSizeTwips(12240, 15840).build() // US Letter
```

Builder: `a4()`, `pageSizeTwips(int, int)`, `marginsTwips(int)`, `marginsCm(double)`,
`marginsInches(double)`, `top/right/bottom/left(int)`, `build()`.

Accessors: `pageWidthTwips()`, `pageHeightTwips()`, `topTwips()`, `rightTwips()`,
`bottomTwips()`, `leftTwips()`, `usableWidthTwips()`, `toSectPr()`.

Constants: `A4_WIDTH_TWIPS` = 11906, `A4_HEIGHT_TWIPS` = 16838,
`DEFAULT_MARGIN_TWIPS` = 851.

`usableWidthTwips()` is page width less both side margins — the width text, images and
tables may occupy.

### `TextStyle`

Run formatting. Immutable, built through a nested builder.

```java
TextStyle.defaults()   // Calibri Light, 20 pt, bold, 1F4E79
TextStyle.body()       // Calibri, 11 pt, regular, 000000

TextStyle.builder()
        .font("Calibri")
        .sizePt(11)
        .bold(true)
        .italic(false)
        .color("#1F4E79")   // '#' optional, any case
        .build();
```

Accessors: `fontFamily()`, `sizePt()`, `bold()`, `italic()`, `colorHex()`, `toRPr()`.

`colorHex()` returns bare uppercase `RRGGBB` — OOXML rejects a leading `#`, so the
builder strips it. Sizes above `MAX_SIZE_PT` (1638) are rejected; that is Word's own
ceiling.

### `RichText`

Plain text carrying a limited set of inline markup on top of one base `TextStyle`.

```java
RichText.of(markup)             // styled against TextStyle.body()
RichText.of(markup, baseStyle)  // styled against an explicit base style
```

Recognised tags: `<b>`/`<strong>` (bold), `<i>`/`<em>` (italic), `<u>` (underline).
Tags nest freely and are matched case-insensitively. Any unrecognised tag, or any
markup that is not well-formed, throws `DocumentGenerationException`.

### `ParagraphStyle`

Paragraph formatting: spacing, alignment, and behaviour at a page boundary.

```java
ParagraphStyle.heading()   // 240 before, 120 after, keepWithNext, keepLines, widowControl
ParagraphStyle.body()      // 0 before, 120 after, widowControl only

ParagraphStyle.builder()
        .spaceBeforeTwips(240)
        .spaceAfterTwips(120)
        .alignment(Alignment.JUSTIFY)
        .keepWithNext(true)
        .keepLines(false)
        .widowControl(true)
        .pageBreakBefore(false)
        .build();
```

Constants: `HEADING_SPACE_BEFORE_TWIPS` = 240, `HEADING_SPACE_AFTER_TWIPS` = 120,
`BODY_SPACE_AFTER_TWIPS` = 120. (240 twips is 12 pt; 120 is 6 pt.)

#### `Alignment`

`LEFT`, `CENTER`, `RIGHT`, `JUSTIFY`. `LEFT` emits no `w:jc` at all, since left is Word's
default.

#### The three page-flow controls do different jobs

| Control | Effect |
| --- | --- |
| `keepWithNext` | Glues this paragraph to the next one. Headings use it. |
| `keepLines` | Keeps all lines of this paragraph together, moving the whole paragraph rather than splitting it. |
| `widowControl` | Allows a split but forbids a lone stranded line at a page top or bottom. |

For body text `widowControl` is right and `keepLines` is usually **wrong** — forcing
whole paragraphs over leaves large gaps at page ends. `ParagraphStyle.body()` reflects
that.

**None of these is a guarantee.** A paragraph taller than the text area is split by Word
regardless of `keepLines`, because the alternative is losing text.

### `TableStyle` and `TableBorderStyle`

```java
TableStyle.defaults()   // header: 1F4E79 1.0pt bottom; body: BFBFBF 0.5pt bottom

TableStyle.builder()
        .headerBorder(TableBorderStyle.bottomOnly("#1F4E79", 1.0))
        .bodyBorder(TableBorderStyle.bottomOnly("#BFBFBF", 0.5))
        .headerText(TextStyle.builder().sizePt(11).bold(true).build())
        .bodyText(TextStyle.body())
        .build();
```

`TableBorderStyle` presets: `none()`, `bottomOnly(colour, widthPt)`, `grid(colour,
widthPt)`, `box(colour, widthPt)`.

Full control:

```java
TableBorderStyle.builder()
        .color("#1F4E79")
        .widthPt(1.0)
        .line(BorderLine.SINGLE)
        .edges(Edge.BOTTOM, Edge.TOP)   // varargs; none means no border
        .build();
```

`BorderLine`: `SINGLE`, `THICK`, `DOUBLE`, `DOTTED`, `DASHED`, `DOT_DASH`. There is no
`NONE` — a border that draws nothing is one with no edges.

`Edge`: `TOP`, `BOTTOM`, `LEFT`, `RIGHT`. Borders are applied to **each cell** in the
group, so:

| Wanted | Set |
| --- | --- |
| Rule under the header | header → `BOTTOM` |
| Rules between body rows | body → `BOTTOM` |
| Full grid | header → all four, **and** body → all four |
| Outer box only | header → `TOP, LEFT, RIGHT`; body → `LEFT, RIGHT, BOTTOM` |

Cell padding is fixed at `Tables.CELL_PADDING_TWIPS` (80 twips, ≈0.14 cm) left and right.

### `Units`

`cmToTwips(double)`, `inchesToTwips(double)`, `pointsToHalfPoints(double)`,
`pointsToEighths(double)`, `twipsToEmu(int)`.

### Lower-level factories

`content.Paragraphs.of(String, TextStyle, ParagraphStyle)` → `org.docx4j.wml.P`
`content.Tables.of(headers, rows, TableStyle, int widthTwips)` → `org.docx4j.wml.Tbl`
`part.ImageParts.svgImage(pkg, svg, png, int widthTwips)` → `org.docx4j.wml.P`

These return raw docx4j objects. Use them if you need to assemble a document by hand;
otherwise prefer the builder.

---

## Recipes

### A section per page

```java
ParagraphStyle newPage = ParagraphStyle.builder()
        .spaceBeforeTwips(ParagraphStyle.HEADING_SPACE_BEFORE_TWIPS)
        .spaceAfterTwips(ParagraphStyle.HEADING_SPACE_AFTER_TWIPS)
        .keepWithNext(true)
        .keepLines(true)
        .pageBreakBefore(true)
        .build();

builder.heading("Appendix", TextStyle.defaults(), newPage);
```

Page breaks are a paragraph property, not inserted content — nothing drifts out of
position when the text above changes.

### Justified body text

```java
ParagraphStyle justified = ParagraphStyle.builder()
        .spaceAfterTwips(ParagraphStyle.BODY_SPACE_AFTER_TWIPS)
        .alignment(Alignment.JUSTIFY)
        .widowControl(true)
        .build();

builder.paragraph(text, TextStyle.body(), justified);
```

### An SVG chart with a PNG fallback

```java
builder.svgImage(svgBytes, pngBytes);
```

Word desktop draws the vector; Word Online and older Word draw the PNG. You supply both —
the library does not rasterise, and does not check that they match. Render the PNG at
roughly twice its display width so it stays sharp.

### A table

```java
builder.table(
        List.of("Region", "Revenue", "Change"),
        List.of(List.of("EMEA", "1 240", "+8%"),
                List.of("APAC", "980", "+21%")),
        TableStyle.defaults());
```

Rows must all have the same length as the headers. Null and blank cells are fine — they
become empty cells, not errors.

### A paragraph ending in a link

```java
builder.paragraph("Revenue grew 12%. ",
        Hyperlink.of("Learn more", "https://example.com/report"));
```

`Hyperlink.of(text, url)` uses Word's own hyperlink look — `#0563C1`, underlined.
`Hyperlink.of(text, url, TextStyle)` overrides it. External URLs only; there is no
support for internal bookmarks or anchors.

Unlike every other `paragraph(...)` overload, these three defer validation and
construction to `build()` — the link's URL is a relationship, which needs the package
that does not exist until then.

### Mixed formatting within a paragraph

```java
builder.paragraph(RichText.of("Some text which needs to be <b>bold</b>."));
```

`paragraph(RichText, ParagraphStyle)` is also available when the paragraph itself needs
explicit spacing, alignment, or page-flow behaviour.

---

## Things that will bite you

These are all things that produce a *wrong* document rather than an obvious failure.

**Word wants SVG 1.1, not SVG 2.** Word's renderer is strict where browsers are lenient,
so an SVG that looks perfect in Chrome can render wrong or not at all. Two common
offenders:

| Offender | Result | SVG 1.1-safe form |
| --- | --- | --- |
| `height="auto"`, or no `width`/`height` | rejected outright | explicit `width` and `height` matching the `viewBox` |
| 8-digit hex like `#444cf71a` | silently falls back to **black** | `fill="#444CF7"` plus `fill-opacity="0.1"` |

**Points, not pixels.** An 11 pt heading is `sizePt(11)`. 11 px would be 8.25 pt and
noticeably smaller.

**Border width is in eighths of a point.** 1 pt is `sz="8"`. This is not the half-point
unit fonts use.

**Adjacent table cells each own their edges.** Setting `LEFT` and `RIGHT` on the body puts
two borders between neighbouring columns. Word collapses them visually and the thicker
wins, so mixing widths across edges can look uneven.

**A "full grid" needs both borders set.** Setting only `bodyBorder` to `grid(...)` leaves
the header row without vertical dividers. Set `headerBorder` too.

**Headings keep with the next paragraph** by default, so they never strand at a page foot.
If you build a custom `ParagraphStyle` for a heading, set `keepWithNext(true)` yourself —
the builder default is `false`.

**A `\n` inside paragraph text does nothing.** It lands raw in `w:t` and Word collapses it
to a single line. There is no line-break support yet; use separate `paragraph` calls.

**`RichText` markup is parsed as XML, not HTML.** Plain-text portions must escape `&`,
`<`, `>` as `&amp;`, `&lt;`, `&gt;`. Only the five XML built-in entities (`&amp;`,
`&lt;`, `&gt;`, `&quot;`, `&apos;`) and numeric character references (e.g. `&#160;`) are
available — HTML named entities like `&nbsp;` are **not** supported and will be
rejected.

**Two adjacent tables would merge into one** in Word. The library emits a spacer paragraph
after every table to prevent it, which is why a document of paragraph-table-paragraph has
four body elements, not three.

---

## Error handling

`DocumentGenerationException extends RuntimeException` is the **only** exception this
library throws deliberately. Catch that one type and you have covered everything.

```java
try {
    byte[] docx = document.toByteArray();
} catch (DocumentGenerationException e) {
    // e.getCause() == null  -> your input was invalid
    // e.getCause() != null  -> docx4j failed underneath
}
```

The presence of a cause is the signal:

| Cause | Meaning | Examples |
| --- | --- | --- |
| absent | invalid input, message names the field | blank heading text, malformed hex colour, a table row of the wrong length (message names the row index), negative margins, font size above 1638 pt |
| present | docx4j or I/O failure, wrapped | serialisation failure, an unreadable PNG |

It is unchecked because neither case is recoverable at the call site: a malformed colour
is a programming error, and a serialisation failure is an environment problem.

Validation deliberately fails closed. Margins wider than the page, and conversions that
would overflow, are rejected rather than silently producing a broken document.

---

## Threading and performance

**`WordDocument` and its builder are not thread safe.** Build one per document. That is
the natural lifetime in a request handler, and construction is cheap.

**The first document in a JVM pays a one-off JAXB context initialisation of roughly one
second.** If latency matters, warm it at startup:

```java
org.docx4j.jaxb.Context.getWmlObjectFactory();
```

For large documents prefer `writeTo(OutputStream)` over `toByteArray()` — it avoids
buffering a second full copy.

---

## Extending the library

The package layout encodes a rule worth keeping:

| Package | Contents | May touch the package? |
| --- | --- | --- |
| `com.example.docx` | `WordDocument`, `DocumentContent`, `Units`, `DocumentGenerationException` | yes |
| `…​.page` | `PageSetup` | produces a `w:sectPr` fragment only |
| `…​.style` | `TextStyle`, `ParagraphStyle`, `TableStyle`, `TableBorderStyle`, `Alignment`, `BorderLine`, `Edge` | **no** |
| `…​.content` | `Paragraphs`, `Tables` | **no** |
| `…​.part` | `ImageParts` | yes — an image is a part plus a relationship |

`content` and `style` are pure functions of their arguments returning JAXB objects, so
they compose freely and need no fixtures. Anything needing `WordprocessingMLPackage`
belongs in `part`.

To add a content type — a list, a page header — add a factory in `content` (or `part` if
it needs the package) and a builder method that appends a `DocumentContent`:

```java
@FunctionalInterface
public interface DocumentContent {
    Object toBodyElement(WordprocessingMLPackage pkg);   // a P or a Tbl
}
```

It returns `Object` because an OOXML body holds `w:p` and `w:tbl` alike, which docx4j
models as `List<Object>`. Deferring to `build()` is what lets images — which need the
package — sit in the same ordered list as text.

---

## What this does not do

Not built, and not planned without a reason:

- Line breaks within a paragraph
- Lists and numbering
- Headers, footers, page numbers, and tables of contents
- Merged table cells, cell shading, per-cell style overrides, explicit per-column widths,
  nested tables, table captions
- Floating or anchored images, image captions, cropping
- Rasterising SVG (you supply the PNG), or validating that an SVG is SVG 1.1-clean
- Template filling
- Any HTTP or Spring layer — this is a library

---

## Verifying output yourself

A `.docx` is a zip. When something looks wrong, look at the XML:

```bash
unzip -l report.docx                              # parts
unzip -p report.docx word/document.xml            # the body
unzip -p report.docx word/_rels/document.xml.rels # image relationships
```

Useful checks:

```bash
unzip -p report.docx word/document.xml | grep -o '<wp:extent[^/]*/>'   # image size in EMU
unzip -p report.docx word/document.xml | grep -c 'w:keepNext'          # headings kept with next
unzip -p report.docx word/document.xml | grep -o '<w:gridCol[^/]*/>'   # table columns
```

---

## Document gallery

A dev tool for seeing every sample document's real Word-rendered look and feel in one
browser tab, without hand-running each `*SampleMain` and opening its output one file at a
time:

```bash
JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn test-compile exec:java@gallery
```

Requires a local LibreOffice install — real rendering fidelity needs a real OOXML layout
engine, which is exactly what this shells out to via a headless LibreOffice process.
`brew install --cask libreoffice` on macOS installs it somewhere the tool auto-detects; if
it isn't found, the command prints an actionable message and exits rather than a stack
trace.

Writes `target/gallery/index.html`, one section per fixture (currently `sample`,
`postmortem`, `rich-text`) with every page rendered as a PNG. Add a fixture by adding one
line to `GalleryFixtures.all()` — no other file needs to change.

This is a review-convenience tool only: it runs on demand, never as part of `mvn test`,
and has no automated pass/fail. Comparing renders against a stored baseline to catch
unintended visual regressions is a deliberately separate, not-yet-built follow-up.
