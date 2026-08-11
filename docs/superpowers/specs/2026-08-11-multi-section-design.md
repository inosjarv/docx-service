# docx-service — Multi-Section Documents

**Date:** 2026-08-11
**Status:** Approved for planning
**Builds on:** [2026-08-10-docx-service-design.md](2026-08-10-docx-service-design.md), [2026-08-11-svg-images-design.md](2026-08-11-svg-images-design.md)

## Goal

Let a document hold any number of headings, body paragraphs and images, in the order
they were added. The count is decided at runtime — five sections or ten, the calling
code is the same loop.

## The problem

`WordDocument.Builder` holds single-valued fields:

```java
private String headingText;
private HeadingStyle headingStyle = HeadingStyle.defaults();
private byte[] svgBytes;
private byte[] pngBytes;
```

So `.heading("A").heading("B")` does not add two headings — the second call overwrites
the first and the document silently contains only "B". There is no error. There is also
no body-text support at all: `content/` contains only `Headings`.

This is the seam the phase-one spec predicted: *"A second heading, a body paragraph, or
a table requires editing `WordDocument.Builder`."*

## Architecture

### Ordered, deferred content

The four fields are replaced by one ordered list:

```java
private final List<DocumentContent> content = new ArrayList<>();
```

`DocumentContent` is a single-method interface at the package root:

```java
@FunctionalInterface
public interface DocumentContent {
    P toParagraph(WordprocessingMLPackage pkg);
}
```

The list cannot hold finished `P` objects. Text paragraphs can be built eagerly, but an
image is a package part plus a relationship, and no package exists until `build()`. A
deferred entry is the only representation that lets images and text interleave in one
ordered list.

It lives at the root rather than in `content/`, because it names
`WordprocessingMLPackage` and the layering rule keeps `content/` and `style/` free of
that type. Implementations for text ignore the argument entirely.

`build()` iterates the list and appends each result. Adding a table later means adding a
factory, not editing `build()`.

### `Headings` merges into `Paragraphs`

A heading paragraph and a body paragraph have identical structure — one `w:p` holding
one `w:r`. They differ only in the styles applied. Two factories would be verbatim
duplication, so `content/Headings.java` is deleted and replaced by:

```java
public static P of(String text, TextStyle textStyle, ParagraphStyle paragraphStyle)
```

"Heading-ness" becomes the styles passed in, not a separate code path.

### `HeadingStyle` becomes `TextStyle`

The class only ever held run formatting — font family, size, bold, italic, colour —
which body text needs identically. The name was wrong from the start. Renaming is safe:
the artifact is `1.0.0-SNAPSHOT` and unmerged, so nothing external depends on it. The
rename touches the class, its test, `WordDocument`, and the README.

### File layout

```
com.example.docx
├── DocumentContent.java       NEW  P toParagraph(WordprocessingMLPackage)
├── WordDocument.java          MOD  List<DocumentContent> replaces four fields
├── Units.java                 —
├── page/PageSetup.java        —
├── part/ImageParts.java       —
├── style/TextStyle.java       REN  from HeadingStyle
├── style/ParagraphStyle.java  NEW  spacing + keepWithNext
├── content/Paragraphs.java    NEW  (text, TextStyle, ParagraphStyle) -> P
└── content/Headings.java      DEL  merged into Paragraphs
```

## API

```java
var builder = WordDocument.builder().pageSetup(PageSetup.a4());

for (Section section : sections) {
    builder.heading(section.title(), HEADING_11_BOLD);
    for (String paragraph : section.paragraphs()) {
        builder.paragraph(paragraph);
    }
}

byte[] docx = builder.build().toByteArray();
```

Every content call appends. Four overloads, each returning `Builder`:

| Method | Styles used |
| --- | --- |
| `heading(String)` | `TextStyle.defaults()`, `ParagraphStyle.heading()` |
| `heading(String, TextStyle)` | given, `ParagraphStyle.heading()` |
| `paragraph(String)` | `TextStyle.body()`, `ParagraphStyle.body()` |
| `paragraph(String, TextStyle)` | given, `ParagraphStyle.body()` |

`svgImage(byte[], byte[])` keeps its signature and also appends.

An 11 pt bold heading — the shape that prompted this work — is:

```java
TextStyle.builder().sizePt(11).bold(true).build()
```

Note **points, not pixels**. Word sizes text in points; 11 px at 96 DPI would be 8.25 pt
and noticeably smaller. `sizePt(11)` emits `w:sz` of 22 half-points.

## `keepWithNext` is the point of `ParagraphStyle`

`ParagraphStyle` carries three things: space before, space after (both twips), and
`keepWithNext`.

`keepWithNext` emits `w:keepNext`, which glues a paragraph to the one after it. Headings
set it by default. Without it a heading lands at the foot of a page with its body
overleaf — a visible defect, and one that becomes certain rather than possible once a
document has ten sections.

Verified against docx4j 17.0.2: `PPrBase.setKeepNext(BooleanDefaultTrue)`,
`PPrBase.setSpacing(PPrBase.Spacing)`, `Spacing.setBefore(BigInteger)` /
`setAfter(BigInteger)`, and `ObjectFactory.createPPr()` / `createPPrBaseSpacing()`.

`w:spacing`'s `w:before` and `w:after` are in twentieths of a point — the same twips the
rest of the project uses, so no new unit is introduced.

## Defaults

| Constant | Values |
| --- | --- |
| `TextStyle.defaults()` | Calibri Light, 20 pt, bold, `1F4E79` — unchanged, so existing behaviour and tests hold |
| `TextStyle.body()` | Calibri, 11 pt, regular, `000000` |
| `ParagraphStyle.heading()` | 240 twips before, 120 after, `keepWithNext = true` |
| `ParagraphStyle.body()` | 0 before, 120 after, `keepWithNext = false` |

240 twips is 12 pt; 120 twips is 6 pt.

## Validation

Text content is validated **eagerly**, at the `heading()` / `paragraph()` call, because
`Paragraphs.of` already runs there to build the paragraph. A blank string fails at the
offending line rather than at `build()`, which is worth having when a loop adds thirty
of them.

Images validate at `build()`, because `ImageParts.svgImage` performs its checks
alongside work that needs the package. The asymmetry is deliberate: making the two match
would mean copying the null/empty byte-array checks into the builder while `ImageParts`
keeps its own — it is public API and must validate regardless — and duplicated
validation drifts. One image per document is also the common case, so the diagnostic
value of failing early is much lower than for text.

`build()` requires at least one content item. This replaces the current "a heading is
required" rule, whose test becomes `requiresAtLeastOneContentItem`. An empty document is
valid OOXML but never what the caller meant.

`ParagraphStyle` rejects negative spacing. Unchanged elsewhere:
`DocumentGenerationException` is the only exception thrown deliberately; validation
failures carry a message naming the field and no cause.

The `pageSetup` a deferred image reads is the one present at `build()` time, not at the
`svgImage()` call, because the lambda reads the builder field when invoked. Setting the
page after adding an image therefore still sizes the image correctly.

## Testing

**Ordering is what this change can break**, so the test that earns its keep builds a
ten-section document — heading plus two paragraphs each — reloads it, and asserts the
body contains 30 paragraphs whose texts appear in exactly the expected sequence. A
change that dropped, duplicated or reordered entries passes every per-paragraph
assertion and fails this one.

Also:
- Two `heading()` calls produce two headings, not one. This is the regression that
  motivated the work; without it the old overwrite could return unnoticed.
- Heading paragraphs carry `w:keepNext` and body paragraphs do not.
- Spacing values reach `w:spacing/@w:before` and `@w:after`.
- A document mixing headings, paragraphs and an image keeps them in insertion order —
  the case eager and deferred entries could diverge on.
- `Paragraphs.of` rejects blank text and null styles; `ParagraphStyle` rejects negative
  spacing.
- The existing SVG, page-setup, units and stream tests continue to pass unchanged.

## Out of scope

Paragraph alignment; line spacing; indentation; lists and numbering; tables; a
`section(title, paragraphs)` convenience wrapper; mixed formatting within one paragraph
(multiple runs); page breaks; and headers or footers.
