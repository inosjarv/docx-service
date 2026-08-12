# docx-service — Hyperlinks

**Date:** 2026-08-12
**Status:** Approved for planning
**Builds on:** [2026-08-11-multi-section-design.md](2026-08-11-multi-section-design.md)

## Goal

Let a paragraph end with a clickable link, in the same line as the body text — "Revenue
grew 12%. **Learn more**" — where only the trailing phrase is a hyperlink. The link is
optional per paragraph.

## Why this needs the package, and what that forces

`content/Paragraphs.of` builds exactly one run and touches no package. A hyperlink is
different in kind: the URL is not inline text, it is a **relationship**
(`word/_rels/document.xml.rels`, `TargetMode="External"`), which only exists once a
`WordprocessingMLPackage` does. `content/` and `style/` are never allowed to reference
that type, so a paragraph carrying a link cannot be built by `Paragraphs`.

This is the same fork `ImageParts` already crossed for images, and it has the same
consequence: **a linked paragraph must defer to `build()`**, unlike a plain
`paragraph(...)` call, which validates and builds at the call site today. That is a real
asymmetry a caller can observe, and it is documented rather than hidden.

## Mechanism, verified against docx4j 17.0.2

Confirmed by building and reloading actual documents before writing this spec:

```java
RelationshipsPart rp = mainDocumentPart.getRelationshipsPart(true);
Relationship rel = new org.docx4j.relationships.ObjectFactory().createRelationship();
rel.setType(Namespaces.HYPERLINK);
rel.setTarget(url);
rel.setTargetMode("External");
rp.addRelationship(rel);              // assigns rel.getId(), e.g. "rId4"

P.Hyperlink link = new P.Hyperlink();
link.setId(rel.getId());
link.getContent().add(run);           // an org.docx4j.wml.R, built the normal way
```

Reload confirms:

```xml
<w:hyperlink w:history="true" r:id="rId3">
    <w:r><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr>
    <w:t xml:space="preserve">Learn more</w:t></w:r>
</w:hyperlink>
```

with the relationship correctly external:

```xml
<Relationship Target="https://example.com" TargetMode="External"
              Type=".../relationships/hyperlink" Id="rId3"/>
```

Two things confirmed by running them, not assumed:

- **A paragraph can mix a plain run, a `P.Hyperlink`, and another plain run**, and the
  order survives a reload (`R, Hyperlink, R`, in that sequence). This is what lets
  "Revenue grew 12%. Learn more" be one paragraph rather than two.
- **Image and hyperlink relationships coexist without collision** — `BinaryPartAbstractImage.createImagePart` and `getRelationshipsPart(true)` assign `rId3`/`rId4` independently and both resolve after reload.

## Architecture

```
style/TextStyle.java        + underline(boolean); + TextStyle.link() preset
content/Paragraphs.java     + public run(text, style) -> R, reused below
content/Hyperlink.java      NEW  immutable value: text, url, TextStyle. No package.
part/Hyperlinks.java        NEW  (pkg, text, TextStyle, ParagraphStyle, Hyperlink) -> P
```

`Hyperlink` is data with validation, the same shape as `TableStyle` — but unlike
`TableStyle`, nothing about it needs a package, so it lives beside `Paragraphs.of`'s
input rather than beside `ImageParts`. The class that actually builds the `w:hyperlink`
is `part/Hyperlinks`, for the same reason `ImageParts` exists: it is the one that needs
`WordprocessingMLPackage`.

`Paragraphs.run(String, TextStyle)` is extracted from the existing `of(...)` method as a
public method, so `Hyperlinks` can build the leading plain-text run without duplicating
`xml:space="preserve"` handling. `of(...)` becomes a thin wrapper: build the run, wrap it
in a paragraph, attach `pPr`.

## Why `TextStyle` gains `underline`, not a new style class

A hyperlink differs from body text only in what `TextStyle` already governs — colour and,
now, underline. A separate `HyperlinkStyle` would duplicate `TextStyle` for no reason.
`underline(boolean)` follows the emit-only-when-true pattern `bold`/`italic` already use,
confirmed against the current `toRPr()`:

```java
if (bold) { ... rPr.setB(on); }
if (italic) { ... rPr.setI(on); }
```

`underline` gets the identical shape, using `RPr.setU(U)` with
`UnderlineEnumeration.SINGLE` — verified in the round trip above.

`TextStyle.link()` is a new preset: `#0563C1`, underlined, otherwise
`TextStyle.body()`'s font and size. `#0563C1` is Word's own current hyperlink blue, not
an invented colour.

## API

```java
builder.paragraph("Revenue grew 12%. ",
        Hyperlink.of("Learn more", "https://example.com/report"));
```

Three `WordDocument.Builder` overloads, mirroring the existing text/paragraph-style
tiers exactly:

```java
paragraph(String text, Hyperlink link)                                       // TextStyle.body(), ParagraphStyle.body()
paragraph(String text, TextStyle textStyle, Hyperlink link)                  // ParagraphStyle.body()
paragraph(String text, TextStyle textStyle, ParagraphStyle paragraphStyle, Hyperlink link)
```

"Optional" is expressed the way it already is everywhere else in this API — by which
overload you call, not a nullable parameter. The existing no-link overloads are
untouched. If `text` is blank, no leading run is emitted; the paragraph is link-only.

```java
Hyperlink.of(text, url)                    // TextStyle.link() default
Hyperlink.of(text, url, TextStyle style)   // explicit style, per "configurable, just in case"
```

## Validation

`Hyperlink.of` needs no package, so its checks are eager, at the call site, the same
principle used everywhere in this codebase for package-free validation:

- link text: blank/null rejected
- url: blank/null rejected; parsed with `java.net.URI` and rejected if that throws
- style: null rejected

Because the *paragraph* containing a `Hyperlink` cannot be built until `build()`, the
**body text**'s blank-check for these three overloads is deferred to `build()` too — an
asymmetry from the plain `paragraph(...)` overloads, which validate immediately. This
mirrors the existing image asymmetry and is stated in the README/DOCUMENTATION.md rather
than left for a caller to discover.

Scope is external URLs only. Internal bookmarks/anchors (`w:anchor`, table of contents
territory) are out of scope, matching what was already excluded.

`DocumentGenerationException` remains the only exception type; validation failures carry
no cause, a `URISyntaxException` from a malformed URL is wrapped with one.

## Testing

The round trip is what matters, same as every prior feature:

- A paragraph built from `paragraph(leadingText, Hyperlink.of(linkText, url))` reloads
  with exactly **two** content items, in order: the leading `R`, then the `Hyperlink`.
  There is no trailing plain run in this API. Assert the reloaded order and that both
  runs' text matches.
- The relationship is present, `TargetMode="External"`, `Target` equals the URL given,
  and the hyperlink's `r:id` resolves to it.
- Blank leading text produces a paragraph containing only the `Hyperlink` — no empty
  leading run.
- `TextStyle.link()` emits `w:color val="0563C1"` and `w:u val="single"`; an explicit
  style on `Hyperlink.of` overrides both.
- Malformed URL and blank link text both throw `DocumentGenerationException` from
  `Hyperlink.of`, eagerly — a test asserts this without ever calling `build()`.
- A document mixing a linked paragraph with the existing plain paragraphs, images and
  tables reloads with everything in insertion order — the case where the deferred
  hyperlink and the eager plain content could diverge.
- Relationship IDs for an image and a hyperlink in the same document are distinct and
  both resolve after reload — the coexistence case verified above.

## Sample and documentation

`SampleMain` gains one linked paragraph, and `DOCUMENTATION.md` gains this feature to its
API reference and its "validation timing" section, naming the new asymmetry explicitly
rather than letting the existing two-case rule go stale.

## Out of scope

Internal bookmarks/anchors; a standalone whole-paragraph link (the entire paragraph is
just the link — not requested, and easy to add later as a fourth overload if it is);
multiple links in one paragraph; links inside table cells; `mailto:` or other scheme
validation beyond `java.net.URI` parseability; link target visited/unvisited colour
variants; removing the underline by default (it stays on, matching Word).
