# docx-service — SVG Images with PNG Fallback

**Date:** 2026-08-11
**Status:** Approved for planning
**Builds on:** [2026-08-10-docx-service-design.md](2026-08-10-docx-service-design.md)

## Goal

Insert an SVG into the generated document so that Word desktop renders the vector
original while Word Online and older Word render a raster fallback.

Delivered in two phases. Phase A closes the findings from phase one's whole-branch
review. Phase B adds the feature. Phase A goes first because two of its findings are
library-packaging defects that adding Batik would make materially worse.

## Phase A — close the phase-one review findings

| # | Finding | Fix |
| --- | --- | --- |
| A1 | `slf4j-simple` is `runtime` scope and `simplelogger.properties` sits in `src/main/resources`, so both leak onto every downstream consumer's classpath, where SLF4J may bind the wrong provider and this library's log level silently applies to their app. | Move `simplelogger.properties` to `src/test/resources/`; change `slf4j-simple` to `test` scope; add `<classpathScope>test</classpathScope>` to `exec-maven-plugin` so `mvn compile exec:java` still has a binding. |
| A2 | `writeTo(OutputStream)` closes the caller's stream — docx4j's `save()` wraps it in a `ZipOutputStream` and closes that. This contradicts the documented servlet use case: closing a `ServletOutputStream` commits the response. Undetected because `ByteArrayOutputStream.close()` is a no-op. | Wrap in a non-closing `FilterOutputStream` whose `close()` only flushes, so the caller retains ownership. Add a test asserting the caller's stream is still open afterwards. |
| A3 | `PageSetup.build()` margin-sum validation uses unguarded `int` addition, so extreme per-side values overflow negative and the `>=` check passes. Verified: `left(Integer.MAX_VALUE).right(Integer.MAX_VALUE)` builds successfully. A validation routine that fails open. | Widen both comparisons to `long`: `if ((long) leftTwips + rightTwips >= pageWidthTwips)`. Add an overflow test. |
| A4 | `writeToMatchesToByteArray` compares whole-ZIP bytes. ZIP entries carry MS-DOS timestamps at 2-second granularity, so two saves straddling a bucket boundary differ. Measured flake rate ~0.034% (6 failures in 17,719 runs). The test is also near-tautological, since `toByteArray()` delegates to `writeTo()`. | Assert both paths reload to a package with the same `sectPr` margins and heading text, rather than comparing raw bytes. |
| A5 | `README.md` claims "`page` and the facade own the package", but `PageSetup` never references `WordprocessingMLPackage` — only `WordDocument` does. It also claims `writeTo` means "the whole file never sits in heap", which overstates the benefit. | Reword to "the facade owns the package; `page` only produces a `w:sectPr` fragment", and to "avoids buffering a second full copy". |

Also fix while in the files: give `Units.check` an upper bound so numeric overflow throws
`DocumentGenerationException` rather than `ArithmeticException` (currently reachable via
`Units.cmToTwips(1e9)` and, worse, via `HeadingStyle.builder().sizePt(1e12).build()`,
which validates successfully and then throws from `toRPr()` at emit time); build two
`HpsMeasure` instances in `toRPr()` instead of aliasing one into both `sz` and `szCs`;
import `Locale.ROOT`; add a trailing newline to `simplelogger.properties`.

## Phase B — SVG with PNG fallback

### How Word actually does it

Word 2016+ does not store an SVG as the image. It stores **both**: the `a:blip`'s
`r:embed` points at a **PNG**, and the SVG is attached as an extension on that blip.

```xml
<a:blip r:embed="rId3">                             <!-- the PNG: universal fallback -->
  <a:extLst>
    <a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">
      <asvg:svgBlip
          xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main"
          r:embed="rId4"/>                          <!-- the SVG -->
    </a:ext>
  </a:extLst>
</a:blip>
```

Renderers that understand the extension draw the vector; everything else draws the PNG
it already had. The fallback is structural, not conditional — there is no branch to
write. This was verified end-to-end by prototype before this spec was written.

### Architecture

Three new files plus additions to `Units` and the facade.

```
com.example.docx
├── Units.java                     + cmToEmu, inchesToEmu (1 inch = 914400 EMU)
├── image/
│   └── SvgRasteriser.java         Batik: (svg bytes, target px width) -> PNG bytes.
│                                  No docx4j at all. Testable standalone.
├── part/
│   └── ImageParts.java            Needs the package. Creates the PNG image part,
│                                  builds the Inline, creates the SVG BinaryPart by
│                                  hand, attaches the extension. Returns a w:p.
└── sample/
    └── DemoChart.java             Generates a bar-chart SVG for the demo.
```

**Why `part/` and not `content/`.** The phase-one layering rule is that `content/` and
`style/` never touch `WordprocessingMLPackage`. Images cannot honour that: an image *is*
a package part plus a relationship. So images get a `part/` package, where needing the
package is the defining characteristic. This extends the rule rather than bending it.

`SvgRasteriser` stays outside both, because rasterising has nothing to do with OOXML.

### API

```java
byte[] docx = WordDocument.builder()
        .pageSetup(PageSetup.a4())
        .heading("Quarterly Report", HeadingStyle.defaults())
        .svgImage(DemoChart.randomBarChart(), 12.0)   // svg bytes, display width in cm
        .build()
        .toByteArray();
```

`svgImage(byte[] svg, double widthCm)` rasterises the SVG, embeds both parts, and
appends one paragraph containing the drawing.

**Sizing.** `createImageInline` sizes the image from the PNG's intrinsic pixel
dimensions and DPI, which is not what the caller asked for. The display size is
therefore set explicitly afterwards on the inline's extent, in EMU:

```java
long cx = Units.cmToEmu(widthCm);
long cy = Math.round(cx * (double) pngHeightPx / pngWidthPx);   // preserve aspect
inline.getExtent().setCx(cx);
inline.getExtent().setCy(cy);
```

Verified: 12 cm emits `<wp:extent cx="4320000" cy="2592000"/>`. Height comes from the
rasterised PNG's own pixel ratio, so the image is never distorted. 1 inch = 914,400 EMU.

The PNG is rasterised at **2× the display width** in pixels, so the fallback stays
sharp on high-DPI screens without inflating the file much. At 96 DPI a 12 cm image is
454 px, so the PNG is rendered 908 px wide.

### Three constraints that produce silently-wrong files

Each of these yields a document that opens without complaint and is wrong. All three
were hit during prototyping.

1. **`DocumentBuilderFactory.setNamespaceAware(true)` is mandatory.** It is `false` by
   default. Without it `asvg:svgBlip` parses as a literal element name in *no*
   namespace and `xmlns:asvg` as an ordinary attribute, emitting
   `<svgBlip xmlns="" … embed="rId4"/>` — no prefix, no namespace, `r:embed` reduced to
   `embed`. Word opens the file and silently ignores the SVG.
2. **docx4j 17.0.2 has no `ImageSvgPart`.** `BinaryPartAbstractImage.createImagePart(…,
   "image/svg+xml", …)` returns a `DefaultXmlPart` — SVG is XML, so docx4j's part
   factory builds an XML part — and the cast to `BinaryPartAbstractImage` throws
   `ClassCastException`. The SVG part must be constructed directly:
   `new BinaryPart(new PartName("/word/media/<name>.svg"))`, then `setBinaryData`,
   `setContentType(new ContentType("image/svg+xml"))`,
   `setRelationshipType(Namespaces.IMAGE)`, and
   `mdp.addTargetPart(part, AddPartBehaviour.RENAME_IF_NAME_EXISTS)`.
3. **The PNG is the primary `r:embed`; the SVG is only the extension.** Reversing them
   breaks every renderer, including Word desktop.

docx4j writes the SVG's content type as an OPC `Override` on the part name rather than
a `Default` on the `svg` extension. Both are valid OPC; no action needed, but a test
asserting a `Default Extension="svg"` entry would fail against a correct file.

### Dependency cost

`batik-transcoder` and `batik-codec` at 1.19 pull **18 additional jars** (batik-anim,
-awt-util, -bridge, -codec, -constants, -css, -dom, -ext, -gvt, -i18n, -parser,
-script, -shared-resources, -svg-dom, -svggen, -transcoder, -util, -xml) plus
`xmlgraphics-commons`. That is a real weight for a library, accepted here because
rasterising arbitrary SVG is the requirement. `batik-script` brings scripting support
that this use case does not need; excluding it is a possible later optimisation, not
part of this phase.

### Error handling

Unchanged contract: `DocumentGenerationException` is the only exception thrown
deliberately. Malformed SVG, a Batik `TranscoderException`, or a non-positive width all
surface as `DocumentGenerationException`, with a cause for the Batik failure and
without one for validation.

### Testing

**`SvgRasteriser`, no docx4j:** output starts with the PNG magic bytes
(`89 50 4E 47`); requested width is honoured; aspect ratio is preserved; malformed SVG
throws `DocumentGenerationException` with the `TranscoderException` as cause; a
non-positive width is rejected.

**Round trip, the one that matters:** build a document with an SVG, reload the bytes,
and assert on the reloaded package — both media parts exist, the SVG part's content
type is `image/svg+xml`, the blip's `r:embed` resolves to the **PNG** part, and the
extension carries uri `{96DAC541-7B7A-43D3-8B79-37D633B846F1}` with an
`asvg:svgBlip` element **in the correct namespace** whose `r:embed` resolves to the SVG
part.

That last assertion must check the namespace, not merely that the string `svgBlip`
appears. The broken first prototype contained the substring `svgBlip` and would have
passed a naive check while being unreadable by Word.

**Sample:** `SampleMain` writes a document containing the heading and a demo chart.

## Out of scope

Floating/anchored images and text wrapping; image captions; multiple images per
document beyond repeated `svgImage` calls; cropping; alt-text customisation beyond a
fixed default; raster input formats (JPEG/PNG passed directly by the caller); excluding
`batik-script`.
