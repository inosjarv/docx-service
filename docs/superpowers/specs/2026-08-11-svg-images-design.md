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

**No rasteriser, and no new dependencies.** The caller supplies both the SVG and its
PNG fallback. Batik was evaluated and rejected: `batik-transcoder` plus `batik-codec`
pull 18 additional jars (including Rhino via `batik-script`), which is disproportionate
weight for a library whose job is assembling OOXML. Rasterising is a build-time or
upstream concern, not a runtime one.

One new file plus additions to `Units` and the facade.

```
com.example.docx
├── Units.java                     + twipsToEmu (1 twip = exactly 635 EMU)
├── page/PageSetup.java            + usableWidthTwips()
└── part/
    └── ImageParts.java            Needs the package. Creates the PNG image part,
                                   builds the Inline, creates the SVG BinaryPart by
                                   hand, attaches the extension. Returns a w:p.
```

**Why `part/` and not `content/`.** The phase-one layering rule is that `content/` and
`style/` never touch `WordprocessingMLPackage`. Images cannot honour that: an image *is*
a package part plus a relationship. So images get a `part/` package, where needing the
package is the defining characteristic. This extends the rule rather than bending it.

PNG pixel dimensions are read with `javax.imageio.ImageIO`, which is in the JDK — no
dependency, and it doubles as validation that the supplied bytes really are an image.

### API

```java
byte[] docx = WordDocument.builder()
        .pageSetup(PageSetup.a4())
        .heading("Quarterly Report", HeadingStyle.defaults())
        .svgImage(svgBytes, pngFallbackBytes)
        .build()
        .toByteArray();
```

`svgImage(byte[] svg, byte[] pngFallback)` embeds both parts and appends one paragraph
containing the drawing. It takes no width: the image is drawn at **half the usable page
width**, with height following the PNG's aspect ratio.

**Sizing.** Usable width is the page width less both margins, so the image tracks
whatever `PageSetup` was configured with rather than any fixed measurement:

```java
int usableTwips = pageSetup.usableWidthTwips();          // pageWidth - left - right
long cx = Units.twipsToEmu(usableTwips / 2);
long cy = Math.round(cx * (double) pngHeightPx / pngWidthPx);   // preserve aspect
inline.getExtent().setCx(cx);
inline.getExtent().setCy(cy);
```

This must be set explicitly, because `createImageInline` otherwise sizes the image from
the PNG's intrinsic pixel dimensions and DPI — which is not what was asked for. Setting
the extent is verified to work: an earlier prototype emitted
`<wp:extent cx="4320000" cy="2592000"/>` for an explicitly requested size.

**1 twip = exactly 635 EMU** (914400 ÷ 1440 divides evenly), so twips→EMU is lossless
integer arithmetic with no rounding to reason about.

Worked example on A4 with the default 851-twip margins and the 1600×1120 demo PNG:

| Quantity | Value |
| --- | --- |
| Page width | 11906 twips |
| Usable width | 11906 − 851 − 851 = 10204 twips |
| Half of usable | 5102 twips ≈ 9.0 cm |
| `cx` | 5102 × 635 = 3,239,770 EMU |
| `cy` | 3,239,770 × 1120/1600 = 2,267,839 EMU ≈ 6.3 cm |

Odd usable widths truncate on the integer division; a half-twip is 1/2880 inch, far
below anything renderable.

The caller owns keeping the PNG visually faithful to the SVG; the library does not
verify that they match, only that both parse.

Guidance for callers, not enforced: render the PNG at roughly 2× its display width so
the fallback stays sharp on high-DPI screens. Half the usable width of A4 is ~340 px at
96 DPI, so the committed 1600 px demo PNG is comfortably oversampled.

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

### The SVG must be SVG 1.1, not SVG 2

Word's SVG renderer is strict. Browsers are not. An SVG that looks perfect in Chrome can
render wrong or not at all in Word, and this is the single likeliest way this feature
disappoints in practice.

The demo asset supplied for this work exhibited both common failures, each found by
rasterising it:

| Source | Problem | SVG 1.1-safe form |
| --- | --- | --- |
| `<svg height="auto" …>` with no `width` | `auto` is SVG2 sizing. Strict parsers reject it outright — Batik errors with *The attribute "height" of the element `<svg>` is invalid*. | `width="2000" height="1400"` matching the `viewBox` |
| `fill="#444cf71a"` | 8-digit hex (`#RRGGBBAA`) is CSS Color 4 / SVG2. Strict parsers cannot read it and fall back to **black**, silently filling the chart area solid. | `fill="#444CF7" fill-opacity="0.102"` |

The second is the more dangerous: it does not fail, it renders wrong. The corrected
asset is committed at `src/test/resources/demo/chart.svg`, with its 1600×1120 PNG
fallback beside it at `demo/chart.png`.

This is a caller responsibility, not something the library enforces — validating SVG
profiles is out of scope. The README documents it, because a caller who hits it will
otherwise blame the library.

### Error handling

Unchanged contract: `DocumentGenerationException` is the only exception thrown
deliberately. Null or empty SVG bytes, PNG bytes `ImageIO` cannot decode, and a
non-positive width all surface as `DocumentGenerationException` — with a cause where one
exists, without one for plain validation.

### Testing

**Validation, no docx4j:** null/empty SVG rejected; undecodable PNG rejected; each
throwing `DocumentGenerationException`.

**`PageSetup.usableWidthTwips()`:** A4 with default margins gives 10204; a setup with
asymmetric margins subtracts both correctly.

**`Units.twipsToEmu`:** 1 twip → 635; 1440 twips → 914400.

**Sizing:** on A4 with default margins and the 1600×1120 demo PNG, the emitted extent is
`cx = 3239770`, `cy = 2267839`. A narrower `PageSetup` produces a proportionally smaller
`cx`, which is the assertion that proves the width actually tracks the page rather than
a constant.

**Round trip, the one that matters:** build a document with an SVG, reload the bytes,
and assert on the reloaded package — both media parts exist, the SVG part's content
type is `image/svg+xml`, the blip's `r:embed` resolves to the **PNG** part, and the
extension carries uri `{96DAC541-7B7A-43D3-8B79-37D633B846F1}` with an
`asvg:svgBlip` element **in the correct namespace** whose `r:embed` resolves to the SVG
part.

That last assertion must check the namespace, not merely that the string `svgBlip`
appears. The broken first prototype contained the substring `svgBlip` and would have
passed a naive check while being unreadable by Word.

**Sample:** `SampleMain` writes a document containing the heading and the demo chart,
loading `demo/chart.svg` and `demo/chart.png` from the classpath.

### Where the demo assets live

`src/test/resources/demo/`, not `src/main/resources/`. Shipping demo art inside the
library JAR is the same defect as A1, one layer over. Fix A1 already gives
`exec-maven-plugin` a `test` classpath scope, so `mvn compile exec:java` reaches them.

`SampleMain` therefore moves from `src/main/java/…/sample/` to
`src/test/java/…/sample/`. It is a demo entry point, not API; leaving it in `main` while
its resources live in `test` would produce a class that throws for any consumer who
called it. After this move the published JAR contains library code only.

## Out of scope

Rasterising SVG at runtime (no Batik — callers supply the PNG); validating that a
supplied SVG is SVG 1.1-clean, or that the PNG matches the SVG; floating/anchored images
and text wrapping; image captions; cropping; alt-text customisation beyond a fixed
default; accepting raster-only input with no SVG.
