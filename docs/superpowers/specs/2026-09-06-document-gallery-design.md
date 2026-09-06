# docx-service — Document Gallery

**Date:** 2026-09-06
**Status:** Approved for planning

## Goal

Let a developer see how every sample/fixture document actually looks — real Word-fidelity
rendering, not docx4j's own approximation — in one browser tab, without hand-running each
`*SampleMain` and opening its output one file at a time. This is phase 1 of a two-phase
plan: a review-convenience tool now, with automated visual-regression detection (comparing
against a stored baseline, failing a test on unintended change) as a deliberately separate
follow-up once the review workflow exists and it's clear what's worth pinning.

Out of scope for this spec: any automated pass/fail. Nothing here runs as part of `mvn
test`; it's an on-demand dev tool, the same way `SampleMain` is today.

## Why LibreOffice, not `toHtml()`

`WordDocument.toHtml()` already produces a self-contained, viewable HTML file, and was
considered and rejected for this purpose: it's docx4j's own XSLT-based approximation, and
this codebase already documents concrete ways it diverges from real Word rendering (the
whole reason `WordDocument`'s `FONT_FACE_CSS`/`SCREEN_LAYOUT_CSS` exist is to paper over
gaps docx4j's conversion leaves). The explicit requirement here is Word-accurate look and
feel, which means rendering through an engine that actually implements OOXML layout —
LibreOffice, via a real `.docx → PDF → PNG` pipeline, is the standard way to get that from
a JVM without installing Word itself.

Two mature libraries do the actual work, chosen instead of hand-rolling `ProcessBuilder`
calls to `soffice`:

- **JODConverter** (`org.jodconverter:jodconverter-local`) drives a headless LibreOffice
  instance — starts it, manages its lifecycle, converts a file, and hands back control —
  without the caller managing the office process by hand.
- **Apache PDFBox** (`org.apache.pdfbox:pdfbox`) rasterizes each page of the resulting PDF
  to a `BufferedImage`/PNG via `PDFRenderer`, at a chosen DPI.

Both are test-scope-only dependencies; the shipped library depends on nothing beyond
docx4j and slf4j, exactly as today.

## Architecture

```
sample/gallery/GalleryFixture.java    NEW  record: name + Supplier<WordDocument>
sample/gallery/GalleryFixtures.java   NEW  the registered list of fixtures
sample/gallery/DocxToImages.java      NEW  (OfficeManager, WordDocument, Path) -> List<Path> page PNGs
sample/gallery/GalleryIndexWriter.java NEW  (fixture name -> page PNG paths) -> index.html content
sample/gallery/GalleryMain.java       NEW  orchestrates all of the above; the runnable entry point
sample/RichTextSampleMain.java        NEW  a small standalone bold/italic/underline/nesting sample
```

`HtmlSampleMain` and `PostmortemHtmlSampleMain` are kept as-is, unchanged — useful as a
quick single-document demo of the HTML export path on its own, separate from the gallery.
`GalleryMain` is purely additive alongside them, not a replacement.

All five new classes live in a new `sample.gallery` subpackage — kept apart from
`sample` because they're infrastructure *for* looking at the samples, not a sample
themselves, and because five small single-purpose files are easier to hold in context
than one that does registration, rendering, and HTML-writing at once.

### `GalleryFixture` and `GalleryFixtures`

```java
public record GalleryFixture(String name, Supplier<WordDocument> builder) {
}
```

```java
public final class GalleryFixtures {
    private GalleryFixtures() { }

    public static List<GalleryFixture> all() {
        return List.of(
                new GalleryFixture("sample", SampleMain::build),
                new GalleryFixture("postmortem", PostmortemSampleMain::build),
                new GalleryFixture("rich-text", RichTextSampleMain::build));
    }
}
```

Adding a fixture later — a table-focused sample, a hyperlink-focused one — is exactly one
line here plus, if it doesn't already exist, a `build()` on some sample main. Nothing else
in the pipeline changes.

### `RichTextSampleMain`

The one new sample document this spec adds: a small, standalone document exercising plain
text, `<b>`, `<i>`, `<u>`, and nesting (`<b>bold <i>and italic</i></b>`) — the rich-text
feature currently has thorough unit/round-trip test coverage but no visual sample of its
own. Follows the exact shape of `SampleMain`/`PostmortemSampleMain`: a package-private
static `build()` returning a `WordDocument`, and a `main()` that writes a `.docx` to
`target/`. No dedicated JUnit test, matching the established convention for sample mains
— rich text's own behavior is already covered by `RichTextTest`/`RichTextDocumentTest`.

### `DocxToImages`

The rendering core, given an already-started `OfficeManager` (started once by
`GalleryMain`, not per fixture — office-process startup is the slow part, at a few seconds,
and paying it once for the whole run rather than once per fixture is the difference
between a usable tool and a slow one):

```java
public static List<Path> renderPages(OfficeManager officeManager, WordDocument document, Path outputDir)
```

1. Writes `document.toByteArray()` to `outputDir/document.docx`.
2. Converts it to `outputDir/document.pdf` via JODConverter's `LocalConverter`, bound to
   the given `OfficeManager`.
3. Loads the PDF with PDFBox, renders each page via `PDFRenderer.renderImageWithDPI(page,
   150, ImageType.RGB)`, and writes each as `outputDir/page-<n>.png`.
4. Returns the ordered list of PNG paths.

150 DPI is the chosen default — legible full-page detail in a browser (roughly
1275×1650px for a US Letter page) without producing multi-megabyte PNGs per page. Both
the intermediate `.docx` and `.pdf` are left in `outputDir` rather than deleted — they're
free byproducts and useful if you want to open the exact file that was rendered.

*(Exact JODConverter/PDFBox method names above reflect the current published API of the
pinned versions below; the implementation plan confirms them against the actual library
once it's added as a dependency, the same way any new external API gets confirmed by
using it rather than assumed from memory.)*

### `GalleryIndexWriter`

A pure function: given the ordered list of `(fixture name, page PNG paths)`, returns the
`index.html` string. One `<section>` per fixture — a heading with its name, then each
page's `<img>` stacked vertically with a "Page N of M" caption — using paths relative to
`target/gallery/`, so the whole directory stays self-contained if copied elsewhere. No
JavaScript, no build step: this is a static file a browser opens directly. Being a pure
string-in-string-out function (given fake paths, it needs no LibreOffice) is what makes it
unit-testable without the rest of the pipeline.

### `GalleryMain`

```java
public static void main(String[] args) {
    // 1. Build and start one shared OfficeManager. On failure (LibreOffice not
    //    found/discoverable), print one clear message and exit(1) — no raw stack trace.
    // 2. For each GalleryFixtures.all() entry:
    //      - build the WordDocument
    //      - DocxToImages.renderPages(...) into target/gallery/<name>/
    //      - on any failure for this one fixture, log it and record a placeholder
    //        instead of aborting the whole run
    // 3. GalleryIndexWriter -> write target/gallery/index.html
    // 4. Stop the OfficeManager (finally-block, always runs)
    // 5. System.out.println("Wrote " + indexPath.toAbsolutePath())
}
```

On a standard Homebrew-cask install, LibreOffice lands at
`/Applications/LibreOffice.app`, one of the paths JODConverter's default office-home
detection already checks — no explicit configuration needed for the common case.

## `HtmlSampleMain` / `PostmortemHtmlSampleMain` stay

Kept deliberately: they're a fast, LibreOffice-free way to demo the HTML export path on
one document, which is a different purpose from the gallery's "compare every fixture's
real rendering at once." Their `html-sample`/`postmortem-html` exec-plugin executions are
untouched; `gallery` is added as a new execution alongside them.

## Maven wiring

```xml
<dependency>
    <groupId>org.jodconverter</groupId>
    <artifactId>jodconverter-local</artifactId>
    <version>4.4.11</version>
    <scope>test</scope>
</dependency>
<dependency>
    <groupId>org.apache.pdfbox</groupId>
    <artifactId>pdfbox</artifactId>
    <version>3.0.8</version>
    <scope>test</scope>
</dependency>
```

A new `gallery` execution is added alongside the existing ones, the same way, running
`com.example.docx.sample.gallery.GalleryMain`:

```
mvn exec:java@gallery
```

## Error handling

- **LibreOffice not found:** `GalleryMain` catches the `OfficeManager` startup failure
  specifically and prints one actionable message (install LibreOffice, or see the new
  "Gallery" section in `DOCUMENTATION.md`) instead of letting JODConverter's own exception
  surface raw. Exits non-zero.
- **One fixture fails to build or render:** caught per-fixture. The run continues; that
  fixture's section in `index.html` shows a visible "failed to render — see console" note
  instead of silently vanishing or aborting every other fixture's output.
- `GalleryIndexWriter` itself does no I/O and has nothing to fail on beyond bad input,
  which the caller (`GalleryMain`) is responsible for not producing.

## Testing

Consistent with `SampleMain`/`PostmortemSampleMain` today: sample mains and the rendering
pipeline that depends on an external program (LibreOffice) get no JUnit tests — the whole
point of this tool is a human looking at its output, and phase 2 is where that becomes an
automated, pinned comparison.

What *is* pure and does get ordinary unit tests: `GalleryFixtures.all()` returns the
expected names in order; `GalleryIndexWriter`, given a small fake list of
`(name, List<Path>)`, produces HTML containing every fixture's heading and `<img
src="...">` for every given path, in order — no LibreOffice involved.

## Documentation

`DOCUMENTATION.md` gains a short "Document gallery" section: what it's for, that it
requires a local LibreOffice install (with the Homebrew-cask path noted as the common
case), how to run it (`mvn exec:java@gallery`), and where the output lands
(`target/gallery/index.html`) — alongside, not replacing, the existing coverage of
`html-sample`/`postmortem-html`.

## Out of scope

Automated visual-regression/snapshot testing (comparing rendered output against a stored
baseline and failing a test on drift) — that is the explicitly deferred phase 2, and
deserves its own spec once this review workflow exists and shows what's actually worth
pinning. Also out of scope: Docker-based LibreOffice (rejected during brainstorming —
native install was preferred); running the gallery as part of `mvn test` or CI; any
styling/theming of `index.html` beyond plain, readable HTML; multi-page-per-row or
thumbnail layouts; deleting the generated `target/gallery/` directory automatically
(it's inside `target/`, so a normal `mvn clean` already handles that).
