# Document Gallery Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A dev tool that renders every sample document through a real headless LibreOffice and writes one `target/gallery/index.html`, so their Word-accurate look and feel can be reviewed in one browser tab without hand-running each `*SampleMain` and opening its output one file at a time.

**Architecture:** A small, extensible fixture registry (`GalleryFixture`/`GalleryFixtures`) pairs a name with a `WordDocument` builder. `DocxToImages` converts one document to PDF via JODConverter driving a shared headless LibreOffice process, then rasterizes each PDF page to PNG via PDFBox. `GalleryIndexWriter` is a pure function turning rendering results into one `index.html`. `GalleryMain` orchestrates all three and is the runnable entry point (`mvn exec:java@gallery`).

**Tech Stack:** Java 25, Maven, JODConverter 4.4.11 (`jodconverter-local`, test-scope), Apache PDFBox 3.0.8 (test-scope), a local LibreOffice install (auto-detected on a standard Homebrew-cask macOS install).

**Spec:** [docs/superpowers/specs/2026-09-06-document-gallery-design.md](../specs/2026-09-06-document-gallery-design.md)

## Global Constraints

- Recognized/rendered fixtures today: `sample` (`SampleMain`), `postmortem` (`PostmortemSampleMain`), `rich-text` (`RichTextSampleMain`, new). Adding another fixture later must be a one-line addition to `GalleryFixtures.all()` — nothing else in the pipeline changes.
- `HtmlSampleMain` and `PostmortemHtmlSampleMain` stay exactly as they are — not touched, not removed. `GalleryMain` is additive.
- No automated pass/fail in this phase: the gallery never runs as part of `mvn test`, and has no stored-baseline comparison. That is an explicitly separate, not-yet-built follow-up.
- Both new dependencies (`jodconverter-local`, `pdfbox`) are `test`-scope only. The shipped library's own dependencies (docx4j, slf4j) are unchanged.
- Render at 150 DPI. PNG per page. Both the intermediate `.docx` and `.pdf` are kept on disk (not deleted) alongside the PNGs.
- LibreOffice not being found is handled with one clear, actionable message (install instructions + a pointer to `DOCUMENTATION.md`) and a non-zero exit — never a raw stack trace.
- One fixture failing to build or render must not abort the run: it's caught, logged, and shown as a visible "failed to render" note in `index.html` in place of that fixture's images.
- Sample mains (including the new `RichTextSampleMain`) get no dedicated JUnit test, matching the existing convention for `SampleMain`/`PostmortemSampleMain` — their content is already covered by the feature's own tests (`RichTextTest`/`RichTextDocumentTest` for rich text).
- `DocxToImages` and `GalleryMain` — the pieces that depend on an external LibreOffice process — get no JUnit test either, verified instead by actually running `GalleryMain` and inspecting its output (Task 4). A test that starts LibreOffice would run on every `mvn test` invocation on any machine that has it installed, which is exactly the "runs as part of `mvn test`" cost the constraint above rules out — even a test that skips gracefully when LibreOffice is absent still pays that cost when it's present.

---

### Task 1: Fixture registry — `RichTextSampleMain`, `GalleryFixture`, `GalleryFixtures`

**Files:**
- Create: `src/test/java/com/example/docx/sample/RichTextSampleMain.java`
- Modify: `src/test/java/com/example/docx/sample/SampleMain.java:114-119`
- Modify: `src/test/java/com/example/docx/sample/PostmortemSampleMain.java:47-52`
- Create: `src/test/java/com/example/docx/sample/gallery/GalleryFixture.java`
- Create: `src/test/java/com/example/docx/sample/gallery/GalleryFixtures.java`
- Test: `src/test/java/com/example/docx/sample/gallery/GalleryFixturesTest.java`

**Interfaces:**
- Consumes: `WordDocument` (existing), `WordDocument.Builder.heading(String)` / `.paragraph(RichText)` (existing), `RichText.of(String)` (existing, `com.example.docx.content.RichText`).
- Produces: `public static WordDocument RichTextSampleMain.build()`; `public record GalleryFixture(String name, Supplier<WordDocument> builder)`; `public static List<GalleryFixture> GalleryFixtures.all()`. Tasks 2-4 do not consume these directly, but Task 4 (`GalleryMain`) calls `GalleryFixtures.all()`.

- [ ] **Step 1: Write the failing test**

Create `src/test/java/com/example/docx/sample/gallery/GalleryFixturesTest.java`:

```java
package com.example.docx.sample.gallery;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

import com.example.docx.WordDocument;
import java.util.List;
import org.junit.jupiter.api.Test;

class GalleryFixturesTest {

    @Test
    void listsTheThreeExpectedFixturesInOrder() {
        List<GalleryFixture> fixtures = GalleryFixtures.all();
        assertEquals(List.of("sample", "postmortem", "rich-text"),
                fixtures.stream().map(GalleryFixture::name).toList());
    }

    @Test
    void everyFixtureBuildsAWordDocumentWithoutThrowing() {
        for (GalleryFixture fixture : GalleryFixtures.all()) {
            WordDocument document = fixture.builder().get();
            assertNotNull(document, fixture.name() + " must build a document");
        }
    }
}
```

- [ ] **Step 2: Run test to verify it fails**

Run: `mvn -q -Dtest=GalleryFixturesTest test`
Expected: FAIL to compile — `GalleryFixture`/`GalleryFixtures` do not exist yet.

- [ ] **Step 3: Create `RichTextSampleMain`**

Create `src/test/java/com/example/docx/sample/RichTextSampleMain.java`:

```java
package com.example.docx.sample;

import com.example.docx.WordDocument;
import com.example.docx.content.RichText;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 * Writes a timestamped {@code target/rich-text-*.docx} so mixed inline formatting can be
 * eyeballed.
 *
 * <p>A small, standalone document — plain text, bold, italic, underline, and nested
 * formatting — rather than the full report {@link SampleMain} builds. The feature itself
 * is already covered by {@code RichTextTest}/{@code RichTextDocumentTest}; this exists
 * purely so the visual result has a sample of its own.
 */
public final class RichTextSampleMain {

    private RichTextSampleMain() {
    }

    public static void main(String[] args) throws IOException {
        WordDocument document = build();

        Path target = Path.of("target", "rich-text-%d.docx".formatted(System.currentTimeMillis()));
        Files.createDirectories(target.getParent());
        try (OutputStream out = Files.newOutputStream(target)) {
            document.writeTo(out);
        }
        System.out.println("Wrote " + target.toAbsolutePath());
    }

    /**
     * Builds the same document {@link #main} writes to disk, so other tools (such as the
     * document gallery) can render it without duplicating the content.
     */
    public static WordDocument build() {
        return WordDocument.builder()
                .heading("Rich text")
                .paragraph(RichText.of("Some text which needs to be <b>bold</b>."))
                .paragraph(RichText.of("Some text which needs to be <i>italic</i>."))
                .paragraph(RichText.of("Some text which needs to be <u>underlined</u>."))
                .paragraph(RichText.of(
                        "Tags nest: <b>bold, <i>bold and italic</i>, back to just bold</b>."))
                .paragraph(RichText.of(
                        "Aliases work the same way: <strong>strong</strong> and <em>em</em>."))
                .build();
    }
}
```

- [ ] **Step 4: Widen `SampleMain.build()` and `PostmortemSampleMain.build()` to `public`**

In `src/test/java/com/example/docx/sample/SampleMain.java`, replace lines 114-119:

```java
    /**
     * Builds the same document {@link #main} writes to disk, so other samples (such as
     * {@code HtmlSampleMain}) can render it in a different format without duplicating
     * the content.
     */
    static WordDocument build() {
```

with:

```java
    /**
     * Builds the same document {@link #main} writes to disk, so other tools (such as
     * {@code HtmlSampleMain} and the document gallery) can render it without duplicating
     * the content.
     */
    public static WordDocument build() {
```

In `src/test/java/com/example/docx/sample/PostmortemSampleMain.java`, replace lines 47-52:

```java
    /**
     * Builds the same document {@link #main} writes to disk, so other samples (such as
     * {@code PostmortemHtmlSampleMain}) can render it in a different format without
     * duplicating the content.
     */
    static WordDocument build() {
```

with:

```java
    /**
     * Builds the same document {@link #main} writes to disk, so other tools (such as
     * {@code PostmortemHtmlSampleMain} and the document gallery) can render it without
     * duplicating the content.
     */
    public static WordDocument build() {
```

Both classes are `public final class`, so this is the only change needed for a
different-package caller to reach `build()`.

- [ ] **Step 5: Create `GalleryFixture`**

Create `src/test/java/com/example/docx/sample/gallery/GalleryFixture.java`:

```java
package com.example.docx.sample.gallery;

import com.example.docx.WordDocument;
import java.util.function.Supplier;

/** One document to render for review: a display name paired with how to build it. */
public record GalleryFixture(String name, Supplier<WordDocument> builder) {
}
```

- [ ] **Step 6: Create `GalleryFixtures`**

Create `src/test/java/com/example/docx/sample/gallery/GalleryFixtures.java`:

```java
package com.example.docx.sample.gallery;

import com.example.docx.sample.PostmortemSampleMain;
import com.example.docx.sample.RichTextSampleMain;
import com.example.docx.sample.SampleMain;
import java.util.List;

/**
 * The documents the gallery renders. Add a fixture here — one line — to have it show up
 * in {@code target/gallery/index.html} the next time {@code GalleryMain} runs.
 */
public final class GalleryFixtures {

    private GalleryFixtures() {
    }

    public static List<GalleryFixture> all() {
        return List.of(
                new GalleryFixture("sample", SampleMain::build),
                new GalleryFixture("postmortem", PostmortemSampleMain::build),
                new GalleryFixture("rich-text", RichTextSampleMain::build));
    }
}
```

- [ ] **Step 7: Run test to verify it passes**

Run: `mvn -q -Dtest=GalleryFixturesTest test`
Expected: PASS, both tests green.

- [ ] **Step 8: Run the full suite**

Run: `mvn -q test`
Expected: PASS, no regressions (the visibility widening is additive, so no existing test should be affected).

- [ ] **Step 9: Commit**

```bash
git add src/test/java/com/example/docx/sample/RichTextSampleMain.java \
        src/test/java/com/example/docx/sample/SampleMain.java \
        src/test/java/com/example/docx/sample/PostmortemSampleMain.java \
        src/test/java/com/example/docx/sample/gallery/GalleryFixture.java \
        src/test/java/com/example/docx/sample/gallery/GalleryFixtures.java \
        src/test/java/com/example/docx/sample/gallery/GalleryFixturesTest.java
git commit -m "feat: add RichTextSampleMain and the gallery fixture registry

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_015AMYuNu7yCy7ja3gmwxaoo"
```

---

### Task 2: `GalleryIndexWriter` and `RenderedFixture`

**Files:**
- Create: `src/test/java/com/example/docx/sample/gallery/RenderedFixture.java`
- Create: `src/test/java/com/example/docx/sample/gallery/GalleryIndexWriter.java`
- Test: `src/test/java/com/example/docx/sample/gallery/GalleryIndexWriterTest.java`

**Interfaces:**
- Consumes: nothing from Task 1 (pure, standalone).
- Produces: `public record RenderedFixture(String name, List<Path> pages, String error)` with `static RenderedFixture success(String name, List<Path> pages)`, `static RenderedFixture failure(String name, String error)`, and `boolean failed()`; `public static String GalleryIndexWriter.write(List<RenderedFixture> fixtures)`. Task 4 (`GalleryMain`) constructs `RenderedFixture` values (one per fixture, from Task 1's `GalleryFixture` and Task 3's `DocxToImages` output) and calls `GalleryIndexWriter.write(...)`.

- [ ] **Step 1: Write the failing test**

Create `src/test/java/com/example/docx/sample/gallery/GalleryIndexWriterTest.java`:

```java
package com.example.docx.sample.gallery;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;

class GalleryIndexWriterTest {

    @Test
    void includesEveryFixtureNameAndItsPageImages() {
        String html = GalleryIndexWriter.write(List.of(
                RenderedFixture.success("sample",
                        List.of(Path.of("sample/page-1.png"), Path.of("sample/page-2.png"))),
                RenderedFixture.success("rich-text",
                        List.of(Path.of("rich-text/page-1.png")))));

        assertTrue(html.contains("sample"));
        assertTrue(html.contains("rich-text"));
        assertTrue(html.contains("sample/page-1.png"));
        assertTrue(html.contains("sample/page-2.png"));
        assertTrue(html.contains("rich-text/page-1.png"));
    }

    @Test
    void showsAFailureNoteInsteadOfImagesWhenAFixtureFailed() {
        String html = GalleryIndexWriter.write(List.of(
                RenderedFixture.failure("broken", "boom")));

        assertTrue(html.contains("broken"));
        assertTrue(html.contains("Failed to render"));
        assertTrue(html.contains("boom"));
        assertFalse(html.contains("<img"));
    }

    @Test
    void escapesHtmlSpecialCharactersInNamesAndErrors() {
        String html = GalleryIndexWriter.write(List.of(
                RenderedFixture.failure("a & b", "<oops>")));

        assertTrue(html.contains("a &amp; b"));
        assertTrue(html.contains("&lt;oops&gt;"));
    }
}
```

- [ ] **Step 2: Run test to verify it fails**

Run: `mvn -q -Dtest=GalleryIndexWriterTest test`
Expected: FAIL to compile — `RenderedFixture`/`GalleryIndexWriter` do not exist yet.

- [ ] **Step 3: Create `RenderedFixture`**

Create `src/test/java/com/example/docx/sample/gallery/RenderedFixture.java`:

```java
package com.example.docx.sample.gallery;

import java.nio.file.Path;
import java.util.List;

/**
 * One fixture's rendering result: its page images on success, or {@code error} (with
 * {@code pages} empty) if rendering failed. Paths are relative to the gallery output
 * directory, exactly as written into {@code index.html}'s {@code <img src>} attributes.
 */
public record RenderedFixture(String name, List<Path> pages, String error) {

    public static RenderedFixture success(String name, List<Path> pages) {
        return new RenderedFixture(name, pages, null);
    }

    public static RenderedFixture failure(String name, String error) {
        return new RenderedFixture(name, List.of(), error);
    }

    public boolean failed() {
        return error != null;
    }
}
```

- [ ] **Step 4: Create `GalleryIndexWriter`**

Create `src/test/java/com/example/docx/sample/gallery/GalleryIndexWriter.java`:

```java
package com.example.docx.sample.gallery;

import java.nio.file.Path;
import java.util.List;

/**
 * Builds the {@code index.html} string for a gallery run: one section per fixture, its
 * page images stacked in order, or a visible failure note if it didn't render. Pure
 * string-in-string-out — no I/O, no LibreOffice — so it's unit-testable on its own.
 */
public final class GalleryIndexWriter {

    private GalleryIndexWriter() {
    }

    public static String write(List<RenderedFixture> fixtures) {
        StringBuilder html = new StringBuilder();
        html.append("<!doctype html>\n<html><head><meta charset=\"utf-8\">")
                .append("<title>Document gallery</title></head><body>\n");

        for (RenderedFixture fixture : fixtures) {
            html.append("<section>\n<h2>").append(escape(fixture.name())).append("</h2>\n");
            if (fixture.failed()) {
                html.append("<p>Failed to render: ").append(escape(fixture.error())).append("</p>\n");
            } else {
                int total = fixture.pages().size();
                for (int i = 0; i < total; i++) {
                    Path page = fixture.pages().get(i);
                    html.append("<p>Page ").append(i + 1).append(" of ").append(total).append("</p>\n")
                            .append("<img src=\"").append(escape(page.toString())).append("\">\n");
                }
            }
            html.append("</section>\n<hr>\n");
        }

        html.append("</body></html>\n");
        return html.toString();
    }

    private static String escape(String text) {
        return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                .replace("\"", "&quot;");
    }
}
```

- [ ] **Step 5: Run test to verify it passes**

Run: `mvn -q -Dtest=GalleryIndexWriterTest test`
Expected: PASS, all three tests green.

- [ ] **Step 6: Run the full suite**

Run: `mvn -q test`
Expected: PASS, no regressions.

- [ ] **Step 7: Commit**

```bash
git add src/test/java/com/example/docx/sample/gallery/RenderedFixture.java \
        src/test/java/com/example/docx/sample/gallery/GalleryIndexWriter.java \
        src/test/java/com/example/docx/sample/gallery/GalleryIndexWriterTest.java
git commit -m "feat: add GalleryIndexWriter, building index.html from rendering results

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_015AMYuNu7yCy7ja3gmwxaoo"
```

---

### Task 3: `DocxToImages` (LibreOffice + PDFBox rendering)

**Files:**
- Modify: `pom.xml`
- Create: `src/test/java/com/example/docx/sample/gallery/DocxToImages.java`

**Interfaces:**
- Consumes: `WordDocument.writeTo(OutputStream)` (existing).
- Produces: `public static List<Path> DocxToImages.renderPages(OfficeManager officeManager, WordDocument document, Path outputDir) throws OfficeException, IOException`. Task 4 (`GalleryMain`) calls this once per fixture, reusing one already-started `OfficeManager` across all calls.

**No JUnit test for this task** (see the Global Constraints entry on `DocxToImages`/
`GalleryMain`) — a test that starts LibreOffice would add that cost to every `mvn test`
run on any machine that has it installed. This class is verified by actually running it
through `GalleryMain` in Task 4 and inspecting the result.

- [ ] **Step 1: Add the two new test-scope dependencies**

In `pom.xml`, inside `<dependencies>`, immediately after the existing `junit-jupiter`
dependency block, add:

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

- [ ] **Step 2: Create `DocxToImages`**

Create `src/test/java/com/example/docx/sample/gallery/DocxToImages.java`:

```java
package com.example.docx.sample.gallery;

import com.example.docx.WordDocument;
import java.awt.image.BufferedImage;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import javax.imageio.ImageIO;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.jodconverter.core.office.OfficeException;
import org.jodconverter.core.office.OfficeManager;
import org.jodconverter.local.LocalConverter;

/**
 * Renders a {@link WordDocument} to one PNG per page, at Word-accurate fidelity, via a
 * real headless LibreOffice: {@code .docx -> .pdf} (JODConverter) then {@code .pdf ->
 * PNG} per page (PDFBox). {@code officeManager} must already be started by the caller —
 * starting one costs a few seconds, so {@code GalleryMain} pays that once for the whole
 * run, not once per fixture.
 */
public final class DocxToImages {

    private static final int RENDER_DPI = 150;

    private DocxToImages() {
    }

    public static List<Path> renderPages(OfficeManager officeManager, WordDocument document, Path outputDir)
            throws OfficeException, IOException {
        Files.createDirectories(outputDir);

        Path docx = outputDir.resolve("document.docx");
        try (OutputStream out = Files.newOutputStream(docx)) {
            document.writeTo(out);
        }

        Path pdf = outputDir.resolve("document.pdf");
        LocalConverter.make(officeManager).convert(docx.toFile()).to(pdf.toFile()).execute();

        List<Path> pages = new ArrayList<>();
        try (PDDocument pdDocument = Loader.loadPDF(pdf.toFile())) {
            PDFRenderer renderer = new PDFRenderer(pdDocument);
            for (int page = 0; page < pdDocument.getNumberOfPages(); page++) {
                BufferedImage image = renderer.renderImageWithDPI(page, RENDER_DPI, ImageType.RGB);
                Path pngPath = outputDir.resolve("page-" + (page + 1) + ".png");
                ImageIO.write(image, "PNG", pngPath.toFile());
                pages.add(pngPath);
            }
        }
        return pages;
    }
}
```

- [ ] **Step 3: Compile it**

Run: `mvn -q test-compile`
Expected: compiles cleanly against the two new dependencies added in Step 1.

- [ ] **Step 4: Run the full suite**

Run: `mvn -q test`
Expected: PASS, no regressions — adding a compiled-but-not-yet-called class changes
nothing about existing behavior.

- [ ] **Step 5: Commit**

```bash
git add pom.xml \
        src/test/java/com/example/docx/sample/gallery/DocxToImages.java
git commit -m "feat: add DocxToImages, rendering a WordDocument to page PNGs via LibreOffice

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_015AMYuNu7yCy7ja3gmwxaoo"
```

Its actual runtime correctness (does it produce the right number of non-empty PNGs
against a real LibreOffice) is confirmed in Task 4, Steps 3-4, the first point it's
actually called from `GalleryMain`.

---

### Task 4: `GalleryMain` and Maven wiring

**Files:**
- Create: `src/test/java/com/example/docx/sample/gallery/GalleryMain.java`
- Modify: `pom.xml`

**Interfaces:**
- Consumes: `GalleryFixtures.all()` (Task 1), `DocxToImages.renderPages(...)` (Task 3), `RenderedFixture.success/failure` and `GalleryIndexWriter.write(...)` (Task 2).
- Produces: nothing further downstream — this is the tool's public entry point.

**No JUnit test for this task**, matching the established convention for `*Main`
orchestrators (`SampleMain`, `PostmortemSampleMain`, `HtmlSampleMain` all have none) — the
whole point of this tool is a human looking at its output, and its three pieces are
already independently tested in Tasks 1-3. Verification here is: run it, and look at what
it produces (Steps 3-4).

- [ ] **Step 1: Create `GalleryMain`**

Create `src/test/java/com/example/docx/sample/gallery/GalleryMain.java`:

```java
package com.example.docx.sample.gallery;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import org.jodconverter.core.office.OfficeException;
import org.jodconverter.core.office.OfficeManager;
import org.jodconverter.local.office.LocalOfficeManager;

/**
 * Renders every {@link GalleryFixtures fixture} through LibreOffice and writes one
 * {@code target/gallery/index.html} so they can all be reviewed in one browser tab,
 * without hand-running each sample main and opening its output one file at a time.
 *
 * <p>Requires a local LibreOffice install (a standard Homebrew-cask install on macOS is
 * auto-detected; see DOCUMENTATION.md). Run with {@code mvn exec:java@gallery}.
 */
public final class GalleryMain {

    private static final Path GALLERY_DIR = Path.of("target", "gallery");

    private GalleryMain() {
    }

    public static void main(String[] args) throws IOException {
        OfficeManager officeManager = LocalOfficeManager.builder().install().build();
        try {
            officeManager.start();
        } catch (OfficeException e) {
            System.err.println("Could not start LibreOffice: " + e.getMessage());
            System.err.println("Install it (e.g. `brew install --cask libreoffice` on "
                    + "macOS) and try again. See DOCUMENTATION.md's \"Document gallery\" "
                    + "section.");
            System.exit(1);
            return;
        }

        try {
            List<RenderedFixture> rendered = new ArrayList<>();
            for (GalleryFixture fixture : GalleryFixtures.all()) {
                rendered.add(render(officeManager, fixture));
            }

            Files.createDirectories(GALLERY_DIR);
            Path indexPath = GALLERY_DIR.resolve("index.html");
            Files.writeString(indexPath, GalleryIndexWriter.write(rendered));
            System.out.println("Wrote " + indexPath.toAbsolutePath());
        } finally {
            try {
                officeManager.stop();
            } catch (OfficeException e) {
                System.err.println("Failed to stop LibreOffice cleanly: " + e.getMessage());
            }
        }
    }

    /**
     * One fixture's build+render, isolated so a single bad fixture (a build that throws,
     * a conversion LibreOffice chokes on) shows up as a visible failure note in the
     * index rather than aborting every other fixture's output. Catching the broad
     * {@code Exception} here is deliberate: a fixture's {@code builder()} can throw an
     * unchecked {@code DocumentGenerationException}, and {@code DocxToImages} throws the
     * checked {@code OfficeException}/{@code IOException} — both must be caught the same
     * way, without aborting the run.
     */
    private static RenderedFixture render(OfficeManager officeManager, GalleryFixture fixture) {
        try {
            List<Path> pages = DocxToImages.renderPages(
                    officeManager, fixture.builder().get(), GALLERY_DIR.resolve(fixture.name()));
            List<Path> relativePages = pages.stream()
                    .map(page -> Path.of(fixture.name()).resolve(page.getFileName()))
                    .toList();
            return RenderedFixture.success(fixture.name(), relativePages);
        } catch (Exception e) {
            System.err.println("Failed to render '" + fixture.name() + "': " + e.getMessage());
            return RenderedFixture.failure(fixture.name(), String.valueOf(e.getMessage()));
        }
    }
}
```

- [ ] **Step 2: Wire the `gallery` exec-plugin execution**

In `pom.xml`, inside the `exec-maven-plugin`'s `<executions>` block, add a new execution
after the existing `postmortem-html` one:

```xml
                    <execution>
                        <id>gallery</id>
                        <goals>
                            <goal>java</goal>
                        </goals>
                        <configuration>
                            <mainClass>com.example.docx.sample.gallery.GalleryMain</mainClass>
                        </configuration>
                    </execution>
```

- [ ] **Step 3: Compile and run it for real**

Run: `mvn -q test-compile`
Expected: compiles cleanly.

Run: `mvn exec:java@gallery`
Expected: prints `Wrote <absolute path>/target/gallery/index.html` with no `[ERROR]`
lines. If LibreOffice isn't installed, it instead prints the actionable message from
Step 1 and exits non-zero — install LibreOffice and re-run before continuing.

- [ ] **Step 4: Look at the result**

Open `target/gallery/index.html` in a browser. Confirm: three sections (`sample`,
`postmortem`, `rich-text`), each showing real rendered page images (not blank, not
broken-image icons), and the rich-text section visibly shows bold/italic/underlined text
matching what `RichTextSampleMain.build()` writes.

- [ ] **Step 5: Run the full suite**

Run: `mvn -q test`
Expected: PASS, no regressions.

- [ ] **Step 6: Commit**

```bash
git add src/test/java/com/example/docx/sample/gallery/GalleryMain.java pom.xml
git commit -m "feat: add GalleryMain, the document gallery's runnable entry point

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_015AMYuNu7yCy7ja3gmwxaoo"
```

---

### Task 5: Documentation

**Files:**
- Modify: `DOCUMENTATION.md`

**Interfaces:** none — this task only adds prose.

- [ ] **Step 1: Add the "Document gallery" section**

In `DOCUMENTATION.md`, immediately after the existing `## Verifying output yourself`
section (at the end of the file), add:

```markdown
---

## Document gallery

A dev tool for seeing every sample document's real Word-rendered look and feel in one
browser tab, without hand-running each `*SampleMain` and opening its output one file at a
time:

```bash
mvn exec:java@gallery
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
```

- [ ] **Step 2: Verify the doc reads correctly**

Read the file section back and confirm it fits the surrounding tone (matches
`## Verifying output yourself` and the rest of the guide's plain, direct style) and that
the `mvn exec:java@gallery` command matches exactly what Task 4 wired up.

- [ ] **Step 3: Commit**

```bash
git add DOCUMENTATION.md
git commit -m "docs: document the gallery dev tool

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_015AMYuNu7yCy7ja3gmwxaoo"
```
