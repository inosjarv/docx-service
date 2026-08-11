# Multi-Section Documents Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Let a document hold any number of headings, body paragraphs and images in insertion order, so a runtime-decided count of sections is just a loop.

**Architecture:** `WordDocument.Builder`'s four single-valued fields are replaced by one ordered `List<DocumentContent>` of deferred entries — deferred because an image is a package part plus a relationship and no package exists until `build()`. `Headings` merges into `Paragraphs` (identical structure, differing only in styles), `HeadingStyle` is renamed `TextStyle` because that is what it always was, and a new `ParagraphStyle` carries spacing and `keepWithNext`.

**Tech Stack:** Java 25, Maven, docx4j 17.0.2, JUnit 6.1.3.

## Global Constraints

- Java release level is exactly `25`. Run Maven as `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test`.
- **No new dependencies.**
- All geometry and spacing is in **twips**. `w:spacing`'s `w:before` / `w:after` are twentieths of a point, i.e. the same twips used elsewhere. 240 twips = 12 pt; 120 twips = 6 pt.
- Font size is in **points**, stored as half-points (`w:sz` = `sizePt × 2`). 11 pt emits `22`. Never pixels.
- Colour is bare uppercase `RRGGBB`, no `#`.
- Every `w:t` carries `xml:space="preserve"`.
- `DocumentGenerationException extends RuntimeException` is the ONLY exception thrown deliberately. Validation failures carry a message naming the field and NO cause; wrapped failures ALWAYS carry a cause.
- Layering: only `page/`, `part/`, `DocumentContent` and `WordDocument` may reference `WordprocessingMLPackage`. `content/` and `style/` stay pure.
- `keepWithNext` must be emitted only when true — absent, never present-and-false.
- A4 is the literal pair `11906 × 16838`; the default margin is the literal `851`; `marginsCm(1.5)` yields 850. Never conflated.
- Do not modify `src/test/resources/demo/`.
- Existing behaviour that must not change: `TextStyle.defaults()` stays Calibri Light / 20 pt / bold / `1F4E79`, and every existing SVG, page-setup, units and stream test must keep passing.

## File Structure

| File | Change | Responsibility |
| --- | --- | --- |
| `src/main/java/com/example/docx/style/HeadingStyle.java` | Rename → `TextStyle.java` | Run formatting, now shared by headings and body |
| `src/test/java/com/example/docx/style/HeadingStyleTest.java` | Rename → `TextStyleTest.java` | |
| `src/main/java/com/example/docx/style/ParagraphStyle.java` | Create | Spacing + `keepWithNext` → `w:pPr` |
| `src/test/java/com/example/docx/style/ParagraphStyleTest.java` | Create | |
| `src/main/java/com/example/docx/content/Paragraphs.java` | Create | `(text, TextStyle, ParagraphStyle)` → `w:p` |
| `src/test/java/com/example/docx/content/ParagraphsTest.java` | Create | |
| `src/main/java/com/example/docx/content/Headings.java` | Delete | Merged into `Paragraphs` |
| `src/test/java/com/example/docx/content/HeadingsTest.java` | Delete | Replaced by `ParagraphsTest` |
| `src/main/java/com/example/docx/DocumentContent.java` | Create | Deferred content entry |
| `src/main/java/com/example/docx/WordDocument.java` | Modify | Ordered list replaces four fields |
| `src/test/java/com/example/docx/MultiSectionTest.java` | Create | Ordering, repeat calls, keepNext |
| `src/test/java/com/example/docx/WordDocumentRoundTripTest.java` | Modify | Drop `requiresAHeading` (superseded) |
| `src/test/java/com/example/docx/sample/SampleMain.java` | Modify | Demo several sections |
| `README.md` | Modify | Document the loop and the new types |

Test counts: 53 now → **62** at the end. `HeadingsTest`'s 4 are replaced by `ParagraphsTest`'s 4; `TextStyleTest` gains 1, `ParagraphStyleTest` adds 4, `MultiSectionTest` adds 5; `WordDocumentRoundTripTest` loses 1.

---

### Task 1: Rename `HeadingStyle` to `TextStyle` and add `TextStyle.body()`

**Files:**
- Rename: `src/main/java/com/example/docx/style/HeadingStyle.java` → `TextStyle.java`
- Rename: `src/test/java/com/example/docx/style/HeadingStyleTest.java` → `TextStyleTest.java`
- Modify: every file referencing `HeadingStyle`

**Interfaces:**
- Consumes: nothing new.
- Produces: `com.example.docx.style.TextStyle` with all of `HeadingStyle`'s existing members unchanged (`builder()`, `defaults()`, `fontFamily()`, `sizePt()`, `bold()`, `italic()`, `colorHex()`, `toRPr()`, `MAX_SIZE_PT`), plus `public static TextStyle body()`.

The class only ever held run formatting — font, size, bold, italic, colour — which body text needs identically. The artifact is `1.0.0-SNAPSHOT` and unmerged, so no external consumer breaks.

- [ ] **Step 1: Rename the type everywhere**

```bash
for f in $(grep -rl "HeadingStyle" src README.md); do
  sed -i '' 's/HeadingStyle/TextStyle/g' "$f"
done
git mv src/main/java/com/example/docx/style/HeadingStyle.java src/main/java/com/example/docx/style/TextStyle.java
git mv src/test/java/com/example/docx/style/HeadingStyleTest.java src/test/java/com/example/docx/style/TextStyleTest.java
```

- [ ] **Step 2: Confirm no references remain**

Run: `grep -rn "HeadingStyle" src README.md`
Expected: no output.

- [ ] **Step 3: Confirm the suite still passes**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B clean test`
Expected: `Tests run: 53, Failures: 0, Errors: 0, Skipped: 0`.

A pure rename must not change behaviour; if a test fails here, the sed hit something it should not have.

- [ ] **Step 4: Write the failing test for `body()`**

Add to `src/test/java/com/example/docx/style/TextStyleTest.java`:

```java
    @Test
    void bodyIsElevenPointRegularBlack() {
        TextStyle body = TextStyle.body();
        assertEquals("Calibri", body.fontFamily());
        assertEquals(11.0, body.sizePt());
        assertEquals(false, body.bold());
        assertEquals(false, body.italic());
        assertEquals("000000", body.colorHex());
        // 11 pt is 22 half-points. Points, never pixels.
        assertEquals(BigInteger.valueOf(22), body.toRPr().getSz().getVal());
    }
```

If `java.math.BigInteger` is not already imported in that file, add `import java.math.BigInteger;`.

- [ ] **Step 5: Run it to confirm it fails**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test -Dtest=TextStyleTest`
Expected: FAIL — compilation error, `cannot find symbol: method body()`.

- [ ] **Step 6: Add `body()`**

In `src/main/java/com/example/docx/style/TextStyle.java`, add immediately after `defaults()`:

```java
    /** Body text: Calibri 11 pt, regular, black. */
    public static TextStyle body() {
        return builder().font("Calibri").sizePt(11).bold(false).color("000000").build();
    }
```

- [ ] **Step 7: Run the suite**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B clean test`
Expected: `Tests run: 54, Failures: 0, Errors: 0, Skipped: 0`.

- [ ] **Step 8: Commit**

```bash
git add -A
git commit -m "refactor: rename HeadingStyle to TextStyle and add body()"
```

---

### Task 2: `ParagraphStyle`

**Files:**
- Create: `src/main/java/com/example/docx/style/ParagraphStyle.java`
- Create: `src/test/java/com/example/docx/style/ParagraphStyleTest.java`

**Interfaces:**
- Consumes: `DocumentGenerationException`.
- Produces: `ParagraphStyle.builder()`, `ParagraphStyle.heading()`, `ParagraphStyle.body()`; accessors `spaceBeforeTwips()`, `spaceAfterTwips()`, `keepWithNext()` returning `int`, `int`, `boolean`; `toPPr()` returning `org.docx4j.wml.PPr`. Builder methods `spaceBeforeTwips(int)`, `spaceAfterTwips(int)`, `keepWithNext(boolean)`, `build()`.

`keepWithNext` emits `w:keepNext`, gluing a paragraph to the one after it. Headings set it, so a heading never strands at the foot of a page with its body overleaf — a certainty rather than a risk once a document has ten sections.

- [ ] **Step 1: Write the failing test**

Create `src/test/java/com/example/docx/style/ParagraphStyleTest.java`:

```java
package com.example.docx.style;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.example.docx.DocumentGenerationException;
import java.math.BigInteger;
import org.docx4j.wml.PPr;
import org.junit.jupiter.api.Test;

class ParagraphStyleTest {

    @Test
    void headingKeepsWithNext() {
        ParagraphStyle style = ParagraphStyle.heading();
        assertTrue(style.keepWithNext());
        assertEquals(240, style.spaceBeforeTwips());
        assertEquals(120, style.spaceAfterTwips());
    }

    @Test
    void bodyDoesNotKeepWithNext() {
        ParagraphStyle style = ParagraphStyle.body();
        assertEquals(false, style.keepWithNext());
        assertEquals(0, style.spaceBeforeTwips());
        assertEquals(120, style.spaceAfterTwips());
    }

    @Test
    void emitsSpacingAndKeepNext() {
        PPr heading = ParagraphStyle.heading().toPPr();
        assertEquals(BigInteger.valueOf(240), heading.getSpacing().getBefore());
        assertEquals(BigInteger.valueOf(120), heading.getSpacing().getAfter());
        assertNotNull(heading.getKeepNext());
        assertTrue(heading.getKeepNext().isVal());

        PPr body = ParagraphStyle.body().toPPr();
        assertNull(body.getKeepNext(), "keepNext must be absent, not present-and-false");
    }

    @Test
    void rejectsNegativeSpacing() {
        assertThrows(DocumentGenerationException.class,
                () -> ParagraphStyle.builder().spaceBeforeTwips(-1).build());
        assertThrows(DocumentGenerationException.class,
                () -> ParagraphStyle.builder().spaceAfterTwips(-1).build());
    }
}
```

- [ ] **Step 2: Run it to confirm it fails**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test -Dtest=ParagraphStyleTest`
Expected: FAIL — compilation error, `cannot find symbol: class ParagraphStyle`.

- [ ] **Step 3: Write the implementation**

Create `src/main/java/com/example/docx/style/ParagraphStyle.java`:

```java
package com.example.docx.style;

import com.example.docx.DocumentGenerationException;
import java.math.BigInteger;
import org.docx4j.jaxb.Context;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.PPr;
import org.docx4j.wml.PPrBase;

/**
 * Immutable paragraph-level formatting: spacing in twips, and keep-with-next.
 *
 * <p>{@code w:spacing}'s before and after are twentieths of a point, which is the same
 * twip the rest of this library uses. 240 twips is 12 pt; 120 twips is 6 pt.
 */
public final class ParagraphStyle {

    public static final int HEADING_SPACE_BEFORE_TWIPS = 240;
    public static final int HEADING_SPACE_AFTER_TWIPS = 120;
    public static final int BODY_SPACE_AFTER_TWIPS = 120;

    private final int spaceBeforeTwips;
    private final int spaceAfterTwips;
    private final boolean keepWithNext;

    private ParagraphStyle(Builder b) {
        this.spaceBeforeTwips = b.spaceBeforeTwips;
        this.spaceAfterTwips = b.spaceAfterTwips;
        this.keepWithNext = b.keepWithNext;
    }

    public static Builder builder() {
        return new Builder();
    }

    /**
     * Heading spacing, keeping with the paragraph that follows.
     *
     * <p>Without {@code keepWithNext} a heading can land at the foot of a page with its
     * body overleaf — which stops being a risk and becomes a certainty once a document
     * has many sections.
     */
    public static ParagraphStyle heading() {
        return builder()
                .spaceBeforeTwips(HEADING_SPACE_BEFORE_TWIPS)
                .spaceAfterTwips(HEADING_SPACE_AFTER_TWIPS)
                .keepWithNext(true)
                .build();
    }

    public static ParagraphStyle body() {
        return builder()
                .spaceBeforeTwips(0)
                .spaceAfterTwips(BODY_SPACE_AFTER_TWIPS)
                .keepWithNext(false)
                .build();
    }

    public int spaceBeforeTwips() {
        return spaceBeforeTwips;
    }

    public int spaceAfterTwips() {
        return spaceAfterTwips;
    }

    public boolean keepWithNext() {
        return keepWithNext;
    }

    /** Builds the {@code w:pPr} for this style. */
    public PPr toPPr() {
        ObjectFactory factory = Context.getWmlObjectFactory();
        PPr pPr = factory.createPPr();

        PPrBase.Spacing spacing = factory.createPPrBaseSpacing();
        spacing.setBefore(BigInteger.valueOf(spaceBeforeTwips));
        spacing.setAfter(BigInteger.valueOf(spaceAfterTwips));
        pPr.setSpacing(spacing);

        // Emitted only when true; an explicit false is not the same as absent.
        if (keepWithNext) {
            BooleanDefaultTrue on = factory.createBooleanDefaultTrue();
            on.setVal(Boolean.TRUE);
            pPr.setKeepNext(on);
        }
        return pPr;
    }

    /** Fluent builder. Every field is defaulted; {@code ParagraphStyle.body()} is valid alone. */
    public static final class Builder {

        private int spaceBeforeTwips = 0;
        private int spaceAfterTwips = BODY_SPACE_AFTER_TWIPS;
        private boolean keepWithNext = false;

        private Builder() {
        }

        public Builder spaceBeforeTwips(int twips) {
            this.spaceBeforeTwips = twips;
            return this;
        }

        public Builder spaceAfterTwips(int twips) {
            this.spaceAfterTwips = twips;
            return this;
        }

        public Builder keepWithNext(boolean keepWithNext) {
            this.keepWithNext = keepWithNext;
            return this;
        }

        public ParagraphStyle build() {
            requireNonNegative(spaceBeforeTwips, "space before");
            requireNonNegative(spaceAfterTwips, "space after");
            return new ParagraphStyle(this);
        }

        private static void requireNonNegative(int value, String field) {
            if (value < 0) {
                throw new DocumentGenerationException(
                        field + " must not be negative, got " + value);
            }
        }
    }
}
```

- [ ] **Step 4: Run the suite**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B clean test`
Expected: `Tests run: 58, Failures: 0, Errors: 0, Skipped: 0`.

- [ ] **Step 5: Commit**

```bash
git add -A
git commit -m "feat: add ParagraphStyle with spacing and keep-with-next"
```

---

### Task 3: `Paragraphs` replaces `Headings`

**Files:**
- Create: `src/main/java/com/example/docx/content/Paragraphs.java`
- Create: `src/test/java/com/example/docx/content/ParagraphsTest.java`
- Delete: `src/main/java/com/example/docx/content/Headings.java`
- Delete: `src/test/java/com/example/docx/content/HeadingsTest.java`
- Modify: `src/main/java/com/example/docx/WordDocument.java` (call site only)

**Interfaces:**
- Consumes: `TextStyle`, `ParagraphStyle`, `DocumentGenerationException`.
- Produces: `public static org.docx4j.wml.P Paragraphs.of(String text, TextStyle textStyle, ParagraphStyle paragraphStyle)` on a final, non-instantiable class. `Headings` no longer exists.

A heading paragraph and a body paragraph are the same structure — one `w:p` holding one `w:r`. They differ only in the styles applied, so two factories would be verbatim duplication.

**Object-model note, verified by running it:** a freshly built run holds a bare `org.docx4j.wml.Text`; one reloaded via `WordprocessingMLPackage.load()` holds a `JAXBElement<Text>`. The test helper below tolerates both, which is why it is safe to copy between test classes.

- [ ] **Step 1: Write the failing test**

Create `src/test/java/com/example/docx/content/ParagraphsTest.java`:

```java
package com.example.docx.content;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

import com.example.docx.DocumentGenerationException;
import com.example.docx.style.ParagraphStyle;
import com.example.docx.style.TextStyle;
import jakarta.xml.bind.JAXBElement;
import java.math.BigInteger;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Text;
import org.junit.jupiter.api.Test;

class ParagraphsTest {

    /** Freshly built runs hold a bare Text; reloaded ones hold a JAXBElement. */
    private static Text textOf(R run) {
        Object first = run.getContent().get(0);
        return (Text) (first instanceof JAXBElement<?> je ? je.getValue() : first);
    }

    @Test
    void buildsOneStyledRunWithParagraphProperties() {
        P p = Paragraphs.of("Overview", TextStyle.defaults(), ParagraphStyle.heading());

        assertEquals(1, p.getContent().size());
        R run = (R) p.getContent().get(0);
        assertEquals("Overview", textOf(run).getValue());
        assertEquals(BigInteger.valueOf(40), run.getRPr().getSz().getVal());
        assertNotNull(p.getPPr().getKeepNext(), "heading style must set keepNext");
        assertEquals(BigInteger.valueOf(240), p.getPPr().getSpacing().getBefore());
    }

    @Test
    void preservesSurroundingWhitespace() {
        P p = Paragraphs.of("  spaced  ", TextStyle.body(), ParagraphStyle.body());
        R run = (R) p.getContent().get(0);
        assertEquals("preserve", textOf(run).getSpace());
        assertEquals("  spaced  ", textOf(run).getValue());
    }

    @Test
    void bodyStyleUsesElevenPointRegular() {
        P p = Paragraphs.of("Body", TextStyle.body(), ParagraphStyle.body());
        R run = (R) p.getContent().get(0);
        assertEquals(BigInteger.valueOf(22), run.getRPr().getSz().getVal(),
                "11 pt is 22 half-points");
        assertEquals("000000", run.getRPr().getColor().getVal());
    }

    @Test
    void rejectsBlankTextAndNullStyles() {
        assertThrows(DocumentGenerationException.class,
                () -> Paragraphs.of("   ", TextStyle.body(), ParagraphStyle.body()));
        assertThrows(DocumentGenerationException.class,
                () -> Paragraphs.of(null, TextStyle.body(), ParagraphStyle.body()));
        assertThrows(DocumentGenerationException.class,
                () -> Paragraphs.of("Text", null, ParagraphStyle.body()));
        assertThrows(DocumentGenerationException.class,
                () -> Paragraphs.of("Text", TextStyle.body(), null));
    }
}
```

- [ ] **Step 2: Run it to confirm it fails**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test -Dtest=ParagraphsTest`
Expected: FAIL — compilation error, `cannot find symbol: class Paragraphs`.

- [ ] **Step 3: Write `Paragraphs`**

Create `src/main/java/com/example/docx/content/Paragraphs.java`:

```java
package com.example.docx.content;

import com.example.docx.DocumentGenerationException;
import com.example.docx.style.ParagraphStyle;
import com.example.docx.style.TextStyle;
import org.docx4j.jaxb.Context;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Text;

/**
 * Builds paragraphs.
 *
 * <p>A heading and a body paragraph share this one factory: both are a {@code w:p}
 * holding a single {@code w:r}, differing only in the styles handed in. A pure function
 * of its arguments — it never touches the package, so it needs no fixtures.
 */
public final class Paragraphs {

    private Paragraphs() {
    }

    /** A {@code w:p} holding one styled {@code w:r} with the given text. */
    public static P of(String text, TextStyle textStyle, ParagraphStyle paragraphStyle) {
        if (text == null || text.isBlank()) {
            throw new DocumentGenerationException("paragraph text must not be blank");
        }
        if (textStyle == null) {
            throw new DocumentGenerationException("text style must not be null");
        }
        if (paragraphStyle == null) {
            throw new DocumentGenerationException("paragraph style must not be null");
        }

        ObjectFactory factory = Context.getWmlObjectFactory();

        Text value = factory.createText();
        value.setValue(text);
        // Without xml:space=preserve, leading and trailing spaces vanish silently.
        value.setSpace("preserve");

        R run = factory.createR();
        run.getContent().add(value);
        run.setRPr(textStyle.toRPr());

        P paragraph = factory.createP();
        paragraph.setPPr(paragraphStyle.toPPr());
        paragraph.getContent().add(run);
        return paragraph;
    }
}
```

- [ ] **Step 4: Point `WordDocument` at it and delete `Headings`**

In `src/main/java/com/example/docx/WordDocument.java`, replace the import

```java
import com.example.docx.content.Headings;
```

with

```java
import com.example.docx.content.Paragraphs;
import com.example.docx.style.ParagraphStyle;
```

and in `build()` replace

```java
                mainDocumentPart.getContent().add(Headings.heading(headingText, headingStyle));
```

with

```java
                mainDocumentPart.getContent().add(
                        Paragraphs.of(headingText, headingStyle, ParagraphStyle.heading()));
```

Then delete the old factory and its test:

```bash
git rm src/main/java/com/example/docx/content/Headings.java
git rm src/test/java/com/example/docx/content/HeadingsTest.java
```

- [ ] **Step 5: Run the suite**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B clean test`
Expected: `Tests run: 58, Failures: 0, Errors: 0, Skipped: 0`.

The count is unchanged from Task 2: `ParagraphsTest`'s four replace `HeadingsTest`'s four.

- [ ] **Step 6: Commit**

```bash
git add -A
git commit -m "refactor: merge Headings into Paragraphs"
```

---

### Task 4: Ordered content list

**Files:**
- Create: `src/main/java/com/example/docx/DocumentContent.java`
- Modify: `src/main/java/com/example/docx/WordDocument.java`
- Create: `src/test/java/com/example/docx/MultiSectionTest.java`
- Modify: `src/test/java/com/example/docx/WordDocumentRoundTripTest.java`

**Interfaces:**
- Consumes: `Paragraphs.of`, `TextStyle.defaults()`, `TextStyle.body()`, `ParagraphStyle.heading()`, `ParagraphStyle.body()`, `ImageParts.svgImage`, `PageSetup.usableWidthTwips()`.
- Produces: `DocumentContent` (functional interface, `P toParagraph(WordprocessingMLPackage pkg)`); `WordDocument.Builder` methods `heading(String)`, `heading(String, TextStyle)`, `paragraph(String)`, `paragraph(String, TextStyle)`, `svgImage(byte[], byte[])`, all returning `Builder` and all appending.

This is the change that fixes the reported bug: today `.heading("A").heading("B")` silently overwrites and yields one heading.

- [ ] **Step 1: Write the failing test**

Create `src/test/java/com/example/docx/MultiSectionTest.java`:

```java
package com.example.docx;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

import com.example.docx.page.PageSetup;
import com.example.docx.style.TextStyle;
import jakarta.xml.bind.JAXBElement;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.Body;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Text;
import org.junit.jupiter.api.Test;

class MultiSectionTest {

    private static byte[] resource(String name) throws IOException {
        try (InputStream in = MultiSectionTest.class.getResourceAsStream(name)) {
            assertNotNull(in, "missing test resource " + name);
            return in.readAllBytes();
        }
    }

    private static Body reload(byte[] bytes) throws Exception {
        return WordprocessingMLPackage.load(new ByteArrayInputStream(bytes))
                .getMainDocumentPart().getJaxbElement().getBody();
    }

    private static List<P> paragraphs(Body body) {
        List<P> out = new ArrayList<>();
        for (Object o : body.getContent()) {
            Object value = (o instanceof JAXBElement<?> je) ? je.getValue() : o;
            if (value instanceof P p) {
                out.add(p);
            }
        }
        return out;
    }

    private static String textOf(P paragraph) {
        StringBuilder sb = new StringBuilder();
        for (Object o : paragraph.getContent()) {
            Object value = (o instanceof JAXBElement<?> je) ? je.getValue() : o;
            if (value instanceof R run) {
                for (Object ro : run.getContent()) {
                    Object rv = (ro instanceof JAXBElement<?> je2) ? je2.getValue() : ro;
                    if (rv instanceof Text t) {
                        sb.append(t.getValue());
                    }
                }
            }
        }
        return sb.toString();
    }

    @Test
    void tenSectionsKeepTheirOrder() throws Exception {
        var builder = WordDocument.builder().pageSetup(PageSetup.a4());
        List<String> expected = new ArrayList<>();

        for (int i = 1; i <= 10; i++) {
            builder.heading("Section " + i, TextStyle.builder().sizePt(11).bold(true).build());
            expected.add("Section " + i);
            for (int j = 1; j <= 2; j++) {
                builder.paragraph("Body " + i + "." + j);
                expected.add("Body " + i + "." + j);
            }
        }

        List<P> ps = paragraphs(reload(builder.build().toByteArray()));
        assertEquals(30, ps.size(), "10 headings + 20 paragraphs");
        // Order is what this change can break; per-paragraph assertions would not catch it.
        assertEquals(expected, ps.stream().map(MultiSectionTest::textOf).toList());
    }

    @Test
    void twoHeadingCallsProduceTwoHeadings() throws Exception {
        byte[] bytes = WordDocument.builder()
                .heading("First")
                .heading("Second")
                .build()
                .toByteArray();

        List<P> ps = paragraphs(reload(bytes));
        assertEquals(2, ps.size(), "the second heading must not replace the first");
        assertEquals("First", textOf(ps.get(0)));
        assertEquals("Second", textOf(ps.get(1)));
    }

    @Test
    void headingsKeepWithNextAndBodyDoesNot() throws Exception {
        byte[] bytes = WordDocument.builder()
                .heading("H")
                .paragraph("B")
                .build()
                .toByteArray();

        List<P> ps = paragraphs(reload(bytes));
        assertNotNull(ps.get(0).getPPr().getKeepNext(), "heading must set w:keepNext");
        assertNull(ps.get(1).getPPr().getKeepNext(), "body must not set w:keepNext");
        assertEquals(BigInteger.valueOf(240), ps.get(0).getPPr().getSpacing().getBefore());
        assertEquals(BigInteger.valueOf(120), ps.get(1).getPPr().getSpacing().getAfter());
    }

    @Test
    void imagesAndTextInterleaveInInsertionOrder() throws Exception {
        byte[] svg = resource("/demo/chart.svg");
        byte[] png = resource("/demo/chart.png");

        byte[] bytes = WordDocument.builder()
                .heading("Before")
                .svgImage(svg, png)
                .paragraph("After")
                .build()
                .toByteArray();

        List<P> ps = paragraphs(reload(bytes));
        assertEquals(3, ps.size());
        // Text entries are built eagerly and images deferred; order must survive that.
        assertEquals("Before", textOf(ps.get(0)));
        assertEquals("", textOf(ps.get(1)), "the middle paragraph holds the image");
        assertEquals("After", textOf(ps.get(2)));
    }

    @Test
    void requiresAtLeastOneContentItem() {
        assertThrows(DocumentGenerationException.class, () -> WordDocument.builder().build());
    }
}
```

- [ ] **Step 2: Run it to confirm it fails**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test -Dtest=MultiSectionTest`
Expected: FAIL — compilation error, `cannot find symbol: method paragraph(java.lang.String)`.

- [ ] **Step 3: Create `DocumentContent`**

Create `src/main/java/com/example/docx/DocumentContent.java`:

```java
package com.example.docx;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.P;

/**
 * One ordered item of document body content.
 *
 * <p>Deferred rather than a finished {@code P}, because an image is a package part plus
 * a relationship and no package exists until {@code build()}. Text implementations
 * ignore the argument.
 *
 * <p>It lives here rather than in {@code content} because it names
 * {@code WordprocessingMLPackage}, which that package is kept free of.
 */
@FunctionalInterface
public interface DocumentContent {
    P toParagraph(WordprocessingMLPackage pkg);
}
```

- [ ] **Step 4: Replace the builder's fields with the list**

In `src/main/java/com/example/docx/WordDocument.java`, add these imports:

```java
import java.util.ArrayList;
import java.util.List;
import org.docx4j.wml.P;
```

Replace the four content fields

```java
        private String headingText;
        private TextStyle headingStyle = TextStyle.defaults();
        private byte[] svgBytes;
        private byte[] pngBytes;
```

with

```java
        private final List<DocumentContent> content = new ArrayList<>();
```

(keep `private PageSetup pageSetup = PageSetup.a4();` as it is).

Replace the existing `heading(String, TextStyle)`, `heading(String)` and `svgImage(byte[], byte[])` methods entirely with:

```java
        /** Appends a heading using {@link TextStyle#defaults()}. */
        public Builder heading(String text) {
            return heading(text, TextStyle.defaults());
        }

        /** Appends a heading. Each call adds one; calls do not replace each other. */
        public Builder heading(String text, TextStyle style) {
            P paragraph = Paragraphs.of(text, style, ParagraphStyle.heading());
            content.add(pkg -> paragraph);
            return this;
        }

        /** Appends a body paragraph using {@link TextStyle#body()}. */
        public Builder paragraph(String text) {
            return paragraph(text, TextStyle.body());
        }

        /** Appends a body paragraph. */
        public Builder paragraph(String text, TextStyle style) {
            P paragraph = Paragraphs.of(text, style, ParagraphStyle.body());
            content.add(pkg -> paragraph);
            return this;
        }

        /**
         * Appends an SVG image with a PNG fallback, drawn at half the usable page width.
         *
         * <p>The SVG must be SVG 1.1: Word's renderer rejects SVG 2 features such as
         * {@code height="auto"} and 8-digit hex colours that browsers accept.
         */
        public Builder svgImage(byte[] svg, byte[] pngFallback) {
            content.add(pkg -> ImageParts.svgImage(
                    pkg, svg, pngFallback, pageSetup.usableWidthTwips() / 2));
            return this;
        }
```

Text is validated here, at the call site, because `Paragraphs.of` runs eagerly — a blank string fails at the offending line rather than at `build()`. The image lambda reads `pageSetup` when invoked, so setting the page after adding an image still sizes it correctly.

- [ ] **Step 5: Replace the guard and the append in `build()`**

Still in `WordDocument.java`, replace

```java
            if (headingText == null || headingText.isBlank()) {
                throw new DocumentGenerationException(
                        "a heading is required; call heading(String) before build()");
            }
```

with

```java
            if (content.isEmpty()) {
                throw new DocumentGenerationException(
                        "a document needs at least one heading, paragraph or image");
            }
```

and replace the body-appending block

```java
                mainDocumentPart.getContent().add(
                        Paragraphs.of(headingText, headingStyle, ParagraphStyle.heading()));

                if (svgBytes != null || pngBytes != null) {
                    mainDocumentPart.getContent().add(ImageParts.svgImage(
                            pkg, svgBytes, pngBytes, pageSetup.usableWidthTwips() / 2));
                }
```

with

```java
                for (DocumentContent item : content) {
                    mainDocumentPart.getContent().add(item.toParagraph(pkg));
                }
```

- [ ] **Step 6: Drop the superseded test**

`WordDocumentRoundTripTest.requiresAHeading` is now covered by
`MultiSectionTest.requiresAtLeastOneContentItem`, and keeping both is duplication.
Delete this whole method from `src/test/java/com/example/docx/WordDocumentRoundTripTest.java`:

```java
    @Test
    void requiresAHeading() {
        assertThrows(DocumentGenerationException.class, () -> WordDocument.builder().build());
    }
```

If `assertThrows` and `DocumentGenerationException` become unused in that file, remove their imports so the build stays warning-free.

- [ ] **Step 7: Run the suite**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B clean test`
Expected: `Tests run: 62, Failures: 0, Errors: 0, Skipped: 0`.

- [ ] **Step 8: Commit**

```bash
git add -A
git commit -m "feat: append content in order so headings and paragraphs repeat"
```

---

### Task 5: Sample and documentation

**Files:**
- Modify: `src/test/java/com/example/docx/sample/SampleMain.java`
- Modify: `README.md`

**Interfaces:**
- Consumes: `WordDocument.Builder.heading/paragraph/svgImage`.
- Produces: nothing other tasks depend on. Terminal task.

- [ ] **Step 1: Show several sections in the sample**

In `src/test/java/com/example/docx/sample/SampleMain.java`, replace the `WordDocument document = ...` statement with:

```java
        TextStyle sectionHeading = TextStyle.builder()
                .font("Calibri")
                .sizePt(11)
                .bold(true)
                .color("#1F4E79")
                .build();

        var builder = WordDocument.builder()
                .pageSetup(PageSetup.builder().a4().marginsTwips(851).build())
                .heading("Quarterly Report", TextStyle.builder()
                        .font("Calibri Light")
                        .sizePt(20)
                        .bold(true)
                        .italic(false)
                        .color("#1F4E79")
                        .build())
                .svgImage(resource("/demo/chart.svg"), resource("/demo/chart.png"));

        for (int section = 1; section <= 3; section++) {
            builder.heading("Section " + section, sectionHeading);
            builder.paragraph("First paragraph of section " + section + ".");
            builder.paragraph("Second paragraph of section " + section + ".");
        }

        WordDocument document = builder.build();
```

Add `import com.example.docx.style.TextStyle;` if the sed in Task 1 did not already leave it (it renames the existing `HeadingStyle` import in place, so it should be present).

- [ ] **Step 2: Run the sample**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test-compile exec:java`
Expected: `BUILD SUCCESS` and `Wrote /…/target/sample.docx`.

- [ ] **Step 3: Verify the sections actually landed**

Run:
```bash
unzip -p target/sample.docx word/document.xml | grep -c "<w:p>"
```
Expected: `11` — one report title, one image paragraph, then 3 sections × (1 heading + 2 paragraphs).

Run:
```bash
unzip -p target/sample.docx word/document.xml | grep -c "w:keepNext"
```
Expected: `4` — the title plus three section headings.

- [ ] **Step 4: Note that a builder is single-use for text**

`heading()` and `paragraph()` build their `P` eagerly and the stored lambda returns that
same instance every time, so two `build()` calls on one `Builder` would hand both
documents the same JAXB nodes. The class is already documented "build one per document",
but make that explicit for the builder too. In
`src/main/java/com/example/docx/WordDocument.java`, extend the `Builder` class Javadoc:

```java
    /**
     * Fluent builder. All docx4j work happens in {@link #build()}.
     *
     * <p>Content calls append; none replaces a previous one. Text paragraphs are built
     * eagerly, so a builder is single-use: calling {@code build()} twice would give both
     * documents the same paragraph objects. Build one document per builder.
     */
```

- [ ] **Step 5: Document it**

In `README.md`, replace the "Using it" example with this. Note `svgImage` is a builder
method, so every content call happens before `build()`:

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

Update the Layout table rows for `…​.style` and `…​.content` to:

```
| `…​.style` | `TextStyle` — run formatting; `ParagraphStyle` — spacing and keep-with-next |
| `…​.content` | `Paragraphs` — stateless paragraph factory, used for headings and body alike |
```

Add to "Things that bite":

```
- **Content is appended, in order.** `heading`, `paragraph` and `svgImage` each add
  one item; none of them replaces a previous call. Build a document by looping.
- **Headings keep with the next paragraph.** `ParagraphStyle.heading()` sets
  `w:keepNext`, so a heading never strands at the foot of a page with its body
  overleaf. Body paragraphs deliberately do not set it.
- **Points, not pixels.** An 11 pt heading is `sizePt(11)`, emitting `w:sz` 22.
  11 px would be 8.25 pt and noticeably smaller.
```

- [ ] **Step 6: Run the full suite**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B clean test`
Expected: `Tests run: 62, Failures: 0, Errors: 0, Skipped: 0` and `BUILD SUCCESS`.

- [ ] **Step 7: Commit**

```bash
git add -A
git commit -m "feat: demo multiple sections in the sample and document the API"
```

---

## Definition of Done

- `mvn clean test` passes with 62 tests.
- `.heading("A").heading("B")` produces two headings, not one.
- A ten-section document reloads with 30 paragraphs whose texts appear in insertion order.
- Heading paragraphs carry `w:keepNext`; body paragraphs do not.
- `grep -rn "HeadingStyle" src README.md` returns nothing.
- `content/Headings.java` no longer exists.
- `mvn test-compile exec:java` writes a `target/sample.docx` containing 11 paragraphs and 4 `w:keepNext` elements.
- No class in `content/` or `style/` references `WordprocessingMLPackage`.
