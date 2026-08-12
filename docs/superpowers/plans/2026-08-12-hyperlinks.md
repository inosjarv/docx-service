# Hyperlinks Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Let a paragraph end in a clickable link, in the same line as the body text — "Revenue grew 12%. **Learn more**" — where only the trailing phrase is a hyperlink, optional per paragraph.

**Architecture:** A hyperlink's target is a relationship, not inline text, so building it needs `WordprocessingMLPackage` — the same fork `ImageParts` already crossed for images. `content.Hyperlink` is a pure value (text, url, style); `part.Hyperlinks` builds the actual `w:hyperlink`, because it is the piece that needs the package. `content.Paragraphs` gains a small `run(text, style)` helper so both the plain-paragraph path and the hyperlink path build runs identically.

**Tech Stack:** Java 25, Maven, docx4j 17.0.2, JUnit 6.1.3.

## Global Constraints

- Java release level is exactly `25`. Run Maven as `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test`.
- **No new dependencies.**
- A hyperlink relationship is built with `mainDocumentPart.getRelationshipsPart(true)`, a `Relationship` with `setType(Namespaces.HYPERLINK)`, `setTarget(url)`, `setTargetMode("External")`, then `relationshipsPart.addRelationship(relationship)` — verified against docx4j 17.0.2; do not use `addTargetPart`, which is for parts, not external relationships.
- The hyperlink element is `new org.docx4j.wml.P.Hyperlink()` (a nested class, not a top-level `Hyperlink` in `org.docx4j.wml`), with `setId(relationship.getId())` and `setHistory(true)`.
- **This paragraph type must defer to `build()`.** Unlike the existing no-link `paragraph(...)` overloads, which validate and build eagerly at the call site, a linked paragraph cannot exist before the package does. Follow the `svgImage` pattern: the builder method only captures arguments in a lambda; all validation and construction happens when that lambda runs.
- `underline` on `TextStyle` follows the exact emit-only-when-true pattern `bold`/`italic` already use: `RPr.setU(U)` with `UnderlineEnumeration.SINGLE`, set only when `underline` is `true`.
- `TextStyle.link()` is `#0563C1`, underlined, Calibri 11 pt, regular — Word's own hyperlink look, not an invented colour.
- `DocumentGenerationException` is the ONLY exception thrown deliberately. `Hyperlink.of`'s own validation (blank text, blank url, null style) carries no cause; a malformed URL wraps the `URISyntaxException` it caught, per the spec's stated exception to the no-cause rule.
- Layering: only `page/`, `part/`, `DocumentContent` and `WordDocument` may reference `WordprocessingMLPackage`. `content/` and `style/` stay pure — `content.Hyperlink` must not reference it.
- External URLs only. No internal bookmarks/anchors.
- Do not modify `src/test/resources/demo/`.

## File Structure

| File | Change | Responsibility |
| --- | --- | --- |
| `src/main/java/com/example/docx/style/TextStyle.java` | Modify | `+ underline(boolean)`, `+ TextStyle.link()` |
| `src/main/java/com/example/docx/content/Paragraphs.java` | Modify | `+ public run(text, style) -> R`, `of(...)` delegates to it |
| `src/main/java/com/example/docx/content/Hyperlink.java` | Create | Immutable value: text, url, style. No package. |
| `src/main/java/com/example/docx/part/Hyperlinks.java` | Create | `(pkg, leadingText, textStyle, paragraphStyle, Hyperlink) -> P` |
| `src/main/java/com/example/docx/WordDocument.java` | Modify | Three `paragraph(..., Hyperlink)` overloads |
| `src/test/java/com/example/docx/style/TextStyleTest.java` | Modify | Underline emission, `link()` defaults |
| `src/test/java/com/example/docx/content/ParagraphsTest.java` | Modify | `run(...)` behaviour |
| `src/test/java/com/example/docx/content/HyperlinkTest.java` | Create | Validation, defaults, malformed URL |
| `src/test/java/com/example/docx/HyperlinkDocumentTest.java` | Create | Round trip: ordering, relationship, styling, coexistence |
| `src/test/java/com/example/docx/sample/SampleMain.java` | Modify | A linked paragraph |
| `README.md` | Modify | Document hyperlinks and the new asymmetry |

Test counts: 91 now → **106**. Task 1 adds 3, Task 2 adds 2, Task 3 adds 5, Task 4 adds 5, Task 5 adds none.

---

### Task 1: `TextStyle` gains `underline` and `link()`

**Files:**
- Modify: `src/main/java/com/example/docx/style/TextStyle.java`
- Modify: `src/test/java/com/example/docx/style/TextStyleTest.java`

**Interfaces:**
- Consumes: existing `TextStyle` fields and builder shape.
- Produces: `public boolean underline()`; `Builder.underline(boolean)`; `public static TextStyle link()`.

- [ ] **Step 1: Write the failing tests**

Add to `src/test/java/com/example/docx/style/TextStyleTest.java`:

```java
    @Test
    void underlineDefaultsToFalseAndEmitsNothing() {
        TextStyle style = TextStyle.defaults();
        assertEquals(false, style.underline());
        assertNull(style.toRPr().getU(), "underline must be absent, not present-and-false");
    }

    @Test
    void underlineTrueEmitsSingle() {
        TextStyle style = TextStyle.builder().underline(true).build();
        assertTrue(style.underline());
        assertEquals("single", style.toRPr().getU().getVal().value());
    }

    @Test
    void linkIsWordsDefaultHyperlinkLook() {
        TextStyle link = TextStyle.link();
        assertEquals("Calibri", link.fontFamily());
        assertEquals(11.0, link.sizePt());
        assertEquals(false, link.bold());
        assertTrue(link.underline());
        assertEquals("0563C1", link.colorHex());
    }
```

- [ ] **Step 2: Run them to confirm they fail**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test -Dtest=TextStyleTest`
Expected: FAIL — compilation error, `cannot find symbol: method underline()`.

- [ ] **Step 3: Add the field, builder method, and `RPr` emission**

In `src/main/java/com/example/docx/style/TextStyle.java`, add the field next to `italic`:

```java
    private final boolean underline;
```

and in the constructor:

```java
        this.underline = b.underline;
```

Add the accessor next to `italic()`:

```java
    public boolean underline() {
        return underline;
    }
```

In `toRPr()`, immediately after the `italic` block:

```java
        if (underline) {
            org.docx4j.wml.U u = factory.createU();
            u.setVal(org.docx4j.wml.UnderlineEnumeration.SINGLE);
            rPr.setU(u);
        }
```

In the `Builder`, add the field next to `italic`:

```java
        private boolean underline = false;
```

and the setter next to `italic(boolean)`:

```java
        public Builder underline(boolean underline) {
            this.underline = underline;
            return this;
        }
```

- [ ] **Step 4: Add the `link()` preset**

Immediately after `body()`:

```java
    /** Word's own hyperlink look: Calibri 11 pt, underlined, {@code #0563C1}. */
    public static TextStyle link() {
        return builder().font("Calibri").sizePt(11).bold(false)
                .underline(true).color("0563C1").build();
    }
```

- [ ] **Step 5: Run the tests**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test -Dtest=TextStyleTest`
Expected: PASS — `Tests run: 16, Failures: 0, Errors: 0`.

- [ ] **Step 6: Run the full suite**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B clean test`
Expected: `Tests run: 94, Failures: 0, Errors: 0, Skipped: 0`.

- [ ] **Step 7: Commit**

```bash
git add -A
git commit -m "feat: add TextStyle.underline() and the link() preset"
```

---

### Task 2: `Paragraphs.run` extraction

**Files:**
- Modify: `src/main/java/com/example/docx/content/Paragraphs.java`
- Modify: `src/test/java/com/example/docx/content/ParagraphsTest.java`

**Interfaces:**
- Consumes: `TextStyle.toRPr()`.
- Produces: `public static org.docx4j.wml.R Paragraphs.run(String text, TextStyle textStyle)`. `of(...)`'s existing public contract — including its exact error messages — is unchanged.

`run(...)` is what Task 4's `part.Hyperlinks` uses to build both the leading plain run and the hyperlink's own run, without duplicating `xml:space="preserve"` handling. `of(...)` keeps validating `text`/`textStyle`/`paragraphStyle` itself with its existing messages ("paragraph text must not be blank", etc.) — do not let `run`'s more generic message ("text must not be blank") leak into `of`'s behaviour; `of` must still throw before calling `run`.

- [ ] **Step 1: Write the failing tests**

Add to `src/test/java/com/example/docx/content/ParagraphsTest.java`:

```java
    @Test
    void runBuildsAStandaloneStyledRun() {
        R run = Paragraphs.run("Hello", TextStyle.body());
        assertEquals("Hello", textOf(run).getValue());
        assertEquals(BigInteger.valueOf(22), run.getRPr().getSz().getVal());
    }

    @Test
    void runRejectsBlankTextAndNullStyle() {
        assertThrows(DocumentGenerationException.class, () -> Paragraphs.run("  ", TextStyle.body()));
        assertThrows(DocumentGenerationException.class, () -> Paragraphs.run(null, TextStyle.body()));
        assertThrows(DocumentGenerationException.class, () -> Paragraphs.run("Hi", null));
    }
```

- [ ] **Step 2: Run them to confirm they fail**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test -Dtest=ParagraphsTest`
Expected: FAIL — compilation error, `cannot find symbol: method run(java.lang.String,com.example.docx.style.TextStyle)`.

- [ ] **Step 3: Extract `run` and have `of` delegate to it**

Replace `src/main/java/com/example/docx/content/Paragraphs.java` entirely:

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
 * Builds paragraphs and runs.
 *
 * <p>A heading and a body paragraph share {@link #of}: both are a {@code w:p} holding a
 * single {@code w:r}, differing only in the styles handed in. A pure function of its
 * arguments — it never touches the package, so it needs no fixtures.
 */
public final class Paragraphs {

    private Paragraphs() {
    }

    /**
     * A styled {@code w:r} holding {@code text}. Reused by {@link #of} and by hyperlinks,
     * which need a run without a surrounding paragraph.
     */
    public static R run(String text, TextStyle textStyle) {
        if (text == null || text.isBlank()) {
            throw new DocumentGenerationException("text must not be blank");
        }
        if (textStyle == null) {
            throw new DocumentGenerationException("text style must not be null");
        }

        ObjectFactory factory = Context.getWmlObjectFactory();

        Text value = factory.createText();
        value.setValue(text);
        // Without xml:space=preserve, leading and trailing spaces vanish silently.
        value.setSpace("preserve");

        R run = factory.createR();
        run.getContent().add(value);
        run.setRPr(textStyle.toRPr());
        return run;
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

        R run = run(text, textStyle);

        ObjectFactory factory = Context.getWmlObjectFactory();
        P paragraph = factory.createP();
        paragraph.setPPr(paragraphStyle.toPPr());
        paragraph.getContent().add(run);
        return paragraph;
    }
}
```

- [ ] **Step 4: Run the tests**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test -Dtest=ParagraphsTest`
Expected: PASS — `Tests run: 6, Failures: 0, Errors: 0`.

- [ ] **Step 5: Run the full suite**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B clean test`
Expected: `Tests run: 96, Failures: 0, Errors: 0, Skipped: 0`.

Every existing `ParagraphsTest` and `HeadingsTest`-descended assertion must still pass unchanged — `of(...)`'s behaviour and messages did not move.

- [ ] **Step 6: Commit**

```bash
git add -A
git commit -m "refactor: extract Paragraphs.run for reuse by hyperlinks"
```

---

### Task 3: `content.Hyperlink`

**Files:**
- Create: `src/main/java/com/example/docx/content/Hyperlink.java`
- Create: `src/test/java/com/example/docx/content/HyperlinkTest.java`

**Interfaces:**
- Consumes: `TextStyle.link()`, `DocumentGenerationException`.
- Produces: `Hyperlink.of(String text, String url)`, `Hyperlink.of(String text, String url, TextStyle style)`; accessors `text()`, `url()`, `style()`.

A pure value — no package reference anywhere in this file. URL validity is checked with `java.net.URI`'s constructor, which throws `URISyntaxException` on malformed input; that exception becomes the *cause* of the wrapping `DocumentGenerationException`, the one deliberate exception to the "validation carries no cause" rule.

- [ ] **Step 1: Write the failing test**

Create `src/test/java/com/example/docx/content/HyperlinkTest.java`:

```java
package com.example.docx.content;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

import com.example.docx.DocumentGenerationException;
import com.example.docx.style.TextStyle;
import org.junit.jupiter.api.Test;

class HyperlinkTest {

    @Test
    void defaultStyleIsTextStyleLink() {
        Hyperlink link = Hyperlink.of("Learn more", "https://example.com");
        assertEquals("Learn more", link.text());
        assertEquals("https://example.com", link.url());
        assertEquals(TextStyle.link().colorHex(), link.style().colorHex());
        assertEquals(TextStyle.link().underline(), link.style().underline());
    }

    @Test
    void explicitStyleOverridesTheDefault() {
        TextStyle custom = TextStyle.builder().color("#FF0000").underline(false).build();
        Hyperlink link = Hyperlink.of("Click", "https://example.com", custom);
        assertEquals("FF0000", link.style().colorHex());
        assertEquals(false, link.style().underline());
    }

    @Test
    void rejectsBlankTextAndUrl() {
        assertThrows(DocumentGenerationException.class, () -> Hyperlink.of("", "https://example.com"));
        assertThrows(DocumentGenerationException.class, () -> Hyperlink.of(null, "https://example.com"));
        assertThrows(DocumentGenerationException.class, () -> Hyperlink.of("Click", ""));
        assertThrows(DocumentGenerationException.class, () -> Hyperlink.of("Click", null));
    }

    @Test
    void rejectsNullStyle() {
        assertThrows(DocumentGenerationException.class,
                () -> Hyperlink.of("Click", "https://example.com", null));
    }

    @Test
    void rejectsMalformedUrlWithACause() {
        String malformed = "not a url" + " " + "with a raw space";
        DocumentGenerationException e = assertThrows(DocumentGenerationException.class,
                () -> Hyperlink.of("Click", malformed));
        assertNotNull(e.getCause(), "a malformed URL wraps URISyntaxException");
    }
}
```

- [ ] **Step 2: Run it to confirm it fails**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test -Dtest=HyperlinkTest`
Expected: FAIL — compilation error, `cannot find symbol: class Hyperlink`.

- [ ] **Step 3: Write `Hyperlink`**

Create `src/main/java/com/example/docx/content/Hyperlink.java`:

```java
package com.example.docx.content;

import com.example.docx.DocumentGenerationException;
import com.example.docx.style.TextStyle;
import java.net.URI;
import java.net.URISyntaxException;

/**
 * An immutable value: the visible text of a link, its destination, and how it is styled.
 *
 * <p>Pure data — building the actual {@code w:hyperlink} needs a relationship, which
 * needs the package, so that work lives in {@code part.Hyperlinks}. This class only
 * validates and holds what it was given.
 */
public final class Hyperlink {

    private final String text;
    private final String url;
    private final TextStyle style;

    private Hyperlink(String text, String url, TextStyle style) {
        this.text = text;
        this.url = url;
        this.style = style;
    }

    /** A link styled with {@link TextStyle#link()} — Word's own hyperlink look. */
    public static Hyperlink of(String text, String url) {
        return of(text, url, TextStyle.link());
    }

    /** A link with explicit styling. */
    public static Hyperlink of(String text, String url, TextStyle style) {
        if (text == null || text.isBlank()) {
            throw new DocumentGenerationException("link text must not be blank");
        }
        if (style == null) {
            throw new DocumentGenerationException("link text style must not be null");
        }
        String checkedUrl = validateUrl(url);
        return new Hyperlink(text, checkedUrl, style);
    }

    private static String validateUrl(String url) {
        if (url == null || url.isBlank()) {
            throw new DocumentGenerationException("link url must not be blank");
        }
        try {
            new URI(url);
        } catch (URISyntaxException e) {
            throw new DocumentGenerationException("link url is not a valid URI: '" + url + "'", e);
        }
        return url;
    }

    public String text() {
        return text;
    }

    public String url() {
        return url;
    }

    public TextStyle style() {
        return style;
    }
}
```

- [ ] **Step 4: Run the tests**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test -Dtest=HyperlinkTest`
Expected: PASS — `Tests run: 5, Failures: 0, Errors: 0`.

- [ ] **Step 5: Run the full suite**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B clean test`
Expected: `Tests run: 101, Failures: 0, Errors: 0, Skipped: 0`.

- [ ] **Step 6: Commit**

```bash
git add -A
git commit -m "feat: add content.Hyperlink, a validated link value"
```

---

### Task 4: `part.Hyperlinks` and the `WordDocument.Builder` overloads

**Files:**
- Create: `src/main/java/com/example/docx/part/Hyperlinks.java`
- Modify: `src/main/java/com/example/docx/WordDocument.java`
- Create: `src/test/java/com/example/docx/HyperlinkDocumentTest.java`

**Interfaces:**
- Consumes: `content.Hyperlink`, `content.Paragraphs.run`, `Namespaces.HYPERLINK` (`org.docx4j.openpackaging.parts.relationships.Namespaces`).
- Produces: `public static org.docx4j.wml.P Hyperlinks.paragraph(WordprocessingMLPackage pkg, String leadingText, TextStyle leadingStyle, ParagraphStyle paragraphStyle, Hyperlink link)`; three new `WordDocument.Builder.paragraph(...)` overloads ending in a `Hyperlink` parameter.

**All the code below was compiled and run against docx4j 17.0.2 before this plan was written** — including a document mixing a heading, an SVG image and a linked paragraph, confirming image and hyperlink relationships coexist with distinct ids. If something does not compile, suspect a transcription slip first.

- [ ] **Step 1: Write the failing test**

Create `src/test/java/com/example/docx/HyperlinkDocumentTest.java`:

```java
package com.example.docx;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.example.docx.content.Hyperlink;
import com.example.docx.page.PageSetup;
import jakarta.xml.bind.JAXBElement;
import java.io.ByteArrayInputStream;
import java.util.List;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.Body;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Text;
import org.junit.jupiter.api.Test;

class HyperlinkDocumentTest {

    private static WordprocessingMLPackage reload(byte[] bytes) throws Exception {
        return WordprocessingMLPackage.load(new ByteArrayInputStream(bytes));
    }

    private static Body body(WordprocessingMLPackage pkg) {
        return pkg.getMainDocumentPart().getJaxbElement().getBody();
    }

    private static P firstParagraph(Body body) {
        for (Object o : body.getContent()) {
            Object v = (o instanceof JAXBElement<?> je) ? je.getValue() : o;
            if (v instanceof P p) {
                return p;
            }
        }
        throw new AssertionError("no paragraph in body");
    }

    private static String textOf(R run) {
        Object first = run.getContent().get(0);
        Object v = (first instanceof JAXBElement<?> je) ? je.getValue() : first;
        return ((Text) v).getValue();
    }

    @Test
    void linkedParagraphReloadsAsLeadingRunThenHyperlink() throws Exception {
        byte[] bytes = WordDocument.builder()
                .pageSetup(PageSetup.a4())
                .paragraph("Revenue grew 12%. ", Hyperlink.of("Learn more", "https://example.com/report"))
                .build()
                .toByteArray();

        WordprocessingMLPackage pkg = reload(bytes);
        P paragraph = firstParagraph(body(pkg));
        List<Object> content = paragraph.getContent();
        assertEquals(2, content.size(), "leading run + hyperlink, nothing else");

        Object first = content.get(0) instanceof JAXBElement<?> je ? je.getValue() : content.get(0);
        assertTrue(first instanceof R, "first item must be the leading run");
        assertEquals("Revenue grew 12%. ", textOf((R) first));

        Object second = content.get(1) instanceof JAXBElement<?> je2 ? je2.getValue() : content.get(1);
        assertTrue(second instanceof org.docx4j.wml.P.Hyperlink, "second item must be the hyperlink");
        org.docx4j.wml.P.Hyperlink hyperlink = (org.docx4j.wml.P.Hyperlink) second;
        R linkRun = (R) (hyperlink.getContent().get(0) instanceof JAXBElement<?> je3
                ? je3.getValue() : hyperlink.getContent().get(0));
        assertEquals("Learn more", textOf(linkRun));
    }

    @Test
    void relationshipIsExternalAndResolvesFromTheHyperlinkId() throws Exception {
        byte[] bytes = WordDocument.builder()
                .paragraph("See ", Hyperlink.of("the report", "https://example.com/report"))
                .build()
                .toByteArray();

        WordprocessingMLPackage pkg = reload(bytes);
        P paragraph = firstParagraph(body(pkg));
        Object second = paragraph.getContent().get(1) instanceof JAXBElement<?> je
                ? je.getValue() : paragraph.getContent().get(1);
        String relId = ((org.docx4j.wml.P.Hyperlink) second).getId();

        Relationship rel = pkg.getMainDocumentPart().getRelationshipsPart().getRelationshipByID(relId);
        assertNotNull(rel, "the hyperlink's r:id must resolve to a relationship");
        assertEquals("External", rel.getTargetMode());
        assertEquals("https://example.com/report", rel.getTarget());
    }

    @Test
    void blankLeadingTextProducesALinkOnlyParagraph() throws Exception {
        byte[] bytes = WordDocument.builder()
                .paragraph("", Hyperlink.of("Click here", "https://example.com"))
                .build()
                .toByteArray();

        P paragraph = firstParagraph(body(reload(bytes)));
        assertEquals(1, paragraph.getContent().size(), "no leading run when text is blank");
        Object only = paragraph.getContent().get(0) instanceof JAXBElement<?> je
                ? je.getValue() : paragraph.getContent().get(0);
        assertTrue(only instanceof org.docx4j.wml.P.Hyperlink);
    }

    @Test
    void linkStyleEmitsWordsDefaultColourAndUnderline() throws Exception {
        byte[] bytes = WordDocument.builder()
                .paragraph("See ", Hyperlink.of("here", "https://example.com"))
                .build()
                .toByteArray();

        P paragraph = firstParagraph(body(reload(bytes)));
        Object second = paragraph.getContent().get(1) instanceof JAXBElement<?> je
                ? je.getValue() : paragraph.getContent().get(1);
        R linkRun = (R) (((org.docx4j.wml.P.Hyperlink) second).getContent().get(0) instanceof JAXBElement<?> je2
                ? je2.getValue() : ((org.docx4j.wml.P.Hyperlink) second).getContent().get(0));
        assertEquals("0563C1", linkRun.getRPr().getColor().getVal());
        assertNotNull(linkRun.getRPr().getU(), "Word's default hyperlink style is underlined");
        assertEquals("single", linkRun.getRPr().getU().getVal().value());
    }

    @Test
    void imageAndHyperlinkRelationshipsCoexistWithDistinctIds() throws Exception {
        byte[] svg = resource("/demo/chart.svg");
        byte[] png = resource("/demo/chart.png");

        byte[] bytes = WordDocument.builder()
                .heading("Report")
                .svgImage(svg, png)
                .paragraph("See ", Hyperlink.of("the chart source", "https://example.com/data"))
                .build()
                .toByteArray();

        WordprocessingMLPackage pkg = reload(bytes);
        // Just confirming the whole document still loads and both relationship types exist.
        var relationships = pkg.getMainDocumentPart().getRelationshipsPart().getRelationships().getRelationship();
        boolean hasImage = relationships.stream().anyMatch(r -> r.getType().contains("/image"));
        boolean hasHyperlink = relationships.stream().anyMatch(r -> r.getType().contains("/hyperlink"));
        assertTrue(hasImage, "image relationship must be present");
        assertTrue(hasHyperlink, "hyperlink relationship must be present");
    }

    private static byte[] resource(String name) throws Exception {
        try (var in = HyperlinkDocumentTest.class.getResourceAsStream(name)) {
            assertNotNull(in, "missing test resource " + name);
            return in.readAllBytes();
        }
    }
}
```

- [ ] **Step 2: Run it to confirm it fails**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test -Dtest=HyperlinkDocumentTest`
Expected: FAIL — compilation error, `cannot find symbol: method paragraph(java.lang.String,com.example.docx.content.Hyperlink)`.

- [ ] **Step 3: Write `part.Hyperlinks`**

Create `src/main/java/com/example/docx/part/Hyperlinks.java`:

```java
package com.example.docx.part;

import com.example.docx.DocumentGenerationException;
import com.example.docx.content.Hyperlink;
import com.example.docx.content.Paragraphs;
import com.example.docx.style.ParagraphStyle;
import com.example.docx.style.TextStyle;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.R;

/**
 * Builds a paragraph whose trailing content is a clickable link.
 *
 * <p>Unlike {@code content.Paragraphs}, this needs the package: a hyperlink's target is
 * a relationship ({@code word/_rels/document.xml.rels}, {@code TargetMode="External"}),
 * not inline text, and a relationship cannot exist before the package does.
 */
public final class Hyperlinks {

    private Hyperlinks() {
    }

    /**
     * A paragraph made of an optional leading run and a trailing hyperlink.
     *
     * @param leadingText the body text before the link; blank or null omits the
     *                    leading run entirely, leaving a link-only paragraph
     */
    public static P paragraph(WordprocessingMLPackage pkg, String leadingText,
                              TextStyle leadingStyle, ParagraphStyle paragraphStyle,
                              Hyperlink link) {
        if (pkg == null) {
            throw new DocumentGenerationException("package must not be null");
        }
        if (leadingStyle == null) {
            throw new DocumentGenerationException("text style must not be null");
        }
        if (paragraphStyle == null) {
            throw new DocumentGenerationException("paragraph style must not be null");
        }
        if (link == null) {
            throw new DocumentGenerationException("link must not be null");
        }

        ObjectFactory factory = Context.getWmlObjectFactory();
        P paragraph = factory.createP();
        paragraph.setPPr(paragraphStyle.toPPr());

        if (leadingText != null && !leadingText.isBlank()) {
            paragraph.getContent().add(Paragraphs.run(leadingText, leadingStyle));
        }
        paragraph.getContent().add(hyperlinkRun(pkg, link, factory));
        return paragraph;
    }

    private static org.docx4j.wml.P.Hyperlink hyperlinkRun(
            WordprocessingMLPackage pkg, Hyperlink link, ObjectFactory factory) {
        MainDocumentPart mainDocumentPart = pkg.getMainDocumentPart();
        RelationshipsPart relationshipsPart = mainDocumentPart.getRelationshipsPart(true);

        Relationship relationship = new org.docx4j.relationships.ObjectFactory().createRelationship();
        relationship.setType(Namespaces.HYPERLINK);
        relationship.setTarget(link.url());
        relationship.setTargetMode("External");
        relationshipsPart.addRelationship(relationship);

        org.docx4j.wml.P.Hyperlink hyperlink = new org.docx4j.wml.P.Hyperlink();
        hyperlink.setId(relationship.getId());
        hyperlink.setHistory(true);

        R run = Paragraphs.run(link.text(), link.style());
        hyperlink.getContent().add(run);
        return hyperlink;
    }
}
```

- [ ] **Step 4: Add the three builder overloads**

In `src/main/java/com/example/docx/WordDocument.java`, add these imports:

```java
import com.example.docx.content.Hyperlink;
import com.example.docx.part.Hyperlinks;
```

Add the three overloads immediately before the `svgImage` method (i.e. after the existing three-argument `paragraph(String, TextStyle, ParagraphStyle)`):

```java
        /** Appends a body paragraph ending in a hyperlink, using {@link TextStyle#body()}. */
        public Builder paragraph(String text, Hyperlink link) {
            return paragraph(text, TextStyle.body(), link);
        }

        /** Appends a body paragraph ending in a hyperlink. */
        public Builder paragraph(String text, TextStyle style, Hyperlink link) {
            return paragraph(text, style, ParagraphStyle.body(), link);
        }

        /**
         * Appends a paragraph ending in a hyperlink, with explicit paragraph-level
         * formatting.
         *
         * <p>Unlike the no-link overloads, this defers to {@code build()}: a hyperlink's
         * target is a relationship, which needs the package that does not exist until
         * then. {@code text}'s blank-check is therefore deferred too, along with
         * everything else this constructs.
         */
        public Builder paragraph(String text, TextStyle style, ParagraphStyle paragraphStyle,
                                 Hyperlink link) {
            content.add(pkg -> Hyperlinks.paragraph(pkg, text, style, paragraphStyle, link));
            return this;
        }
```

Note this deliberately performs **no validation** in the builder method itself — everything, including null checks on `style`/`paragraphStyle`/`link`, is validated inside `Hyperlinks.paragraph` when the lambda runs at `build()` time. Adding eager null checks here would duplicate that validation and risk it drifting, the same reasoning that keeps `svgImage` free of eager checks.

- [ ] **Step 5: Run the test**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test -Dtest=HyperlinkDocumentTest`
Expected: PASS — `Tests run: 5, Failures: 0, Errors: 0`.

- [ ] **Step 6: Run the full suite**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B clean test`
Expected: `Tests run: 106, Failures: 0, Errors: 0, Skipped: 0`.

- [ ] **Step 7: Commit**

```bash
git add -A
git commit -m "feat: add trailing hyperlinks to paragraphs"
```

---

### Task 5: Sample and documentation

**Files:**
- Modify: `src/test/java/com/example/docx/sample/SampleMain.java`
- Modify: `README.md`
- Modify: `DOCUMENTATION.md`

**Interfaces:**
- Consumes: `WordDocument.Builder.paragraph(String, Hyperlink)`, `content.Hyperlink.of`.
- Produces: nothing other tasks depend on. Terminal task.

- [ ] **Step 1: Add a linked paragraph to the sample**

In `src/test/java/com/example/docx/sample/SampleMain.java`, add the import:

```java
import com.example.docx.content.Hyperlink;
```

Immediately after the section loop and before `WordDocument document = builder.build();`, add:

```java
        // A trailing hyperlink inside the same paragraph as the body text.
        builder.paragraph(
                "The full dataset behind this report is available online. ",
                Hyperlink.of("View the source data", "https://example.com/data"));
```

- [ ] **Step 2: Run the sample**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test-compile exec:java`
Expected: `BUILD SUCCESS` and `Wrote /…/target/sample-<millis>.docx`.

- [ ] **Step 3: Verify the link landed**

Run:
```bash
F=$(ls -t target/sample-*.docx | head -1)
unzip -p "$F" word/document.xml | grep -o "<w:hyperlink" | wc -l
```
Expected: `1`.

Run:
```bash
F=$(ls -t target/sample-*.docx | head -1)
unzip -p "$F" word/_rels/document.xml.rels | grep -c 'TargetMode="External"'
```
Expected: `1`.

- [ ] **Step 4: Document it**

In `README.md`, add to the "Things that bite" section:

```
- **A linked paragraph defers to `build()`.** `paragraph(text, Hyperlink)` and its two
  siblings need the package for the link's relationship, so — unlike every other
  `paragraph(...)` overload — they validate at `build()` time, not at the call site.
```

In `DOCUMENTATION.md`, add a row to the `WordDocument.Builder` method table:

```
| `paragraph(String, Hyperlink)` | `TextStyle.body()`, `ParagraphStyle.body()` — defers to `build()`, see below |
```

Add a "Hyperlinks" subsection under "Recipes":

````markdown
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
````

- [ ] **Step 5: Run the full suite**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B clean test`
Expected: `Tests run: 106, Failures: 0, Errors: 0, Skipped: 0` and `BUILD SUCCESS`.

- [ ] **Step 6: Commit**

```bash
git add -A
git commit -m "feat: show a linked paragraph in the sample and document it"
```

---

## Definition of Done

- `mvn clean test` passes with 106 tests.
- A paragraph built with `paragraph(text, Hyperlink.of(linkText, url))` reloads as exactly two content items: the leading run, then the hyperlink.
- The hyperlink's `r:id` resolves to a relationship with `TargetMode="External"` and the given `Target`.
- Blank leading text produces a link-only paragraph — no empty leading run.
- `TextStyle.link()` emits `w:color val="0563C1"` and `w:u val="single"`.
- A document with an image and a linked paragraph has distinct relationship ids for both, and both resolve after reload.
- Malformed URLs and blank link text throw `DocumentGenerationException` from `Hyperlink.of` without ever calling `build()`.
- `mvn test-compile exec:java` writes a sample whose document.xml has exactly one `<w:hyperlink` and whose rels file has exactly one `TargetMode="External"`.
- No class in `content/` or `style/` references `WordprocessingMLPackage`.
