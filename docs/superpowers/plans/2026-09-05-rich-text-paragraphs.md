# Rich Text Paragraphs Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Let callers pass a string containing a limited set of inline HTML tags (e.g. `Some text which needs to be <b>bold</b>`) and get a paragraph with correctly mixed run formatting, instead of every paragraph being forced into one style for its entire text.

**Architecture:** A new immutable value type, `content.RichText`, parses the tagged string eagerly into an ordered list of (text, resolved `TextStyle`) spans, using the JDK's built-in DOM parser (the string is wrapped in a synthetic root element and parsed as XML — no new dependency). `Paragraphs` gets a new overload that turns a `RichText` into a multi-run `w:p`, and `WordDocument.Builder` gets matching `paragraph(RichText, ...)` overloads. This is purely additive: the existing `paragraph(String, ...)` overloads, which treat their argument as literal text, are untouched.

**Tech Stack:** Java 25, docx4j, `javax.xml.parsers` (JDK DOM parser), JUnit 5. No new Maven dependency.

**Spec:** No separate spec document — this is a bounded change (an addition to code that already exists: `Paragraphs`, `WordDocument.Builder`). The design was agreed in chat on 2026-09-05 and is summarized in Global Constraints below.

## Global Constraints

- Recognized tags are exactly `<b>` / `<strong>` (bold), `<i>` / `<em>` (italic), `<u>` (underline) — matched case-insensitively on the tag name, freely nestable (e.g. `<b>bold <i>and italic</i></b>` produces a run that is both). No other tag (e.g. `<br>`, `<span>`) is supported.
- Any problem with the markup — an unrecognized tag, an unclosed tag, a stray unescaped `&`/`<`/`>` in the plain text — throws `DocumentGenerationException` immediately. There is no best-effort recovery, matching every other validation already in this codebase (`TextStyle`, `Hyperlink`, `Paragraphs` all fail fast on bad input).
- Because parsing is XML-based, plain text portions must escape XML's special characters the same way HTML requires (`&amp;`, `&lt;`, `&gt;`). This is a real constraint on callers and must be documented on the new class, not just in this plan.
- This is new, explicit API surface (`RichText` plus `paragraph(RichText, ...)` overloads) parallel to the existing `Hyperlink` plus `paragraph(String, ..., Hyperlink)` overloads. The existing `paragraph(String, ...)` overloads are not touched and do not sniff their argument for tags — a literal `<` in plain text keeps rendering as-is.
- No new Maven dependency: parsing uses `javax.xml.parsers.DocumentBuilderFactory`/`DocumentBuilder`, already on the JDK classpath. Because the input is untrusted request data, the factory must be hardened against XXE: DOCTYPE declarations and external entities disabled.

---

### Task 1: `RichText` value type

**Files:**
- Create: `src/main/java/com/example/docx/content/RichText.java`
- Test: `src/test/java/com/example/docx/content/RichTextTest.java`

**Interfaces:**
- Consumes: `com.example.docx.DocumentGenerationException` (existing), `com.example.docx.style.TextStyle` (existing — `builder()`, `body()`, and its getters `fontFamily()`, `sizePt()`, `bold()`, `italic()`, `underline()`, `colorHex()`), `com.example.docx.content.Paragraphs.run(String, TextStyle)` (existing, same package).
- Produces: `RichText.of(String markup)`, `RichText.of(String markup, TextStyle baseStyle)`, both returning `RichText` or throwing `DocumentGenerationException`; instance method `List<org.docx4j.wml.R> toRuns()`. Task 2 calls both.

- [ ] **Step 1: Write the failing tests**

Create `src/test/java/com/example/docx/content/RichTextTest.java`:

```java
package com.example.docx.content;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.example.docx.DocumentGenerationException;
import com.example.docx.style.TextStyle;
import jakarta.xml.bind.JAXBElement;
import java.util.List;
import org.docx4j.wml.R;
import org.docx4j.wml.Text;
import org.junit.jupiter.api.Test;

class RichTextTest {

    private static Text textOf(R run) {
        Object first = run.getContent().get(0);
        return (Text) (first instanceof JAXBElement<?> je ? je.getValue() : first);
    }

    @Test
    void plainTextProducesOneRunWithNoInlineFormatting() {
        List<R> runs = RichText.of("Hello world").toRuns();
        assertEquals(1, runs.size());
        assertEquals("Hello world", textOf(runs.get(0)).getValue());
        assertNull(runs.get(0).getRPr().getB());
    }

    @Test
    void boldTagSplitsIntoThreeRunsWithTheMiddleOneBold() {
        List<R> runs = RichText.of("Some text which needs to be <b>bold</b>.").toRuns();
        assertEquals(3, runs.size());

        assertEquals("Some text which needs to be ", textOf(runs.get(0)).getValue());
        assertNull(runs.get(0).getRPr().getB());

        assertEquals("bold", textOf(runs.get(1)).getValue());
        assertTrue(runs.get(1).getRPr().getB().isVal());

        assertEquals(".", textOf(runs.get(2)).getValue());
        assertNull(runs.get(2).getRPr().getB());
    }

    @Test
    void strongEmAndUAreAcceptedAliases() {
        List<R> runs = RichText.of("<strong>a</strong> <em>b</em> <u>c</u>").toRuns();
        assertEquals(3, runs.size());
        assertTrue(runs.get(0).getRPr().getB().isVal());
        assertTrue(runs.get(1).getRPr().getI().isVal());
        assertEquals("single", runs.get(2).getRPr().getU().getVal().value());
    }

    @Test
    void tagMatchingIsCaseInsensitive() {
        List<R> runs = RichText.of("<B>bold</B>").toRuns();
        assertEquals(1, runs.size());
        assertTrue(runs.get(0).getRPr().getB().isVal());
    }

    @Test
    void nestedTagsCombineFormatting() {
        List<R> runs = RichText.of("<b>bold <i>and italic</i></b>").toRuns();
        assertEquals(2, runs.size());

        assertEquals("bold ", textOf(runs.get(0)).getValue());
        assertTrue(runs.get(0).getRPr().getB().isVal());
        assertNull(runs.get(0).getRPr().getI());

        assertEquals("and italic", textOf(runs.get(1)).getValue());
        assertTrue(runs.get(1).getRPr().getB().isVal());
        assertTrue(runs.get(1).getRPr().getI().isVal());
    }

    @Test
    void whitespaceBetweenTagsIsKeptNotDropped() {
        List<R> runs = RichText.of("<b>Hello</b> <i>world</i>").toRuns();
        assertEquals(2, runs.size());
        assertEquals("Hello ", textOf(runs.get(0)).getValue());
        assertEquals("world", textOf(runs.get(1)).getValue());
    }

    @Test
    void explicitBaseStyleCarriesThroughAndTagsAddOnTop() {
        TextStyle base = TextStyle.builder().font("Arial").sizePt(14).bold(false).italic(true).build();
        R run = RichText.of("<b>x</b>", base).toRuns().get(0);
        assertEquals("Arial", run.getRPr().getRFonts().getAscii());
        assertTrue(run.getRPr().getI().isVal(), "base style's italic must survive");
        assertTrue(run.getRPr().getB().isVal(), "the tag must add bold on top");
    }

    @Test
    void escapedAmpersandIsAccepted() {
        R run = RichText.of("Fish &amp; Chips").toRuns().get(0);
        assertEquals("Fish & Chips", textOf(run).getValue());
    }

    @Test
    void rejectsBlankMarkupAndNullBaseStyle() {
        assertThrows(DocumentGenerationException.class, () -> RichText.of(null));
        assertThrows(DocumentGenerationException.class, () -> RichText.of(""));
        assertThrows(DocumentGenerationException.class, () -> RichText.of("   "));
        assertThrows(DocumentGenerationException.class, () -> RichText.of("text", null));
    }

    @Test
    void rejectsAnUnrecognisedTag() {
        assertThrows(DocumentGenerationException.class, () -> RichText.of("<span>x</span>"));
    }

    @Test
    void rejectsUnclosedTags() {
        assertThrows(DocumentGenerationException.class, () -> RichText.of("<b>bold"));
    }

    @Test
    void rejectsAnUnescapedAmpersand() {
        assertThrows(DocumentGenerationException.class, () -> RichText.of("Fish & Chips"));
    }
}
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `mvn -q -Dtest=RichTextTest test`
Expected: compile failure — `RichText` does not exist yet.

- [ ] **Step 3: Write the implementation**

Create `src/main/java/com/example/docx/content/RichText.java`:

```java
package com.example.docx.content;

import com.example.docx.DocumentGenerationException;
import com.example.docx.style.TextStyle;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import org.docx4j.wml.R;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

/**
 * Text carrying a limited set of inline markup on top of one base {@link TextStyle}:
 * {@code <b>}/{@code <strong>} for bold, {@code <i>}/{@code <em>} for italic, {@code <u>}
 * for underline. Tags nest freely — {@code <b>bold <i>and italic</i></b>} yields a run
 * that is both — and tag names are matched case-insensitively, though an opening and
 * closing tag must still agree in case with each other (XML itself is case-sensitive).
 *
 * <p>Markup is parsed as XML, so the plain-text portions must escape the characters XML
 * gives special meaning: write {@code &amp;}, {@code &lt;}, {@code &gt;} rather than bare
 * {@code &}, {@code <}, {@code >} — the same rule HTML itself imposes. An unrecognised
 * tag, an unclosed tag, or any other well-formedness problem throws
 * {@link DocumentGenerationException} rather than guessing at intent.
 *
 * <p>Parses eagerly at {@link #of}: the result is a fixed, already-validated sequence of
 * (text, style) spans, one {@code w:r} per span once {@link #toRuns()} runs.
 */
public final class RichText {

    private final List<Span> spans;

    private RichText(List<Span> spans) {
        this.spans = spans;
    }

    /** Markup styled against {@link TextStyle#body()}. */
    public static RichText of(String markup) {
        return of(markup, TextStyle.body());
    }

    /** Markup styled against an explicit base style; tags only add bold/italic/underline on top. */
    public static RichText of(String markup, TextStyle baseStyle) {
        if (markup == null || markup.isBlank()) {
            throw new DocumentGenerationException("rich text must not be blank");
        }
        if (baseStyle == null) {
            throw new DocumentGenerationException("base text style must not be null");
        }

        List<Span> raw = new ArrayList<>();
        collectSpans(parse(markup).getDocumentElement(), baseStyle, raw);
        List<Span> spans = foldWhitespaceOnlySpans(raw);

        if (spans.isEmpty() || spans.stream().anyMatch(s -> s.text().isBlank())) {
            throw new DocumentGenerationException(
                    "rich text must contain visible, non-whitespace text: '" + markup + "'");
        }
        return new RichText(spans);
    }

    /** One {@code w:r} per parsed span, in document order. */
    public List<R> toRuns() {
        List<R> runs = new ArrayList<>();
        for (Span span : spans) {
            runs.add(Paragraphs.run(span.text(), span.style()));
        }
        return runs;
    }

    private record Span(String text, TextStyle style) {
    }

    private static Document parse(String markup) {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        try {
            // This parses untrusted request data: no DTDs, no external entities (XXE).
            factory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
            factory.setFeature("http://xml.org/sax/features/external-general-entities", false);
            factory.setFeature("http://xml.org/sax/features/external-parameter-entities", false);
            factory.setXIncludeAware(false);
            factory.setExpandEntityReferences(false);
            DocumentBuilder builder = factory.newDocumentBuilder();
            String wrapped = "<root>" + markup + "</root>";
            return builder.parse(new ByteArrayInputStream(wrapped.getBytes(StandardCharsets.UTF_8)));
        } catch (ParserConfigurationException e) {
            throw new IllegalStateException("XML parser misconfigured", e);
        } catch (SAXException | IOException e) {
            throw new DocumentGenerationException(
                    "rich text is not well-formed markup: '" + markup + "'", e);
        }
    }

    private static void collectSpans(Node node, TextStyle style, List<Span> spans) {
        NodeList children = node.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            Node child = children.item(i);
            switch (child.getNodeType()) {
                case Node.TEXT_NODE -> {
                    String text = child.getNodeValue();
                    if (!text.isEmpty()) {
                        spans.add(new Span(text, style));
                    }
                }
                case Node.ELEMENT_NODE ->
                        collectSpans(child, styleFor(child.getNodeName(), style), spans);
                default -> throw new DocumentGenerationException(
                        "unsupported markup content: '" + child.getNodeName() + "'");
            }
        }
    }

    private static TextStyle styleFor(String tagName, TextStyle base) {
        TextStyle.Builder builder = TextStyle.builder()
                .font(base.fontFamily())
                .sizePt(base.sizePt())
                .bold(base.bold())
                .italic(base.italic())
                .underline(base.underline())
                .color(base.colorHex());
        switch (tagName.toLowerCase(Locale.ROOT)) {
            case "b", "strong" -> builder.bold(true);
            case "i", "em" -> builder.italic(true);
            case "u" -> builder.underline(true);
            default -> throw new DocumentGenerationException(
                    "unrecognised rich text tag: '" + tagName + "'");
        }
        return builder.build();
    }

    /**
     * A text node holding only whitespace (the gap between two adjacent tags) has no
     * style of its own and would otherwise become a run {@link Paragraphs#run} rejects
     * as blank. Folding it into whichever span comes before it — or after it, for
     * whitespace at the very start — keeps every span non-blank without discarding any
     * of the whitespace.
     */
    private static List<Span> foldWhitespaceOnlySpans(List<Span> raw) {
        List<Span> result = new ArrayList<>(raw);
        for (int i = 0; i < result.size(); i++) {
            Span span = result.get(i);
            if (!span.text().isBlank() || result.size() == 1) {
                continue;
            }
            if (i > 0) {
                Span prev = result.get(i - 1);
                result.set(i - 1, new Span(prev.text() + span.text(), prev.style()));
            } else {
                Span next = result.get(i + 1);
                result.set(i + 1, new Span(span.text() + next.text(), next.style()));
            }
            result.remove(i);
            i--;
        }
        return result;
    }
}
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `mvn -q -Dtest=RichTextTest test`
Expected: PASS, all 12 tests green.

- [ ] **Step 5: Commit**

```bash
git add src/main/java/com/example/docx/content/RichText.java src/test/java/com/example/docx/content/RichTextTest.java
git commit -m "feat: add content.RichText, parsing a limited inline-markup subset into styled spans

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_015AMYuNu7yCy7ja3gmwxaoo"
```

---

### Task 2: `Paragraphs.of(RichText, ParagraphStyle)`

**Files:**
- Modify: `src/main/java/com/example/docx/content/Paragraphs.java`
- Modify: `src/test/java/com/example/docx/content/ParagraphsTest.java` (add tests; existing tests and the `textOf` helper already there are untouched)

**Interfaces:**
- Consumes: `RichText.of(String)` and `RichText#toRuns()` from Task 1; `ParagraphStyle#toPPr()` (existing).
- Produces: `Paragraphs.of(RichText richText, ParagraphStyle paragraphStyle)` returning `org.docx4j.wml.P`, or throwing `DocumentGenerationException` on a null argument. Task 3 calls this.

- [ ] **Step 1: Write the failing tests**

Add to `src/test/java/com/example/docx/content/ParagraphsTest.java`, inside the `ParagraphsTest` class (after the existing `runRejectsBlankTextAndNullStyle` test) — also add `assertTrue` to the existing static-import block at the top of the file:

```java
    @Test
    void buildsOneParagraphWithOneRunPerRichTextSpan() {
        P p = Paragraphs.of(
                RichText.of("Some text which needs to be <b>bold</b>."), ParagraphStyle.body());

        assertEquals(3, p.getContent().size());
        R bold = (R) p.getContent().get(1);
        assertEquals("bold", textOf(bold).getValue());
        assertTrue(bold.getRPr().getB().isVal());
    }

    @Test
    void richTextParagraphUsesTheGivenParagraphStyle() {
        P p = Paragraphs.of(RichText.of("Overview"), ParagraphStyle.heading());
        assertNotNull(p.getPPr().getKeepNext(), "heading paragraph style must set keepNext");
    }

    @Test
    void ofRichTextRejectsNulls() {
        assertThrows(DocumentGenerationException.class,
                () -> Paragraphs.of((RichText) null, ParagraphStyle.body()));
        assertThrows(DocumentGenerationException.class,
                () -> Paragraphs.of(RichText.of("Text"), null));
    }
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `mvn -q -Dtest=ParagraphsTest test`
Expected: compile failure — no `Paragraphs.of(RichText, ParagraphStyle)` overload yet.

- [ ] **Step 3: Write the implementation**

In `src/main/java/com/example/docx/content/Paragraphs.java`, add this method after `of(String, TextStyle, ParagraphStyle)` (no new imports needed — `RichText` is in the same package):

```java
    /** A {@code w:p} whose runs come from a {@link RichText}'s parsed spans. */
    public static P of(RichText richText, ParagraphStyle paragraphStyle) {
        if (richText == null) {
            throw new DocumentGenerationException("rich text must not be null");
        }
        if (paragraphStyle == null) {
            throw new DocumentGenerationException("paragraph style must not be null");
        }

        ObjectFactory factory = Context.getWmlObjectFactory();
        P paragraph = factory.createP();
        paragraph.setPPr(paragraphStyle.toPPr());
        paragraph.getContent().addAll(richText.toRuns());
        return paragraph;
    }
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `mvn -q -Dtest=ParagraphsTest test`
Expected: PASS, all tests green (existing + 3 new).

- [ ] **Step 5: Commit**

```bash
git add src/main/java/com/example/docx/content/Paragraphs.java src/test/java/com/example/docx/content/ParagraphsTest.java
git commit -m "feat: add Paragraphs.of(RichText, ParagraphStyle) building multi-run paragraphs

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_015AMYuNu7yCy7ja3gmwxaoo"
```

---

### Task 3: `WordDocument.Builder.paragraph(RichText, ...)`

**Files:**
- Modify: `src/main/java/com/example/docx/WordDocument.java`
- Create: `src/test/java/com/example/docx/RichTextDocumentTest.java` (mirrors the existing `HyperlinkDocumentTest.java` pattern — a document-level round-trip test file per content addition)

**Interfaces:**
- Consumes: `RichText.of(String)` from Task 1, `Paragraphs.of(RichText, ParagraphStyle)` from Task 2, `ParagraphStyle.body()`/`ParagraphStyle.heading()` (existing).
- Produces: `WordDocument.Builder.paragraph(RichText richText)` and `paragraph(RichText richText, ParagraphStyle paragraphStyle)`, both returning `Builder`. This is the public entry point end users call — no later task depends on it.

- [ ] **Step 1: Write the failing tests**

Create `src/test/java/com/example/docx/RichTextDocumentTest.java`:

```java
package com.example.docx;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.example.docx.content.RichText;
import com.example.docx.style.ParagraphStyle;
import jakarta.xml.bind.JAXBElement;
import java.io.ByteArrayInputStream;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.Body;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Text;
import org.junit.jupiter.api.Test;

class RichTextDocumentTest {

    private static Body reload(byte[] bytes) throws Exception {
        WordprocessingMLPackage pkg = WordprocessingMLPackage.load(new ByteArrayInputStream(bytes));
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
    void richTextParagraphReloadsAsThreeRunsWithTheMiddleOneBold() throws Exception {
        byte[] bytes = WordDocument.builder()
                .paragraph(RichText.of("Some text which needs to be <b>bold</b>."))
                .build()
                .toByteArray();

        P p = firstParagraph(reload(bytes));
        assertEquals(3, p.getContent().size());

        R first = (R) p.getContent().get(0);
        assertEquals("Some text which needs to be ", textOf(first));
        assertNull(first.getRPr().getB());

        R bold = (R) p.getContent().get(1);
        assertEquals("bold", textOf(bold));
        assertTrue(bold.getRPr().getB().isVal());

        R last = (R) p.getContent().get(2);
        assertEquals(".", textOf(last));
        assertNull(last.getRPr().getB());
    }

    @Test
    void richTextParagraphDefaultsToBodyParagraphStyle() throws Exception {
        byte[] bytes = WordDocument.builder()
                .paragraph(RichText.of("Plain body text"))
                .build()
                .toByteArray();

        P p = firstParagraph(reload(bytes));
        assertNull(p.getPPr().getKeepNext(), "body paragraphs must not set keepNext");
    }

    @Test
    void richTextParagraphAcceptsExplicitParagraphStyle() throws Exception {
        byte[] bytes = WordDocument.builder()
                .paragraph(RichText.of("Heading-styled text"), ParagraphStyle.heading())
                .build()
                .toByteArray();

        P p = firstParagraph(reload(bytes));
        assertNotNull(p.getPPr().getKeepNext(), "explicit heading paragraph style must set keepNext");
    }
}
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `mvn -q -Dtest=RichTextDocumentTest test`
Expected: compile failure — `WordDocument.Builder` has no `paragraph(RichText)` overload yet.

- [ ] **Step 3: Write the implementation**

In `src/main/java/com/example/docx/WordDocument.java`:

Add the import, alongside the existing `content` imports near the top of the file:

```java
import com.example.docx.content.RichText;
```

Add these two methods to `Builder`, directly after `paragraph(String text, TextStyle style, ParagraphStyle paragraphStyle)` and before the hyperlink `paragraph(...)` overloads:

```java
        /** Appends a body paragraph built from {@link RichText}, using {@link ParagraphStyle#body()}. */
        public Builder paragraph(RichText richText) {
            return paragraph(richText, ParagraphStyle.body());
        }

        /** Appends a body paragraph built from {@link RichText} with explicit paragraph-level formatting. */
        public Builder paragraph(RichText richText, ParagraphStyle paragraphStyle) {
            P paragraph = Paragraphs.of(richText, paragraphStyle);
            content.add(pkg -> paragraph);
            return this;
        }
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `mvn -q -Dtest=RichTextDocumentTest test`
Expected: PASS, all 3 tests green.

- [ ] **Step 5: Run the full suite**

Run: `mvn -q test`
Expected: PASS, no regressions in any existing test class.

- [ ] **Step 6: Commit**

```bash
git add src/main/java/com/example/docx/WordDocument.java src/test/java/com/example/docx/RichTextDocumentTest.java
git commit -m "feat: add WordDocument.Builder.paragraph(RichText, ...) overloads

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_015AMYuNu7yCy7ja3gmwxaoo"
```
