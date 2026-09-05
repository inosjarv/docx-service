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
            builder.setErrorHandler(new org.xml.sax.helpers.DefaultHandler());
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
