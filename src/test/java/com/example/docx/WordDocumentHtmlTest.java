package com.example.docx;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.example.docx.page.PageSetup;
import com.example.docx.style.TableStyle;
import com.example.docx.style.TextStyle;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.StringReader;
import java.nio.charset.StandardCharsets;
import java.util.List;
import javax.xml.parsers.DocumentBuilderFactory;
import org.junit.jupiter.api.Test;
import org.w3c.dom.Document;
import org.xml.sax.InputSource;

class WordDocumentHtmlTest {

    private static byte[] resource(String name) throws IOException {
        try (InputStream in = WordDocumentHtmlTest.class.getResourceAsStream(name)) {
            assertNotNull(in, "missing test resource " + name);
            return in.readAllBytes();
        }
    }

    /** Parses as XML to catch malformed XHTML; returns the source string unchanged. */
    private static String assertWellFormedXml(String html) throws Exception {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        factory.setNamespaceAware(true);
        Document parsed = factory.newDocumentBuilder().parse(new InputSource(new StringReader(html)));
        assertNotNull(parsed.getDocumentElement());
        return html;
    }

    @Test
    void toHtmlContainsHeadingAndParagraphText() throws Exception {
        WordDocument doc = WordDocument.builder()
                .pageSetup(PageSetup.a4())
                .heading("Quarterly Report", TextStyle.defaults())
                .paragraph("Revenue grew across every region.")
                .build();

        String html = assertWellFormedXml(doc.toHtml());

        assertTrue(html.contains("Quarterly Report"), "heading text missing from " + html);
        assertTrue(html.contains("Revenue grew across every region."), "paragraph text missing");
    }

    @Test
    void toHtmlRendersTableContent() throws Exception {
        WordDocument doc = WordDocument.builder()
                .pageSetup(PageSetup.a4())
                .heading("Report")
                .table(
                        List.of("Region", "Revenue"),
                        List.of(List.of("EMEA", "1 240")),
                        TableStyle.builder().build())
                .build();

        String html = assertWellFormedXml(doc.toHtml());

        assertTrue(html.contains("<table"), "expected a <table> element in " + html);
        assertTrue(html.contains("Region") && html.contains("Revenue") && html.contains("EMEA"),
                "table cell text missing from " + html);
    }

    @Test
    void toHtmlEmbedsImagesAsDataUris() throws Exception {
        WordDocument doc = WordDocument.builder()
                .pageSetup(PageSetup.a4())
                .heading("Report")
                .svgImage(resource("/demo/chart.svg"), resource("/demo/chart.png"))
                .build();

        String html = assertWellFormedXml(doc.toHtml());

        assertTrue(html.contains("data:image/"), "expected an embedded data: image URI in " + html);
        assertFalse(html.contains("word/media"), "image should not reference the docx's internal part path");
    }

    @Test
    void toHtmlGivesTheBodyOnScreenMargin() throws Exception {
        WordDocument doc = WordDocument.builder().heading("Report").build();

        String html = assertWellFormedXml(doc.toHtml());

        // docx4j's own generated CSS only sets margin inside @page {...}, a print-media
        // rule browsers ignore on screen — so screen margin must come from us.
        assertTrue(html.matches("(?s).*body\\s*\\{[^}]*margin[^}]*\\}.*"),
                "expected an on-screen body margin rule in " + html);
    }

    @Test
    void toHtmlLetsTablesShrinkToFitTheViewport() throws Exception {
        WordDocument doc = WordDocument.builder()
                .heading("Report")
                .table(List.of("Region"), List.of(List.of("EMEA")), TableStyle.builder().build())
                .build();

        String html = assertWellFormedXml(doc.toHtml());

        // docx4j hard-codes an absolute inch width on every <table> with no responsive
        // fallback, which is what overflows a narrower viewport.
        assertTrue(html.matches("(?s).*table\\s*\\{[^}]*max-width\\s*:\\s*100%[^}]*\\}.*"),
                "expected a responsive max-width rule for tables in " + html);
    }

    @Test
    void toHtmlAlignsTableHeadersLikeTheirBody() throws Exception {
        WordDocument doc = WordDocument.builder()
                .heading("Report")
                .table(List.of("Region"), List.of(List.of("EMEA")), TableStyle.builder().build())
                .build();

        String html = assertWellFormedXml(doc.toHtml());

        // Neither <th> nor <td> carries a text-align: the library's default
        // Alignment.LEFT emits no w:jc at all, so both cell kinds are unstyled in the
        // same way. Left alone, every browser's UA stylesheet still centers <th> text
        // while leaving <td> left-aligned, so headers visibly drift from body content
        // unless something neutralises that default explicitly.
        assertTrue(html.matches("(?s).*th\\s*\\{[^}]*text-align\\s*:\\s*inherit[^}]*\\}.*"),
                "expected a rule aligning <th> text like its <td> siblings in " + html);
    }

    @Test
    void writeHtmlToAndToHtmlProduceEquivalentOutput() throws Exception {
        WordDocument doc = WordDocument.builder()
                .pageSetup(PageSetup.a4())
                .heading("Quarterly Report", TextStyle.defaults())
                .build();

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        doc.writeHtmlTo(out);
        String streamed = out.toString(StandardCharsets.UTF_8);

        assertTrue(streamed.contains("Quarterly Report"));
        assertTrue(doc.toHtml().contains("Quarterly Report"));
    }

    @Test
    void writeHtmlToLeavesTheCallersStreamOpen() {
        WordDocument doc = WordDocument.builder()
                .heading("Title", TextStyle.defaults())
                .build();

        class TrackingStream extends ByteArrayOutputStream {
            boolean closed;

            @Override
            public void close() {
                closed = true;
            }
        }

        TrackingStream out = new TrackingStream();
        doc.writeHtmlTo(out);

        assertTrue(out.size() > 0, "should have written bytes");
        assertFalse(out.closed, "writeHtmlTo must not close the caller's stream");
    }

    @Test
    void rejectsNullOutputStream() {
        WordDocument doc = WordDocument.builder().heading("Title").build();
        assertThrows(DocumentGenerationException.class, () -> doc.writeHtmlTo(null));
    }
}
