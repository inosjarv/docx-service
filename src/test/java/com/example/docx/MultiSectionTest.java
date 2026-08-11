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
