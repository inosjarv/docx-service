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
