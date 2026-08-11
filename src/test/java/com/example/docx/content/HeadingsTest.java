package com.example.docx.content;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import com.example.docx.DocumentGenerationException;
import com.example.docx.style.TextStyle;
import jakarta.xml.bind.JAXBElement;
import java.math.BigInteger;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Text;
import org.junit.jupiter.api.Test;

class HeadingsTest {

    /** Freshly built runs hold a bare Text; reloaded ones hold a JAXBElement. */
    private static Text textOf(R run) {
        Object first = run.getContent().get(0);
        return (Text) (first instanceof JAXBElement<?> je ? je.getValue() : first);
    }

    @Test
    void buildsParagraphWithOneStyledRun() {
        P p = Headings.heading("Quarterly Report", TextStyle.defaults());

        assertEquals(1, p.getContent().size());
        R run = (R) p.getContent().get(0);

        assertEquals("Quarterly Report", textOf(run).getValue());
        assertEquals(BigInteger.valueOf(40), run.getRPr().getSz().getVal());
        assertEquals("1F4E79", run.getRPr().getColor().getVal());
    }

    @Test
    void preservesSurroundingWhitespace() {
        P p = Headings.heading("  spaced  ", TextStyle.defaults());
        R run = (R) p.getContent().get(0);
        assertEquals("preserve", textOf(run).getSpace());
        assertEquals("  spaced  ", textOf(run).getValue());
    }

    @Test
    void rejectsBlankText() {
        assertThrows(DocumentGenerationException.class,
                () -> Headings.heading("   ", TextStyle.defaults()));
        assertThrows(DocumentGenerationException.class,
                () -> Headings.heading(null, TextStyle.defaults()));
    }

    @Test
    void rejectsNullStyle() {
        assertThrows(DocumentGenerationException.class,
                () -> Headings.heading("Title", null));
    }
}
