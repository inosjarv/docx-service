package com.example.docx.content;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.example.docx.DocumentGenerationException;
import com.example.docx.style.ParagraphStyle;
import com.example.docx.style.TextStyle;
import jakarta.xml.bind.JAXBElement;
import java.math.BigInteger;
import java.util.List;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Text;
import org.junit.jupiter.api.Test;

class ParagraphsTest {

    /** Freshly built runs hold a bare Text; reloaded ones hold a JAXBElement. */
    private static Text textOf(R run) {
        return textAt(run.getContent(), 0);
    }

    /** Same unwrapping as {@link #textOf}, for a content item that isn't the first. */
    private static Text textAt(List<Object> content, int index) {
        Object item = content.get(index);
        return (Text) (item instanceof JAXBElement<?> je ? je.getValue() : item);
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

    @Test
    void runSplitsTabsIntoRealTabElementsBetweenTextSegments() {
        R run = Paragraphs.run("Label:\tValue", TextStyle.body());

        assertEquals(3, run.getContent().size());
        assertEquals("Label:", textAt(run.getContent(), 0).getValue());
        assertTrue(run.getContent().get(1) instanceof org.docx4j.wml.R.Tab,
                "a tab character must become a real <w:tab/>, not literal text");
        assertEquals("Value", textAt(run.getContent(), 2).getValue());
    }

    @Test
    void runWithConsecutiveOrEdgeTabsEmitsNoEmptyTextSegments() {
        R leading = Paragraphs.run("\tValue", TextStyle.body());
        assertEquals(2, leading.getContent().size(), "no empty Text before the tab");
        assertTrue(leading.getContent().get(0) instanceof org.docx4j.wml.R.Tab);
        assertEquals("Value", textAt(leading.getContent(), 1).getValue());

        R trailing = Paragraphs.run("Label:\t", TextStyle.body());
        assertEquals(2, trailing.getContent().size(), "no empty Text after the tab");
        assertEquals("Label:", textAt(trailing.getContent(), 0).getValue());
        assertTrue(trailing.getContent().get(1) instanceof org.docx4j.wml.R.Tab);

        R consecutive = Paragraphs.run("A\t\tB", TextStyle.body());
        assertEquals(4, consecutive.getContent().size(), "two tabs, nothing empty between them");
        assertTrue(consecutive.getContent().get(1) instanceof org.docx4j.wml.R.Tab);
        assertTrue(consecutive.getContent().get(2) instanceof org.docx4j.wml.R.Tab);
    }

    @Test
    void runWithoutTabsStillProducesExactlyOneTextSegment() {
        R run = Paragraphs.run("No tabs here", TextStyle.body());
        assertEquals(1, run.getContent().size());
        assertEquals("No tabs here", textOf(run).getValue());
    }

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
}
