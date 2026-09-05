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
