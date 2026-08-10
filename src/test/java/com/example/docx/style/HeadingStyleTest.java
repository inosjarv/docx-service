package com.example.docx.style;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.example.docx.DocumentGenerationException;
import java.math.BigInteger;
import org.docx4j.wml.RPr;
import org.junit.jupiter.api.Test;

class HeadingStyleTest {

    @Test
    void defaultsMatchTheSpec() {
        HeadingStyle style = HeadingStyle.defaults();
        assertEquals("Calibri Light", style.fontFamily());
        assertEquals(20.0, style.sizePt());
        assertTrue(style.bold());
        assertEquals(false, style.italic());
        assertEquals("1F4E79", style.colorHex());
    }

    @Test
    void stripsLeadingHashAndUppercases() {
        assertEquals("1F4E79", HeadingStyle.builder().color("#1f4e79").build().colorHex());
        assertEquals("1F4E79", HeadingStyle.builder().color("1f4e79").build().colorHex());
        assertEquals("ABCDEF", HeadingStyle.builder().color("#AbCdEf").build().colorHex());
    }

    @Test
    void emitsHalfPointSize() {
        RPr rPr = HeadingStyle.builder().sizePt(20).build().toRPr();
        assertEquals(BigInteger.valueOf(40), rPr.getSz().getVal());
        assertEquals(BigInteger.valueOf(40), rPr.getSzCs().getVal());
    }

    @Test
    void emitsFontOnAsciiAndHAnsi() {
        RPr rPr = HeadingStyle.builder().font("Arial").build().toRPr();
        assertEquals("Arial", rPr.getRFonts().getAscii());
        assertEquals("Arial", rPr.getRFonts().getHAnsi());
    }

    @Test
    void emitsColourWithoutHash() {
        RPr rPr = HeadingStyle.builder().color("#1F4E79").build().toRPr();
        assertEquals("1F4E79", rPr.getColor().getVal());
    }

    @Test
    void boldAndItalicAreOmittedWhenFalse() {
        RPr rPr = HeadingStyle.builder().bold(false).italic(false).build().toRPr();
        assertNull(rPr.getB());
        assertNull(rPr.getI());
    }

    @Test
    void boldAndItalicArePresentWhenTrue() {
        RPr rPr = HeadingStyle.builder().bold(true).italic(true).build().toRPr();
        assertTrue(rPr.getB().isVal());
        assertTrue(rPr.getI().isVal());
    }

    @Test
    void rejectsMalformedColour() {
        assertThrows(DocumentGenerationException.class,
                () -> HeadingStyle.builder().color("blue").build());
        assertThrows(DocumentGenerationException.class,
                () -> HeadingStyle.builder().color("#12345").build());
        assertThrows(DocumentGenerationException.class,
                () -> HeadingStyle.builder().color("#1234567").build());
        assertThrows(DocumentGenerationException.class,
                () -> HeadingStyle.builder().color("#12345G").build());
        assertThrows(DocumentGenerationException.class,
                () -> HeadingStyle.builder().color(null).build());
    }

    @Test
    void rejectsNonPositiveSize() {
        assertThrows(DocumentGenerationException.class,
                () -> HeadingStyle.builder().sizePt(0).build());
        assertThrows(DocumentGenerationException.class,
                () -> HeadingStyle.builder().sizePt(-4).build());
    }

    @Test
    void rejectsBlankFont() {
        assertThrows(DocumentGenerationException.class,
                () -> HeadingStyle.builder().font("   ").build());
        assertThrows(DocumentGenerationException.class,
                () -> HeadingStyle.builder().font(null).build());
    }
}
