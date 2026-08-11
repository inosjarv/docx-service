package com.example.docx.style;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotSame;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.example.docx.DocumentGenerationException;
import java.math.BigInteger;
import org.docx4j.wml.RPr;
import org.junit.jupiter.api.Test;

class TextStyleTest {

    @Test
    void defaultsMatchTheSpec() {
        TextStyle style = TextStyle.defaults();
        assertEquals("Calibri Light", style.fontFamily());
        assertEquals(20.0, style.sizePt());
        assertTrue(style.bold());
        assertEquals(false, style.italic());
        assertEquals("1F4E79", style.colorHex());
    }

    @Test
    void stripsLeadingHashAndUppercases() {
        assertEquals("1F4E79", TextStyle.builder().color("#1f4e79").build().colorHex());
        assertEquals("1F4E79", TextStyle.builder().color("1f4e79").build().colorHex());
        assertEquals("ABCDEF", TextStyle.builder().color("#AbCdEf").build().colorHex());
    }

    @Test
    void emitsHalfPointSize() {
        RPr rPr = TextStyle.builder().sizePt(20).build().toRPr();
        assertEquals(BigInteger.valueOf(40), rPr.getSz().getVal());
        assertEquals(BigInteger.valueOf(40), rPr.getSzCs().getVal());
    }

    @Test
    void emitsFontOnAsciiAndHAnsi() {
        RPr rPr = TextStyle.builder().font("Arial").build().toRPr();
        assertEquals("Arial", rPr.getRFonts().getAscii());
        assertEquals("Arial", rPr.getRFonts().getHAnsi());
    }

    @Test
    void emitsColourWithoutHash() {
        RPr rPr = TextStyle.builder().color("#1F4E79").build().toRPr();
        assertEquals("1F4E79", rPr.getColor().getVal());
    }

    @Test
    void boldAndItalicAreOmittedWhenFalse() {
        RPr rPr = TextStyle.builder().bold(false).italic(false).build().toRPr();
        assertNull(rPr.getB());
        assertNull(rPr.getI());
    }

    @Test
    void boldAndItalicArePresentWhenTrue() {
        RPr rPr = TextStyle.builder().bold(true).italic(true).build().toRPr();
        assertTrue(rPr.getB().isVal());
        assertTrue(rPr.getI().isVal());
    }

    @Test
    void rejectsMalformedColour() {
        assertThrows(DocumentGenerationException.class,
                () -> TextStyle.builder().color("blue").build());
        assertThrows(DocumentGenerationException.class,
                () -> TextStyle.builder().color("#12345").build());
        assertThrows(DocumentGenerationException.class,
                () -> TextStyle.builder().color("#1234567").build());
        assertThrows(DocumentGenerationException.class,
                () -> TextStyle.builder().color("#12345G").build());
        assertThrows(DocumentGenerationException.class,
                () -> TextStyle.builder().color(null).build());
    }

    @Test
    void rejectsNonPositiveSize() {
        assertThrows(DocumentGenerationException.class,
                () -> TextStyle.builder().sizePt(0).build());
        assertThrows(DocumentGenerationException.class,
                () -> TextStyle.builder().sizePt(-4).build());
    }

    @Test
    void rejectsBlankFont() {
        assertThrows(DocumentGenerationException.class,
                () -> TextStyle.builder().font("   ").build());
        assertThrows(DocumentGenerationException.class,
                () -> TextStyle.builder().font(null).build());
    }

    @Test
    void rejectsSizesBeyondWhatWordSupports() {
        // ST_HpsMeasure is capped at 1638 pt in Word. Rejecting at build() keeps a
        // validated TextStyle from throwing later at toRPr() time.
        assertThrows(DocumentGenerationException.class,
                () -> TextStyle.builder().sizePt(1639).build());
        assertThrows(DocumentGenerationException.class,
                () -> TextStyle.builder().sizePt(1e12).build());
        assertEquals(1638.0, TextStyle.builder().sizePt(1638).build().sizePt());
    }

    @Test
    void sizeAndComplexScriptSizeAreIndependentObjects() {
        RPr rPr = TextStyle.builder().sizePt(20).build().toRPr();
        assertNotSame(rPr.getSz(), rPr.getSzCs(),
                "sz and szCs must not share one mutable HpsMeasure");
        assertEquals(rPr.getSz().getVal(), rPr.getSzCs().getVal());
    }

    @Test
    void bodyIsElevenPointRegularBlack() {
        TextStyle body = TextStyle.body();
        assertEquals("Calibri", body.fontFamily());
        assertEquals(11.0, body.sizePt());
        assertEquals(false, body.bold());
        assertEquals(false, body.italic());
        assertEquals("000000", body.colorHex());
        // 11 pt is 22 half-points. Points, never pixels.
        assertEquals(BigInteger.valueOf(22), body.toRPr().getSz().getVal());
    }
}
