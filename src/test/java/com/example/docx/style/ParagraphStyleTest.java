package com.example.docx.style;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.example.docx.DocumentGenerationException;
import java.math.BigInteger;
import org.docx4j.wml.PPr;
import org.junit.jupiter.api.Test;

class ParagraphStyleTest {

    @Test
    void headingKeepsWithNext() {
        ParagraphStyle style = ParagraphStyle.heading();
        assertTrue(style.keepWithNext());
        assertEquals(240, style.spaceBeforeTwips());
        assertEquals(120, style.spaceAfterTwips());
    }

    @Test
    void bodyDoesNotKeepWithNext() {
        ParagraphStyle style = ParagraphStyle.body();
        assertEquals(false, style.keepWithNext());
        assertEquals(0, style.spaceBeforeTwips());
        assertEquals(120, style.spaceAfterTwips());
    }

    @Test
    void emitsSpacingAndKeepNext() {
        PPr heading = ParagraphStyle.heading().toPPr();
        assertEquals(BigInteger.valueOf(240), heading.getSpacing().getBefore());
        assertEquals(BigInteger.valueOf(120), heading.getSpacing().getAfter());
        assertNotNull(heading.getKeepNext());
        assertTrue(heading.getKeepNext().isVal());

        PPr body = ParagraphStyle.body().toPPr();
        assertNull(body.getKeepNext(), "keepNext must be absent, not present-and-false");
    }

    @Test
    void rejectsNegativeSpacing() {
        assertThrows(DocumentGenerationException.class,
                () -> ParagraphStyle.builder().spaceBeforeTwips(-1).build());
        assertThrows(DocumentGenerationException.class,
                () -> ParagraphStyle.builder().spaceAfterTwips(-1).build());
    }
}
