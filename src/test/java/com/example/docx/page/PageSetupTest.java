package com.example.docx.page;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import com.example.docx.DocumentGenerationException;
import java.math.BigInteger;
import org.docx4j.wml.SectPr;
import org.junit.jupiter.api.Test;

class PageSetupTest {

    @Test
    void a4UsesLiteralWordDimensions() {
        PageSetup setup = PageSetup.a4();
        assertEquals(11906, setup.pageWidthTwips());
        assertEquals(16838, setup.pageHeightTwips());
    }

    @Test
    void defaultMarginsAre851Twips() {
        PageSetup setup = PageSetup.a4();
        assertEquals(851, setup.topTwips());
        assertEquals(851, setup.rightTwips());
        assertEquals(851, setup.bottomTwips());
        assertEquals(851, setup.leftTwips());
    }

    @Test
    void marginsCmIsNotTheSameAs851() {
        PageSetup setup = PageSetup.builder().a4().marginsCm(1.5).build();
        assertEquals(850, setup.topTwips());
    }

    @Test
    void perSideMarginsOverrideTheBulkSetter() {
        PageSetup setup = PageSetup.builder().a4().marginsTwips(851).top(1440).build();
        assertEquals(1440, setup.topTwips());
        assertEquals(851, setup.bottomTwips());
    }

    @Test
    void marginsInches() {
        PageSetup setup = PageSetup.builder().a4().marginsInches(1.0).build();
        assertEquals(1440, setup.leftTwips());
    }

    @Test
    void emitsSectPrWithSizeAndMargins() {
        SectPr sectPr = PageSetup.a4().toSectPr();
        assertEquals(BigInteger.valueOf(11906), sectPr.getPgSz().getW());
        assertEquals(BigInteger.valueOf(16838), sectPr.getPgSz().getH());
        assertEquals(BigInteger.valueOf(851), sectPr.getPgMar().getTop());
        assertEquals(BigInteger.valueOf(851), sectPr.getPgMar().getRight());
        assertEquals(BigInteger.valueOf(851), sectPr.getPgMar().getBottom());
        assertEquals(BigInteger.valueOf(851), sectPr.getPgMar().getLeft());
    }

    @Test
    void rejectsNegativeMargin() {
        assertThrows(DocumentGenerationException.class,
                () -> PageSetup.builder().a4().top(-1).build());
    }

    @Test
    void rejectsMarginsWiderThanThePage() {
        assertThrows(DocumentGenerationException.class,
                () -> PageSetup.builder().a4().left(6000).right(6000).build());
    }

    @Test
    void rejectsMarginsTallerThanThePage() {
        assertThrows(DocumentGenerationException.class,
                () -> PageSetup.builder().a4().top(9000).bottom(9000).build());
    }

    @Test
    void rejectsNonPositivePageSize() {
        assertThrows(DocumentGenerationException.class,
                () -> PageSetup.builder().pageSizeTwips(0, 16838).build());
    }
}
