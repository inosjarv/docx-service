package com.example.docx.style;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.example.docx.DocumentGenerationException;
import java.util.EnumSet;
import org.docx4j.wml.TcPrInner;
import org.junit.jupiter.api.Test;

class TableBorderStyleTest {

    @Test
    void noneDrawsNothing() {
        TableBorderStyle none = TableBorderStyle.none();
        assertTrue(none.isEmpty());
        assertNull(none.toTcBorders(), "no edges means no w:tcBorders element at all");
    }

    @Test
    void bottomOnlySetsJustTheBottomEdge() {
        TableBorderStyle style = TableBorderStyle.bottomOnly("#1F4E79", 1.0);
        assertEquals(EnumSet.of(Edge.BOTTOM), style.edges());

        TcPrInner.TcBorders borders = style.toTcBorders();
        assertNotNull(borders.getBottom());
        assertNull(borders.getTop());
        assertNull(borders.getLeft());
        assertNull(borders.getRight());
    }

    @Test
    void gridSetsAllFourEdges() {
        assertEquals(EnumSet.allOf(Edge.class), TableBorderStyle.grid("#000000", 0.5).edges());
    }

    @Test
    void boxSetsLeftAndRight() {
        assertEquals(EnumSet.of(Edge.LEFT, Edge.RIGHT),
                TableBorderStyle.box("#000000", 0.5).edges());
    }

    @Test
    void colourIsNormalisedToBareUppercase() {
        assertEquals("1F4E79", TableBorderStyle.bottomOnly("#1f4e79", 1.0).colorHex());
        assertEquals("ABCDEF", TableBorderStyle.bottomOnly("abcdef", 1.0).colorHex());
    }

    @Test
    void defaultLineIsSingle() {
        assertEquals(BorderLine.SINGLE, TableBorderStyle.bottomOnly("#000000", 1.0).line());
        assertEquals("single", TableBorderStyle.bottomOnly("#000000", 1.0)
                .toTcBorders().getBottom().getVal().value());
    }

    @Test
    void edgesAreDefensivelyCopied() {
        TableBorderStyle style = TableBorderStyle.grid("#000000", 1.0);
        style.edges().clear();
        assertEquals(4, style.edges().size(),
                "mutating the returned set must not affect the style");
    }

    @Test
    void rejectsBadInput() {
        assertThrows(DocumentGenerationException.class,
                () -> TableBorderStyle.builder().color("blue").build());
        assertThrows(DocumentGenerationException.class,
                () -> TableBorderStyle.builder().color(null).build());
        assertThrows(DocumentGenerationException.class,
                () -> TableBorderStyle.builder().line(null).build());
        // Width only matters when something is actually drawn.
        assertThrows(DocumentGenerationException.class,
                () -> TableBorderStyle.builder().edges(Edge.BOTTOM).widthPt(0).build());
    }
}
