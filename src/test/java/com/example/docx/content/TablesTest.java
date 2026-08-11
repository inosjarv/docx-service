package com.example.docx.content;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.example.docx.DocumentGenerationException;
import com.example.docx.style.BorderLine;
import com.example.docx.style.Edge;
import com.example.docx.style.TableBorderStyle;
import com.example.docx.style.TableStyle;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import org.docx4j.wml.P;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Tr;
import org.junit.jupiter.api.Test;

class TablesTest {

    private static final List<String> HEADERS = List.of("Region", "Revenue", "Change");
    private static final List<List<String>> ROWS = List.of(
            List.of("EMEA", "1 240", "+8%"),
            List.of("APAC", "980", "+21%"));

    private static final TableStyle STYLE = TableStyle.builder()
            .headerBorder(TableBorderStyle.builder()
                    .color("#1F4E79").widthPt(1.0).line(BorderLine.SINGLE)
                    .edges(Edge.BOTTOM).build())
            .bodyBorder(TableBorderStyle.builder()
                    .color("#BFBFBF").widthPt(0.5).line(BorderLine.DOTTED)
                    .edges(Edge.BOTTOM).build())
            .build();

    /** Freshly built rows and cells are bare, not JAXBElement-wrapped. */
    private static Tr rowAt(Tbl table, int index) {
        return (Tr) table.getContent().get(index);
    }

    private static Tc cellAt(Tr row, int index) {
        return (Tc) row.getContent().get(index);
    }

    @Test
    void gridHasOneColumnPerHeaderAndWidthsSumExactly() {
        Tbl table = Tables.of(HEADERS, ROWS, STYLE, 10204);
        var columns = table.getTblGrid().getGridCol();
        assertEquals(3, columns.size());
        int sum = columns.stream().mapToInt(c -> c.getW().intValue()).sum();
        // The rounding remainder goes to the last column so the sum is exact.
        assertEquals(10204, sum, "grid must sum to the requested table width");
    }

    @Test
    void everyCellContainsAParagraphIncludingBlankAndNullOnes() {
        List<List<String>> withBlanks = new ArrayList<>();
        withBlanks.add(Arrays.asList("EMEA", "", null));
        Tbl table = Tables.of(HEADERS, withBlanks, STYLE, 10204);

        for (Object rowObject : table.getContent()) {
            Tr row = (Tr) rowObject;
            for (Object cellObject : row.getContent()) {
                Tc cell = (Tc) cellObject;
                assertTrue(cell.getContent().stream().anyMatch(o -> o instanceof P),
                        "an empty w:tc makes the document unopenable");
            }
        }
    }

    @Test
    void headerAndBodyCarryTheirOwnBorders() {
        Tbl table = Tables.of(HEADERS, ROWS, STYLE, 10204);

        var header = cellAt(rowAt(table, 0), 0).getTcPr().getTcBorders();
        assertNotNull(header.getBottom());
        assertEquals("1F4E79", header.getBottom().getColor());
        assertEquals(BigInteger.valueOf(8), header.getBottom().getSz(), "1.0 pt is 8 eighths");
        assertNull(header.getTop(), "only BOTTOM was requested");
        assertNull(header.getLeft());
        assertNull(header.getRight());

        var body = cellAt(rowAt(table, 1), 0).getTcPr().getTcBorders();
        assertEquals("BFBFBF", body.getBottom().getColor());
        assertEquals(BigInteger.valueOf(4), body.getBottom().getSz(), "0.5 pt is 4 eighths");
        assertEquals("dotted", body.getBottom().getVal().value());
    }

    @Test
    void headerRowRepeatsAcrossPagesAndBodyRowsDoNot() {
        Tbl table = Tables.of(HEADERS, ROWS, STYLE, 10204);

        assertNotNull(rowAt(table, 0).getTrPr(), "header needs trPr");
        boolean repeats = rowAt(table, 0).getTrPr().getCnfStyleOrDivIdOrGridBefore().stream()
                .anyMatch(e -> e.getName().getLocalPart().equals("tblHeader"));
        assertTrue(repeats, "header row must carry w:tblHeader");

        assertNull(rowAt(table, 1).getTrPr(), "body rows must not repeat");
    }

    @Test
    void noEdgesMeansNoTcBordersElement() {
        TableStyle bare = TableStyle.builder()
                .headerBorder(TableBorderStyle.none())
                .bodyBorder(TableBorderStyle.none())
                .build();
        Tbl table = Tables.of(HEADERS, ROWS, bare, 10204);
        assertNull(cellAt(rowAt(table, 0), 0).getTcPr().getTcBorders());
    }

    @Test
    void rejectsMismatchedRowLengthNamingTheRow() {
        List<List<String>> bad = List.of(List.of("EMEA", "1 240"));
        DocumentGenerationException e = assertThrows(DocumentGenerationException.class,
                () -> Tables.of(HEADERS, bad, STYLE, 10204));
        assertTrue(e.getMessage().contains("row 0"),
                "the message must name the offending row, got: " + e.getMessage());
    }

    @Test
    void rejectsEmptyHeadersNullRowsAndBadWidth() {
        assertThrows(DocumentGenerationException.class,
                () -> Tables.of(List.of(), ROWS, STYLE, 10204));
        assertThrows(DocumentGenerationException.class,
                () -> Tables.of(HEADERS, null, STYLE, 10204));
        assertThrows(DocumentGenerationException.class,
                () -> Tables.of(HEADERS, ROWS, STYLE, 0));
        assertThrows(DocumentGenerationException.class,
                () -> Tables.of(HEADERS, ROWS, null, 10204));
    }
}
