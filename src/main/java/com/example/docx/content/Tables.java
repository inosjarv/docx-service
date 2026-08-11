package com.example.docx.content;

import com.example.docx.DocumentGenerationException;
import com.example.docx.style.ParagraphStyle;
import com.example.docx.style.TableBorderStyle;
import com.example.docx.style.TableStyle;
import com.example.docx.style.TextStyle;
import java.math.BigInteger;
import java.util.List;
import org.docx4j.jaxb.Context;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.CTTblCellMar;
import org.docx4j.wml.CTTblLayoutType;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.STTblLayoutType;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.TblGrid;
import org.docx4j.wml.TblGridCol;
import org.docx4j.wml.TblPr;
import org.docx4j.wml.TblWidth;
import org.docx4j.wml.Tc;
import org.docx4j.wml.TcPr;
import org.docx4j.wml.TcPrInner;
import org.docx4j.wml.Tr;
import org.docx4j.wml.TrPr;

/**
 * Builds {@code w:tbl} tables. A pure function of its arguments — no package needed,
 * because column widths arrive as an {@code int}.
 *
 * <p>Two OOXML rules are load-bearing:
 * <ol>
 *   <li>{@code w:tblGrid} is mandatory, with a fixed layout. Without both, Word
 *       auto-fits to content and the same table renders differently everywhere.</li>
 *   <li>A cell must contain at least one paragraph. An empty {@code w:tc} makes the
 *       document unopenable — not misrendered, unopenable.</li>
 * </ol>
 */
public final class Tables {

    /** Left and right cell padding, in twips. */
    public static final int CELL_PADDING_TWIPS = 80;

    /** {@code dxa} is twentieths of a point, i.e. the twips used everywhere else. */
    private static final String TWIPS = "dxa";

    /** Body spacing looks wrong inside a cell, so cell paragraphs are compact. */
    private static final ParagraphStyle CELL_PARAGRAPH = ParagraphStyle.builder()
            .spaceBeforeTwips(0)
            .spaceAfterTwips(0)
            .build();

    private Tables() {
    }

    /** A table whose columns share {@code widthTwips} equally. */
    public static Tbl of(List<String> headers, List<List<String>> rows,
                         TableStyle style, int widthTwips) {
        validate(headers, rows, style, widthTwips);

        ObjectFactory factory = Context.getWmlObjectFactory();
        int[] columnWidths = columnWidths(widthTwips, headers.size());

        Tbl table = factory.createTbl();
        table.setTblPr(tableProperties(factory, widthTwips));
        table.setTblGrid(grid(factory, columnWidths));

        table.getContent().add(row(factory, headers, style.headerText(),
                style.headerBorder(), columnWidths, true));
        for (List<String> values : rows) {
            table.getContent().add(row(factory, values, style.bodyText(),
                    style.bodyBorder(), columnWidths, false));
        }
        return table;
    }

    private static void validate(List<String> headers, List<List<String>> rows,
                                 TableStyle style, int widthTwips) {
        if (headers == null || headers.isEmpty()) {
            throw new DocumentGenerationException("a table needs at least one header column");
        }
        if (rows == null) {
            throw new DocumentGenerationException("table rows must not be null");
        }
        if (style == null) {
            throw new DocumentGenerationException("table style must not be null");
        }
        if (widthTwips <= 0) {
            throw new DocumentGenerationException(
                    "table width must be greater than zero twips, got " + widthTwips);
        }
        for (int i = 0; i < rows.size(); i++) {
            List<String> row = rows.get(i);
            if (row == null) {
                throw new DocumentGenerationException("row " + i + " must not be null");
            }
            if (row.size() != headers.size()) {
                throw new DocumentGenerationException("row " + i + " has " + row.size()
                        + " cells but the table has " + headers.size() + " columns");
            }
        }
    }

    /** Even split, with the rounding remainder given to the last column so the sum is exact. */
    private static int[] columnWidths(int widthTwips, int columns) {
        int[] widths = new int[columns];
        int each = widthTwips / columns;
        for (int i = 0; i < columns; i++) {
            widths[i] = each;
        }
        widths[columns - 1] += widthTwips - each * columns;
        return widths;
    }

    private static TblPr tableProperties(ObjectFactory factory, int widthTwips) {
        TblPr properties = factory.createTblPr();
        properties.setTblW(width(factory, widthTwips));

        // Fixed layout makes Word honour the grid instead of auto-fitting to content.
        CTTblLayoutType layout = factory.createCTTblLayoutType();
        layout.setType(STTblLayoutType.FIXED);
        properties.setTblLayout(layout);

        properties.setTblCellMar(cellMargins(factory));
        return properties;
    }

    private static CTTblCellMar cellMargins(ObjectFactory factory) {
        CTTblCellMar margins = factory.createCTTblCellMar();
        margins.setLeft(width(factory, CELL_PADDING_TWIPS));
        margins.setRight(width(factory, CELL_PADDING_TWIPS));
        margins.setTop(width(factory, 0));
        margins.setBottom(width(factory, 0));
        return margins;
    }

    private static TblWidth width(ObjectFactory factory, int twips) {
        TblWidth width = factory.createTblWidth();
        width.setType(TWIPS);
        width.setW(BigInteger.valueOf(twips));
        return width;
    }

    private static TblGrid grid(ObjectFactory factory, int[] columnWidths) {
        TblGrid grid = factory.createTblGrid();
        for (int columnWidth : columnWidths) {
            TblGridCol column = factory.createTblGridCol();
            column.setW(BigInteger.valueOf(columnWidth));
            grid.getGridCol().add(column);
        }
        return grid;
    }

    private static Tr row(ObjectFactory factory, List<String> values, TextStyle textStyle,
                          TableBorderStyle border, int[] columnWidths, boolean header) {
        Tr row = factory.createTr();
        if (header) {
            // Repeats the header at the top of every page the table spans.
            TrPr properties = factory.createTrPr();
            BooleanDefaultTrue on = factory.createBooleanDefaultTrue();
            on.setVal(Boolean.TRUE);
            properties.getCnfStyleOrDivIdOrGridBefore().add(factory.createCTTrPrBaseTblHeader(on));
            row.setTrPr(properties);
        }
        for (int i = 0; i < values.size(); i++) {
            row.getContent().add(cell(factory, values.get(i), textStyle, border, columnWidths[i]));
        }
        return row;
    }

    private static Tc cell(ObjectFactory factory, String text, TextStyle textStyle,
                           TableBorderStyle border, int widthTwips) {
        Tc cell = factory.createTc();

        TcPr properties = factory.createTcPr();
        properties.setTcW(width(factory, widthTwips));

        TcPrInner.TcBorders borders = border.toTcBorders();
        if (borders != null) {
            properties.setTcBorders(borders);
        }
        cell.setTcPr(properties);

        // A cell with no paragraph makes the document unopenable, so a blank cell still
        // gets one. The guard below mirrors Paragraphs.of's blank rejection: if the text
        // is blank (empty or whitespace-only), build the empty paragraph directly rather
        // than routing to Paragraphs.of, which would throw DocumentGenerationException.
        String value = text == null ? "" : text;
        cell.getContent().add(value.isBlank()
                ? emptyParagraph(factory)
                : Paragraphs.of(value, textStyle, CELL_PARAGRAPH));
        return cell;
    }

    private static P emptyParagraph(ObjectFactory factory) {
        P paragraph = factory.createP();
        paragraph.setPPr(CELL_PARAGRAPH.toPPr());
        return paragraph;
    }
}
