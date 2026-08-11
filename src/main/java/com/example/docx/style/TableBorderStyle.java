package com.example.docx.style;

import com.example.docx.DocumentGenerationException;
import com.example.docx.Units;
import java.math.BigInteger;
import java.util.EnumSet;
import java.util.Locale;
import java.util.Set;
import java.util.regex.Pattern;
import org.docx4j.jaxb.Context;
import org.docx4j.wml.CTBorder;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.TcPrInner;

/**
 * Immutable border description for one group of cells: which edges are drawn, in what
 * colour, width and line style.
 *
 * <p>Borders are per cell edge rather than OOXML's {@code insideH}/{@code insideV},
 * because {@code w:tblBorders} is table-wide and cannot give the header row a border
 * different from the body's. At cell level {@code insideH} and {@code insideV} do not
 * mean what their names suggest — they apply to a cell's own internal splits.
 */
public final class TableBorderStyle {

    private static final Pattern HEX = Pattern.compile("[0-9A-Fa-f]{6}");

    private final String colorHex;
    private final double widthPt;
    private final BorderLine line;
    private final Set<Edge> edges;

    private TableBorderStyle(Builder b) {
        this.colorHex = b.colorHex;
        this.widthPt = b.widthPt;
        this.line = b.line;
        this.edges = EnumSet.copyOf(b.edges.isEmpty() ? EnumSet.noneOf(Edge.class) : b.edges);
    }

    public static Builder builder() {
        return new Builder();
    }

    /** No border at all. */
    public static TableBorderStyle none() {
        return builder().build();
    }

    /** A rule along the bottom of each cell — a header underline, or rules between rows. */
    public static TableBorderStyle bottomOnly(String colorHex, double widthPt) {
        return builder().color(colorHex).widthPt(widthPt).edges(Edge.BOTTOM).build();
    }

    /** All four edges of every cell. */
    public static TableBorderStyle grid(String colorHex, double widthPt) {
        return builder().color(colorHex).widthPt(widthPt)
                .edges(Edge.TOP, Edge.BOTTOM, Edge.LEFT, Edge.RIGHT).build();
    }

    /** Left and right edges only. */
    public static TableBorderStyle box(String colorHex, double widthPt) {
        return builder().color(colorHex).widthPt(widthPt).edges(Edge.LEFT, Edge.RIGHT).build();
    }

    public String colorHex() {
        return colorHex;
    }

    public double widthPt() {
        return widthPt;
    }

    public BorderLine line() {
        return line;
    }

    /** A copy — mutating it does not affect this style. */
    public Set<Edge> edges() {
        return edges.isEmpty() ? EnumSet.noneOf(Edge.class) : EnumSet.copyOf(edges);
    }

    public boolean isEmpty() {
        return edges.isEmpty();
    }

    /**
     * The {@code w:tcBorders} for this style, or {@code null} when nothing is drawn.
     *
     * <p>Null rather than an empty element, so the caller omits {@code w:tcBorders}
     * entirely.
     */
    public TcPrInner.TcBorders toTcBorders() {
        if (edges.isEmpty()) {
            return null;
        }
        ObjectFactory factory = Context.getWmlObjectFactory();
        TcPrInner.TcBorders borders = new TcPrInner.TcBorders();
        if (edges.contains(Edge.TOP)) {
            borders.setTop(border(factory));
        }
        if (edges.contains(Edge.BOTTOM)) {
            borders.setBottom(border(factory));
        }
        if (edges.contains(Edge.LEFT)) {
            borders.setLeft(border(factory));
        }
        if (edges.contains(Edge.RIGHT)) {
            borders.setRight(border(factory));
        }
        return borders;
    }

    /** A fresh border each time — never shared, since callers may mutate the graph. */
    private CTBorder border(ObjectFactory factory) {
        CTBorder border = factory.createCTBorder();
        border.setVal(line.stBorder());
        border.setColor(colorHex);
        // w:sz on a border is in eighths of a point — not half-points, not twips.
        border.setSz(BigInteger.valueOf(Units.pointsToEighths(widthPt)));
        border.setSpace(BigInteger.ZERO);
        return border;
    }

    /** Fluent builder. With no edges set, the result draws nothing. */
    public static final class Builder {

        private String colorHex = "000000";
        private double widthPt = 0.5;
        private BorderLine line = BorderLine.SINGLE;
        private final Set<Edge> edges = EnumSet.noneOf(Edge.class);

        private Builder() {
        }

        /** Accepts {@code #RRGGBB} or {@code RRGGBB}, any case. */
        public Builder color(String hex) {
            this.colorHex = normaliseColour(hex);
            return this;
        }

        public Builder widthPt(double widthPt) {
            this.widthPt = widthPt;
            return this;
        }

        public Builder line(BorderLine line) {
            if (line == null) {
                throw new DocumentGenerationException("border line must not be null");
            }
            this.line = line;
            return this;
        }

        /** Which edges to draw. Passing none leaves the border empty. */
        public Builder edges(Edge... edges) {
            if (edges == null) {
                throw new DocumentGenerationException("edges must not be null");
            }
            for (Edge edge : edges) {
                if (edge == null) {
                    throw new DocumentGenerationException("edge must not be null");
                }
                this.edges.add(edge);
            }
            return this;
        }

        public TableBorderStyle build() {
            if (!edges.isEmpty() && !(widthPt > 0)) {
                throw new DocumentGenerationException(
                        "border width must be greater than zero points, got " + widthPt);
            }
            return new TableBorderStyle(this);
        }

        private static String normaliseColour(String hex) {
            if (hex == null) {
                throw new DocumentGenerationException("border colour must not be null");
            }
            String bare = hex.startsWith("#") ? hex.substring(1) : hex;
            if (!HEX.matcher(bare).matches()) {
                throw new DocumentGenerationException(
                        "border colour must be 6 hex digits, optionally prefixed with '#', got '"
                                + hex + "'");
            }
            return bare.toUpperCase(Locale.ROOT);
        }
    }
}
