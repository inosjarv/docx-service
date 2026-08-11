package com.example.docx.page;

import com.example.docx.DocumentGenerationException;
import com.example.docx.Units;
import java.math.BigInteger;
import org.docx4j.jaxb.Context;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.SectPr;

/**
 * Immutable page geometry: size and the four margins, all in twips.
 *
 * <p>A4's dimensions are literals, not conversions. Deriving 297 mm gives 16838.7,
 * which rounds to 16839 and does not match the value Word writes.
 */
public final class PageSetup {

    /** A4 portrait width in twips, as written by Word. */
    public static final int A4_WIDTH_TWIPS = 11906;

    /** A4 portrait height in twips, as written by Word. */
    public static final int A4_HEIGHT_TWIPS = 16838;

    /** Default margin, ~1.5 cm. Note {@code marginsCm(1.5)} yields 850, not this. */
    public static final int DEFAULT_MARGIN_TWIPS = 851;

    private final int pageWidthTwips;
    private final int pageHeightTwips;
    private final int topTwips;
    private final int rightTwips;
    private final int bottomTwips;
    private final int leftTwips;

    private PageSetup(Builder b) {
        this.pageWidthTwips = b.pageWidthTwips;
        this.pageHeightTwips = b.pageHeightTwips;
        this.topTwips = b.topTwips;
        this.rightTwips = b.rightTwips;
        this.bottomTwips = b.bottomTwips;
        this.leftTwips = b.leftTwips;
    }

    public static Builder builder() {
        return new Builder();
    }

    /** A4 portrait with the default 851-twip margins. */
    public static PageSetup a4() {
        return builder().a4().build();
    }

    public int pageWidthTwips() {
        return pageWidthTwips;
    }

    public int pageHeightTwips() {
        return pageHeightTwips;
    }

    public int topTwips() {
        return topTwips;
    }

    public int rightTwips() {
        return rightTwips;
    }

    public int bottomTwips() {
        return bottomTwips;
    }

    public int leftTwips() {
        return leftTwips;
    }

    /** Page width less both side margins: the width text and images may occupy. */
    public int usableWidthTwips() {
        return pageWidthTwips - leftTwips - rightTwips;
    }

    /** Builds the {@code w:sectPr} describing this page. */
    public SectPr toSectPr() {
        ObjectFactory factory = Context.getWmlObjectFactory();
        SectPr sectPr = factory.createSectPr();

        SectPr.PgSz pgSz = factory.createSectPrPgSz();
        pgSz.setW(BigInteger.valueOf(pageWidthTwips));
        pgSz.setH(BigInteger.valueOf(pageHeightTwips));
        sectPr.setPgSz(pgSz);

        SectPr.PgMar pgMar = factory.createSectPrPgMar();
        pgMar.setTop(BigInteger.valueOf(topTwips));
        pgMar.setRight(BigInteger.valueOf(rightTwips));
        pgMar.setBottom(BigInteger.valueOf(bottomTwips));
        pgMar.setLeft(BigInteger.valueOf(leftTwips));
        sectPr.setPgMar(pgMar);

        return sectPr;
    }

    /** Fluent builder. Every field is defaulted; {@code PageSetup.a4()} is valid alone. */
    public static final class Builder {

        private int pageWidthTwips = A4_WIDTH_TWIPS;
        private int pageHeightTwips = A4_HEIGHT_TWIPS;
        private int topTwips = DEFAULT_MARGIN_TWIPS;
        private int rightTwips = DEFAULT_MARGIN_TWIPS;
        private int bottomTwips = DEFAULT_MARGIN_TWIPS;
        private int leftTwips = DEFAULT_MARGIN_TWIPS;

        private Builder() {
        }

        public Builder a4() {
            this.pageWidthTwips = A4_WIDTH_TWIPS;
            this.pageHeightTwips = A4_HEIGHT_TWIPS;
            return this;
        }

        public Builder pageSizeTwips(int width, int height) {
            this.pageWidthTwips = width;
            this.pageHeightTwips = height;
            return this;
        }

        public Builder marginsTwips(int all) {
            this.topTwips = all;
            this.rightTwips = all;
            this.bottomTwips = all;
            this.leftTwips = all;
            return this;
        }

        public Builder marginsCm(double all) {
            return marginsTwips(Units.cmToTwips(all));
        }

        public Builder marginsInches(double all) {
            return marginsTwips(Units.inchesToTwips(all));
        }

        public Builder top(int twips) {
            this.topTwips = twips;
            return this;
        }

        public Builder right(int twips) {
            this.rightTwips = twips;
            return this;
        }

        public Builder bottom(int twips) {
            this.bottomTwips = twips;
            return this;
        }

        public Builder left(int twips) {
            this.leftTwips = twips;
            return this;
        }

        public PageSetup build() {
            requirePositive(pageWidthTwips, "page width");
            requirePositive(pageHeightTwips, "page height");
            requireNonNegative(topTwips, "top margin");
            requireNonNegative(rightTwips, "right margin");
            requireNonNegative(bottomTwips, "bottom margin");
            requireNonNegative(leftTwips, "left margin");

            if ((long) leftTwips + rightTwips >= pageWidthTwips) {
                throw new DocumentGenerationException(
                        "left + right margins (" + ((long) leftTwips + rightTwips)
                                + " twips) leave no width on a " + pageWidthTwips + "-twip page");
            }
            if ((long) topTwips + bottomTwips >= pageHeightTwips) {
                throw new DocumentGenerationException(
                        "top + bottom margins (" + ((long) topTwips + bottomTwips)
                                + " twips) leave no height on a " + pageHeightTwips + "-twip page");
            }
            return new PageSetup(this);
        }

        private static void requirePositive(int value, String field) {
            if (value <= 0) {
                throw new DocumentGenerationException(
                        field + " must be greater than zero twips, got " + value);
            }
        }

        private static void requireNonNegative(int value, String field) {
            if (value < 0) {
                throw new DocumentGenerationException(
                        field + " must not be negative, got " + value);
            }
        }
    }
}
