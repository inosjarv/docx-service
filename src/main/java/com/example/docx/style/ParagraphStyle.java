package com.example.docx.style;

import com.example.docx.DocumentGenerationException;
import java.math.BigInteger;
import org.docx4j.jaxb.Context;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.Jc;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.PPr;
import org.docx4j.wml.PPrBase;

/**
 * Immutable paragraph-level formatting: spacing, alignment, and how the paragraph
 * behaves at a page boundary.
 *
 * <p>{@code w:spacing}'s before and after are twentieths of a point, which is the same
 * twip the rest of this library uses. 240 twips is 12 pt; 120 twips is 6 pt.
 *
 * <p>Three separate controls govern page flow, and they do different jobs:
 * <ul>
 *   <li>{@code keepWithNext} glues this paragraph to the one after it. Headings use it
 *       so they never sit alone at the foot of a page.</li>
 *   <li>{@code keepLines} keeps every line of this paragraph on one page, pushing the
 *       whole paragraph over rather than splitting it.</li>
 *   <li>{@code widowControl} allows a split but forbids stranding a single line at the
 *       top or bottom of a page.</li>
 * </ul>
 *
 * <p>For body text, {@code widowControl} is the right default and {@code keepLines} is
 * usually wrong: forcing whole paragraphs to move leaves large gaps at page ends.
 *
 * <p><strong>None of these is a guarantee.</strong> A paragraph taller than the text
 * area is broken by Word regardless of {@code keepLines}, because the alternative is
 * losing text.
 */
public final class ParagraphStyle {

    public static final int HEADING_SPACE_BEFORE_TWIPS = 240;
    public static final int HEADING_SPACE_AFTER_TWIPS = 120;
    public static final int BODY_SPACE_AFTER_TWIPS = 120;

    private final int spaceBeforeTwips;
    private final int spaceAfterTwips;
    private final boolean keepWithNext;
    private final boolean keepLines;
    private final boolean widowControl;
    private final boolean pageBreakBefore;
    private final Alignment alignment;

    private ParagraphStyle(Builder b) {
        this.spaceBeforeTwips = b.spaceBeforeTwips;
        this.spaceAfterTwips = b.spaceAfterTwips;
        this.keepWithNext = b.keepWithNext;
        this.keepLines = b.keepLines;
        this.widowControl = b.widowControl;
        this.pageBreakBefore = b.pageBreakBefore;
        this.alignment = b.alignment;
    }

    public static Builder builder() {
        return new Builder();
    }

    /**
     * Heading spacing, kept with the paragraph that follows and never split itself.
     *
     * <p>Without {@code keepWithNext} a heading can land at the foot of a page with its
     * body overleaf — which stops being a risk and becomes a certainty once a document
     * has many sections.
     */
    public static ParagraphStyle heading() {
        return builder()
                .spaceBeforeTwips(HEADING_SPACE_BEFORE_TWIPS)
                .spaceAfterTwips(HEADING_SPACE_AFTER_TWIPS)
                .keepWithNext(true)
                .keepLines(true)
                .widowControl(true)
                .build();
    }

    /**
     * Body spacing. Splits across pages when it must, but never leaves a single line
     * stranded.
     */
    public static ParagraphStyle body() {
        return builder()
                .spaceBeforeTwips(0)
                .spaceAfterTwips(BODY_SPACE_AFTER_TWIPS)
                .keepWithNext(false)
                .keepLines(false)
                .widowControl(true)
                .build();
    }

    public int spaceBeforeTwips() {
        return spaceBeforeTwips;
    }

    public int spaceAfterTwips() {
        return spaceAfterTwips;
    }

    public boolean keepWithNext() {
        return keepWithNext;
    }

    public boolean keepLines() {
        return keepLines;
    }

    public boolean widowControl() {
        return widowControl;
    }

    public boolean pageBreakBefore() {
        return pageBreakBefore;
    }

    public Alignment alignment() {
        return alignment;
    }

    /** Builds the {@code w:pPr} for this style. */
    public PPr toPPr() {
        ObjectFactory factory = Context.getWmlObjectFactory();
        PPr pPr = factory.createPPr();

        PPrBase.Spacing spacing = factory.createPPrBaseSpacing();
        spacing.setBefore(BigInteger.valueOf(spaceBeforeTwips));
        spacing.setAfter(BigInteger.valueOf(spaceAfterTwips));
        pPr.setSpacing(spacing);

        // Left is Word's own default, so LEFT emits nothing rather than an explicit
        // w:jc — same principle as the toggles below.
        if (alignment.jcValue() != null) {
            Jc jc = factory.createJc();
            jc.setVal(alignment.jcValue());
            pPr.setJc(jc);
        }

        // Each is emitted only when true; an explicit false is not the same as absent.
        if (keepWithNext) {
            pPr.setKeepNext(on(factory));
        }
        if (keepLines) {
            pPr.setKeepLines(on(factory));
        }
        if (widowControl) {
            pPr.setWidowControl(on(factory));
        }
        if (pageBreakBefore) {
            pPr.setPageBreakBefore(on(factory));
        }
        return pPr;
    }

    /** A fresh {@code w:val="true"} toggle. Never shared — callers may mutate the graph. */
    private static BooleanDefaultTrue on(ObjectFactory factory) {
        BooleanDefaultTrue flag = factory.createBooleanDefaultTrue();
        flag.setVal(Boolean.TRUE);
        return flag;
    }

    /** Fluent builder. Every field is defaulted; {@code ParagraphStyle.body()} is valid alone. */
    public static final class Builder {

        private int spaceBeforeTwips = 0;
        private int spaceAfterTwips = BODY_SPACE_AFTER_TWIPS;
        private boolean keepWithNext = false;
        private boolean keepLines = false;
        private boolean widowControl = true;
        private boolean pageBreakBefore = false;
        private Alignment alignment = Alignment.LEFT;

        private Builder() {
        }

        public Builder spaceBeforeTwips(int twips) {
            this.spaceBeforeTwips = twips;
            return this;
        }

        public Builder spaceAfterTwips(int twips) {
            this.spaceAfterTwips = twips;
            return this;
        }

        /** Glue this paragraph to the one after it. */
        public Builder keepWithNext(boolean keepWithNext) {
            this.keepWithNext = keepWithNext;
            return this;
        }

        /**
         * Keep every line of this paragraph on one page.
         *
         * <p>Usually wrong for body text — it leaves gaps at page ends. Prefer
         * {@code widowControl}. Cannot help a paragraph taller than the text area.
         */
        public Builder keepLines(boolean keepLines) {
            this.keepLines = keepLines;
            return this;
        }

        /** Forbid a single line stranded at the top or bottom of a page. On by default. */
        public Builder widowControl(boolean widowControl) {
            this.widowControl = widowControl;
            return this;
        }

        /** Start this paragraph on a new page. */
        public Builder pageBreakBefore(boolean pageBreakBefore) {
            this.pageBreakBefore = pageBreakBefore;
            return this;
        }

        public Builder alignment(Alignment alignment) {
            if (alignment == null) {
                throw new DocumentGenerationException("alignment must not be null");
            }
            this.alignment = alignment;
            return this;
        }

        public ParagraphStyle build() {
            requireNonNegative(spaceBeforeTwips, "space before");
            requireNonNegative(spaceAfterTwips, "space after");
            return new ParagraphStyle(this);
        }

        private static void requireNonNegative(int value, String field) {
            if (value < 0) {
                throw new DocumentGenerationException(
                        field + " must not be negative, got " + value);
            }
        }
    }
}
