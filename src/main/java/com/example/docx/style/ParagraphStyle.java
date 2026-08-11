package com.example.docx.style;

import com.example.docx.DocumentGenerationException;
import java.math.BigInteger;
import org.docx4j.jaxb.Context;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.PPr;
import org.docx4j.wml.PPrBase;

/**
 * Immutable paragraph-level formatting: spacing in twips, and keep-with-next.
 *
 * <p>{@code w:spacing}'s before and after are twentieths of a point, which is the same
 * twip the rest of this library uses. 240 twips is 12 pt; 120 twips is 6 pt.
 */
public final class ParagraphStyle {

    public static final int HEADING_SPACE_BEFORE_TWIPS = 240;
    public static final int HEADING_SPACE_AFTER_TWIPS = 120;
    public static final int BODY_SPACE_AFTER_TWIPS = 120;

    private final int spaceBeforeTwips;
    private final int spaceAfterTwips;
    private final boolean keepWithNext;

    private ParagraphStyle(Builder b) {
        this.spaceBeforeTwips = b.spaceBeforeTwips;
        this.spaceAfterTwips = b.spaceAfterTwips;
        this.keepWithNext = b.keepWithNext;
    }

    public static Builder builder() {
        return new Builder();
    }

    /**
     * Heading spacing, keeping with the paragraph that follows.
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
                .build();
    }

    public static ParagraphStyle body() {
        return builder()
                .spaceBeforeTwips(0)
                .spaceAfterTwips(BODY_SPACE_AFTER_TWIPS)
                .keepWithNext(false)
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

    /** Builds the {@code w:pPr} for this style. */
    public PPr toPPr() {
        ObjectFactory factory = Context.getWmlObjectFactory();
        PPr pPr = factory.createPPr();

        PPrBase.Spacing spacing = factory.createPPrBaseSpacing();
        spacing.setBefore(BigInteger.valueOf(spaceBeforeTwips));
        spacing.setAfter(BigInteger.valueOf(spaceAfterTwips));
        pPr.setSpacing(spacing);

        // Emitted only when true; an explicit false is not the same as absent.
        if (keepWithNext) {
            BooleanDefaultTrue on = factory.createBooleanDefaultTrue();
            on.setVal(Boolean.TRUE);
            pPr.setKeepNext(on);
        }
        return pPr;
    }

    /** Fluent builder. Every field is defaulted; {@code ParagraphStyle.body()} is valid alone. */
    public static final class Builder {

        private int spaceBeforeTwips = 0;
        private int spaceAfterTwips = BODY_SPACE_AFTER_TWIPS;
        private boolean keepWithNext = false;

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

        public Builder keepWithNext(boolean keepWithNext) {
            this.keepWithNext = keepWithNext;
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
