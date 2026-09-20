package com.example.docx.style;

import com.example.docx.DocumentGenerationException;
import com.example.docx.Units;
import org.docx4j.wml.STLineSpacingRule;

/**
 * How far apart the lines within a paragraph sit -- distinct from
 * {@link ParagraphStyle}'s {@code spaceBefore}/{@code spaceAfter}, which space
 * paragraphs from each other.
 *
 * <p>{@code multiple} tracks Word's own "single / 1.5 lines / double / Multiple" option:
 * the gap scales with the run's font size. {@code exactPt} instead pins an exact leading
 * in points, the same regardless of font size -- the right tool for matching a fixed
 * design spec rather than a proportion.
 */
public final class LineSpacing {

    private final int lineTwips;
    private final STLineSpacingRule rule;

    private LineSpacing(int lineTwips, STLineSpacingRule rule) {
        this.lineTwips = lineTwips;
        this.rule = rule;
    }

    /** E.g. 1.0 single, 1.15, 1.5, 2.0 double. Scales with the run's font size. */
    public static LineSpacing multiple(double multiple) {
        if (!Double.isFinite(multiple) || multiple <= 0) {
            throw new DocumentGenerationException(
                    "line spacing multiple must be positive, got " + multiple);
        }
        // w:line is in twentieths of a single line; 240 is Word's single spacing.
        int twips = (int) Math.round(multiple * 240.0);
        return new LineSpacing(twips, STLineSpacingRule.AUTO);
    }

    /** A fixed leading in points, independent of the run's font size. */
    public static LineSpacing exactPt(double points) {
        return new LineSpacing(Units.pointsToTwips(points), STLineSpacingRule.EXACT);
    }

    int lineTwips() {
        return lineTwips;
    }

    STLineSpacingRule rule() {
        return rule;
    }
}
