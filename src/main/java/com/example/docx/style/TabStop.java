package com.example.docx.style;

import com.example.docx.DocumentGenerationException;

/**
 * One custom tab stop: the position a {@code <w:tab/>} in this paragraph jumps to, and
 * how the text after it lines up against that position.
 *
 * <p>Word's own default tab stops sit every half inch. Set these when text needs to
 * land at an exact column instead -- e.g. right-aligning a value against a label.
 *
 * <p>Only takes effect in the {@code .docx} output. This library's HTML export renders
 * every {@code <w:tab/>} as three fixed spaces regardless of position or alignment --
 * a limitation in docx4j's own HTML conversion, not something configurable here.
 */
public final class TabStop {

    private final int positionTwips;
    private final TabStopAlignment alignment;

    private TabStop(int positionTwips, TabStopAlignment alignment) {
        this.positionTwips = positionTwips;
        this.alignment = alignment;
    }

    public static TabStop at(int positionTwips, TabStopAlignment alignment) {
        if (positionTwips < 0) {
            throw new DocumentGenerationException(
                    "tab stop position must not be negative, got " + positionTwips);
        }
        if (alignment == null) {
            throw new DocumentGenerationException("tab stop alignment must not be null");
        }
        return new TabStop(positionTwips, alignment);
    }

    int positionTwips() {
        return positionTwips;
    }

    TabStopAlignment alignment() {
        return alignment;
    }
}
