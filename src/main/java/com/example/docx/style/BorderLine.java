package com.example.docx.style;

import org.docx4j.wml.STBorder;

/**
 * Table border line styles.
 *
 * <p>A curated subset. {@code STBorder} carries roughly 180 values, the large majority
 * of them decorative page-border art, which has no place in a table API.
 *
 * <p>There is deliberately no {@code NONE}: a border that draws nothing is expressed by
 * having no edges, and a second way to say it would be one too many.
 */
public enum BorderLine {

    SINGLE(STBorder.SINGLE),
    THICK(STBorder.THICK),
    DOUBLE(STBorder.DOUBLE),
    DOTTED(STBorder.DOTTED),
    DASHED(STBorder.DASHED),
    DOT_DASH(STBorder.DOT_DASH);

    private final STBorder value;

    BorderLine(STBorder value) {
        this.value = value;
    }

    STBorder stBorder() {
        return value;
    }
}
