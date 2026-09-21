package com.example.docx.style;

import org.docx4j.wml.STTabJc;

/**
 * How text lines up against a {@link TabStop}'s position.
 *
 * <p>A curated subset of {@code STTabJc}: the four kinds someone actually reaches for
 * (left/center/right/decimal), not the full set (bar tabs, list-numbering tabs) OOXML
 * also defines.
 */
public enum TabStopAlignment {

    LEFT(STTabJc.LEFT),
    CENTER(STTabJc.CENTER),
    RIGHT(STTabJc.RIGHT),
    /** Numbers line up on their decimal point. */
    DECIMAL(STTabJc.DECIMAL);

    private final STTabJc value;

    TabStopAlignment(STTabJc value) {
        this.value = value;
    }

    STTabJc stTabJc() {
        return value;
    }
}
