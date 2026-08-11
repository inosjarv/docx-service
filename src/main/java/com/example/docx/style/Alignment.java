package com.example.docx.style;

import org.docx4j.wml.JcEnumeration;

/** Horizontal paragraph alignment. */
public enum Alignment {

    /** Word's own default. Emits no {@code w:jc} at all. */
    LEFT(null),
    CENTER(JcEnumeration.CENTER),
    RIGHT(JcEnumeration.RIGHT),
    /** Flush to both margins. OOXML calls this {@code both}, not {@code justify}. */
    JUSTIFY(JcEnumeration.BOTH);

    private final JcEnumeration value;

    Alignment(JcEnumeration value) {
        this.value = value;
    }

    /** The OOXML value, or {@code null} when nothing should be emitted. */
    JcEnumeration jcValue() {
        return value;
    }
}
