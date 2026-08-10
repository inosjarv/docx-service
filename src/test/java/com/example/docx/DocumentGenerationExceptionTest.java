package com.example.docx;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertSame;

import org.junit.jupiter.api.Test;

class DocumentGenerationExceptionTest {

    @Test
    void validationFailureCarriesMessageAndNoCause() {
        DocumentGenerationException e = new DocumentGenerationException("colour must be RRGGBB");
        assertEquals("colour must be RRGGBB", e.getMessage());
        assertNull(e.getCause());
    }

    @Test
    void wrappedFailureCarriesCause() {
        Throwable cause = new IllegalStateException("boom");
        DocumentGenerationException e = new DocumentGenerationException("save failed", cause);
        assertSame(cause, e.getCause());
    }

    @Test
    void isUnchecked() {
        assertEquals(true, RuntimeException.class.isAssignableFrom(DocumentGenerationException.class));
    }
}
