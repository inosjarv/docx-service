package com.example.docx;

/**
 * The only exception this module throws deliberately.
 *
 * <p>Validation failures carry a message naming the offending field and no cause.
 * Failures wrapped from docx4j always carry a cause.
 */
public class DocumentGenerationException extends RuntimeException {

    private static final long serialVersionUID = 1L;

    public DocumentGenerationException(String message) {
        super(message);
    }

    public DocumentGenerationException(String message, Throwable cause) {
        super(message, cause);
    }
}
