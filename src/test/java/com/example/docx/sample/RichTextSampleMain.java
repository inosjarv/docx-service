package com.example.docx.sample;

import com.example.docx.WordDocument;
import com.example.docx.content.RichText;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 * Writes a timestamped {@code target/rich-text-*.docx} so mixed inline formatting can be
 * eyeballed.
 *
 * <p>A small, standalone document — plain text, bold, italic, underline, and nested
 * formatting — rather than the full report {@link SampleMain} builds. The feature itself
 * is already covered by {@code RichTextTest}/{@code RichTextDocumentTest}; this exists
 * purely so the visual result has a sample of its own.
 */
public final class RichTextSampleMain {

    private RichTextSampleMain() {
    }

    public static void main(String[] args) throws IOException {
        WordDocument document = build();

        Path target = Path.of("target", "rich-text-%d.docx".formatted(System.currentTimeMillis()));
        Files.createDirectories(target.getParent());
        try (OutputStream out = Files.newOutputStream(target)) {
            document.writeTo(out);
        }
        System.out.println("Wrote " + target.toAbsolutePath());
    }

    /**
     * Builds the same document {@link #main} writes to disk, so other tools (such as the
     * document gallery) can render it without duplicating the content.
     */
    public static WordDocument build() {
        return WordDocument.builder()
                .heading("Rich text")
                .paragraph(RichText.of("Some text which needs to be <b>bold</b>."))
                .paragraph(RichText.of("Some text which needs to be <i>italic</i>."))
                .paragraph(RichText.of("Some text which needs to be <u>underlined</u>."))
                .paragraph(RichText.of(
                        "Tags nest: <b>bold, <i>bold and italic</i>, back to just bold</b>."))
                .paragraph(RichText.of(
                        "Aliases work the same way: <strong>strong</strong> and <em>em</em>."))
                .build();
    }
}
