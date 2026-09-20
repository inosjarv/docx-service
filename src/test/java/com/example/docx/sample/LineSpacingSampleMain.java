package com.example.docx.sample;

import com.example.docx.WordDocument;
import com.example.docx.style.LineSpacing;
import com.example.docx.style.ParagraphStyle;
import com.example.docx.style.TextStyle;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 * Writes a timestamped {@code target/line-spacing-*.docx} so {@link LineSpacing} can be
 * eyeballed.
 *
 * <p>The same three-line paragraph repeated under five headings, each paragraph carrying
 * a different {@link ParagraphStyle#lineSpacing()} — everything else held constant so the
 * only thing that changes from one block to the next is line height.
 */
public final class LineSpacingSampleMain {

    /** Long enough to wrap onto three lines at A4 width, so line spacing is visible. */
    private static final String SAMPLE_PARAGRAPH =
            "Line spacing controls the gap between the lines inside a single paragraph, "
                    + "which is a different knob from the space before or after the paragraph "
                    + "itself. This block wraps onto three lines on its own so the difference "
                    + "between settings is visible without needing several paragraphs.";

    private LineSpacingSampleMain() {
    }

    public static void main(String[] args) throws IOException {
        WordDocument document = build();

        Path target = Path.of("target", "line-spacing-%d.docx".formatted(System.currentTimeMillis()));
        Files.createDirectories(target.getParent());
        try (OutputStream out = Files.newOutputStream(target)) {
            document.writeTo(out);
        }
        System.out.println("Wrote " + target.toAbsolutePath());
    }

    /**
     * Builds the same document {@link #main} writes to disk, so other tools (such as
     * {@code LineSpacingHtmlSampleMain} and the document gallery) can render it without
     * duplicating the content.
     */
    public static WordDocument build() {
        var builder = WordDocument.builder().heading("Line spacing");

        addExample(builder, "Default (Word's own single spacing)", null);
        addExample(builder, "Multiple: 1.15", LineSpacing.multiple(1.15));
        addExample(builder, "Multiple: 1.5", LineSpacing.multiple(1.5));
        addExample(builder, "Multiple: 2.0 (double)", LineSpacing.multiple(2.0));
        addExample(builder, "Exact: 20pt", LineSpacing.exactPt(20));

        return builder.build();
    }

    /** One labelled block: a small heading naming the setting, then the sample paragraph. */
    private static void addExample(WordDocument.Builder builder, String label, LineSpacing lineSpacing) {
        builder.heading(label, TextStyle.builder().font("Calibri").sizePt(12).bold(true).build());

        ParagraphStyle.Builder paragraphStyle = ParagraphStyle.builder()
                .spaceAfterTwips(ParagraphStyle.BODY_SPACE_AFTER_TWIPS);
        if (lineSpacing != null) {
            paragraphStyle.lineSpacing(lineSpacing);
        }

        builder.paragraph(SAMPLE_PARAGRAPH, TextStyle.body(), paragraphStyle.build());
    }
}
