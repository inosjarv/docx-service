package com.example.docx.sample;

import com.example.docx.WordDocument;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 * Writes a timestamped {@code target/line-spacing-*.html} so the HTML export can be
 * eyeballed in a browser.
 *
 * <p>Renders the exact same document as {@link LineSpacingSampleMain} through
 * {@link WordDocument#writeHtmlTo} instead of {@link WordDocument#writeTo}, so the two
 * outputs are directly comparable.
 */
public final class LineSpacingHtmlSampleMain {

    private LineSpacingHtmlSampleMain() {
    }

    public static void main(String[] args) throws IOException {
        WordDocument document = LineSpacingSampleMain.build();

        Path target = Path.of("target", "line-spacing-%d.html".formatted(System.currentTimeMillis()));
        Files.createDirectories(target.getParent());
        try (OutputStream out = Files.newOutputStream(target)) {
            document.writeHtmlTo(out);
        }
        System.out.println("Wrote " + target.toAbsolutePath());
    }
}
