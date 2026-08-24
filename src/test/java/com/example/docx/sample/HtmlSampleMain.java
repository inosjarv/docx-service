package com.example.docx.sample;

import com.example.docx.WordDocument;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 * Writes a timestamped {@code target/sample-*.html} so the HTML export can be
 * eyeballed in a browser.
 *
 * <p>Renders the exact same document as {@link SampleMain} — heading, chart, table,
 * three report sections and a trailing hyperlink — through {@link WordDocument#writeHtmlTo}
 * instead of {@link WordDocument#writeTo}, so the two outputs are directly comparable.
 */
public final class HtmlSampleMain {

    private HtmlSampleMain() {
    }

    public static void main(String[] args) throws IOException {
        WordDocument document = SampleMain.build();

        Path target = Path.of("target", "sample-%d.html".formatted(System.currentTimeMillis()));
        Files.createDirectories(target.getParent());
        try (OutputStream out = Files.newOutputStream(target)) {
            document.writeHtmlTo(out);
        }
        System.out.println("Wrote " + target.toAbsolutePath());
    }
}
