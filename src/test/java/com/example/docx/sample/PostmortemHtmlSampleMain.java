package com.example.docx.sample;

import com.example.docx.WordDocument;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 * Writes a timestamped {@code target/postmortem-*.html} so the HTML export can be
 * eyeballed on a bigger, multi-page document.
 *
 * <p>Renders the exact same document as {@link PostmortemSampleMain} through
 * {@link WordDocument#writeHtmlTo} instead of {@link WordDocument#writeTo}, so the two
 * outputs are directly comparable.
 */
public final class PostmortemHtmlSampleMain {

    private PostmortemHtmlSampleMain() {
    }

    public static void main(String[] args) throws IOException {
        WordDocument document = PostmortemSampleMain.build();

        Path target = Path.of("target", "postmortem-%d.html".formatted(System.currentTimeMillis()));
        Files.createDirectories(target.getParent());
        try (OutputStream out = Files.newOutputStream(target)) {
            document.writeHtmlTo(out);
        }
        System.out.println("Wrote " + target.toAbsolutePath());
    }
}
