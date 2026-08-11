package com.example.docx.sample;

import com.example.docx.WordDocument;
import com.example.docx.page.PageSetup;
import com.example.docx.style.TextStyle;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UncheckedIOException;
import java.nio.file.Files;
import java.nio.file.Path;

/** Writes {@code target/sample.docx} so the output can be opened and eyeballed. */
public final class SampleMain {

    private SampleMain() {
    }

    public static void main(String[] args) throws IOException {
        TextStyle sectionHeading = TextStyle.builder()
                .font("Calibri")
                .sizePt(11)
                .bold(true)
                .color("#1F4E79")
                .build();

        var builder = WordDocument.builder()
                .pageSetup(PageSetup.builder().a4().marginsTwips(851).build())
                .heading("Quarterly Report", TextStyle.builder()
                        .font("Calibri Light")
                        .sizePt(20)
                        .bold(true)
                        .italic(false)
                        .color("#1F4E79")
                        .build())
                .svgImage(resource("/demo/chart.svg"), resource("/demo/chart.png"));

        for (int section = 1; section <= 3; section++) {
            builder.heading("Section " + section, sectionHeading);
            builder.paragraph("First paragraph of section " + section + ".");
            builder.paragraph("Second paragraph of section " + section + ".");
        }

        WordDocument document = builder.build();

        Path target = Path.of("target", "sample.docx");
        Files.createDirectories(target.getParent());
        try (OutputStream out = Files.newOutputStream(target)) {
            document.writeTo(out);
        }
        System.out.println("Wrote " + target.toAbsolutePath());
    }

    private static byte[] resource(String name) {
        try (InputStream in = SampleMain.class.getResourceAsStream(name)) {
            if (in == null) {
                throw new IllegalStateException("missing demo resource " + name);
            }
            return in.readAllBytes();
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
    }
}
