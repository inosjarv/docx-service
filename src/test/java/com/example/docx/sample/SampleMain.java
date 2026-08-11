package com.example.docx.sample;

import com.example.docx.WordDocument;
import com.example.docx.page.PageSetup;
import com.example.docx.style.HeadingStyle;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

/** Writes {@code target/sample.docx} so the output can be opened and eyeballed. */
public final class SampleMain {

    private SampleMain() {
    }

    public static void main(String[] args) throws IOException {
        WordDocument document = WordDocument.builder()
                .pageSetup(PageSetup.builder()
                        .a4()
                        .marginsTwips(851)
                        .build())
                .heading("Quarterly Report", HeadingStyle.builder()
                        .font("Calibri Light")
                        .sizePt(20)
                        .bold(true)
                        .italic(false)
                        .color("#1F4E79")
                        .build())
                .build();

        Path target = Path.of("target", "sample.docx");
        Files.createDirectories(target.getParent());
        try (OutputStream out = Files.newOutputStream(target)) {
            document.writeTo(out);
        }
        System.out.println("Wrote " + target.toAbsolutePath());
    }
}
