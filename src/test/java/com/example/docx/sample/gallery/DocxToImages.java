package com.example.docx.sample.gallery;

import com.example.docx.WordDocument;
import java.awt.image.BufferedImage;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import javax.imageio.ImageIO;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.jodconverter.core.office.OfficeException;
import org.jodconverter.core.office.OfficeManager;
import org.jodconverter.local.LocalConverter;

/**
 * Renders a {@link WordDocument} to one PNG per page, at Word-accurate fidelity, via a
 * real headless LibreOffice: {@code .docx -> .pdf} (JODConverter) then {@code .pdf ->
 * PNG} per page (PDFBox). {@code officeManager} must already be started by the caller —
 * starting one costs a few seconds, so {@code GalleryMain} pays that once for the whole
 * run, not once per fixture.
 */
public final class DocxToImages {

    private static final int RENDER_DPI = 150;

    private DocxToImages() {
    }

    public static List<Path> renderPages(OfficeManager officeManager, WordDocument document, Path outputDir)
            throws OfficeException, IOException {
        Files.createDirectories(outputDir);

        Path docx = outputDir.resolve("document.docx");
        try (OutputStream out = Files.newOutputStream(docx)) {
            document.writeTo(out);
        }

        Path pdf = outputDir.resolve("document.pdf");
        LocalConverter.make(officeManager).convert(docx.toFile()).to(pdf.toFile()).execute();

        List<Path> pages = new ArrayList<>();
        try (PDDocument pdDocument = Loader.loadPDF(pdf.toFile())) {
            PDFRenderer renderer = new PDFRenderer(pdDocument);
            for (int page = 0; page < pdDocument.getNumberOfPages(); page++) {
                BufferedImage image = renderer.renderImageWithDPI(page, RENDER_DPI, ImageType.RGB);
                Path pngPath = outputDir.resolve("page-" + (page + 1) + ".png");
                ImageIO.write(image, "PNG", pngPath.toFile());
                pages.add(pngPath);
            }
        }
        return pages;
    }
}
