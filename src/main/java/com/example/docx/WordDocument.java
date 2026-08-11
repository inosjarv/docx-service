package com.example.docx;

import com.example.docx.content.Paragraphs;
import com.example.docx.page.PageSetup;
import com.example.docx.part.ImageParts;
import com.example.docx.style.ParagraphStyle;
import com.example.docx.style.TextStyle;
import java.io.ByteArrayOutputStream;
import java.io.FilterOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Body;

/**
 * Fluent facade over a single Word document.
 *
 * <p>Wraps mutable docx4j state and is <strong>not thread safe</strong>. Build one
 * per document; construction is cheap and that is the natural lifetime in a request
 * handler.
 *
 * <p>The first document generated in a JVM pays a one-off JAXB context
 * initialisation of roughly one second. Warm it at startup with
 * {@code Context.getWmlObjectFactory()} if latency matters.
 */
public final class WordDocument {

    private final WordprocessingMLPackage pkg;

    private WordDocument(WordprocessingMLPackage pkg) {
        this.pkg = pkg;
    }

    public static Builder builder() {
        return new Builder();
    }

    /** Serialises the document. */
    public byte[] toByteArray() {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        writeTo(out);
        return out.toByteArray();
    }

    /** Serialises the document to {@code out}, without buffering the whole file. */
    public void writeTo(OutputStream out) {
        if (out == null) {
            throw new DocumentGenerationException("output stream must not be null");
        }
        try {
            // docx4j wraps the stream in a ZipOutputStream and closes it on save.
            // Closing a caller's ServletOutputStream commits the response, so the
            // stream is shielded and the caller keeps ownership.
            pkg.save(new FilterOutputStream(out) {
                @Override
                public void write(byte[] b, int off, int len) throws IOException {
                    out.write(b, off, len);
                }

                @Override
                public void close() throws IOException {
                    flush();
                }
            });
        } catch (Docx4JException e) {
            throw new DocumentGenerationException("failed to serialise the document", e);
        }
    }

    /** Fluent builder. All docx4j work happens in {@link #build()}. */
    public static final class Builder {

        private PageSetup pageSetup = PageSetup.a4();
        private String headingText;
        private TextStyle headingStyle = TextStyle.defaults();
        private byte[] svgBytes;
        private byte[] pngBytes;

        private Builder() {
        }

        public Builder pageSetup(PageSetup pageSetup) {
            if (pageSetup == null) {
                throw new DocumentGenerationException("page setup must not be null");
            }
            this.pageSetup = pageSetup;
            return this;
        }

        /** Sets the heading text and its style. */
        public Builder heading(String text, TextStyle style) {
            if (style == null) {
                throw new DocumentGenerationException("heading style must not be null");
            }
            this.headingText = text;
            this.headingStyle = style;
            return this;
        }

        /** Sets the heading text, keeping {@link TextStyle#defaults()}. */
        public Builder heading(String text) {
            this.headingText = text;
            return this;
        }

        /**
         * Adds an SVG image with a PNG fallback, drawn at half the usable page width.
         *
         * <p>The SVG must be SVG 1.1: Word's renderer rejects SVG 2 features such as
         * {@code height="auto"} and 8-digit hex colours that browsers accept.
         */
        public Builder svgImage(byte[] svg, byte[] pngFallback) {
            this.svgBytes = svg;
            this.pngBytes = pngFallback;
            return this;
        }

        public WordDocument build() {
            if (headingText == null || headingText.isBlank()) {
                throw new DocumentGenerationException(
                        "a heading is required; call heading(String) before build()");
            }
            try {
                WordprocessingMLPackage pkg = WordprocessingMLPackage.createPackage();
                MainDocumentPart mainDocumentPart = pkg.getMainDocumentPart();

                Body body = mainDocumentPart.getJaxbElement().getBody();
                body.setSectPr(pageSetup.toSectPr());

                mainDocumentPart.getContent().add(
                        Paragraphs.of(headingText, headingStyle, ParagraphStyle.heading()));

                if (svgBytes != null || pngBytes != null) {
                    mainDocumentPart.getContent().add(ImageParts.svgImage(
                            pkg, svgBytes, pngBytes, pageSetup.usableWidthTwips() / 2));
                }

                return new WordDocument(pkg);
            } catch (Docx4JException e) {
                throw new DocumentGenerationException("failed to create the document package", e);
            }
        }
    }
}
