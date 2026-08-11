package com.example.docx;

import com.example.docx.content.Paragraphs;
import com.example.docx.content.Tables;
import com.example.docx.page.PageSetup;
import com.example.docx.part.ImageParts;
import com.example.docx.style.ParagraphStyle;
import com.example.docx.style.TableStyle;
import com.example.docx.style.TextStyle;
import java.io.ByteArrayOutputStream;
import java.io.FilterOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Body;
import org.docx4j.wml.P;
import org.docx4j.wml.Tbl;

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

    /**
     * Fluent builder. All docx4j work happens in {@link #build()}.
     *
     * <p>Content calls append; none replaces a previous one. Text paragraphs are built
     * eagerly, so a builder is single-use: calling {@code build()} twice would give both
     * documents the same paragraph objects. Build one document per builder.
     */
    public static final class Builder {

        private PageSetup pageSetup = PageSetup.a4();
        private final List<DocumentContent> content = new ArrayList<>();

        private Builder() {
        }

        public Builder pageSetup(PageSetup pageSetup) {
            if (pageSetup == null) {
                throw new DocumentGenerationException("page setup must not be null");
            }
            this.pageSetup = pageSetup;
            return this;
        }

        /** Appends a heading using {@link TextStyle#defaults()}. */
        public Builder heading(String text) {
            return heading(text, TextStyle.defaults());
        }

        /** Appends a heading. Each call adds one; calls do not replace each other. */
        public Builder heading(String text, TextStyle style) {
            return heading(text, style, ParagraphStyle.heading());
        }

        /** Appends a heading with explicit paragraph-level formatting. */
        public Builder heading(String text, TextStyle style, ParagraphStyle paragraphStyle) {
            P paragraph = Paragraphs.of(text, style, paragraphStyle);
            content.add(pkg -> paragraph);
            return this;
        }

        /** Appends a body paragraph using {@link TextStyle#body()}. */
        public Builder paragraph(String text) {
            return paragraph(text, TextStyle.body());
        }

        /** Appends a body paragraph. */
        public Builder paragraph(String text, TextStyle style) {
            return paragraph(text, style, ParagraphStyle.body());
        }

        /** Appends a body paragraph with explicit paragraph-level formatting. */
        public Builder paragraph(String text, TextStyle style, ParagraphStyle paragraphStyle) {
            P paragraph = Paragraphs.of(text, style, paragraphStyle);
            content.add(pkg -> paragraph);
            return this;
        }

        /**
         * Appends an SVG image with a PNG fallback, drawn at half the usable page width.
         *
         * <p>The SVG must be SVG 1.1: Word's renderer rejects SVG 2 features such as
         * {@code height="auto"} and 8-digit hex colours that browsers accept.
         */
        public Builder svgImage(byte[] svg, byte[] pngFallback) {
            content.add(pkg -> ImageParts.svgImage(
                    pkg, svg, pngFallback, pageSetup.usableWidthTwips() / 2));
            return this;
        }

        /**
         * Appends a table spanning the full usable page width.
         *
         * <p>The width is read from the current page setup <em>now</em>, not at
         * {@code build()}. Call {@link #pageSetup(PageSetup)} before this, or use the
         * four-argument overload with an explicit width — otherwise a later
         * {@code pageSetup} call leaves the table sized for the old page.
         */
        public Builder table(List<String> headers, List<List<String>> rows, TableStyle style) {
            return table(headers, rows, style, pageSetup.usableWidthTwips());
        }

        /** Appends a table of the given width, followed by a spacer paragraph. */
        public Builder table(List<String> headers, List<List<String>> rows,
                             TableStyle style, int widthTwips) {
            Tbl table = Tables.of(headers, rows, style, widthTwips);
            content.add(pkg -> table);
            // Two adjacent tables merge into one in Word, and a body ending in a table
            // rather than a paragraph is irregular. A spacer prevents both.
            P spacer = Context.getWmlObjectFactory().createP();
            content.add(pkg -> spacer);
            return this;
        }

        public WordDocument build() {
            if (content.isEmpty()) {
                throw new DocumentGenerationException(
                        "a document needs at least one heading, paragraph or image");
            }
            try {
                WordprocessingMLPackage pkg = WordprocessingMLPackage.createPackage();
                MainDocumentPart mainDocumentPart = pkg.getMainDocumentPart();

                Body body = mainDocumentPart.getJaxbElement().getBody();
                body.setSectPr(pageSetup.toSectPr());

                for (DocumentContent item : content) {
                    mainDocumentPart.getContent().add(item.toBodyElement(pkg));
                }

                return new WordDocument(pkg);
            } catch (Docx4JException e) {
                throw new DocumentGenerationException("failed to create the document package", e);
            }
        }
    }
}
