package com.example.docx.content;

import com.example.docx.DocumentGenerationException;
import com.example.docx.style.TextStyle;
import org.docx4j.jaxb.Context;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Text;

/**
 * Builds heading paragraphs.
 *
 * <p>A pure function of its arguments: it never touches the package, so it composes
 * freely and needs no fixtures.
 */
public final class Headings {

    private Headings() {
    }

    /** A {@code w:p} holding one styled {@code w:r} with the given text. */
    public static P heading(String text, TextStyle style) {
        if (text == null || text.isBlank()) {
            throw new DocumentGenerationException("heading text must not be blank");
        }
        if (style == null) {
            throw new DocumentGenerationException("heading style must not be null");
        }

        ObjectFactory factory = Context.getWmlObjectFactory();

        Text value = factory.createText();
        value.setValue(text);
        // Without xml:space=preserve, leading and trailing spaces vanish silently.
        value.setSpace("preserve");

        R run = factory.createR();
        run.getContent().add(value);
        run.setRPr(style.toRPr());

        P paragraph = factory.createP();
        paragraph.getContent().add(run);
        return paragraph;
    }
}
