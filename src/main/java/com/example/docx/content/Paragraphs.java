package com.example.docx.content;

import com.example.docx.DocumentGenerationException;
import com.example.docx.style.ParagraphStyle;
import com.example.docx.style.TextStyle;
import org.docx4j.jaxb.Context;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Text;

/**
 * Builds paragraphs and runs.
 *
 * <p>A heading and a body paragraph share {@link #of}: both are a {@code w:p} holding a
 * single {@code w:r}, differing only in the styles handed in. A pure function of its
 * arguments — it never touches the package, so it needs no fixtures.
 */
public final class Paragraphs {

    private Paragraphs() {
    }

    /**
     * A styled {@code w:r} holding {@code text}. Reused by {@link #of} and by hyperlinks,
     * which need a run without a surrounding paragraph.
     */
    public static R run(String text, TextStyle textStyle) {
        if (text == null || text.isBlank()) {
            throw new DocumentGenerationException("text must not be blank");
        }
        if (textStyle == null) {
            throw new DocumentGenerationException("text style must not be null");
        }

        ObjectFactory factory = Context.getWmlObjectFactory();

        Text value = factory.createText();
        value.setValue(text);
        // Without xml:space=preserve, leading and trailing spaces vanish silently.
        value.setSpace("preserve");

        R run = factory.createR();
        run.getContent().add(value);
        run.setRPr(textStyle.toRPr());
        return run;
    }

    /** A {@code w:p} holding one styled {@code w:r} with the given text. */
    public static P of(String text, TextStyle textStyle, ParagraphStyle paragraphStyle) {
        if (text == null || text.isBlank()) {
            throw new DocumentGenerationException("paragraph text must not be blank");
        }
        if (textStyle == null) {
            throw new DocumentGenerationException("text style must not be null");
        }
        if (paragraphStyle == null) {
            throw new DocumentGenerationException("paragraph style must not be null");
        }

        R run = run(text, textStyle);

        ObjectFactory factory = Context.getWmlObjectFactory();
        P paragraph = factory.createP();
        paragraph.setPPr(paragraphStyle.toPPr());
        paragraph.getContent().add(run);
        return paragraph;
    }

    /** A {@code w:p} whose runs come from a {@link RichText}'s parsed spans. */
    public static P of(RichText richText, ParagraphStyle paragraphStyle) {
        if (richText == null) {
            throw new DocumentGenerationException("rich text must not be null");
        }
        if (paragraphStyle == null) {
            throw new DocumentGenerationException("paragraph style must not be null");
        }

        ObjectFactory factory = Context.getWmlObjectFactory();
        P paragraph = factory.createP();
        paragraph.setPPr(paragraphStyle.toPPr());
        paragraph.getContent().addAll(richText.toRuns());
        return paragraph;
    }
}
