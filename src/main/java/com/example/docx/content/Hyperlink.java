package com.example.docx.content;

import com.example.docx.DocumentGenerationException;
import com.example.docx.style.TextStyle;
import java.net.URI;
import java.net.URISyntaxException;

/**
 * An immutable value: the visible text of a link, its destination, and how it is styled.
 *
 * <p>Pure data — building the actual {@code w:hyperlink} needs a relationship, which
 * needs the package, so that work lives in {@code part.Hyperlinks}. This class only
 * validates and holds what it was given.
 */
public final class Hyperlink {

    private final String text;
    private final String url;
    private final TextStyle style;

    private Hyperlink(String text, String url, TextStyle style) {
        this.text = text;
        this.url = url;
        this.style = style;
    }

    /** A link styled with {@link TextStyle#link()} — Word's own hyperlink look. */
    public static Hyperlink of(String text, String url) {
        return of(text, url, TextStyle.link());
    }

    /** A link with explicit styling. */
    public static Hyperlink of(String text, String url, TextStyle style) {
        if (text == null || text.isBlank()) {
            throw new DocumentGenerationException("link text must not be blank");
        }
        if (style == null) {
            throw new DocumentGenerationException("link text style must not be null");
        }
        String checkedUrl = validateUrl(url);
        return new Hyperlink(text, checkedUrl, style);
    }

    private static String validateUrl(String url) {
        if (url == null || url.isBlank()) {
            throw new DocumentGenerationException("link url must not be blank");
        }
        try {
            new URI(url);
        } catch (URISyntaxException e) {
            throw new DocumentGenerationException("link url is not a valid URI: '" + url + "'", e);
        }
        return url;
    }

    public String text() {
        return text;
    }

    public String url() {
        return url;
    }

    public TextStyle style() {
        return style;
    }
}
