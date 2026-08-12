package com.example.docx.content;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

import com.example.docx.DocumentGenerationException;
import com.example.docx.style.TextStyle;
import org.junit.jupiter.api.Test;

class HyperlinkTest {

    @Test
    void defaultStyleIsTextStyleLink() {
        Hyperlink link = Hyperlink.of("Learn more", "https://example.com");
        assertEquals("Learn more", link.text());
        assertEquals("https://example.com", link.url());
        assertEquals(TextStyle.link().colorHex(), link.style().colorHex());
        assertEquals(TextStyle.link().underline(), link.style().underline());
    }

    @Test
    void explicitStyleOverridesTheDefault() {
        TextStyle custom = TextStyle.builder().color("#FF0000").underline(false).build();
        Hyperlink link = Hyperlink.of("Click", "https://example.com", custom);
        assertEquals("FF0000", link.style().colorHex());
        assertEquals(false, link.style().underline());
    }

    @Test
    void rejectsBlankTextAndUrl() {
        assertThrows(DocumentGenerationException.class, () -> Hyperlink.of("", "https://example.com"));
        assertThrows(DocumentGenerationException.class, () -> Hyperlink.of(null, "https://example.com"));
        assertThrows(DocumentGenerationException.class, () -> Hyperlink.of("Click", ""));
        assertThrows(DocumentGenerationException.class, () -> Hyperlink.of("Click", null));
    }

    @Test
    void rejectsNullStyle() {
        assertThrows(DocumentGenerationException.class,
                () -> Hyperlink.of("Click", "https://example.com", null));
    }

    @Test
    void rejectsMalformedUrlWithACause() {
        String malformed = "not a url" + " " + "with a raw space";
        DocumentGenerationException e = assertThrows(DocumentGenerationException.class,
                () -> Hyperlink.of("Click", malformed));
        assertNotNull(e.getCause(), "a malformed URL wraps URISyntaxException");
    }
}
