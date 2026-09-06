package com.example.docx.sample.gallery;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;

class GalleryIndexWriterTest {

    @Test
    void includesEveryFixtureNameAndItsPageImages() {
        String html = GalleryIndexWriter.write(List.of(
                RenderedFixture.success("sample",
                        List.of(Path.of("sample/page-1.png"), Path.of("sample/page-2.png"))),
                RenderedFixture.success("rich-text",
                        List.of(Path.of("rich-text/page-1.png")))));

        assertTrue(html.contains("sample"));
        assertTrue(html.contains("rich-text"));
        assertTrue(html.contains("sample/page-1.png"));
        assertTrue(html.contains("sample/page-2.png"));
        assertTrue(html.contains("rich-text/page-1.png"));
    }

    @Test
    void showsAFailureNoteInsteadOfImagesWhenAFixtureFailed() {
        String html = GalleryIndexWriter.write(List.of(
                RenderedFixture.failure("broken", "boom")));

        assertTrue(html.contains("broken"));
        assertTrue(html.contains("Failed to render"));
        assertTrue(html.contains("boom"));
        assertFalse(html.contains("<img"));
    }

    @Test
    void escapesHtmlSpecialCharactersInNamesAndErrors() {
        String html = GalleryIndexWriter.write(List.of(
                RenderedFixture.failure("a & b", "<oops>")));

        assertTrue(html.contains("a &amp; b"));
        assertTrue(html.contains("&lt;oops&gt;"));
    }
}
