package com.example.docx.sample.gallery;

import java.nio.file.Path;
import java.util.List;

/**
 * One fixture's rendering result: its page images on success, or {@code error} (with
 * {@code pages} empty) if rendering failed. Paths are relative to the gallery output
 * directory, exactly as written into {@code index.html}'s {@code <img src>} attributes.
 */
public record RenderedFixture(String name, List<Path> pages, String error) {

    public static RenderedFixture success(String name, List<Path> pages) {
        return new RenderedFixture(name, pages, null);
    }

    public static RenderedFixture failure(String name, String error) {
        return new RenderedFixture(name, List.of(), error);
    }

    public boolean failed() {
        return error != null;
    }
}
