package com.example.docx.sample.gallery;

import java.nio.file.Path;
import java.util.List;

/**
 * Builds the {@code index.html} string for a gallery run: one section per fixture, its
 * page images stacked in order, or a visible failure note if it didn't render. Pure
 * string-in-string-out — no I/O, no LibreOffice — so it's unit-testable on its own.
 */
public final class GalleryIndexWriter {

    private GalleryIndexWriter() {
    }

    public static String write(List<RenderedFixture> fixtures) {
        StringBuilder html = new StringBuilder();
        html.append("<!doctype html>\n<html><head><meta charset=\"utf-8\">")
                .append("<title>Document gallery</title></head><body>\n");

        for (RenderedFixture fixture : fixtures) {
            html.append("<section>\n<h2>").append(escape(fixture.name())).append("</h2>\n");
            if (fixture.failed()) {
                html.append("<p>Failed to render: ").append(escape(fixture.error())).append("</p>\n");
            } else {
                int total = fixture.pages().size();
                for (int i = 0; i < total; i++) {
                    Path page = fixture.pages().get(i);
                    html.append("<p>Page ").append(i + 1).append(" of ").append(total).append("</p>\n")
                            .append("<img src=\"").append(escape(page.toString())).append("\">\n");
                }
            }
            html.append("</section>\n<hr>\n");
        }

        html.append("</body></html>\n");
        return html.toString();
    }

    private static String escape(String text) {
        return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                .replace("\"", "&quot;");
    }
}
