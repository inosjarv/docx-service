package com.example.docx.sample.gallery;

import com.example.docx.WordDocument;
import java.util.function.Supplier;

/** One document to render for review: a display name paired with how to build it. */
public record GalleryFixture(String name, Supplier<WordDocument> builder) {
}
