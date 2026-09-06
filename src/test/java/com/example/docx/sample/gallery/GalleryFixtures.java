package com.example.docx.sample.gallery;

import com.example.docx.sample.PostmortemSampleMain;
import com.example.docx.sample.RichTextSampleMain;
import com.example.docx.sample.SampleMain;
import java.util.List;

/**
 * The documents the gallery renders. Add a fixture here — one line — to have it show up
 * in {@code target/gallery/index.html} the next time {@code GalleryMain} runs.
 */
public final class GalleryFixtures {

    private GalleryFixtures() {
    }

    public static List<GalleryFixture> all() {
        return List.of(
                new GalleryFixture("sample", SampleMain::build),
                new GalleryFixture("postmortem", PostmortemSampleMain::build),
                new GalleryFixture("rich-text", RichTextSampleMain::build));
    }
}
