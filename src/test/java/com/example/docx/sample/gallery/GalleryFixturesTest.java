package com.example.docx.sample.gallery;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

import com.example.docx.WordDocument;
import java.util.List;
import org.junit.jupiter.api.Test;

class GalleryFixturesTest {

    @Test
    void listsTheThreeExpectedFixturesInOrder() {
        List<GalleryFixture> fixtures = GalleryFixtures.all();
        assertEquals(List.of("sample", "postmortem", "rich-text"),
                fixtures.stream().map(GalleryFixture::name).toList());
    }

    @Test
    void everyFixtureBuildsAWordDocumentWithoutThrowing() {
        for (GalleryFixture fixture : GalleryFixtures.all()) {
            WordDocument document = fixture.builder().get();
            assertNotNull(document, fixture.name() + " must build a document");
        }
    }
}
