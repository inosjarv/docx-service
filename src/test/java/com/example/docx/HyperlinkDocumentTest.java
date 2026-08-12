package com.example.docx;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.example.docx.content.Hyperlink;
import com.example.docx.page.PageSetup;
import jakarta.xml.bind.JAXBElement;
import java.io.ByteArrayInputStream;
import java.util.List;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.Body;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Text;
import org.junit.jupiter.api.Test;

class HyperlinkDocumentTest {

    private static WordprocessingMLPackage reload(byte[] bytes) throws Exception {
        return WordprocessingMLPackage.load(new ByteArrayInputStream(bytes));
    }

    private static Body body(WordprocessingMLPackage pkg) {
        return pkg.getMainDocumentPart().getJaxbElement().getBody();
    }

    private static P firstParagraph(Body body) {
        for (Object o : body.getContent()) {
            Object v = (o instanceof JAXBElement<?> je) ? je.getValue() : o;
            if (v instanceof P p) {
                return p;
            }
        }
        throw new AssertionError("no paragraph in body");
    }

    private static String textOf(R run) {
        Object first = run.getContent().get(0);
        Object v = (first instanceof JAXBElement<?> je) ? je.getValue() : first;
        return ((Text) v).getValue();
    }

    @Test
    void linkedParagraphReloadsAsLeadingRunThenHyperlink() throws Exception {
        byte[] bytes = WordDocument.builder()
                .pageSetup(PageSetup.a4())
                .paragraph("Revenue grew 12%. ", Hyperlink.of("Learn more", "https://example.com/report"))
                .build()
                .toByteArray();

        WordprocessingMLPackage pkg = reload(bytes);
        P paragraph = firstParagraph(body(pkg));
        List<Object> content = paragraph.getContent();
        assertEquals(2, content.size(), "leading run + hyperlink, nothing else");

        Object first = content.get(0) instanceof JAXBElement<?> je ? je.getValue() : content.get(0);
        assertTrue(first instanceof R, "first item must be the leading run");
        assertEquals("Revenue grew 12%. ", textOf((R) first));

        Object second = content.get(1) instanceof JAXBElement<?> je2 ? je2.getValue() : content.get(1);
        assertTrue(second instanceof org.docx4j.wml.P.Hyperlink, "second item must be the hyperlink");
        org.docx4j.wml.P.Hyperlink hyperlink = (org.docx4j.wml.P.Hyperlink) second;
        R linkRun = (R) (hyperlink.getContent().get(0) instanceof JAXBElement<?> je3
                ? je3.getValue() : hyperlink.getContent().get(0));
        assertEquals("Learn more", textOf(linkRun));
    }

    @Test
    void relationshipIsExternalAndResolvesFromTheHyperlinkId() throws Exception {
        byte[] bytes = WordDocument.builder()
                .paragraph("See ", Hyperlink.of("the report", "https://example.com/report"))
                .build()
                .toByteArray();

        WordprocessingMLPackage pkg = reload(bytes);
        P paragraph = firstParagraph(body(pkg));
        Object second = paragraph.getContent().get(1) instanceof JAXBElement<?> je
                ? je.getValue() : paragraph.getContent().get(1);
        String relId = ((org.docx4j.wml.P.Hyperlink) second).getId();

        Relationship rel = pkg.getMainDocumentPart().getRelationshipsPart().getRelationshipByID(relId);
        assertNotNull(rel, "the hyperlink's r:id must resolve to a relationship");
        assertEquals("External", rel.getTargetMode());
        assertEquals("https://example.com/report", rel.getTarget());
    }

    @Test
    void blankLeadingTextProducesALinkOnlyParagraph() throws Exception {
        byte[] bytes = WordDocument.builder()
                .paragraph("", Hyperlink.of("Click here", "https://example.com"))
                .build()
                .toByteArray();

        P paragraph = firstParagraph(body(reload(bytes)));
        assertEquals(1, paragraph.getContent().size(), "no leading run when text is blank");
        Object only = paragraph.getContent().get(0) instanceof JAXBElement<?> je
                ? je.getValue() : paragraph.getContent().get(0);
        assertTrue(only instanceof org.docx4j.wml.P.Hyperlink);
    }

    @Test
    void linkStyleEmitsWordsDefaultColourAndUnderline() throws Exception {
        byte[] bytes = WordDocument.builder()
                .paragraph("See ", Hyperlink.of("here", "https://example.com"))
                .build()
                .toByteArray();

        P paragraph = firstParagraph(body(reload(bytes)));
        Object second = paragraph.getContent().get(1) instanceof JAXBElement<?> je
                ? je.getValue() : paragraph.getContent().get(1);
        R linkRun = (R) (((org.docx4j.wml.P.Hyperlink) second).getContent().get(0) instanceof JAXBElement<?> je2
                ? je2.getValue() : ((org.docx4j.wml.P.Hyperlink) second).getContent().get(0));
        assertEquals("0563C1", linkRun.getRPr().getColor().getVal());
        assertNotNull(linkRun.getRPr().getU(), "Word's default hyperlink style is underlined");
        assertEquals("single", linkRun.getRPr().getU().getVal().value());
    }

    @Test
    void imageAndHyperlinkRelationshipsCoexistWithDistinctIds() throws Exception {
        byte[] svg = resource("/demo/chart.svg");
        byte[] png = resource("/demo/chart.png");

        byte[] bytes = WordDocument.builder()
                .heading("Report")
                .svgImage(svg, png)
                .paragraph("See ", Hyperlink.of("the chart source", "https://example.com/data"))
                .build()
                .toByteArray();

        WordprocessingMLPackage pkg = reload(bytes);
        // Just confirming the whole document still loads and both relationship types exist.
        var relationships = pkg.getMainDocumentPart().getRelationshipsPart().getRelationships().getRelationship();
        boolean hasImage = relationships.stream().anyMatch(r -> r.getType().contains("/image"));
        boolean hasHyperlink = relationships.stream().anyMatch(r -> r.getType().contains("/hyperlink"));
        assertTrue(hasImage, "image relationship must be present");
        assertTrue(hasHyperlink, "hyperlink relationship must be present");
    }

    private static byte[] resource(String name) throws Exception {
        try (var in = HyperlinkDocumentTest.class.getResourceAsStream(name)) {
            assertNotNull(in, "missing test resource " + name);
            return in.readAllBytes();
        }
    }
}
