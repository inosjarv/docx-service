package com.example.docx.part;

import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.example.docx.DocumentGenerationException;
import com.example.docx.WordDocument;
import com.example.docx.page.PageSetup;
import com.example.docx.style.HeadingStyle;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.PartName;
import org.junit.jupiter.api.Test;

class ImagePartsTest {

    private static byte[] resource(String name) throws IOException {
        try (InputStream in = ImagePartsTest.class.getResourceAsStream(name)) {
            assertNotNull(in, "missing test resource " + name);
            return in.readAllBytes();
        }
    }

    private static byte[] svg() throws IOException {
        return resource("/demo/chart.svg");
    }

    private static byte[] png() throws IOException {
        return resource("/demo/chart.png");
    }

    private static String documentXml(PageSetup setup) throws Exception {
        byte[] bytes = WordDocument.builder()
                .pageSetup(setup)
                .heading("Quarterly Report", HeadingStyle.defaults())
                .svgImage(svg(), png())
                .build()
                .toByteArray();
        return WordprocessingMLPackage.load(new ByteArrayInputStream(bytes))
                .getMainDocumentPart().getXML();
    }

    @Test
    void bothMediaPartsExistWithCorrectContentTypes() throws Exception {
        byte[] bytes = WordDocument.builder()
                .pageSetup(PageSetup.a4())
                .heading("Quarterly Report", HeadingStyle.defaults())
                .svgImage(svg(), png())
                .build()
                .toByteArray();

        WordprocessingMLPackage re =
                WordprocessingMLPackage.load(new ByteArrayInputStream(bytes));

        Map<String, String> media = new HashMap<>();
        for (Map.Entry<PartName, Part> e : re.getParts().getParts().entrySet()) {
            String name = e.getKey().getName();
            if (name.startsWith("/word/media/")) {
                media.put(name, e.getValue().getContentType());
            }
        }

        assertTrue(media.keySet().stream().anyMatch(n -> n.endsWith(".png")),
                "png part missing, got " + media);
        assertTrue(media.keySet().stream().anyMatch(n -> n.endsWith(".svg")),
                "svg part missing, got " + media);
        assertTrue(media.values().stream().anyMatch(c -> c.startsWith("image/svg+xml")),
                "svg content type missing, got " + media);
    }

    @Test
    void svgBlipCarriesItsNamespaceAndPrefixedEmbed() throws Exception {
        String xml = documentXml(PageSetup.a4());

        assertTrue(xml.contains("{96DAC541-7B7A-43D3-8B79-37D633B846F1}"),
                "svg extension uri missing");
        // The namespace, not just the local name. An unnamespaced svgBlip still
        // contains the string "svgBlip" and is silently ignored by Word.
        assertTrue(xml.contains("http://schemas.microsoft.com/office/drawing/2016/SVG/main"),
                "asvg namespace missing");
        assertTrue(xml.contains(":svgBlip"), "svgBlip must carry its namespace prefix");
        assertTrue(xml.contains("r:embed"), "r:embed must keep its prefix");
    }

    @Test
    void imageIsHalfTheUsablePageWidth() throws Exception {
        // A4 usable = 11906 - 851 - 851 = 10204; half = 5102; cx = 5102 * 635.
        // PNG is 1600x1120, so cy = cx * 1120 / 1600.
        String xml = documentXml(PageSetup.a4());
        assertTrue(xml.contains("cx=\"3239770\""), "expected cx 3239770 in " + extent(xml));
        assertTrue(xml.contains("cy=\"2267839\""), "expected cy 2267839 in " + extent(xml));
    }

    @Test
    void widthTracksThePageRatherThanAConstant() throws Exception {
        // usable = 11906 - 2000 - 2000 = 7906; half = 3953; cx = 3953 * 635 = 2510155.
        String xml = documentXml(PageSetup.builder().a4().marginsTwips(2000).build());
        assertTrue(xml.contains("cx=\"2510155\""),
                "a narrower page must yield a narrower image, got " + extent(xml));
    }

    private static String extent(String xml) {
        return xml.replaceAll("(?s).*(<wp:extent[^/]*/>).*", "$1");
    }

    @Test
    void rejectsBadInput() throws Exception {
        WordprocessingMLPackage pkg = WordprocessingMLPackage.createPackage();
        assertThrows(DocumentGenerationException.class,
                () -> ImageParts.svgImage(pkg, new byte[0], png(), 5102));
        assertThrows(DocumentGenerationException.class,
                () -> ImageParts.svgImage(pkg, svg(), null, 5102));
        assertThrows(DocumentGenerationException.class,
                () -> ImageParts.svgImage(pkg, svg(), "not a png".getBytes(), 5102));
        assertThrows(DocumentGenerationException.class,
                () -> ImageParts.svgImage(pkg, svg(), png(), 0));
    }
}
