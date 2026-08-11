package com.example.docx;

import static org.junit.jupiter.api.Assertions.assertEquals;

import com.example.docx.page.PageSetup;
import com.example.docx.style.TableStyle;
import jakarta.xml.bind.JAXBElement;
import java.io.ByteArrayInputStream;
import java.util.List;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.Body;
import org.docx4j.wml.P;
import org.docx4j.wml.Tbl;
import org.junit.jupiter.api.Test;

class TableDocumentTest {

    private static final List<String> HEADERS = List.of("Region", "Revenue");
    private static final List<List<String>> ROWS = List.of(List.of("EMEA", "1 240"));

    private static Body reload(byte[] bytes) throws Exception {
        return WordprocessingMLPackage.load(new ByteArrayInputStream(bytes))
                .getMainDocumentPart().getJaxbElement().getBody();
    }

    /** A reloaded body may wrap its children in JAXBElement; a fresh one does not. */
    private static List<Object> unwrapped(Body body) {
        return body.getContent().stream()
                .map(o -> (o instanceof JAXBElement<?> je) ? je.getValue() : o)
                .toList();
    }

    private static List<String> shape(Body body) {
        return unwrapped(body).stream()
                .map(o -> o instanceof Tbl ? "TBL" : o instanceof P ? "P" : "?")
                .toList();
    }

    @Test
    void tableSurvivesTheRoundTripAtFullUsableWidth() throws Exception {
        byte[] bytes = WordDocument.builder()
                .pageSetup(PageSetup.a4())
                .heading("Report")
                .table(HEADERS, ROWS, TableStyle.defaults())
                .build()
                .toByteArray();

        Tbl table = (Tbl) unwrapped(reload(bytes)).stream()
                .filter(o -> o instanceof Tbl).findFirst().orElseThrow();

        int sum = table.getTblGrid().getGridCol().stream()
                .mapToInt(c -> c.getW().intValue()).sum();
        assertEquals(PageSetup.a4().usableWidthTwips(), sum,
                "the default table fills the usable page width");
    }

    @Test
    void aSpacerParagraphFollowsEveryTable() throws Exception {
        byte[] bytes = WordDocument.builder()
                .paragraph("Before")
                .table(HEADERS, ROWS, TableStyle.defaults())
                .paragraph("After")
                .build()
                .toByteArray();

        // Two adjacent tables merge in Word, and a body ending in a table is irregular,
        // so a spacer paragraph follows each one.
        assertEquals(List.of("P", "TBL", "P", "P"), shape(reload(bytes)),
                "expected paragraph, table, spacer, paragraph");
    }

    @Test
    void tableWidthReflectsPageSetupSetAfterTable() throws Exception {
        PageSetup wide = PageSetup.builder()
                .pageSizeTwips(20000, 15840)
                .marginsTwips(851)
                .build();

        byte[] bytes = WordDocument.builder()
                .table(HEADERS, ROWS, TableStyle.defaults())
                .pageSetup(wide)
                .build()
                .toByteArray();

        Tbl table = (Tbl) unwrapped(reload(bytes)).stream()
                .filter(o -> o instanceof Tbl).findFirst().orElseThrow();

        int sum = table.getTblGrid().getGridCol().stream()
                .mapToInt(c -> c.getW().intValue()).sum();
        assertEquals(wide.usableWidthTwips(), sum,
                "table() called before pageSetup() should still size to the final page setup");
    }
}
