package com.example.docx;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.example.docx.page.PageSetup;
import com.example.docx.style.HeadingStyle;
import jakarta.xml.bind.JAXBElement;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.math.BigInteger;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.Body;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.SectPr;
import org.docx4j.wml.Text;
import org.junit.jupiter.api.Test;

class WordDocumentRoundTripTest {

    private static byte[] sample() {
        return WordDocument.builder()
                .pageSetup(PageSetup.a4())
                .heading("Quarterly Report", HeadingStyle.defaults())
                .build()
                .toByteArray();
    }

    private static Body reload(byte[] bytes) throws Exception {
        WordprocessingMLPackage pkg =
                WordprocessingMLPackage.load(new ByteArrayInputStream(bytes));
        return pkg.getMainDocumentPart().getJaxbElement().getBody();
    }

    private static P firstParagraph(Body body) {
        for (Object o : body.getContent()) {
            Object value = (o instanceof JAXBElement<?> je) ? je.getValue() : o;
            if (value instanceof P p) {
                return p;
            }
        }
        throw new AssertionError("no paragraph in body");
    }

    @Test
    void marginsSurviveTheRoundTrip() throws Exception {
        SectPr sectPr = reload(sample()).getSectPr();
        assertNotNull(sectPr);
        assertEquals(BigInteger.valueOf(851), sectPr.getPgMar().getTop());
        assertEquals(BigInteger.valueOf(851), sectPr.getPgMar().getRight());
        assertEquals(BigInteger.valueOf(851), sectPr.getPgMar().getBottom());
        assertEquals(BigInteger.valueOf(851), sectPr.getPgMar().getLeft());
    }

    @Test
    void pageSizeSurvivesTheRoundTrip() throws Exception {
        SectPr sectPr = reload(sample()).getSectPr();
        assertEquals(BigInteger.valueOf(11906), sectPr.getPgSz().getW());
        assertEquals(BigInteger.valueOf(16838), sectPr.getPgSz().getH());
    }

    @Test
    void headingSurvivesTheRoundTrip() throws Exception {
        P p = firstParagraph(reload(sample()));
        R run = (R) p.getContent().get(0);

        assertEquals(BigInteger.valueOf(40), run.getRPr().getSz().getVal());
        assertEquals("1F4E79", run.getRPr().getColor().getVal());
        assertTrue(run.getRPr().getB().isVal());
        assertEquals("Calibri Light", run.getRPr().getRFonts().getAscii());

        Text text = (Text) ((JAXBElement<?>) run.getContent().get(0)).getValue();
        assertEquals("Quarterly Report", text.getValue());
    }

    @Test
    void customStyleSurvivesTheRoundTrip() throws Exception {
        byte[] bytes = WordDocument.builder()
                .pageSetup(PageSetup.builder().a4().marginsInches(1.0).build())
                .heading("Custom", HeadingStyle.builder()
                        .font("Arial")
                        .sizePt(14)
                        .bold(false)
                        .italic(true)
                        .color("#C00000")
                        .build())
                .build()
                .toByteArray();

        Body body = reload(bytes);
        assertEquals(BigInteger.valueOf(1440), body.getSectPr().getPgMar().getTop());

        R run = (R) firstParagraph(body).getContent().get(0);
        assertEquals("Arial", run.getRPr().getRFonts().getAscii());
        assertEquals(BigInteger.valueOf(28), run.getRPr().getSz().getVal());
        assertEquals("C00000", run.getRPr().getColor().getVal());
        assertTrue(run.getRPr().getI().isVal());
    }

    @Test
    void writeToAndToByteArrayProduceEquivalentDocuments() throws Exception {
        WordDocument doc = WordDocument.builder()
                .pageSetup(PageSetup.a4())
                .heading("Quarterly Report", HeadingStyle.defaults())
                .build();

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        doc.writeTo(out);

        // Not a byte comparison: ZIP entries carry 2-second-granularity timestamps,
        // so two saves can legitimately differ. Compare what the document means.
        Body streamed = reload(out.toByteArray());
        Body buffered = reload(doc.toByteArray());

        assertEquals(streamed.getSectPr().getPgMar().getTop(),
                buffered.getSectPr().getPgMar().getTop());
        assertEquals(streamed.getSectPr().getPgSz().getW(),
                buffered.getSectPr().getPgSz().getW());

        R streamedRun = (R) firstParagraph(streamed).getContent().get(0);
        R bufferedRun = (R) firstParagraph(buffered).getContent().get(0);
        assertEquals(text(streamedRun), text(bufferedRun));
    }

    private static String text(R run) {
        Object first = run.getContent().get(0);
        Text t = (Text) (first instanceof JAXBElement<?> je ? je.getValue() : first);
        return t.getValue();
    }

    @Test
    void headingDefaultsToTheDefaultStyle() throws Exception {
        byte[] bytes = WordDocument.builder().heading("Title").build().toByteArray();
        R run = (R) firstParagraph(reload(bytes)).getContent().get(0);
        assertEquals("Calibri Light", run.getRPr().getRFonts().getAscii());
    }

    @Test
    void requiresAHeading() {
        assertThrows(DocumentGenerationException.class, () -> WordDocument.builder().build());
    }

    @Test
    void rejectsNullOutputStream() {
        WordDocument doc = WordDocument.builder().heading("Title").build();
        assertThrows(DocumentGenerationException.class, () -> doc.writeTo(null));
    }
}
