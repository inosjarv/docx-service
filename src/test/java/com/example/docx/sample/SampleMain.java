package com.example.docx.sample;

import com.example.docx.WordDocument;
import com.example.docx.page.PageSetup;
import com.example.docx.style.Alignment;
import com.example.docx.style.BorderLine;
import com.example.docx.style.Edge;
import com.example.docx.style.ParagraphStyle;
import com.example.docx.style.TableBorderStyle;
import com.example.docx.style.TableStyle;
import com.example.docx.style.TextStyle;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UncheckedIOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

/**
 * Writes a timestamped {@code target/sample-*.docx} so the output can be eyeballed.
 *
 * <p>Deliberately long enough to run over several pages, so page-flow behaviour is
 * visible: sections after the first start on a fresh page, headings never strand at a
 * page foot, and body paragraphs split where they must without leaving a lone line.
 */
public final class SampleMain {

    /** One colour per section, shared by its heading and its body text. */
    private static final List<String> SECTION_COLOURS = List.of("#1F4E79", "#843C0C", "#375623");

    private static final List<String> SECTION_TITLES =
            List.of("Revenue", "Operating costs", "Outlook");

    /** Long enough to wrap and to push the document over a page boundary. */
    private static final List<List<String>> SECTION_BODIES = List.of(
            List.of(
                    "Revenue grew twelve per cent quarter on quarter, driven largely by renewals "
                            + "in the enterprise tier and a smaller but growing contribution from "
                            + "self-serve. Expansion within existing accounts accounted for rather "
                            + "more of the increase than new logos did, which is the healthier of "
                            + "the two shapes at this stage of the year and suggests the retention "
                            + "work done over the previous two quarters is beginning to show.",
                    "Regional performance was uneven. EMEA finished ahead of plan on the strength "
                            + "of two large renewals that had been forecast to slip into the "
                            + "following quarter, while APAC grew faster in percentage terms from "
                            + "a much smaller base. The Americas were broadly flat, which we "
                            + "attribute to a hiring pause in the first six weeks rather than to "
                            + "any change in demand.",
                    "Average contract value rose modestly. The increase came almost entirely from "
                            + "existing customers moving up a tier rather than from list-price "
                            + "changes, and the discounting rate was within the band agreed at the "
                            + "start of the year. Net revenue retention finished slightly above "
                            + "the internal target.",
                    "Collections were unremarkable, which is the outcome we want. Days sales "
                            + "outstanding moved by less than a day and no account of material "
                            + "size moved into the ninety-day bucket. The one exception has since "
                            + "been resolved and did not require escalation."),
            List.of(
                    "Operating costs rose more slowly than revenue for the third consecutive "
                            + "quarter. Headcount was the largest single line as expected, though "
                            + "the increase was concentrated in engineering rather than spread "
                            + "evenly across functions, reflecting the hiring plan agreed at the "
                            + "start of the year and the decision to defer two commercial roles.",
                    "Infrastructure spend fell in absolute terms despite higher usage, following "
                            + "the migration completed in the second quarter. We expect that "
                            + "benefit to be largely one-off and would not model a repeat in the "
                            + "coming period; the underlying growth in usage continues and will "
                            + "reassert itself once the migration saving has been absorbed.",
                    "Professional services costs were higher than planned. Most of the variance "
                            + "sits in two implementations that ran longer than scoped, and the "
                            + "scoping process has since been changed to require a written "
                            + "technical review before a start date is committed.",
                    "Travel and events returned to roughly pre-pandemic levels. This was "
                            + "anticipated and budgeted, and the return on the two conferences we "
                            + "sponsored is being measured against pipeline created rather than "
                            + "against leads captured, which we consider the more honest metric."),
            List.of(
                    "The outlook for the coming quarter is cautiously positive. The renewal book "
                            + "is smaller than the one just closed, so growth will depend more on "
                            + "new business than it has recently, and new business carries a "
                            + "longer and less predictable cycle. We have not changed the annual "
                            + "guidance on the strength of one good quarter.",
                    "Two risks are worth naming. The first is concentration: the top ten accounts "
                            + "remain a large share of recurring revenue, and the loss of any one "
                            + "of them would be material to the year. The second is the pipeline's "
                            + "dependence on a single channel, which we are actively working to "
                            + "diversify but which will take more than one quarter to change.",
                    "Against those, two things are working in our favour. Gross margin has "
                            + "improved for four consecutive quarters and shows no sign of "
                            + "reversing, and the product roadmap for the next two releases is "
                            + "already committed and staffed, which removes a source of "
                            + "uncertainty that troubled the previous year.",
                    "We will report again at the end of the quarter. The reporting pack will "
                            + "carry the same structure as this one, with the addition of a "
                            + "cohort view that several readers have asked for and which is now "
                            + "possible following the data migration."));

    private SampleMain() {
    }

    public static void main(String[] args) throws IOException {
        var builder = WordDocument.builder()
                .pageSetup(PageSetup.builder().a4().marginsTwips(851).build())
                .heading("Quarterly Report", TextStyle.builder()
                        .font("Calibri Light")
                        .sizePt(20)
                        .bold(true)
                        .italic(false)
                        .color("#1F4E79")
                        .build())
                .svgImage(resource("/demo/chart.svg"), resource("/demo/chart.png"))
                .table(
                        List.of("Region", "Revenue", "Change", "Share"),
                        List.of(
                                List.of("EMEA", "1 240", "+8%", "42%"),
                                List.of("APAC", "980", "+21%", "33%"),
                                List.of("Americas", "740", "0%", "25%")),
                        TableStyle.builder()
                                // A single rule under the header, nothing else.
                                .headerBorder(TableBorderStyle.builder()
                                        .color("#1F4E79").widthPt(1.0)
                                        .line(BorderLine.SINGLE).edges(Edge.BOTTOM).build())
                                // Hairline rules between rows.
                                .bodyBorder(TableBorderStyle.builder()
                                        .color("#BFBFBF").widthPt(0.5)
                                        .line(BorderLine.SINGLE).edges(Edge.BOTTOM).build())
                                .build());

        for (int section = 0; section < SECTION_TITLES.size(); section++) {
            String colour = SECTION_COLOURS.get(section);

            // The first section flows on after the chart; the rest open a new page.
            ParagraphStyle headingStyle = ParagraphStyle.builder()
                    .spaceBeforeTwips(ParagraphStyle.HEADING_SPACE_BEFORE_TWIPS)
                    .spaceAfterTwips(ParagraphStyle.HEADING_SPACE_AFTER_TWIPS)
                    .keepWithNext(true)
                    .keepLines(true)
                    .pageBreakBefore(section > 0)
                    .build();

            builder.heading(SECTION_TITLES.get(section), TextStyle.builder()
                    .font("Calibri")
                    .sizePt(11)
                    .bold(true)
                    .color(colour)
                    .build(), headingStyle);

            TextStyle bodyText = TextStyle.builder()
                    .font("Calibri")
                    .sizePt(11)
                    .bold(false)
                    .color(colour)
                    .build();

            // Justified, free to split across pages, but never stranding a single line.
            ParagraphStyle bodyParagraph = ParagraphStyle.builder()
                    .spaceAfterTwips(ParagraphStyle.BODY_SPACE_AFTER_TWIPS)
                    .alignment(Alignment.JUSTIFY)
                    .widowControl(true)
                    .build();

            for (String paragraph : SECTION_BODIES.get(section)) {
                builder.paragraph(paragraph, bodyText, bodyParagraph);
            }
        }

        WordDocument document = builder.build();

        Path target = Path.of("target", "sample-%d.docx".formatted(System.currentTimeMillis()));
        Files.createDirectories(target.getParent());
        try (OutputStream out = Files.newOutputStream(target)) {
            document.writeTo(out);
        }
        System.out.println("Wrote " + target.toAbsolutePath());
    }

    private static byte[] resource(String name) {
        try (InputStream in = SampleMain.class.getResourceAsStream(name)) {
            if (in == null) {
                throw new IllegalStateException("missing demo resource " + name);
            }
            return in.readAllBytes();
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
    }
}
