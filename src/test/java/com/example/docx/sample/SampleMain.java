package com.example.docx.sample;

import com.example.docx.WordDocument;
import com.example.docx.page.PageSetup;
import com.example.docx.style.Alignment;
import com.example.docx.style.ParagraphStyle;
import com.example.docx.style.TextStyle;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UncheckedIOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

/** Writes a timestamped {@code target/sample-*.docx} so the output can be eyeballed. */
public final class SampleMain {

    /** One colour per section, shared by its heading and its body text. */
    private static final List<String> SECTION_COLOURS = List.of("#1F4E79", "#843C0C", "#375623");

    private static final List<String> SECTION_TITLES =
            List.of("Revenue", "Operating costs", "Outlook");

    /** Long enough to wrap over several lines, so justification is visible. */
    private static final List<List<String>> SECTION_BODIES = List.of(
            List.of(
                    "Revenue grew twelve per cent quarter on quarter, driven largely by renewals in "
                            + "the enterprise tier and a smaller but growing contribution from "
                            + "self-serve. Expansion within existing accounts accounted for rather "
                            + "more of the increase than new logos did, which is the healthier of "
                            + "the two shapes at this stage.",
                    "Regional performance was uneven. EMEA finished ahead of plan on the strength "
                            + "of two large renewals that had been forecast to slip, while APAC "
                            + "grew faster in percentage terms from a much smaller base. The "
                            + "Americas were broadly flat."),
            List.of(
                    "Operating costs rose more slowly than revenue for the third consecutive "
                            + "quarter. Headcount was the largest single line as expected, though "
                            + "the increase was concentrated in engineering rather than spread "
                            + "evenly, reflecting the hiring plan agreed at the start of the year.",
                    "Infrastructure spend fell in absolute terms despite higher usage, following "
                            + "the migration completed in the second quarter. We expect that "
                            + "benefit to be largely one-off and would not model a repeat."),
            List.of(
                    "The outlook for the coming quarter is cautiously positive. The renewal book "
                            + "is smaller than the one just closed, so growth will depend more on "
                            + "new business than it has recently, and new business carries a "
                            + "longer and less predictable cycle.",
                    "Two risks are worth naming. The first is concentration: the top ten accounts "
                            + "remain a large share of recurring revenue. The second is the "
                            + "pipeline's dependence on a single channel, which we are actively "
                            + "working to diversify."));

    private SampleMain() {
    }

    public static void main(String[] args) throws IOException {
        // Justified body text, flush to both margins.
        ParagraphStyle justifiedBody = ParagraphStyle.builder()
                .spaceAfterTwips(120)
                .alignment(Alignment.JUSTIFY)
                .build();

        var builder = WordDocument.builder()
                .pageSetup(PageSetup.builder().a4().marginsTwips(851).build())
                .heading("Quarterly Report", TextStyle.builder()
                        .font("Calibri Light")
                        .sizePt(20)
                        .bold(true)
                        .italic(false)
                        .color("#1F4E79")
                        .build())
                .svgImage(resource("/demo/chart.svg"), resource("/demo/chart.png"));

        for (int section = 0; section < SECTION_TITLES.size(); section++) {
            String colour = SECTION_COLOURS.get(section);

            builder.heading(SECTION_TITLES.get(section), TextStyle.builder()
                    .font("Calibri")
                    .sizePt(11)
                    .bold(true)
                    .color(colour)
                    .build());

            TextStyle bodyStyle = TextStyle.builder()
                    .font("Calibri")
                    .sizePt(11)
                    .bold(false)
                    .color(colour)
                    .build();

            for (String paragraph : SECTION_BODIES.get(section)) {
                builder.paragraph(paragraph, bodyStyle, justifiedBody);
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
