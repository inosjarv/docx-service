package com.example.docx.sample;

import com.example.docx.WordDocument;
import com.example.docx.content.Hyperlink;
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
 * Writes a timestamped {@code target/postmortem-*.docx} so the output can be
 * eyeballed.
 *
 * <p>Bigger and a different genre than {@link SampleMain}'s quarterly report: an
 * incident postmortem, long enough to clear four A4 pages. It exists mainly to
 * demonstrate one specific behaviour reliably rather than by chance — see the long
 * paragraph inside "Root Cause Analysis" below, which is deliberately written long
 * and placed so it is forced to split across a page boundary.
 */
public final class PostmortemSampleMain {

    private PostmortemSampleMain() {
    }

    public static void main(String[] args) throws IOException {
        WordDocument document = build();

        Path target = Path.of("target", "postmortem-%d.docx".formatted(System.currentTimeMillis()));
        Files.createDirectories(target.getParent());
        try (OutputStream out = Files.newOutputStream(target)) {
            document.writeTo(out);
        }
        System.out.println("Wrote " + target.toAbsolutePath());
    }

    /**
     * Builds the same document {@link #main} writes to disk, so other tools (such as
     * {@code PostmortemHtmlSampleMain} and the document gallery) can render it without
     * duplicating the content.
     */
    public static WordDocument build() {
        var builder = WordDocument.builder()
                .pageSetup(PageSetup.builder().a4().marginsTwips(851).build())
                .heading("Incident Postmortem: Checkout Service Outage", TextStyle.builder()
                        .font("Calibri Light")
                        .sizePt(20)
                        .bold(true)
                        .italic(false)
                        .color("#1F4E79")
                        .build())
                .svgImage(resource("/demo/chart.svg"), resource("/demo/chart.png"));

        TextStyle captionText = TextStyle.builder()
                .font("Calibri")
                .sizePt(9)
                .italic(true)
                .color("#595959")
                .build();
        ParagraphStyle captionParagraph = ParagraphStyle.builder()
                .spaceAfterTwips(ParagraphStyle.BODY_SPACE_AFTER_TWIPS)
                .alignment(Alignment.CENTER)
                .build();
        builder.paragraph(
                "Error rate and request volume across the checkout payment endpoint, "
                        + "07:00-10:00 UTC on 14 May.",
                captionText, captionParagraph);

        TextStyle plainBody = TextStyle.builder()
                .font("Calibri")
                .sizePt(11)
                .bold(false)
                .color("#000000")
                .build();

        ParagraphStyle bodyParagraph = ParagraphStyle.builder()
                .spaceAfterTwips(ParagraphStyle.BODY_SPACE_AFTER_TWIPS)
                .alignment(Alignment.JUSTIFY)
                .widowControl(true)
                .build();

        // --- Executive summary (no heading; flows straight under the chart) ----
        builder.paragraph(
                "On the morning of 14 May, the checkout service returned elevated error "
                        + "rates for just under two hours, preventing a share of customers "
                        + "from completing purchases. The incident was detected by automated "
                        + "alerting eleven minutes after the first error spike and resolved by "
                        + "rolling back a configuration change that had shipped the previous "
                        + "evening. No data was lost and no customer payment information was "
                        + "exposed at any point.",
                plainBody, bodyParagraph);
        builder.paragraph(
                "This report sets out the detection and response timeline, the root cause "
                        + "as currently understood, the customers affected, and the "
                        + "remediation work already under way. It is written for the wider "
                        + "engineering organisation, not only for the team that owns the "
                        + "checkout service, because two of the follow-up actions touch shared "
                        + "infrastructure other teams depend on.",
                plainBody, bodyParagraph);

        ParagraphStyle headingNoBreak = ParagraphStyle.builder()
                .spaceBeforeTwips(ParagraphStyle.HEADING_SPACE_BEFORE_TWIPS)
                .spaceAfterTwips(ParagraphStyle.HEADING_SPACE_AFTER_TWIPS)
                .keepWithNext(true)
                .keepLines(true)
                .pageBreakBefore(false)
                .build();
        ParagraphStyle headingNewPage = ParagraphStyle.builder()
                .spaceBeforeTwips(ParagraphStyle.HEADING_SPACE_BEFORE_TWIPS)
                .spaceAfterTwips(ParagraphStyle.HEADING_SPACE_AFTER_TWIPS)
                .keepWithNext(true)
                .keepLines(true)
                .pageBreakBefore(true)
                .build();

        // --- Timeline of events -------------------------------------------------
        String timelineColour = "#177E89";
        builder.heading("Timeline of Events", sectionHeadingText(timelineColour), headingNoBreak);
        builder.table(
                List.of("Time (UTC)", "Event", "Owner"),
                List.of(
                        List.of("07:42",
                                "Configuration change deployed to the checkout service "
                                        + "(release 4.118).",
                                "Checkout team"),
                        List.of("08:15",
                                "Error rate on the checkout payment endpoint begins climbing "
                                        + "above baseline.",
                                "—"),
                        List.of("08:26",
                                "Automated alert fires for elevated 5xx rate on "
                                        + "/checkout/payment.",
                                "PagerDuty"),
                        List.of("08:29",
                                "On-call engineer acknowledges the page and begins triage.",
                                "J. Alvarez"),
                        List.of("08:41",
                                "Incident declared; incident channel opened; status page "
                                        + "updated to \"investigating\".",
                                "J. Alvarez"),
                        List.of("08:53",
                                "Root cause narrowed to the 07:42 configuration change; "
                                        + "rollback proposed.",
                                "Checkout team"),
                        List.of("09:05", "Rollback approved and deployment started.",
                                "Checkout team"),
                        List.of("09:19",
                                "Rollback complete; error rate begins returning to baseline.",
                                "Checkout team"),
                        List.of("09:31",
                                "Error rate confirmed at baseline for ten consecutive "
                                        + "minutes; incident downgraded.",
                                "J. Alvarez"),
                        List.of("09:38",
                                "Status page updated to \"resolved\"; incident channel "
                                        + "archived.",
                                "J. Alvarez")),
                sectionTableStyle(timelineColour));

        // --- Detection and response ----------------------------------------------
        String detectionColour = "#843C0C";
        TextStyle detectionText = sectionBodyText(detectionColour);
        builder.heading("Detection and Response", sectionHeadingText(detectionColour), headingNewPage);
        builder.paragraph(
                "The alert that caught this incident was a standard error-rate threshold "
                        + "on the checkout payment endpoint, not anything specific to the "
                        + "change that caused it. It fired within eleven minutes of the first "
                        + "elevated errors, which is within the target for a customer-facing "
                        + "endpoint of this kind, and the on-call engineer acknowledged it "
                        + "within three minutes.",
                detectionText, bodyParagraph);
        builder.paragraph(
                "Triage moved quickly because the deployment log was checked first, before "
                        + "any code-level investigation began. The 07:42 release was the only "
                        + "change in the preceding two hours, which narrowed the search "
                        + "considerably and is the main reason time-to-rollback was as short "
                        + "as it was.",
                detectionText, bodyParagraph);
        builder.paragraph(
                "The rollback itself went smoothly and is not a source of concern. The "
                        + "concern, addressed in the sections that follow, is why a change "
                        + "with this effect was able to reach production without being caught "
                        + "earlier.",
                detectionText, bodyParagraph);

        // --- Root cause analysis --------------------------------------------------
        String rootCauseColour = "#375623";
        TextStyle rootCauseText = sectionBodyText(rootCauseColour);
        builder.heading("Root Cause Analysis", sectionHeadingText(rootCauseColour), headingNewPage);
        builder.paragraph(
                "The proximate cause was straightforward once found; the reasons it was "
                        + "not found sooner are less so, and are the more useful part of this "
                        + "analysis.",
                rootCauseText, bodyParagraph);

        // The paragraph below is the page-split demonstration this sample exists
        // for. "Root Cause Analysis" opens on a fresh page (pageBreakBefore(true)
        // on its heading), so this paragraph can start as early as the very top of
        // an otherwise-empty page — there is no guarantee of only "a little room
        // left" to lean on. Instead it is written long enough (roughly a thousand
        // words) to exceed a full page's text-area capacity on its own, so it is
        // forced to spill onto the next page no matter where on the page it
        // starts. It uses the same widowControl(true)/no-keepLines body style as
        // every other paragraph in this document — keepLines would force the
        // whole thing onto one page and defeat the point.
        String longRootCauseParagraph =
                "The 07:42 release changed the timeout applied to calls from the "
                        + "checkout service to the payment gateway, reducing it from eight "
                        + "hundred milliseconds to two hundred as part of an unrelated "
                        + "effort to tighten latency budgets across the payments path ahead "
                        + "of a planned traffic increase next quarter. In isolation the "
                        + "change was reasonable: the payment gateway's own published "
                        + "service level objective is a ninety-fifth percentile response "
                        + "time under one hundred and fifty milliseconds, and two hundred "
                        + "milliseconds already carries a comfortable margin above that "
                        + "figure under normal conditions. What the change did not account "
                        + "for is that the gateway's response time is not stable across the "
                        + "day; it has a known, well-documented diurnal pattern in which "
                        + "response times rise for roughly ninety minutes around the daily "
                        + "settlement batch, a scheduled job that reconciles the previous "
                        + "day's transactions and runs shortly after seven each morning. "
                        + "During that window the gateway's ninety-fifth percentile "
                        + "response time regularly exceeds two hundred and fifty "
                        + "milliseconds, and has done so consistently for at least the last "
                        + "six months of available metrics, which the team reviewed after "
                        + "the fact. The new two-hundred-millisecond timeout therefore "
                        + "began cutting off a meaningful share of otherwise-successful "
                        + "payment calls every single morning during the settlement window, "
                        + "and this particular morning was unremarkable in that respect; "
                        + "nothing about it was worse than any other weekday. The requests "
                        + "that timed out were retried once by the checkout service's "
                        + "existing retry logic, which is correct behaviour and worked as "
                        + "designed, but a second timeout at two hundred milliseconds "
                        + "during an already-slow window frequently failed as well, and a "
                        + "second failure is surfaced to the customer as a checkout error "
                        + "rather than retried further, which is also correct behaviour "
                        + "given the retry budget in place. The result was not a broken "
                        + "payment gateway and not a broken checkout service in isolation, "
                        + "but an interaction between a latency budget set from an "
                        + "aggregate metric and a load pattern that aggregate metric does "
                        + "not describe. A closer look at the metrics also shows this is "
                        + "not the first time the settlement window has caused trouble: "
                        + "three months earlier, an unrelated batch job that also runs "
                        + "during the same window briefly starved database connections used "
                        + "by an internal reporting service, and was resolved without "
                        + "customer impact only because that service happened to fail "
                        + "closed rather than open. That earlier incident was reviewed "
                        + "internally at the time, but its scope was limited to the "
                        + "reporting service itself and never produced any "
                        + "organisation-wide guidance about the settlement window as a "
                        + "recurring risk period, so the two incidents share a root "
                        + "condition that nobody had connected until this postmortem was "
                        + "being written. Monitoring did not surface today's problem "
                        + "earlier for a related reason: the checkout service's dashboards "
                        + "track error rate and latency as all-day rolling averages, the "
                        + "same shape as the gateway's own published SLO, so a "
                        + "ninety-minute daily spike is smoothed into insignificance "
                        + "against twenty-four hours of otherwise-normal traffic. An "
                        + "on-call engineer glancing at the primary dashboard on any given "
                        + "morning would have seen nothing unusual, because nothing about "
                        + "the aggregate view was unusual; the elevated error rate during "
                        + "the settlement window has, in effect, been present in the data "
                        + "since the timeout change shipped in March, without ever once "
                        + "crossing an alerting threshold tuned for the day as a whole "
                        + "rather than for any particular hour of it. A retrospective query "
                        + "the team ran after the fact puts the affected share at roughly "
                        + "four per cent of checkout attempts during the window on a "
                        + "typical morning, which is a small enough fraction of total daily "
                        + "volume to stay under every existing alert, but a large enough "
                        + "fraction of the settlement window specifically that today's "
                        + "incident was, in some sense, overdue rather than anomalous. It "
                        + "is worth being precise about what that four per cent means in "
                        + "practice: on an average morning it was a few dozen customers "
                        + "seeing a transient error and, in most cases, succeeding on a "
                        + "manual retry a minute or two later, which is unpleasant but not "
                        + "the kind of failure that generates support tickets in volume. "
                        + "Today crossed into a different category only because the "
                        + "settlement batch itself ran roughly eighteen minutes longer than "
                        + "usual, for reasons still being investigated separately by the "
                        + "data platform team, which extended the affected window and "
                        + "pushed the failure rate high enough to trip the standard "
                        + "error-rate alert rather than staying invisible within it. The "
                        + "settlement-window diurnal pattern was known to the payments "
                        + "platform team, who maintain a dashboard that shows it clearly, "
                        + "but it was not known to the checkout team making the timeout "
                        + "change, and there is no automated check today that would have "
                        + "flagged a proposed timeout against that specific historical "
                        + "pattern before the change shipped. The review that approved the "
                        + "change checked the timeout against the gateway's published SLO, "
                        + "which is the standard practice for this kind of change and would "
                        + "ordinarily be sufficient, but the SLO itself is an all-day "
                        + "aggregate and does not capture the settlement window as a "
                        + "distinct period, so a reviewer following the standard practice "
                        + "exactly had no way to catch this from the SLO document alone. "
                        + "None of this points to an individual mistake worth dwelling on. "
                        + "The engineer who made the timeout change followed the documented "
                        + "review process, the reviewer who approved it checked the "
                        + "document that process says to check, and the on-call rotation "
                        + "responded correctly to every signal the monitoring stack was "
                        + "configured to produce; nobody in this sequence of events did "
                        + "anything other than what their role asked of them. The gap sits "
                        + "entirely in the space between those individually reasonable "
                        + "actions: a review process anchored to an all-day SLO, a "
                        + "monitoring stack averaged over that same all-day window, and a "
                        + "scheduled batch job whose effect on the payments path was "
                        + "documented in one place that nobody making this particular kind "
                        + "of change was in the habit of consulting. This is the part of "
                        + "the incident we consider systemic rather than a one-off mistake: "
                        + "the information that would have prevented this existed, was "
                        + "accurate, and was even visible on a dashboard someone on another "
                        + "team looks at regularly, but it was not in a form that surfaced "
                        + "itself to the person making a decision that depended on it.";
        builder.paragraph(longRootCauseParagraph, rootCauseText, bodyParagraph);

        builder.paragraph(
                "Two changes are proposed as a result, both covered under Remediation and "
                        + "Follow-up below: surfacing the settlement-window pattern directly "
                        + "in the latency-budget review checklist, and moving the "
                        + "retry-then-fail behaviour to a circuit breaker that would have "
                        + "contained this to a brief spike rather than a two-hour incident.",
                rootCauseText, bodyParagraph);

        // --- Customer impact --------------------------------------------------
        String impactColour = "#5B2C6F";
        builder.heading("Customer Impact", sectionHeadingText(impactColour), headingNewPage);
        builder.table(
                List.of("Region", "Users affected", "Compensation"),
                List.of(
                        List.of("EMEA", "~2,100", "Automatic 10% credit issued"),
                        List.of("Americas", "~640", "Automatic 10% credit issued"),
                        List.of("APAC", "~310", "Automatic 10% credit issued"),
                        List.of("Enterprise accounts", "0",
                                "Not affected — routed via dedicated gateway pool")),
                sectionTableStyle(impactColour));

        // --- Remediation and follow-up ---------------------------------------------
        String remediationColour = "#B7950B";
        TextStyle remediationText = sectionBodyText(remediationColour);
        builder.heading("Remediation and Follow-up", sectionHeadingText(remediationColour), headingNewPage);
        builder.table(
                List.of("Action", "Owner", "Due date", "Status"),
                List.of(
                        List.of(
                                "Add settlement-window pattern to the latency-budget review "
                                        + "checklist",
                                "Payments platform team", "2026-05-21", "Done"),
                        List.of(
                                "Replace retry-then-fail with a circuit breaker on the "
                                        + "payment gateway client",
                                "Checkout team", "2026-06-04", "In progress"),
                        List.of(
                                "Publish settlement-window metrics as a first-class SLO "
                                        + "dimension, not just a dashboard",
                                "Payments platform team", "2026-06-11", "In progress"),
                        List.of(
                                "Add an automated check that flags new timeouts below the "
                                        + "99th percentile of the last 90 days",
                                "Platform tooling team", "2026-06-18", "Not started"),
                        List.of(
                                "Backfill a synthetic settlement-window load test into the "
                                        + "checkout pre-release suite",
                                "Checkout team", "2026-06-25", "Not started"),
                        List.of("Share this postmortem at the cross-team reliability review",
                                "Incident commander", "2026-05-16", "Done")),
                sectionTableStyle(remediationColour));
        builder.paragraph(
                "None of the follow-up actions above are blocked on each other, and the "
                        + "two rated \"Done\" were completed within a week of the incident. "
                        + "The remaining items are tracked in the reliability backlog under "
                        + "the epic linked below, with the same due dates shown here.",
                remediationText, bodyParagraph);

        // A trailing hyperlink inside the same paragraph as the body text.
        builder.paragraph(
                "The full incident runbook, including the complete alert history and the "
                        + "underlying metrics referenced above, is available online. ",
                Hyperlink.of("View the runbook",
                        "https://example.com/runbooks/checkout-outage-2026-05-14"));

        return builder.build();
    }

    /** Bold 11pt Calibri heading text in the given section colour. */
    private static TextStyle sectionHeadingText(String colour) {
        return TextStyle.builder().font("Calibri").sizePt(11).bold(true).color(colour).build();
    }

    /** Plain 11pt Calibri body text in the given section colour. */
    private static TextStyle sectionBodyText(String colour) {
        return TextStyle.builder().font("Calibri").sizePt(11).bold(false).color(colour).build();
    }

    /** A table bordered in the section colour on the header, hairline grey on the body. */
    private static TableStyle sectionTableStyle(String colour) {
        return TableStyle.builder()
                .headerBorder(TableBorderStyle.builder()
                        .color(colour).widthPt(1.0)
                        .line(BorderLine.SINGLE).edges(Edge.BOTTOM).build())
                .bodyBorder(TableBorderStyle.builder()
                        .color("#BFBFBF").widthPt(0.5)
                        .line(BorderLine.SINGLE).edges(Edge.BOTTOM).build())
                .build();
    }

    private static byte[] resource(String name) {
        try (InputStream in = PostmortemSampleMain.class.getResourceAsStream(name)) {
            if (in == null) {
                throw new IllegalStateException("missing demo resource " + name);
            }
            return in.readAllBytes();
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
    }
}
