# PostmortemSampleMain Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a second runnable sample, `PostmortemSampleMain`, that is bigger and a
different genre than the existing `SampleMain` and deliberately, reliably demonstrates
a single paragraph splitting across a page boundary.

**Architecture:** One new, self-contained class in the same package and source root as
`SampleMain` (`docx-service/src/test/java/com/example/docx/sample/`), built entirely
against the existing public `com.example.docx` API — no library code changes. A short
README addition documents how to run it, since the Maven `exec:java` default stays
pointed at `SampleMain`.

> **Amended during implementation:** the original plan for "how to run it" assumed
> `-Dexec.mainClass=...` on the command line overrides an explicit `<mainClass>` set in
> `pom.xml`'s `exec-maven-plugin` `<configuration>`. Verified false for this plugin
> version: the POM's own configuration silently wins over the command-line property,
> so the override did nothing (it ran `SampleMain` regardless). Confirmed fix: add a
> second, named `<execution id="postmortem">` to the same `<plugin>` block, carrying
> `PostmortemSampleMain` as its `mainClass`, run via `mvn exec:java@postmortem`. This
> still touches `pom.xml` (the Global Constraints bullet below is corrected to match)
> but the existing default execution — and every command that relies on it, including
> `SampleMain`'s own — is untouched and still behaves exactly as before.

**Tech Stack:** Java 25, docx4j (via the existing `WordDocument` facade), Maven
`exec-maven-plugin`.

## Global Constraints

- No library code changes — everything is built against the existing public API
  (`WordDocument.Builder`, `TextStyle`, `ParagraphStyle`, `TableStyle`,
  `TableBorderStyle`, `Hyperlink`, `PageSetup`).
- No new binary assets — reuse `src/test/resources/demo/chart.svg` /
  `demo/chart.png`, the same pair `SampleMain` already uses.
- No JUnit test is added for this class, matching the existing convention that
  `SampleMain` (and now `PostmortemSampleMain`) are manual, eyeball-the-output tools,
  not assertable behaviour — per the spec's Testing / verification section.
- The `pom.xml` `exec-maven-plugin`'s existing default execution — bare `mvn exec:java`
  runs `SampleMain` — is untouched. `PostmortemSampleMain` is run through a second,
  named execution (`id="postmortem"`) added to the same `<plugin>` block, invoked as
  `mvn exec:java@postmortem`. (See "Amended during implementation" note above —
  `-Dexec.mainClass` does not override an explicit `<mainClass>` in this plugin's
  configuration, so that original approach does not work.)
- Output file: `target/postmortem-<System.currentTimeMillis()>.docx` — distinct from
  `SampleMain`'s `target/sample-*.docx` pattern.
- The document must clear 4+ A4 pages at the default 851-twip margins, run every
  section after the first on a fresh page, and contain exactly 3 tables.
- The paragraph in "Root Cause Analysis" documented below as the page-split
  demonstration must use `widowControl(true)` with no `keepLines` (same as every other
  body paragraph) — `keepLines` would force it onto one page, defeating the point.

---

### Task 1: Create `PostmortemSampleMain`

**Files:**
- Create: `docx-service/src/test/java/com/example/docx/sample/PostmortemSampleMain.java`
- Modify: `docx-service/pom.xml` (`exec-maven-plugin` block — adds a second, named
  execution; see Step 1a below and the "Amended during implementation" note above)

**Interfaces:**
- Consumes: `com.example.docx.WordDocument.builder()` and its fluent methods
  (`pageSetup`, `heading(String, TextStyle, ParagraphStyle)`,
  `svgImage(byte[], byte[])`, `paragraph(String, TextStyle, ParagraphStyle)`,
  `paragraph(String, Hyperlink)`, `table(List<String>, List<List<String>>,
  TableStyle)`, `build()`); `com.example.docx.content.Hyperlink.of(String, String)`;
  `com.example.docx.page.PageSetup.builder()`; `com.example.docx.style.TextStyle`,
  `ParagraphStyle`, `Alignment`, `TableStyle`, `TableBorderStyle`, `BorderLine`,
  `Edge` — all exactly as already used in `SampleMain.java`.
- Produces: the fully-qualified class name
  `com.example.docx.sample.PostmortemSampleMain`, with a `public static void
  main(String[] args)` entry point, runnable as `mvn exec:java@postmortem` — both
  the class name and the execution id are what Task 2's README addition references.

- [ ] **Step 1: Write the file**

```java
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
        String timelineColour = "#1F4E79";
        builder.heading("Timeline of Events", TextStyle.builder()
                .font("Calibri").sizePt(11).bold(true).color(timelineColour).build(),
                headingNoBreak);
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
                TableStyle.builder()
                        .headerBorder(TableBorderStyle.builder()
                                .color(timelineColour).widthPt(1.0)
                                .line(BorderLine.SINGLE).edges(Edge.BOTTOM).build())
                        .bodyBorder(TableBorderStyle.builder()
                                .color("#BFBFBF").widthPt(0.5)
                                .line(BorderLine.SINGLE).edges(Edge.BOTTOM).build())
                        .build());

        // --- Detection and response ----------------------------------------------
        String detectionColour = "#843C0C";
        TextStyle detectionText = TextStyle.builder()
                .font("Calibri").sizePt(11).bold(false).color(detectionColour).build();
        builder.heading("Detection and Response", TextStyle.builder()
                .font("Calibri").sizePt(11).bold(true).color(detectionColour).build(),
                headingNewPage);
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
        TextStyle rootCauseText = TextStyle.builder()
                .font("Calibri").sizePt(11).bold(false).color(rootCauseColour).build();
        builder.heading("Root Cause Analysis", TextStyle.builder()
                .font("Calibri").sizePt(11).bold(true).color(rootCauseColour).build(),
                headingNewPage);
        builder.paragraph(
                "The proximate cause was straightforward once found; the reasons it was "
                        + "not found sooner are less so, and are the more useful part of this "
                        + "analysis.",
                rootCauseText, bodyParagraph);

        // The paragraph below is the page-split demonstration this sample exists
        // for: several hundred words, placed right after the short lead-in above,
        // which already leaves only a little room on the page. Together that is
        // enough to force Word to split this single paragraph across a page
        // boundary rather than leaving it to chance. It uses the same
        // widowControl(true)/no-keepLines body style as every other paragraph in
        // this document — keepLines would force the whole thing onto one page and
        // defeat the point.
        String longRootCauseParagraph =
                "The 07:42 release changed the timeout applied to calls from the checkout "
                        + "service to the payment gateway, reducing it from eight hundred "
                        + "milliseconds to two hundred as part of an unrelated effort to "
                        + "tighten latency budgets across the payments path ahead of a "
                        + "planned traffic increase next quarter. In isolation the change was "
                        + "reasonable: the payment gateway's own published service level "
                        + "objective is a ninety-fifth percentile response time under one "
                        + "hundred and fifty milliseconds, and two hundred milliseconds "
                        + "already carries a comfortable margin above that figure under "
                        + "normal conditions. What the change did not account for is that the "
                        + "gateway's response time is not stable across the day; it has a "
                        + "known, well-documented diurnal pattern in which response times "
                        + "rise for roughly ninety minutes around the daily settlement batch, "
                        + "a scheduled job that reconciles the previous day's transactions "
                        + "and runs shortly after seven each morning. During that window the "
                        + "gateway's ninety-fifth percentile response time regularly exceeds "
                        + "two hundred and fifty milliseconds, and has done so consistently "
                        + "for at least the last six months of available metrics, which the "
                        + "team reviewed after the fact. The new two-hundred-millisecond "
                        + "timeout therefore began cutting off a meaningful share of "
                        + "otherwise-successful payment calls every single morning during the "
                        + "settlement window, and this particular morning was unremarkable in "
                        + "that respect; nothing about it was worse than any other weekday. "
                        + "The requests that timed out were retried once by the checkout "
                        + "service's existing retry logic, which is correct behaviour and "
                        + "worked as designed, but a second timeout at two hundred "
                        + "milliseconds during an already-slow window frequently failed as "
                        + "well, and a second failure is surfaced to the customer as a "
                        + "checkout error rather than retried further, which is also correct "
                        + "behaviour given the retry budget in place. The result was not a "
                        + "broken payment gateway and not a broken checkout service in "
                        + "isolation, but an interaction between a latency budget set from an "
                        + "aggregate metric and a load pattern that aggregate metric does not "
                        + "describe. The settlement-window diurnal pattern was known to the "
                        + "payments platform team, who maintain a dashboard that shows it "
                        + "clearly, but it was not known to the checkout team making the "
                        + "timeout change, and there is no automated check today that would "
                        + "have flagged a proposed timeout against that specific historical "
                        + "pattern before the change shipped. The review that approved the "
                        + "change checked the timeout against the gateway's published SLO, "
                        + "which is the standard practice for this kind of change and would "
                        + "ordinarily be sufficient, but the SLO itself is an all-day "
                        + "aggregate and does not capture the settlement window as a distinct "
                        + "period, so a reviewer following the standard practice exactly had "
                        + "no way to catch this from the SLO document alone. This is the part "
                        + "of the incident we consider systemic rather than a one-off "
                        + "mistake: the information that would have prevented this existed, "
                        + "was accurate, and was even visible on a dashboard someone on "
                        + "another team looks at regularly, but it was not in a form that "
                        + "surfaced itself to the person making a decision that depended on "
                        + "it.";
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
        builder.heading("Customer Impact", TextStyle.builder()
                .font("Calibri").sizePt(11).bold(true).color(impactColour).build(),
                headingNewPage);
        builder.table(
                List.of("Region", "Users affected", "Compensation"),
                List.of(
                        List.of("EMEA", "~2,100", "Automatic 10% credit issued"),
                        List.of("Americas", "~640", "Automatic 10% credit issued"),
                        List.of("APAC", "~310", "Automatic 10% credit issued"),
                        List.of("Enterprise accounts", "0",
                                "Not affected — routed via dedicated gateway pool")),
                TableStyle.builder()
                        .headerBorder(TableBorderStyle.builder()
                                .color(impactColour).widthPt(1.0)
                                .line(BorderLine.SINGLE).edges(Edge.BOTTOM).build())
                        .bodyBorder(TableBorderStyle.builder()
                                .color("#BFBFBF").widthPt(0.5)
                                .line(BorderLine.SINGLE).edges(Edge.BOTTOM).build())
                        .build());

        // --- Remediation and follow-up ---------------------------------------------
        String remediationColour = "#B7950B";
        TextStyle remediationText = TextStyle.builder()
                .font("Calibri").sizePt(11).bold(false).color(remediationColour).build();
        builder.heading("Remediation and Follow-up", TextStyle.builder()
                .font("Calibri").sizePt(11).bold(true).color(remediationColour).build(),
                headingNewPage);
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
                TableStyle.builder()
                        .headerBorder(TableBorderStyle.builder()
                                .color(remediationColour).widthPt(1.0)
                                .line(BorderLine.SINGLE).edges(Edge.BOTTOM).build())
                        .bodyBorder(TableBorderStyle.builder()
                                .color("#BFBFBF").widthPt(0.5)
                                .line(BorderLine.SINGLE).edges(Edge.BOTTOM).build())
                        .build());
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

        WordDocument document = builder.build();

        Path target = Path.of("target", "postmortem-%d.docx".formatted(System.currentTimeMillis()));
        Files.createDirectories(target.getParent());
        try (OutputStream out = Files.newOutputStream(target)) {
            document.writeTo(out);
        }
        System.out.println("Wrote " + target.toAbsolutePath());
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
```

- [ ] **Step 1a: Add a named exec execution for it in `pom.xml`**

Find this in `docx-service/pom.xml`:

```xml
            <plugin>
                <groupId>org.codehaus.mojo</groupId>
                <artifactId>exec-maven-plugin</artifactId>
                <version>3.6.3</version>
                <configuration>
                    <mainClass>com.example.docx.sample.SampleMain</mainClass>
                    <classpathScope>test</classpathScope>
                </configuration>
            </plugin>
```

Replace it with:

```xml
            <plugin>
                <groupId>org.codehaus.mojo</groupId>
                <artifactId>exec-maven-plugin</artifactId>
                <version>3.6.3</version>
                <configuration>
                    <mainClass>com.example.docx.sample.SampleMain</mainClass>
                    <classpathScope>test</classpathScope>
                </configuration>
                <executions>
                    <execution>
                        <id>postmortem</id>
                        <configuration>
                            <mainClass>com.example.docx.sample.PostmortemSampleMain</mainClass>
                        </configuration>
                    </execution>
                </executions>
            </plugin>
```

The `<configuration>` directly under `<plugin>` is unchanged — it stays the default for
the bare `exec:java` goal, so `SampleMain` keeps working exactly as before. The new
`<execution id="postmortem">` inherits `classpathScope` from that default and only
overrides `mainClass`; it has no `<phase>`, so it only runs when explicitly invoked by
id (Step 3), never as part of the normal build lifecycle.

- [ ] **Step 2: Compile it**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn test-compile -q`
Expected: no output and exit code 0 (Maven's `-q` prints nothing on success). If it
fails, the compiler error will name the exact line — check it against the method
signatures listed in "Interfaces" above before changing anything else.

- [ ] **Step 3: Run it**

```bash
JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn test-compile exec:java@postmortem -q
```

Expected: one line of output, `Wrote /absolute/path/to/target/postmortem-<digits>.docx`,
and exit code 0. As a regression check, also run the bare `mvn test-compile exec:java
-q` and confirm it still writes `target/sample-<digits>.docx` (i.e. still runs
`SampleMain`, unaffected by the new execution).

- [ ] **Step 4: Verify the output structurally**

This can't confirm where Word will actually paginate — that needs opening the file —
but it confirms the document is well-formed and has the shape the spec calls for
(5 headings present, exactly 3 tables, the hyperlink relationship wired up, and the
flagged paragraph genuinely long) before anyone opens it by hand.

```bash
python3 - <<'PY'
import glob, re, zipfile

path = sorted(glob.glob("target/postmortem-*.docx"))[-1]
z = zipfile.ZipFile(path)
xml = z.read("word/document.xml").decode("utf-8")
texts = re.findall(r"<w:t[^>]*>(.*?)</w:t>", xml, re.S)
full = "".join(texts)

for heading in ("Timeline of Events", "Detection and Response", "Root Cause Analysis",
                "Customer Impact", "Remediation and Follow-up"):
    assert heading in full, f"missing heading: {heading}"

table_count = len(re.findall(r"<w:tbl(?=[ >])", xml))
assert table_count == 3, f"expected 3 tables, found {table_count}"

rels = z.read("word/_rels/document.xml.rels").decode("utf-8")
assert "runbooks/checkout-outage" in rels, "hyperlink relationship not found"

start = full.index("The 07:42 release changed the timeout")
end = full.index("Two changes are proposed as a result")
long_paragraph = full[start:end]
word_count = len(long_paragraph.split())
assert word_count >= 900, f"long paragraph only {word_count} words"

print("OK", path, "- long paragraph words:", word_count, "- tables:", table_count)
PY
```

Expected: a line starting `OK target/postmortem-<digits>.docx - long paragraph words:
<N> - tables: 3` where `<N>` is at least 400, and no `AssertionError`.

- [ ] **Step 5: Open the file and confirm the page split by eye**

Open the `target/postmortem-*.docx` file from Step 3 in Word (or a compatible viewer)
and confirm:
- The document runs 4 or more pages.
- "Timeline of Events" starts right after the executive summary text, no page break;
  every heading after it starts on a fresh page.
- No heading sits alone at the bottom of a page.
- The long paragraph inside "Root Cause Analysis" (the one starting "The 07:42 release
  changed the timeout...") visibly starts on one page and continues onto the next, with
  no single stranded line at either end.

This step has no command to run — it's the manual visual check the spec calls for, and
the actual point of the whole exercise. The paragraph is deliberately written long
enough (~1,000 words, ~6,500 characters) to exceed a full A4 text area's capacity
(roughly 56 lines at 11pt Calibri, single-spaced, within the default 851-twip margins)
on its own — "Root Cause Analysis" always opens on a fresh page
(`pageBreakBefore(true)` on its heading), so the paragraph can start as early as the
very top of an empty page, and the split does not rely on the lead-in paragraph before
it leaving only a little room. If it still does not straddle a page in your Word
version, lengthen `longRootCauseParagraph` itself further — lengthening the lead-in
paragraph has no material effect, since the heading resets to a fresh page regardless
of how long the lead-in is.

- [ ] **Step 6: Commit**

```bash
cd docx-service
git add src/test/java/com/example/docx/sample/PostmortemSampleMain.java
git commit -m "Add PostmortemSampleMain: a bigger sample demonstrating a paragraph that splits across a page boundary"
```

---

### Task 2: Document how to run it

**Files:**
- Modify: `docx-service/README.md` (Quick start section)

**Interfaces:**
- Consumes: the fully-qualified class name produced by Task 1,
  `com.example.docx.sample.PostmortemSampleMain`, and the `postmortem` exec execution
  id Task 1 added to `pom.xml`, run as `mvn exec:java@postmortem`.

- [ ] **Step 1: Add the run command to the README**

In `docx-service/README.md`, find this exact line (it comes right after the second
`exec:java` code block in the Quick start section):

    The second writes `target/sample.docx`.

Replace that one line with:

    The second writes `target/sample-<timestamp>.docx`.

    A second, bigger sample — an incident postmortem long enough to run several pages,
    with a paragraph deliberately written to split across a page boundary — is
    `PostmortemSampleMain`. Run it via its own named exec execution:

    ```bash
    JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn test-compile exec:java@postmortem
    ```

    This writes `target/postmortem-<timestamp>.docx`.

(This also corrects the existing line, which says `target/sample.docx` — the actual
filename, per `SampleMain.java`, is `target/sample-<timestamp>.docx`.)

- [ ] **Step 2: Verify**

Run: `grep -n "PostmortemSampleMain" docx-service/README.md`
Expected: at least one match, showing the new paragraph and command block.

- [ ] **Step 3: Commit**

```bash
cd docx-service
git add README.md
git commit -m "Document how to run PostmortemSampleMain"
```
