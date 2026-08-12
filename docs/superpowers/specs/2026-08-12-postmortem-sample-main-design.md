# docx-service — PostmortemSampleMain

**Date:** 2026-08-12
**Status:** Approved for planning

## Goal

Add a second runnable sample, alongside the existing `SampleMain`, that:

1. Is visibly bigger and different in genre — an incident postmortem report rather than
   a quarterly financial report — so the two samples don't read as duplicates of the
   same content with different words.
2. Deliberately demonstrates a single paragraph splitting across a page boundary, in a
   way that isn't left to chance. `SampleMain`'s body paragraphs are already free to
   split (`widowControl(true)`, no `keepLines`), but nothing in it is written long
   enough, or placed close enough to a page foot, to make a mid-paragraph split
   *reliable* to observe on open. This sample fixes that.

No library code changes. This is content authored against the existing public API
(`WordDocument.Builder`, `TextStyle`, `ParagraphStyle`, `TableStyle`, `Hyperlink`) —
the same API `SampleMain` uses.

## File and naming

- `docx-service/src/test/java/com/example/docx/sample/PostmortemSampleMain.java`
- Same package and source root as `SampleMain` (test sources — stays out of the jar,
  per the existing convention documented in the README's Layout table).
- Output: `target/postmortem-<System.currentTimeMillis()>.docx` — a distinct filename
  pattern from `SampleMain`'s `target/sample-*.docx`, so running both leaves both files
  on disk instead of one overwriting the other.
- No JUnit test is added for it. `SampleMain` has none either; both are manual,
  eyeball-the-output tools by design, not assertable behaviour.

`pom.xml`'s `exec-maven-plugin` default `mainClass` stays `SampleMain` — unchanged, so
`mvn test-compile exec:java` keeps working exactly as documented today. Running the new
sample instead is a one-off override:

```bash
JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn test-compile exec:java \
    -Dexec.mainClass=com.example.docx.sample.PostmortemSampleMain
```

This override line is added to the README's Quick start section, directly under the
existing `exec:java` example.

## Document content

Theme: **"Incident Postmortem: Checkout Service Outage"** — an internal engineering
report. Chosen because postmortems naturally contain both structured data (a timeline,
an impact table) and long unstructured prose (root cause analysis), which is exactly
the mix needed to make the long-paragraph demonstration read as natural content rather
than as a deliberately-inflated wall of text.

Structure, top to bottom:

1. **Title heading** — "Incident Postmortem: Checkout Service Outage", styled the same
   way as `SampleMain`'s title (Calibri Light, 20pt, bold, coloured).
2. **Chart image** — the existing `demo/chart.svg` / `demo/chart.png` pair, reused
   as-is (no new binary assets), captioned in context as the error-rate/request-volume
   graph during the incident window.
3. **Executive Summary** — no heading (flows directly under the image, like
   `SampleMain`'s table does), 2 short paragraphs: what broke, how bad, how long.
4. **Timeline of Events** (heading + table) — columns Time / Event / Owner, ~10 rows,
   walking from first alert to resolution.
5. **Detection and Response** (heading + 2–3 normal-length paragraphs) — how the
   incident was found and the immediate response.
6. **Root Cause Analysis** (heading + paragraphs) — contains the deliberately long
   paragraph. See below.
7. **Customer Impact** (heading + table) — columns Region / Users affected /
   Compensation.
8. **Remediation and Follow-up** (heading + table + 1–2 closing paragraphs) — columns
   Action / Owner / Due date / Status.
9. **Trailing hyperlink paragraph** — "The full incident runbook is available
   online. **View the runbook**", same pattern as `SampleMain`'s trailing hyperlink
   (`builder.paragraph(text, Hyperlink.of(...))`).

Each section heading uses `keepWithNext(true)` + `keepLines(true)` so a heading never
strands alone at a page foot, and `pageBreakBefore(true)` on every section after the
first — the same pattern `SampleMain` uses. One colour per section, shared by that
section's heading and body text — a 6-entry palette defined locally in
`PostmortemSampleMain` (not shared with `SampleMain`, whose palette only has 3 entries).

Sized so the whole document clears 4+ A4 pages at the default margins used in
`SampleMain` (`PageSetup.builder().a4().marginsTwips(851).build()`) — three tables, one
image, and enough body paragraphs across six sections to get there without needing the
long paragraph to carry that weight by itself.

## The page-spanning paragraph

Inside **Root Cause Analysis**:

- A short lead-in paragraph (1–2 sentences) is placed first — deliberately, so the long
  paragraph that follows starts with only a little vertical room left on the page,
  rather than at the top of a fresh page where a shorter document might let it fit
  entirely.
- The long paragraph itself: several hundred words of root-cause narrative (a few
  hundred words longer than any paragraph in `SampleMain`), long enough that even
  starting near the top of a page it would still overflow the remaining text area.
- Styled identically to every other body paragraph — `widowControl(true)`,
  `Alignment.JUSTIFY`, no `keepLines`. Per `DOCUMENTATION.md`, `keepLines` is wrong for
  body text (it would force the whole paragraph onto one page, leaving a gap, or — per
  the "none of these is a guarantee" note — get silently overridden anyway once the
  paragraph is taller than the text area). Leaving `keepLines` off is what lets Word
  split it.
- A code comment directly above the paragraph's text constant names it as the
  page-split demonstration, so it's not just long by accident — a reader of the source
  can find it immediately.

This placement (short paragraph eating most of the remaining page, long paragraph
right after it) makes the split reliable rather than a matter of getting lucky with
where earlier content happened to land.

## Testing / verification

None beyond what `SampleMain` already gets: run the `exec:java` override above, open
the resulting `target/postmortem-*.docx` in Word (or a compatible viewer), and confirm
visually that:

- The document runs 4+ pages.
- Every section after the first starts on a fresh page.
- No heading is stranded alone at the bottom of a page.
- The flagged long paragraph in Root Cause Analysis visibly starts on one page and
  continues onto the next, with no lone stranded line at either end (widow/orphan
  control working as expected).

This is a manual, visual check — consistent with how `SampleMain` itself is verified
today. No automated assertion is added, since page layout is a Word rendering concern
docx4j does not compute and this project does not attempt to simulate.
