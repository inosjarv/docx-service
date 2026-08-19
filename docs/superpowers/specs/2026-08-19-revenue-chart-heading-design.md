# docx-service — Small heading above the chart in SampleMain

**Date:** 2026-08-19
**Status:** Approved for planning

## Goal

`SampleMain`'s chart image currently has nothing labelling it — the title heading
("Quarterly Report") is immediately followed by the SVG image with no indication of
what the chart shows. Add a small heading directly above it.

`PostmortemSampleMain` is out of scope: it already has a caption below its image,
added in a previous session, and the user confirmed only `SampleMain` needs this
change.

## Change

In `SampleMain.main()`, insert one more `.heading(...)` call in the existing builder
chain, between the title heading and `.svgImage(...)`:

```java
var builder = WordDocument.builder()
        .pageSetup(PageSetup.builder().a4().marginsTwips(851).build())
        .heading("Quarterly Report", TextStyle.builder()
                .font("Calibri Light")
                .sizePt(20)
                .bold(true)
                .italic(false)
                .color("#1F4E79")
                .build())
        .heading("Revenue Trend", TextStyle.builder()
                .font("Calibri")
                .sizePt(11)
                .bold(true)
                .color("#1F4E79")
                .build())
        .svgImage(resource("/demo/chart.svg"), resource("/demo/chart.png"))
        .table(...)
```

- Text: `"Revenue Trend"`.
- Style: Calibri 11pt bold, `#1F4E79` — the same size/weight/colour as the section
  headings later in the file (`SECTION_TITLES` loop), so it reads as a minor label
  rather than competing with the 20pt document title above it.
- Paragraph style: the two-argument `.heading(text, style)` overload, which defaults
  to `ParagraphStyle.heading()` — standard heading spacing, `keepWithNext`,
  `keepLines`, `widowControl`. No `pageBreakBefore` — it must flow directly under the
  title, not start a new page.

No other files change. No library code changes. No new tests — matches the existing
convention that `SampleMain` and `PostmortemSampleMain` are manual, eyeball-the-output
samples, not assertable behaviour.

## Verification

Run `mvn test-compile exec:java` and open the resulting `target/sample-*.docx`:
confirm "Revenue Trend" appears directly under "Quarterly Report" and directly above
the chart image, in the smaller bold navy style described above.
