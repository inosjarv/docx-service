# docx-service — Phase One Design

**Date:** 2026-08-10
**Status:** Approved for planning

## Goal

A Java library that generates a Word `.docx` from a backend. Phase one produces a
document with configurable page margins and exactly one configurable heading —
nothing else.

Scope is deliberate. Phase one exists to prove the docx4j plumbing renders what we
expect and to establish the layering that later phases extend. Paragraphs, tables,
images, headers, footers, lists and templates are all out of scope.

## Context

`docx-generator/` already exists in this workspace as a ~2,400-line docx4j library.
It is left untouched. `docx-service/` is a new, independent module; `docx-generator`
serves only as a reference for docx4j pitfalls already discovered.

## Deliverable

A Maven JAR plus a runnable `main` that writes a sample file to disk for visual
inspection. No web framework, no HTTP layer.

## Stack

| Component | Version | Notes |
| --- | --- | --- |
| Java | 25 | `/usr/bin/java` → Temurin 25.0.1 |
| Maven compiler | 3.15.0 | `<maven.compiler.release>25</maven.compiler.release>` |
| docx4j | 11.5.3 | `org.docx4j:docx4j-JAXB-ReferenceImpl` |
| JUnit | 6.1.3 | via `org.junit:junit-bom` |
| Surefire | 3.5.6 | |
| exec-maven-plugin | 3.6.3 | runs the sample `main` |
| slf4j-api | 2.0.17 | `slf4j-simple` at `runtime` scope, for the sample only |

`docx4j-JAXB-ReferenceImpl` pulls in `docx4j-core` and the Glassfish JAXB runtime.
The bare `org.docx4j:docx4j` artifact is abandoned at 6.1.2 and must not be used.
Inside an application server that already provides JAXB, swap the dependency for
`docx4j-JAXB-MOXy`.

`release=25` fixes the bytecode and API level regardless of which JVM Maven itself
runs on. Toolchain pinning is intentionally not configured.

## Architecture

Package root `com.example.docx`. Seven files:

```
com.example.docx
├── WordDocument.java              fluent facade; the only mutable, stateful class
├── Units.java                     twips ↔ cm/inch, points → half-points
├── DocumentGenerationException.java  the module's single unchecked exception
├── page/
│   └── PageSetup.java             immutable; page size + 4 margins → w:sectPr
├── style/
│   └── HeadingStyle.java          immutable; font, size, bold, italic, colour → w:rPr
├── content/
│   └── Headings.java              stateless factory: (text, style) → org.docx4j.wml.P
└── sample/
    └── SampleMain.java            writes target/sample.docx
```

### The layering rule

`content/` and `style/` never reference `WordprocessingMLPackage`. They are pure
functions of their arguments returning JAXB objects, so they compose freely and need
no test fixtures. `page/` and the facade own the package, because they create parts
and relationships.

This is what makes the module extensible: phase two adds `content/Paragraphs`,
`content/Tables` and `style/ParagraphStyle` as new files, without editing anything
that already exists.

### Component responsibilities

**`Units`** — static conversions only. `cmToTwips`, `inchesToTwips`, `pointsToHalfPoints`.
Rounds half-up to the nearest whole twip.

**`PageSetup`** — immutable value carrying page width, page height and four margins,
all in twips. Built through a nested builder. Knows how to emit a `w:sectPr`.

**`HeadingStyle`** — immutable value carrying font family, size in points, bold,
italic and colour. Built through a nested builder. Knows how to emit a `w:rPr`.

**`Headings`** — one static method producing a `w:p` containing a single `w:r` with
the heading text and the style's `w:rPr`.

**`DocumentGenerationException`** — unchecked. The single exception type the module
throws, for both invalid configuration and wrapped docx4j failures.

**`WordDocument`** — fluent facade. Creates the package, applies the `sectPr`, appends
the heading paragraph, and serialises. Wraps mutable docx4j state and is **not thread
safe**; build one per document.

## API

```java
byte[] docx = WordDocument.builder()
        .pageSetup(PageSetup.builder()
                .a4()
                .marginsTwips(851)
                .build())
        .heading("Quarterly Report", HeadingStyle.builder()
                .font("Calibri Light")
                .sizePt(20)
                .bold(true)
                .italic(false)
                .color("#1F4E79")
                .build())
        .build()
        .toByteArray();
```

Every field on both nested builders is defaulted, so `PageSetup.a4()` and
`HeadingStyle.defaults()` are valid one-liners.

`PageSetup.Builder` offers `a4()`, `marginsTwips(int)`, `marginsCm(double)`,
`marginsInches(double)`, and per-side `top()`, `right()`, `bottom()`, `left()`.

`WordDocument` exposes both `toByteArray()` and `writeTo(OutputStream)`. The latter
lets a controller stream to the response without the whole file sitting in heap.

## Defaults

| Setting | Default | Twips |
| --- | --- | --- |
| Page size | A4 portrait | 11906 × 16838 |
| Margins, all four sides | 851 twips (≈ 1.5 cm) | 851 |
| Font | Calibri Light | — |
| Size | 20 pt | `sz` = 40 half-points |
| Bold | true | — |
| Italic | false | — |
| Colour | `1F4E79` | — |

Note on units: the value 851 is **twips** (1/1440 inch), not pixels. OOXML expresses
all page geometry in twips; 851 pixels at 96 DPI would be 8.86 inches, wider than an
entire A4 page.

### Constants are literals, not computed

Two rounding facts, both of which will silently fail the round-trip test if the
implementer derives the values instead:

- **A4 is the literal pair `11906 × 16838`**, the exact values Word writes. Deriving
  it gives 297 mm → 297 / 25.4 × 1440 = 16838.7 → **16839**, which is off by one.
  `PageSetup.a4()` must hardcode both numbers.
- **851 twips is 1.5009 cm, not exactly 1.5 cm.** `marginsCm(1.5)` correctly computes
  850 (1.5 / 2.54 × 1440 = 850.39). The default margin is specified as the literal
  `851` so it matches the value Word produces; the 0.02 mm difference from 850 is
  invisible in print. The defaults, the sample and the round-trip test all use
  `marginsTwips(851)` and assert 851. Do not assume `marginsCm(1.5)` yields it.

`Units` conversions are therefore used only for caller-supplied cm and inch values,
never to produce the built-in constants.

## Encoding rules

Three conversions are silent-corruption traps and are handled in exactly one place each:

- **Colour** is stored and emitted as bare `RRGGBB`. OOXML rejects a leading `#`.
  The builder accepts `#1F4E79` or `1F4E79`, case-insensitive, and normalises.
- **Font size** is stored in half-points. `sz` = `sizePt × 2`, so 20 pt emits `40`.
  Both `w:sz` and `w:szCs` are set, so the size holds for complex-script runs too.
- **Text** is written with `xml:space="preserve"` on every `w:t`. Without it, leading
  and trailing spaces vanish silently.

The heading also sets `w:rFonts` on both `ascii` and `hAnsi`.

The heading uses direct formatting only. It carries no built-in `Heading1` style id,
so it will not appear in Word's Navigation pane. That is an accepted phase-one
trade-off: the requirement is full control over font, size, weight, slant and colour,
and direct formatting delivers that identically in every renderer regardless of the
document stylesheet. Adding a style id alongside is a phase-two option.

## Error handling

`DocumentGenerationException extends RuntimeException` is the module's single
exception type. Everything this library throws deliberately is an instance of it, so
a caller wraps a generation call in one `catch` and is done.

**Invalid configuration**, detected in the nested builders and thrown from `build()`:
blank or null heading text; font size not greater than zero; negative margins;
margins whose sum exceeds the corresponding page dimension; malformed hex colour;
blank font family. Validating here means bad input fails before docx4j is touched, so
the stack trace points at calling code rather than at JAXB internals. A malformed
colour is rejected rather than silently defaulted to black.

**docx4j failure**, wrapped by the facade: `createPackage()` and `save()` throw
checked `Docx4JException`. Propagating it would force `throws` onto every method in
the fluent chain and make the API unusable. It is wrapped, always preserving the
cause.

Validation failures carry a message naming the offending field and value, and no
cause. Wrapped docx4j failures always carry a cause. That distinction is what a
caller inspects when it needs to tell the two apart; no separate subclass is
introduced for it.

Unchecked, because neither case is recoverable at the call site — a malformed colour
is a programming error, and a docx4j serialisation failure is an environment problem.

## Testing

JUnit 6.1.3, three tiers.

**Value tests**, no docx4j involved:
- `Units`: 1.5 cm → 850 twips; 2.54 cm → 1440 twips; 1.0 inch → 1440 twips;
  20 pt → 40 half-points.
- `PageSetup.a4()` reports width 11906 and height 16838 exactly.
- `HeadingStyle`: hex normalisation accepts both `#RRGGBB` and `RRGGBB`, any case.
- Every validation rejection listed under Error handling throws
  `DocumentGenerationException`, with a message naming the offending field.

**Round-trip test** — the one that matters. Build the document, then
`WordprocessingMLPackage.load()` the produced bytes back and assert on the reloaded
object: `sectPr` margins are 851 on all four sides, page size is 11906 × 16838, the
heading run's `rPr` carries `sz` 40, bold true, colour `1F4E79`, and the text matches.
This proves the file is valid and re-readable. The characteristic OOXML failure is a
document Word refuses to open, which asserting on setters cannot catch.

**Stream equivalence**: `toByteArray()` and `writeTo(OutputStream)` produce identical
bytes.

## Running it

```bash
mvn test
```

```bash
mvn compile exec:java
```

The second writes `target/sample.docx` from `SampleMain`.

## Known characteristics

- `WordDocument` is not thread safe. One instance per document; that is the natural
  lifetime in a request handler and construction is cheap.
- The first document generated in a JVM pays a one-off JAXB context initialisation of
  roughly one second. If latency matters, warm it at startup with
  `Context.getWmlObjectFactory()`.

## Out of scope for phase one

Paragraphs, body text, tables, images, lists and numbering, headers and footers, page
numbers, table of contents, multiple headings, heading levels, template filling,
document metadata, landscape orientation, and any HTTP or Spring layer.
