# docx-service

Generates Word (`.docx`) documents from a Java backend with
[docx4j](https://www.docx4java.org/). Java 25, Maven, no web framework.

Phase one produces a document with configurable A4 page margins and one
configurable heading. Nothing else.

## Quick start

```bash
JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn test
```

```bash
JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn test-compile exec:java
```

The second writes `target/sample.docx`.

## Using it

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

Both nested builders default every field, so `PageSetup.a4()` and
`HeadingStyle.defaults()` are valid alone.

For large documents prefer `writeTo(out)` over `toByteArray()`: it avoids
buffering a second full copy of the file. `writeTo` does not close the stream
you give it, so a controller keeps ownership of the response.

## Layout

| Package | Contents |
| --- | --- |
| `com.example.docx` | `WordDocument` (facade), `Units`, `DocumentGenerationException` |
| `…​.page` | `PageSetup` — page size and margins |
| `…​.style` | `HeadingStyle` — run formatting |
| `…​.content` | `Headings` — stateless paragraph factory |
| `…​.sample` | Runnable `main` |

`content` and `style` never touch `WordprocessingMLPackage`, so they are pure
functions of their arguments and need no fixtures. The facade owns the
package; `page` only produces a `w:sectPr` fragment for it.

## Things that bite

- **Twips, not pixels.** All page geometry is in twips (1/1440 inch). The
  default margin of 851 twips is ~1.5 cm.
- **851 ≠ `marginsCm(1.5)`.** The cm path computes 850 (1.5 / 2.54 × 1440 =
  850.39). The 0.02 mm difference is invisible, but the numbers are not equal.
- **A4 is a literal `11906 × 16838`.** Deriving 297 mm gives 16839, which is
  not what Word writes.
- **Half-points.** `w:sz` is twice the point size: 20 pt emits `40`.
- **Bare hex.** OOXML rejects `#` in `w:color`; the builder strips it.
- **`xml:space="preserve"`** is set on every `w:t`, or leading and trailing
  spaces vanish silently.
- **Latest docx4j is 17.0.2, not 11.5.x.** `search.maven.org` reports 11.5.3
  and is stale; check `maven-metadata.xml` on repo1 instead. In 17.x the `w:*`
  classes come from `docx4j-generated-objects`, not `docx4j-openxml-objects`.

## Notes

- `WordDocument` is not thread safe. Build one per document.
- The first document in a JVM pays a one-off ~1s JAXB context initialisation.
  Warm it at startup with `Context.getWmlObjectFactory()` if latency matters.
- The heading uses direct formatting only and carries no `Heading1` style id,
  so it does not appear in Word's Navigation pane. That is a deliberate
  phase-one trade-off.
