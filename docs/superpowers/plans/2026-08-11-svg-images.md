# SVG Images with PNG Fallback — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Close the nine phase-one review findings, then insert an SVG into the document with a PNG fallback that Word Online and older Word render automatically.

**Architecture:** Word stores an SVG as a PNG blip carrying the SVG as an `asvg:svgBlip` extension, so the fallback is structural rather than conditional. The caller supplies both the SVG and the PNG, so the library gains no new dependency. Image handling lives in a new `part/` package, because an image is a package part plus a relationship and therefore cannot honour the purity rule that governs `content/` and `style/`.

**Tech Stack:** Java 25, Maven, docx4j 17.0.2, JUnit 6.1.3, slf4j 2.0.18, `javax.imageio` (JDK).

## Global Constraints

- Java release level is exactly `25`. Run Maven as `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test`.
- **No new dependencies.** Batik was evaluated and rejected (18 extra jars). PNG dimensions come from `javax.imageio.ImageIO`, which is in the JDK.
- All page geometry is in **twips**. A4 is the literal pair `11906 × 16838`; the default margin is the literal `851`. `marginsCm(1.5)` yields 850, not 851 — never conflate them.
- **1 twip = exactly 635 EMU** (914400 ÷ 1440). Integer arithmetic, no rounding.
- `DocumentGenerationException extends RuntimeException` is the ONLY exception thrown deliberately. Validation failures carry a message naming the field and NO cause; wrapped failures ALWAYS carry a cause.
- Layering: only `page/`, `part/` and `WordDocument` may reference `WordprocessingMLPackage`. `content/` and `style/` stay pure.
- **`DocumentBuilderFactory.setNamespaceAware(true)` is mandatory** when building the `asvg:svgBlip` element. It is `false` by default, and without it the element lands in no namespace and Word silently ignores the SVG.
- **docx4j 17.0.2 has no `ImageSvgPart`.** `createImagePart(..., "image/svg+xml", ...)` returns a `DefaultXmlPart` and `ClassCastException`s. The SVG part must be built by hand as a `BinaryPart`.
- **The PNG is the primary `r:embed`; the SVG is only the extension.** Reversing them breaks every renderer.
- The SVG extension URI is `{96DAC541-7B7A-43D3-8B79-37D633B846F1}`; the namespace is `http://schemas.microsoft.com/office/drawing/2016/SVG/main`.
- Demo assets already exist at `src/test/resources/demo/chart.svg` and `demo/chart.png` (1600×1120). Do not regenerate them.
- The published JAR must contain library code only — no demo assets, no logging config, no sample class.

## File Structure

| File | Change | Responsibility |
| --- | --- | --- |
| `pom.xml` | Modify | slf4j-simple → `test`; exec plugin → `classpathScope=test` |
| `src/main/resources/simplelogger.properties` | Move → `src/test/resources/` | Sample logging only |
| `src/main/java/com/example/docx/sample/SampleMain.java` | Move → `src/test/java/…` | Demo entry point, not API |
| `src/main/java/com/example/docx/WordDocument.java` | Modify | Non-closing stream wrapper; `svgImage` builder method |
| `src/main/java/com/example/docx/Units.java` | Modify | `twipsToEmu`; overflow guard |
| `src/main/java/com/example/docx/page/PageSetup.java` | Modify | `usableWidthTwips()`; `long` margin-sum checks |
| `src/main/java/com/example/docx/style/HeadingStyle.java` | Modify | Separate `HpsMeasure`s; `Locale` import; size ceiling |
| `src/main/java/com/example/docx/part/ImageParts.java` | Create | PNG part + SVG part + blip extension → `w:p` |
| `src/test/java/com/example/docx/StreamOwnershipTest.java` | Create | Caller's stream stays open |
| `src/test/java/com/example/docx/part/ImagePartsTest.java` | Create | Media parts, namespace, extent sizing |
| `README.md` | Modify | Correct layering and heap claims; document SVG 1.1 requirement |

---

### Task 1: Stop the library JAR shipping non-library files

**Files:**
- Modify: `pom.xml`
- Move: `src/main/resources/simplelogger.properties` → `src/test/resources/simplelogger.properties`
- Move: `src/main/java/com/example/docx/sample/SampleMain.java` → `src/test/java/com/example/docx/sample/SampleMain.java`

**Interfaces:**
- Consumes: nothing.
- Produces: `SampleMain` is now a test-source class; `mvn compile exec:java` runs with the test classpath.

`slf4j-simple` at `runtime` scope propagates transitively, so every consumer of this library gets it on their classpath and SLF4J may bind it instead of their own Logback or Log4j. `simplelogger.properties` in `src/main/resources` lands at the JAR root, where `SimpleLogger` reads it and applies this library's log level to their application. Both are packaging defects, not preferences.

- [ ] **Step 1: Move the two files**

```bash
git mv src/main/resources/simplelogger.properties src/test/resources/simplelogger.properties
mkdir -p src/test/java/com/example/docx/sample
git mv src/main/java/com/example/docx/sample/SampleMain.java src/test/java/com/example/docx/sample/SampleMain.java
rmdir src/main/resources src/main/java/com/example/docx/sample 2>/dev/null || true
```

- [ ] **Step 2: Add the trailing newline the file is missing**

```bash
printf 'org.slf4j.simpleLogger.defaultLogLevel=warn\n' > src/test/resources/simplelogger.properties
```

- [ ] **Step 3: Rescope slf4j-simple and give exec the test classpath**

In `pom.xml`, change the `slf4j-simple` dependency scope from `runtime` to `test`:

```xml
        <dependency>
            <groupId>org.slf4j</groupId>
            <artifactId>slf4j-simple</artifactId>
            <version>${slf4j.version}</version>
            <scope>test</scope>
        </dependency>
```

And add `<classpathScope>test</classpathScope>` to the exec plugin, so the sample can still find its logging binding and (later) its demo assets:

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

- [ ] **Step 4: Verify the tests still pass**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B clean test`
Expected: `Tests run: 40, Failures: 0, Errors: 0, Skipped: 0` and `BUILD SUCCESS`.

- [ ] **Step 5: Verify the sample still runs**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test-compile exec:java`
Expected: `BUILD SUCCESS` and a line `Wrote /…/target/sample.docx`.

Note `test-compile` rather than `compile`: `SampleMain` is now a test source.

- [ ] **Step 6: Verify the JAR is clean**

Run:
```bash
JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B -DskipTests package && unzip -l target/docx-service-1.0.0-SNAPSHOT.jar
```
Expected: the listing contains `com/example/docx/**` class files and **none** of `simplelogger.properties`, `SampleMain.class`, or `demo/`.

- [ ] **Step 7: Commit**

```bash
git add -A
git commit -m "fix: keep logging binding and sample out of the published jar"
```

---

### Task 2: Stop `writeTo` closing the caller's stream

**Files:**
- Modify: `src/main/java/com/example/docx/WordDocument.java`
- Create: `src/test/java/com/example/docx/StreamOwnershipTest.java`
- Modify: `src/test/java/com/example/docx/WordDocumentRoundTripTest.java` (replace `writeToMatchesToByteArray`)

**Interfaces:**
- Consumes: `WordDocument.writeTo(OutputStream)`, `toByteArray()`.
- Produces: unchanged signatures; `writeTo` no longer closes its argument.

`pkg.save(out)` wraps the stream in a `ZipOutputStream` and closes it. The spec documents `writeTo` for streaming to a servlet response, and closing a `ServletOutputStream` commits the response and blocks any later error handling. No test caught it because `ByteArrayOutputStream.close()` is a no-op.

- [ ] **Step 1: Write the failing test**

Create `src/test/java/com/example/docx/StreamOwnershipTest.java`:

```java
package com.example.docx;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.example.docx.style.HeadingStyle;
import java.io.ByteArrayOutputStream;
import org.junit.jupiter.api.Test;

class StreamOwnershipTest {

    /** ByteArrayOutputStream.close() is a no-op, so track the call explicitly. */
    private static final class TrackingStream extends ByteArrayOutputStream {
        boolean closed;

        @Override
        public void close() {
            closed = true;
        }
    }

    @Test
    void writeToLeavesTheCallersStreamOpen() {
        WordDocument doc = WordDocument.builder()
                .heading("Title", HeadingStyle.defaults())
                .build();

        TrackingStream out = new TrackingStream();
        doc.writeTo(out);

        assertTrue(out.size() > 0, "should have written bytes");
        assertFalse(out.closed, "writeTo must not close the caller's stream");
    }
}
```

- [ ] **Step 2: Run it to confirm it fails**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test -Dtest=StreamOwnershipTest`
Expected: FAIL — `writeTo must not close the caller's stream ==> expected: <false> but was: <true>`.

- [ ] **Step 3: Shield the stream**

In `src/main/java/com/example/docx/WordDocument.java`, add these imports beside the existing `java.io.ByteArrayOutputStream`:

```java
import java.io.FilterOutputStream;
import java.io.IOException;
```

Then replace the body of the `try` in `writeTo`:

```java
        try {
            // docx4j wraps the stream in a ZipOutputStream and closes it on save.
            // Closing a caller's ServletOutputStream commits the response, so the
            // stream is shielded and the caller keeps ownership.
            pkg.save(new FilterOutputStream(out) {
                @Override
                public void close() throws IOException {
                    flush();
                }
            });
        } catch (Docx4JException e) {
            throw new DocumentGenerationException("failed to serialise the document", e);
        }
```

- [ ] **Step 4: Run it to confirm it passes**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test -Dtest=StreamOwnershipTest`
Expected: PASS — `Tests run: 1, Failures: 0, Errors: 0`.

- [ ] **Step 5: Replace the timestamp-flaky equivalence test**

`writeToMatchesToByteArray` compares whole-ZIP bytes. ZIP entries carry MS-DOS timestamps at 2-second granularity, so two saves straddling a bucket boundary differ — measured at 6 failures in 17,719 runs. It is also near-tautological, since `toByteArray()` delegates to `writeTo()`.

In `src/test/java/com/example/docx/WordDocumentRoundTripTest.java`, replace the whole `writeToMatchesToByteArray` method with:

```java
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
```

If `assertArrayEquals` is now unused in that file, remove its static import.

- [ ] **Step 6: Run the full suite**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B clean test`
Expected: `Tests run: 41, Failures: 0, Errors: 0, Skipped: 0`.

- [ ] **Step 7: Commit**

```bash
git add -A
git commit -m "fix: writeTo no longer closes the caller's stream"
```

---

### Task 3: Close the two fail-open validation gaps

**Files:**
- Modify: `src/main/java/com/example/docx/page/PageSetup.java`
- Modify: `src/main/java/com/example/docx/Units.java`
- Modify: `src/main/java/com/example/docx/style/HeadingStyle.java`
- Modify: `src/test/java/com/example/docx/page/PageSetupTest.java`
- Modify: `src/test/java/com/example/docx/UnitsTest.java`
- Modify: `src/test/java/com/example/docx/style/HeadingStyleTest.java`

**Interfaces:**
- Consumes: existing `PageSetup.Builder`, `Units`, `HeadingStyle.Builder`.
- Produces: unchanged signatures. `Units.roundHalfUp` gains a range guard; `HeadingStyle` rejects sizes above 1638 pt.

Two separate fail-open defects. `PageSetup` sums margins as `int`, so extreme values wrap negative and pass the `>=` check — verified: `left(Integer.MAX_VALUE).right(Integer.MAX_VALUE)` builds successfully. And `Units` throws `ArithmeticException`, a type outside the module's contract; worse, `HeadingStyle.builder().sizePt(1e12).build()` validates fine and then throws at `toRPr()` time, so an object that passed validation blows up later.

- [ ] **Step 1: Write the failing tests**

Add to `src/test/java/com/example/docx/page/PageSetupTest.java`:

```java
    @Test
    void marginSumValidationSurvivesIntegerOverflow() {
        assertThrows(DocumentGenerationException.class,
                () -> PageSetup.builder().a4()
                        .left(Integer.MAX_VALUE).right(Integer.MAX_VALUE).build());
        assertThrows(DocumentGenerationException.class,
                () -> PageSetup.builder().a4()
                        .top(2_000_000_000).bottom(2_000_000_000).build());
    }
```

Add to `src/test/java/com/example/docx/UnitsTest.java`:

```java
    @Test
    void oversizedConversionsThrowTheModuleException() {
        assertThrows(DocumentGenerationException.class, () -> Units.cmToTwips(1e9));
        assertThrows(DocumentGenerationException.class, () -> Units.inchesToTwips(1e9));
        assertThrows(DocumentGenerationException.class, () -> Units.pointsToHalfPoints(1e12));
    }
```

Add to `src/test/java/com/example/docx/style/HeadingStyleTest.java`:

```java
    @Test
    void rejectsSizesBeyondWhatWordSupports() {
        // ST_HpsMeasure is capped at 1638 pt in Word. Rejecting at build() keeps a
        // validated HeadingStyle from throwing later at toRPr() time.
        assertThrows(DocumentGenerationException.class,
                () -> HeadingStyle.builder().sizePt(1639).build());
        assertThrows(DocumentGenerationException.class,
                () -> HeadingStyle.builder().sizePt(1e12).build());
        assertEquals(1638.0, HeadingStyle.builder().sizePt(1638).build().sizePt());
    }
```

- [ ] **Step 2: Run them to confirm they fail**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test`
Expected: FAIL — the `PageSetup` and `HeadingStyle` cases fail with "Expected DocumentGenerationException to be thrown, but nothing was thrown"; the `Units` cases fail with an unexpected `ArithmeticException`.

- [ ] **Step 3: Widen the margin-sum comparisons**

In `src/main/java/com/example/docx/page/PageSetup.java`, change the two comparisons in `build()` to compute in `long`:

```java
            if ((long) leftTwips + rightTwips >= pageWidthTwips) {
                throw new DocumentGenerationException(
                        "left + right margins (" + ((long) leftTwips + rightTwips)
                                + " twips) leave no width on a " + pageWidthTwips + "-twip page");
            }
            if ((long) topTwips + bottomTwips >= pageHeightTwips) {
                throw new DocumentGenerationException(
                        "top + bottom margins (" + ((long) topTwips + bottomTwips)
                                + " twips) leave no height on a " + pageHeightTwips + "-twip page");
            }
```

- [ ] **Step 4: Guard the conversion range**

In `src/main/java/com/example/docx/Units.java`, replace `roundHalfUp` with a version that reports through the module's own exception, and pass the unit through from each caller:

```java
    public static int cmToTwips(double cm) {
        check(cm, "centimetres");
        return roundHalfUp(cm / CM_PER_INCH * TWIPS_PER_INCH, "centimetres");
    }

    public static int inchesToTwips(double inches) {
        check(inches, "inches");
        return roundHalfUp(inches * TWIPS_PER_INCH, "inches");
    }

    public static int pointsToHalfPoints(double points) {
        check(points, "points");
        return roundHalfUp(points * 2.0, "points");
    }
```

```java
    private static int roundHalfUp(double value, String unit) {
        long rounded = Math.round(value);
        if (rounded > Integer.MAX_VALUE) {
            throw new DocumentGenerationException(
                    unit + " converts to " + rounded + ", which exceeds the maximum of "
                            + Integer.MAX_VALUE);
        }
        return (int) rounded;
    }
```

- [ ] **Step 5: Cap the font size at build time**

In `src/main/java/com/example/docx/style/HeadingStyle.java`, add the constant beside the other defaults:

```java
    /** Word's practical ceiling for {@code ST_HpsMeasure}. */
    public static final double MAX_SIZE_PT = 1638.0;
```

and extend the size check inside `build()`:

```java
            if (!(sizePt > 0) || !Double.isFinite(sizePt)) {
                throw new DocumentGenerationException(
                        "font size must be greater than zero points, got " + sizePt);
            }
            if (sizePt > MAX_SIZE_PT) {
                throw new DocumentGenerationException(
                        "font size must not exceed " + MAX_SIZE_PT + " points, got " + sizePt);
            }
```

- [ ] **Step 6: Run the suite**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B clean test`
Expected: `Tests run: 44, Failures: 0, Errors: 0, Skipped: 0`.

- [ ] **Step 7: Commit**

```bash
git add -A
git commit -m "fix: margin and conversion validation no longer fail open"
```

---

### Task 4: Hygiene — aliasing, import, and README accuracy

**Files:**
- Modify: `src/main/java/com/example/docx/style/HeadingStyle.java`
- Modify: `README.md`

**Interfaces:**
- Consumes: nothing new.
- Produces: `toRPr()` returns an object graph with no shared mutable nodes.

`toRPr()` assigns the same `HpsMeasure` instance to both `sz` and `szCs`, so a caller mutating one silently mutates the other. Marshalling is unaffected, so this is hygiene rather than a bug — but it hands callers a graph with hidden aliasing. The README meanwhile states two things that are not true of the code.

- [ ] **Step 1: Write the failing test**

Add to `src/test/java/com/example/docx/style/HeadingStyleTest.java`:

```java
    @Test
    void sizeAndComplexScriptSizeAreIndependentObjects() {
        RPr rPr = HeadingStyle.builder().sizePt(20).build().toRPr();
        assertNotSame(rPr.getSz(), rPr.getSzCs(),
                "sz and szCs must not share one mutable HpsMeasure");
        assertEquals(rPr.getSz().getVal(), rPr.getSzCs().getVal());
    }
```

Add the import `static org.junit.jupiter.api.Assertions.assertNotSame;`.

- [ ] **Step 2: Run it to confirm it fails**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test -Dtest=HeadingStyleTest`
Expected: FAIL — `sz and szCs must not share one mutable HpsMeasure`.

- [ ] **Step 3: Build two instances and import `Locale`**

In `src/main/java/com/example/docx/style/HeadingStyle.java`, replace the size block in `toRPr()`:

```java
        int halfPoints = Units.pointsToHalfPoints(sizePt);
        HpsMeasure size = factory.createHpsMeasure();
        size.setVal(BigInteger.valueOf(halfPoints));
        rPr.setSz(size);

        HpsMeasure complexScriptSize = factory.createHpsMeasure();
        complexScriptSize.setVal(BigInteger.valueOf(halfPoints));
        rPr.setSzCs(complexScriptSize);
```

Add `import java.util.Locale;` to the imports, and change the fully-qualified use to `bare.toUpperCase(Locale.ROOT);`.

- [ ] **Step 4: Run it to confirm it passes**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test -Dtest=HeadingStyleTest`
Expected: PASS.

- [ ] **Step 5: Correct the README's two false claims**

In `README.md`, replace:

```
For large documents prefer `writeTo(out)` over `toByteArray()`, so the whole
file never sits in heap.
```

with:

```
For large documents prefer `writeTo(out)` over `toByteArray()`: it avoids
buffering a second full copy of the file. `writeTo` does not close the stream
you give it, so a controller keeps ownership of the response.
```

And replace:

```
`content` and `style` never touch `WordprocessingMLPackage`, so they are pure
functions of their arguments and need no fixtures. `page` and the facade own
the package.
```

with:

```
`content` and `style` never touch `WordprocessingMLPackage`, so they are pure
functions of their arguments and need no fixtures. The facade owns the
package; `page` only produces a `w:sectPr` fragment for it.
```

- [ ] **Step 6: Run the suite and commit**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B clean test`
Expected: `Tests run: 45, Failures: 0, Errors: 0, Skipped: 0`.

```bash
git add -A
git commit -m "fix: drop rPr aliasing and correct README claims"
```

---

### Task 5: Sizing foundations — `twipsToEmu` and `usableWidthTwips`

**Files:**
- Modify: `src/main/java/com/example/docx/Units.java`
- Modify: `src/main/java/com/example/docx/page/PageSetup.java`
- Modify: `src/test/java/com/example/docx/UnitsTest.java`
- Modify: `src/test/java/com/example/docx/page/PageSetupTest.java`

**Interfaces:**
- Consumes: existing `Units`, `PageSetup`.
- Produces: `public static long Units.twipsToEmu(int twips)` and `public int PageSetup.usableWidthTwips()`. Task 6 uses both.

- [ ] **Step 1: Write the failing tests**

Add to `src/test/java/com/example/docx/UnitsTest.java`:

```java
    @Test
    void twipsToEmu() {
        // 914400 / 1440 = 635 exactly, so this conversion never rounds.
        assertEquals(635L, Units.twipsToEmu(1));
        assertEquals(914400L, Units.twipsToEmu(1440));
        assertEquals(0L, Units.twipsToEmu(0));
        assertEquals(3_239_770L, Units.twipsToEmu(5102));
    }

    @Test
    void twipsToEmuRejectsNegatives() {
        assertThrows(DocumentGenerationException.class, () -> Units.twipsToEmu(-1));
    }
```

Add to `src/test/java/com/example/docx/page/PageSetupTest.java`:

```java
    @Test
    void usableWidthSubtractsBothSideMargins() {
        assertEquals(10204, PageSetup.a4().usableWidthTwips());       // 11906 - 851 - 851
        assertEquals(7906, PageSetup.builder().a4().marginsTwips(2000).build()
                .usableWidthTwips());                                 // 11906 - 4000
        assertEquals(10906, PageSetup.builder().a4().left(600).right(400).build()
                .usableWidthTwips());                                 // 11906 - 1000
    }
```

- [ ] **Step 2: Run them to confirm they fail**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test`
Expected: FAIL — compilation errors, `cannot find symbol: method twipsToEmu(int)` and `cannot find symbol: method usableWidthTwips()`.

- [ ] **Step 3: Add `twipsToEmu`**

In `src/main/java/com/example/docx/Units.java`, add the constant beside the others:

```java
    private static final int EMU_PER_TWIP = 635;
```

and the method, above the private helpers:

```java
    /** 1 twip is exactly 635 EMU, since 914400 / 1440 divides evenly. */
    public static long twipsToEmu(int twips) {
        if (twips < 0) {
            throw new DocumentGenerationException("twips must not be negative, got " + twips);
        }
        return (long) twips * EMU_PER_TWIP;
    }
```

- [ ] **Step 4: Add `usableWidthTwips`**

In `src/main/java/com/example/docx/page/PageSetup.java`, add above `toSectPr()`:

```java
    /** Page width less both side margins: the width text and images may occupy. */
    public int usableWidthTwips() {
        return pageWidthTwips - leftTwips - rightTwips;
    }
```

- [ ] **Step 5: Run the suite**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B clean test`
Expected: `Tests run: 48, Failures: 0, Errors: 0, Skipped: 0`.

- [ ] **Step 6: Commit**

```bash
git add -A
git commit -m "feat: add twipsToEmu and usableWidthTwips"
```

---

### Task 6: `ImageParts` — SVG with PNG fallback

**Files:**
- Create: `src/main/java/com/example/docx/part/ImageParts.java`
- Modify: `src/main/java/com/example/docx/WordDocument.java`
- Create: `src/test/java/com/example/docx/part/ImagePartsTest.java`

**Interfaces:**
- Consumes: `Units.twipsToEmu(int)`, `PageSetup.usableWidthTwips()`, `DocumentGenerationException`.
- Produces: `public static org.docx4j.wml.P ImageParts.svgImage(WordprocessingMLPackage pkg, byte[] svg, byte[] png, int widthTwips)` and `WordDocument.Builder.svgImage(byte[] svg, byte[] pngFallback)`.

All of the code below was compiled and run against docx4j 17.0.2 before this plan was written; the expected extent values are measured, not calculated on paper.

- [ ] **Step 1: Write the failing test**

Create `src/test/java/com/example/docx/part/ImagePartsTest.java`:

```java
package com.example.docx.part;

import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.example.docx.DocumentGenerationException;
import com.example.docx.WordDocument;
import com.example.docx.page.PageSetup;
import com.example.docx.style.HeadingStyle;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.PartName;
import org.junit.jupiter.api.Test;

class ImagePartsTest {

    private static byte[] resource(String name) throws IOException {
        try (InputStream in = ImagePartsTest.class.getResourceAsStream(name)) {
            assertNotNull(in, "missing test resource " + name);
            return in.readAllBytes();
        }
    }

    private static byte[] svg() throws IOException {
        return resource("/demo/chart.svg");
    }

    private static byte[] png() throws IOException {
        return resource("/demo/chart.png");
    }

    private static String documentXml(PageSetup setup) throws Exception {
        byte[] bytes = WordDocument.builder()
                .pageSetup(setup)
                .heading("Quarterly Report", HeadingStyle.defaults())
                .svgImage(svg(), png())
                .build()
                .toByteArray();
        return WordprocessingMLPackage.load(new ByteArrayInputStream(bytes))
                .getMainDocumentPart().getXML();
    }

    @Test
    void bothMediaPartsExistWithCorrectContentTypes() throws Exception {
        byte[] bytes = WordDocument.builder()
                .pageSetup(PageSetup.a4())
                .heading("Quarterly Report", HeadingStyle.defaults())
                .svgImage(svg(), png())
                .build()
                .toByteArray();

        WordprocessingMLPackage re =
                WordprocessingMLPackage.load(new ByteArrayInputStream(bytes));

        Map<String, String> media = new HashMap<>();
        for (Map.Entry<PartName, Part> e : re.getParts().getParts().entrySet()) {
            String name = e.getKey().getName();
            if (name.startsWith("/word/media/")) {
                media.put(name, e.getValue().getContentType());
            }
        }

        assertTrue(media.keySet().stream().anyMatch(n -> n.endsWith(".png")),
                "png part missing, got " + media);
        assertTrue(media.keySet().stream().anyMatch(n -> n.endsWith(".svg")),
                "svg part missing, got " + media);
        assertTrue(media.values().stream().anyMatch(c -> c.startsWith("image/svg+xml")),
                "svg content type missing, got " + media);
    }

    @Test
    void svgBlipCarriesItsNamespaceAndPrefixedEmbed() throws Exception {
        String xml = documentXml(PageSetup.a4());

        assertTrue(xml.contains("{96DAC541-7B7A-43D3-8B79-37D633B846F1}"),
                "svg extension uri missing");
        // The namespace, not just the local name. An unnamespaced svgBlip still
        // contains the string "svgBlip" and is silently ignored by Word.
        assertTrue(xml.contains("http://schemas.microsoft.com/office/drawing/2016/SVG/main"),
                "asvg namespace missing");
        assertTrue(xml.contains(":svgBlip"), "svgBlip must carry its namespace prefix");
        assertTrue(xml.contains("r:embed"), "r:embed must keep its prefix");
    }

    @Test
    void imageIsHalfTheUsablePageWidth() throws Exception {
        // A4 usable = 11906 - 851 - 851 = 10204; half = 5102; cx = 5102 * 635.
        // PNG is 1600x1120, so cy = cx * 1120 / 1600.
        String xml = documentXml(PageSetup.a4());
        assertTrue(xml.contains("cx=\"3239770\""), "expected cx 3239770 in " + extent(xml));
        assertTrue(xml.contains("cy=\"2267839\""), "expected cy 2267839 in " + extent(xml));
    }

    @Test
    void widthTracksThePageRatherThanAConstant() throws Exception {
        // usable = 11906 - 2000 - 2000 = 7906; half = 3953; cx = 3953 * 635 = 2510155.
        String xml = documentXml(PageSetup.builder().a4().marginsTwips(2000).build());
        assertTrue(xml.contains("cx=\"2510155\""),
                "a narrower page must yield a narrower image, got " + extent(xml));
    }

    private static String extent(String xml) {
        return xml.replaceAll("(?s).*(<wp:extent[^/]*/>).*", "$1");
    }

    @Test
    void rejectsBadInput() throws Exception {
        WordprocessingMLPackage pkg = WordprocessingMLPackage.createPackage();
        assertThrows(DocumentGenerationException.class,
                () -> ImageParts.svgImage(pkg, new byte[0], png(), 5102));
        assertThrows(DocumentGenerationException.class,
                () -> ImageParts.svgImage(pkg, svg(), null, 5102));
        assertThrows(DocumentGenerationException.class,
                () -> ImageParts.svgImage(pkg, svg(), "not a png".getBytes(), 5102));
        assertThrows(DocumentGenerationException.class,
                () -> ImageParts.svgImage(pkg, svg(), png(), 0));
    }
}
```

- [ ] **Step 2: Run it to confirm it fails**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test -Dtest=ImagePartsTest`
Expected: FAIL — compilation errors, `cannot find symbol: class ImageParts` and `cannot find symbol: method svgImage(byte[],byte[])`.

- [ ] **Step 3: Write `ImageParts`**

Create `src/main/java/com/example/docx/part/ImageParts.java`:

```java
package com.example.docx.part;

import com.example.docx.DocumentGenerationException;
import com.example.docx.Units;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.StringReader;
import javax.imageio.ImageIO;
import javax.xml.parsers.DocumentBuilderFactory;
import org.docx4j.dml.CTBlip;
import org.docx4j.dml.CTOfficeArtExtension;
import org.docx4j.dml.CTOfficeArtExtensionList;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.contenttype.ContentType;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPart;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart.AddPartBehaviour;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.w3c.dom.Element;
import org.xml.sax.InputSource;

/**
 * Builds image paragraphs.
 *
 * <p>Unlike {@code content} and {@code style}, this needs the package: an image is a
 * part plus a relationship, so it cannot be a pure function of its arguments.
 *
 * <p>An SVG is stored the way Word stores one — the blip's {@code r:embed} points at a
 * PNG, and the SVG rides along as an {@code asvg:svgBlip} extension. Renderers that
 * understand the extension draw the vector; the rest draw the PNG they already had.
 */
public final class ImageParts {

    /** Identifies the SVG blip extension. Defined by Microsoft; do not change. */
    private static final String SVG_EXTENSION_URI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";

    private static final String ASVG_NS =
            "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
    private static final String RELATIONSHIPS_NS =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private ImageParts() {
    }

    /**
     * A paragraph holding the SVG, drawn {@code widthTwips} wide with the height taken
     * from the PNG's aspect ratio.
     */
    public static P svgImage(WordprocessingMLPackage pkg, byte[] svg, byte[] png, int widthTwips) {
        if (pkg == null) {
            throw new DocumentGenerationException("package must not be null");
        }
        if (svg == null || svg.length == 0) {
            throw new DocumentGenerationException("svg bytes must not be empty");
        }
        if (png == null || png.length == 0) {
            throw new DocumentGenerationException("png fallback bytes must not be empty");
        }
        if (widthTwips <= 0) {
            throw new DocumentGenerationException(
                    "image width must be greater than zero twips, got " + widthTwips);
        }

        BufferedImage raster = decodePng(png);
        MainDocumentPart mainDocumentPart = pkg.getMainDocumentPart();

        try {
            // The PNG is the blip's primary image: the fallback every renderer understands.
            BinaryPartAbstractImage pngPart = BinaryPartAbstractImage.createImagePart(pkg, png);
            Inline inline = pngPart.createImageInline("image", "Image", 1, 2, false);

            // docx4j has no ImageSvgPart, and its part factory treats image/svg+xml as
            // XML and hands back a DefaultXmlPart, so the SVG part is built by hand.
            BinaryPart svgPart = new BinaryPart(new PartName("/word/media/image.svg"));
            svgPart.setBinaryData(svg);
            svgPart.setContentType(new ContentType("image/svg+xml"));
            svgPart.setRelationshipType(Namespaces.IMAGE);
            Relationship svgRelationship = mainDocumentPart.addTargetPart(
                    svgPart, AddPartBehaviour.RENAME_IF_NAME_EXISTS);

            attachSvgExtension(inline, svgRelationship.getId());
            setExtent(inline, widthTwips, raster.getWidth(), raster.getHeight());

            ObjectFactory factory = Context.getWmlObjectFactory();
            Drawing drawing = factory.createDrawing();
            drawing.getAnchorOrInline().add(inline);
            R run = factory.createR();
            run.getContent().add(drawing);
            P paragraph = factory.createP();
            paragraph.getContent().add(run);
            return paragraph;
        } catch (DocumentGenerationException e) {
            throw e;
        } catch (Exception e) {
            throw new DocumentGenerationException("failed to embed the image", e);
        }
    }

    private static BufferedImage decodePng(byte[] png) {
        try {
            BufferedImage image = ImageIO.read(new ByteArrayInputStream(png));
            if (image == null) {
                throw new DocumentGenerationException("png fallback is not a readable image");
            }
            return image;
        } catch (IOException e) {
            throw new DocumentGenerationException("failed to read the png fallback", e);
        }
    }

    private static void attachSvgExtension(Inline inline, String svgRelationshipId)
            throws Exception {
        CTBlip blip = inline.getGraphic().getGraphicData().getPic().getBlipFill().getBlip();

        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        // Mandatory, and false by default. Left off, the prefixed name parses as a
        // literal element in NO namespace and Word silently ignores the SVG.
        factory.setNamespaceAware(true);
        Element svgBlip = factory.newDocumentBuilder()
                .parse(new InputSource(new StringReader(
                        "<asvg:svgBlip xmlns:asvg=\"" + ASVG_NS
                                + "\" xmlns:r=\"" + RELATIONSHIPS_NS
                                + "\" r:embed=\"" + svgRelationshipId + "\"/>")))
                .getDocumentElement();

        CTOfficeArtExtension extension = new CTOfficeArtExtension();
        extension.setUri(SVG_EXTENSION_URI);
        extension.setAny(svgBlip);

        CTOfficeArtExtensionList extensions = new CTOfficeArtExtensionList();
        extensions.getExt().add(extension);
        blip.setExtLst(extensions);
    }

    private static void setExtent(Inline inline, int widthTwips, int pngWidthPx, int pngHeightPx) {
        // createImageInline sizes from the PNG's intrinsic dimensions and DPI, which is
        // not what the caller asked for, so the extent is set explicitly.
        long cx = Units.twipsToEmu(widthTwips);
        long cy = Math.round(cx * (double) pngHeightPx / pngWidthPx);
        inline.getExtent().setCx(cx);
        inline.getExtent().setCy(cy);
    }
}
```

- [ ] **Step 4: Wire it into the facade**

In `src/main/java/com/example/docx/WordDocument.java`, add the import:

```java
import com.example.docx.part.ImageParts;
```

Add two fields to `Builder`, beside `headingText`:

```java
        private byte[] svgBytes;
        private byte[] pngBytes;
```

Add the builder method above `build()`:

```java
        /**
         * Adds an SVG image with a PNG fallback, drawn at half the usable page width.
         *
         * <p>The SVG must be SVG 1.1: Word's renderer rejects SVG 2 features such as
         * {@code height="auto"} and 8-digit hex colours that browsers accept.
         */
        public Builder svgImage(byte[] svg, byte[] pngFallback) {
            this.svgBytes = svg;
            this.pngBytes = pngFallback;
            return this;
        }
```

And append the image inside the existing `try` in `build()`, immediately after the heading is added:

```java
                mainDocumentPart.getContent().add(Headings.heading(headingText, headingStyle));

                if (svgBytes != null || pngBytes != null) {
                    mainDocumentPart.getContent().add(ImageParts.svgImage(
                            pkg, svgBytes, pngBytes, pageSetup.usableWidthTwips() / 2));
                }
```

Testing the fields for null rather than calling unconditionally keeps `svgImage` optional; passing only one of the two still reaches `ImageParts` and is rejected there with a message naming which is missing.

- [ ] **Step 5: Run the test to confirm it passes**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test -Dtest=ImagePartsTest`
Expected: PASS — `Tests run: 5, Failures: 0, Errors: 0`.

- [ ] **Step 6: Run the full suite**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B clean test`
Expected: `Tests run: 53, Failures: 0, Errors: 0, Skipped: 0`.

- [ ] **Step 7: Commit**

```bash
git add -A
git commit -m "feat: embed SVG images with a PNG fallback"
```

---

### Task 7: Sample and documentation

**Files:**
- Modify: `src/test/java/com/example/docx/sample/SampleMain.java`
- Modify: `README.md`

**Interfaces:**
- Consumes: `WordDocument.Builder.svgImage(byte[], byte[])`.
- Produces: nothing other tasks depend on. Terminal task.

- [ ] **Step 1: Add the chart to the sample**

Replace `src/test/java/com/example/docx/sample/SampleMain.java` with:

```java
package com.example.docx.sample;

import com.example.docx.WordDocument;
import com.example.docx.page.PageSetup;
import com.example.docx.style.HeadingStyle;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UncheckedIOException;
import java.nio.file.Files;
import java.nio.file.Path;

/** Writes {@code target/sample.docx} so the output can be opened and eyeballed. */
public final class SampleMain {

    private SampleMain() {
    }

    public static void main(String[] args) throws IOException {
        WordDocument document = WordDocument.builder()
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
                .svgImage(resource("/demo/chart.svg"), resource("/demo/chart.png"))
                .build();

        Path target = Path.of("target", "sample.docx");
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
```

- [ ] **Step 2: Run the sample**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test-compile exec:java`
Expected: `BUILD SUCCESS` and `Wrote /…/target/sample.docx`.

- [ ] **Step 3: Verify the output really carries both images**

Run:
```bash
unzip -l target/sample.docx | grep media
```
Expected: two entries under `word/media/` — one `.png` and one `.svg`.

Run:
```bash
unzip -p target/sample.docx word/document.xml | grep -c "svgBlip"
```
Expected: `1`.

Run:
```bash
unzip -p target/sample.docx word/document.xml | grep -o '<wp:extent[^/]*/>'
```
Expected: `<wp:extent cx="3239770" cy="2267839"/>`.

- [ ] **Step 4: Document the feature and the SVG 1.1 requirement**

In `README.md`, add to the "Using it" example, after the `.heading(...)` call:

```java
        .svgImage(svgBytes, pngFallbackBytes)
```

Add a row to the Layout table, after the `…​.content` row:

```
| `…​.part` | `ImageParts` — image parts and the SVG blip extension |
```

Change the `…​.sample` row to:

```
| `…​.sample` | Runnable `main` (test sources, so it stays out of the jar) |
```

Add to "Things that bite":

```
- **Word wants SVG 1.1, not SVG 2.** Word's renderer is strict where browsers
  are lenient, so an SVG that looks right in Chrome can render wrong or not at
  all. Two common offenders: `height="auto"` (SVG 2 sizing — rejected outright),
  and 8-digit hex colours like `#444cf71a` (CSS Color 4 — silently falls back to
  **black**). Use explicit `width`/`height`, and `fill="#444CF7"` with a separate
  `fill-opacity`.
- **SVGs carry a PNG twin.** `svgImage` embeds both: the PNG is what the blip
  points at, and the SVG rides along as an extension. Word desktop draws the
  vector, Word Online and older Word draw the PNG. You supply both — the library
  does not rasterise, and does not check that they match.
```

Add to "Notes":

```
- Images are drawn at half the usable page width (page width less both margins),
  with height from the PNG's aspect ratio.
```

- [ ] **Step 5: Run the full suite one last time**

Run: `JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B clean test`
Expected: `Tests run: 53, Failures: 0, Errors: 0, Skipped: 0` and `BUILD SUCCESS`.

- [ ] **Step 6: Commit**

```bash
git add -A
git commit -m "feat: show the SVG chart in the sample and document it"
```

---

## Definition of Done

- `mvn clean test` passes with 53 tests.
- `mvn test-compile exec:java` writes `target/sample.docx` containing one `.png` and one `.svg` under `word/media/`, exactly one `svgBlip` element, and `<wp:extent cx="3239770" cy="2267839"/>`.
- The `svgBlip` element carries the `asvg` prefix and its namespace; `r:embed` keeps its prefix.
- `mvn -DskipTests package` produces a JAR containing no `simplelogger.properties`, no `SampleMain.class`, and no `demo/` resources.
- `writeTo` leaves the caller's stream open.
- `PageSetup.builder().a4().left(Integer.MAX_VALUE).right(Integer.MAX_VALUE).build()` throws `DocumentGenerationException`.
- The document opens in Word without a repair prompt, and the chart renders.
