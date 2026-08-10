# docx-service Phase One Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A Java 25 library that generates a Word `.docx` containing configurable A4 page margins and exactly one configurable heading.

**Architecture:** Layered packages under `com.example.docx`. `style/` and `content/` hold pure functions of their arguments that return JAXB objects and never touch `WordprocessingMLPackage`; `page/` and the `WordDocument` facade own the package because they create parts. Immutable value objects are built through nested static builders, so later phases add files rather than editing existing ones.

**Tech Stack:** Java 25, Maven, docx4j 17.0.2 (`docx4j-JAXB-ReferenceImpl`), JUnit 6.1.3, slf4j 2.0.18.

## Global Constraints

- Java release level is exactly `25` (`<maven.compiler.release>25</maven.compiler.release>`).
- docx4j is `org.docx4j:docx4j-JAXB-ReferenceImpl:17.0.2`. Do **not** use `org.docx4j:docx4j` (abandoned at 6.1.2) and do **not** trust `search.maven.org`, whose index reports 11.5.3 as latest and is stale by a major line.
- In docx4j 17.x the `w:*` model classes come from `docx4j-generated-objects`, not `docx4j-openxml-objects`. Both arrive transitively; declare neither.
- All page geometry is in **twips** (1/1440 inch). Never pixels.
- A4 is the literal pair `11906 × 16838`. Do not derive it — 297 mm computes to 16838.7 → 16839, which is wrong.
- The default margin is the literal `851` twips. `marginsCm(1.5)` correctly yields **850**, not 851. Never substitute one for the other.
- Colour is stored and emitted as bare `RRGGBB`, uppercase, no `#`.
- Font size is stored in half-points: `sz` = `sizePt × 2`.
- Every `w:t` sets `xml:space="preserve"`.
- `DocumentGenerationException extends RuntimeException` is the **only** exception type this module throws deliberately. Validation failures carry a message naming the field and **no cause**; wrapped docx4j failures **always** carry a cause.
- Package root is `com.example.docx`. Module directory is `docx-service/`.
- Run Maven with `JAVA_HOME=$(/usr/libexec/java_home -v 25)`.

## File Structure

| File | Responsibility |
| --- | --- |
| `pom.xml` | Dependencies, Java 25, surefire, exec |
| `src/main/java/com/example/docx/Units.java` | twips ↔ cm/inch, points → half-points |
| `src/main/java/com/example/docx/DocumentGenerationException.java` | the module's single unchecked exception |
| `src/main/java/com/example/docx/page/PageSetup.java` | immutable page size + margins; emits `w:sectPr` |
| `src/main/java/com/example/docx/style/HeadingStyle.java` | immutable run formatting; emits `w:rPr` |
| `src/main/java/com/example/docx/content/Headings.java` | `(text, style) → org.docx4j.wml.P` |
| `src/main/java/com/example/docx/WordDocument.java` | fluent facade; owns the package, serialises |
| `src/main/java/com/example/docx/sample/SampleMain.java` | writes `target/sample.docx` |
| `src/test/java/com/example/docx/UnitsTest.java` | conversion values |
| `src/test/java/com/example/docx/page/PageSetupTest.java` | A4 constants, margin validation |
| `src/test/java/com/example/docx/style/HeadingStyleTest.java` | hex normalisation, validation |
| `src/test/java/com/example/docx/WordDocumentRoundTripTest.java` | reload and assert; stream equivalence |

---

### Task 1: Project skeleton and `DocumentGenerationException`

**Files:**
- Create: `docx-service/pom.xml`
- Create: `docx-service/src/main/java/com/example/docx/DocumentGenerationException.java`
- Test: `docx-service/src/test/java/com/example/docx/DocumentGenerationExceptionTest.java`

**Interfaces:**
- Consumes: nothing.
- Produces: `DocumentGenerationException(String message)` and `DocumentGenerationException(String message, Throwable cause)`, both public. Every later task throws this type.

- [ ] **Step 1: Create the POM**

Create `docx-service/pom.xml`:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<project xmlns="http://maven.apache.org/POM/4.0.0"
         xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
         xsi:schemaLocation="http://maven.apache.org/POM/4.0.0 https://maven.apache.org/xsd/maven-4.0.0.xsd">
    <modelVersion>4.0.0</modelVersion>

    <groupId>com.example</groupId>
    <artifactId>docx-service</artifactId>
    <version>1.0.0-SNAPSHOT</version>
    <packaging>jar</packaging>
    <name>docx-service</name>
    <description>Word (.docx) generation on docx4j: page margins and one heading</description>

    <properties>
        <maven.compiler.release>25</maven.compiler.release>
        <project.build.sourceEncoding>UTF-8</project.build.sourceEncoding>
        <docx4j.version>17.0.2</docx4j.version>
        <slf4j.version>2.0.18</slf4j.version>
        <junit.version>6.1.3</junit.version>
    </properties>

    <dependencyManagement>
        <dependencies>
            <dependency>
                <groupId>org.junit</groupId>
                <artifactId>junit-bom</artifactId>
                <version>${junit.version}</version>
                <type>pom</type>
                <scope>import</scope>
            </dependency>
        </dependencies>
    </dependencyManagement>

    <dependencies>
        <dependency>
            <groupId>org.docx4j</groupId>
            <artifactId>docx4j-JAXB-ReferenceImpl</artifactId>
            <version>${docx4j.version}</version>
        </dependency>
        <dependency>
            <groupId>org.slf4j</groupId>
            <artifactId>slf4j-api</artifactId>
            <version>${slf4j.version}</version>
        </dependency>
        <dependency>
            <groupId>org.slf4j</groupId>
            <artifactId>slf4j-simple</artifactId>
            <version>${slf4j.version}</version>
            <scope>runtime</scope>
        </dependency>
        <dependency>
            <groupId>org.junit.jupiter</groupId>
            <artifactId>junit-jupiter</artifactId>
            <scope>test</scope>
        </dependency>
    </dependencies>

    <build>
        <plugins>
            <plugin>
                <groupId>org.apache.maven.plugins</groupId>
                <artifactId>maven-compiler-plugin</artifactId>
                <version>3.15.0</version>
            </plugin>
            <plugin>
                <groupId>org.apache.maven.plugins</groupId>
                <artifactId>maven-surefire-plugin</artifactId>
                <version>3.5.6</version>
            </plugin>
            <plugin>
                <groupId>org.codehaus.mojo</groupId>
                <artifactId>exec-maven-plugin</artifactId>
                <version>3.6.3</version>
                <configuration>
                    <mainClass>com.example.docx.sample.SampleMain</mainClass>
                </configuration>
            </plugin>
        </plugins>
    </build>
</project>
```

- [ ] **Step 2: Write the failing test**

Create `docx-service/src/test/java/com/example/docx/DocumentGenerationExceptionTest.java`:

```java
package com.example.docx;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertSame;

import org.junit.jupiter.api.Test;

class DocumentGenerationExceptionTest {

    @Test
    void validationFailureCarriesMessageAndNoCause() {
        DocumentGenerationException e = new DocumentGenerationException("colour must be RRGGBB");
        assertEquals("colour must be RRGGBB", e.getMessage());
        assertNull(e.getCause());
    }

    @Test
    void wrappedFailureCarriesCause() {
        Throwable cause = new IllegalStateException("boom");
        DocumentGenerationException e = new DocumentGenerationException("save failed", cause);
        assertSame(cause, e.getCause());
    }

    @Test
    void isUnchecked() {
        assertEquals(true, RuntimeException.class.isAssignableFrom(DocumentGenerationException.class));
    }
}
```

- [ ] **Step 3: Run test to verify it fails**

Run: `cd docx-service && JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test`
Expected: FAIL — compilation error, `cannot find symbol: class DocumentGenerationException`.

- [ ] **Step 4: Write minimal implementation**

Create `docx-service/src/main/java/com/example/docx/DocumentGenerationException.java`:

```java
package com.example.docx;

/**
 * The only exception this module throws deliberately.
 *
 * <p>Validation failures carry a message naming the offending field and no cause.
 * Failures wrapped from docx4j always carry a cause.
 */
public class DocumentGenerationException extends RuntimeException {

    private static final long serialVersionUID = 1L;

    public DocumentGenerationException(String message) {
        super(message);
    }

    public DocumentGenerationException(String message, Throwable cause) {
        super(message, cause);
    }
}
```

- [ ] **Step 5: Run test to verify it passes**

Run: `cd docx-service && JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test`
Expected: PASS — `Tests run: 3, Failures: 0, Errors: 0, Skipped: 0` and `BUILD SUCCESS`.

- [ ] **Step 6: Commit**

```bash
git add docx-service/pom.xml docx-service/src/main/java/com/example/docx/DocumentGenerationException.java docx-service/src/test/java/com/example/docx/DocumentGenerationExceptionTest.java
git commit -m "feat: project skeleton and DocumentGenerationException"
```

---

### Task 2: `Units` conversions

**Files:**
- Create: `docx-service/src/main/java/com/example/docx/Units.java`
- Test: `docx-service/src/test/java/com/example/docx/UnitsTest.java`

**Interfaces:**
- Consumes: `DocumentGenerationException` from Task 1.
- Produces: `public static int cmToTwips(double cm)`, `public static int inchesToTwips(double inches)`, `public static int pointsToHalfPoints(double points)`. All are `static` on a final, non-instantiable class.

Rounding is half-up. Verified expectations: `1.5 cm → 850` (1.5 / 2.54 × 1440 = 850.39), `2.54 cm → 1440`, `1.0 inch → 1440`, `20 pt → 40`.

- [ ] **Step 1: Write the failing test**

Create `docx-service/src/test/java/com/example/docx/UnitsTest.java`:

```java
package com.example.docx;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import org.junit.jupiter.api.Test;

class UnitsTest {

    @Test
    void cmToTwips() {
        assertEquals(850, Units.cmToTwips(1.5));
        assertEquals(1440, Units.cmToTwips(2.54));
        assertEquals(0, Units.cmToTwips(0.0));
    }

    @Test
    void inchesToTwips() {
        assertEquals(1440, Units.inchesToTwips(1.0));
        assertEquals(720, Units.inchesToTwips(0.5));
    }

    @Test
    void pointsToHalfPoints() {
        assertEquals(40, Units.pointsToHalfPoints(20));
        assertEquals(23, Units.pointsToHalfPoints(11.5));
    }

    @Test
    void rejectsNonFiniteInput() {
        assertThrows(DocumentGenerationException.class, () -> Units.cmToTwips(Double.NaN));
        assertThrows(DocumentGenerationException.class, () -> Units.inchesToTwips(Double.POSITIVE_INFINITY));
    }

    @Test
    void rejectsNegativeInput() {
        assertThrows(DocumentGenerationException.class, () -> Units.cmToTwips(-1.0));
        assertThrows(DocumentGenerationException.class, () -> Units.pointsToHalfPoints(-1.0));
    }
}
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd docx-service && JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test`
Expected: FAIL — compilation error, `cannot find symbol: class Units`.

- [ ] **Step 3: Write minimal implementation**

Create `docx-service/src/main/java/com/example/docx/Units.java`:

```java
package com.example.docx;

/**
 * Conversions into the units OOXML actually stores.
 *
 * <p>Page geometry is in twips (1/1440 inch); font sizes are in half-points.
 * These helpers convert caller-supplied values only. Built-in constants such as
 * A4's dimensions are literals elsewhere, because deriving them rounds wrong.
 */
public final class Units {

    private static final double TWIPS_PER_INCH = 1440.0;
    private static final double CM_PER_INCH = 2.54;

    private Units() {
    }

    public static int cmToTwips(double cm) {
        check(cm, "centimetres");
        return roundHalfUp(cm / CM_PER_INCH * TWIPS_PER_INCH);
    }

    public static int inchesToTwips(double inches) {
        check(inches, "inches");
        return roundHalfUp(inches * TWIPS_PER_INCH);
    }

    public static int pointsToHalfPoints(double points) {
        check(points, "points");
        return roundHalfUp(points * 2.0);
    }

    private static void check(double value, String unit) {
        if (!Double.isFinite(value)) {
            throw new DocumentGenerationException(unit + " must be a finite number, got " + value);
        }
        if (value < 0) {
            throw new DocumentGenerationException(unit + " must not be negative, got " + value);
        }
    }

    private static int roundHalfUp(double value) {
        return Math.toIntExact(Math.round(value));
    }
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd docx-service && JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test`
Expected: PASS — `Tests run: 8, Failures: 0, Errors: 0, Skipped: 0`.

- [ ] **Step 5: Commit**

```bash
git add docx-service/src/main/java/com/example/docx/Units.java docx-service/src/test/java/com/example/docx/UnitsTest.java
git commit -m "feat: add Units twip and half-point conversions"
```

---

### Task 3: `PageSetup`

**Files:**
- Create: `docx-service/src/main/java/com/example/docx/page/PageSetup.java`
- Test: `docx-service/src/test/java/com/example/docx/page/PageSetupTest.java`

**Interfaces:**
- Consumes: `Units`, `DocumentGenerationException` from Tasks 1–2.
- Produces:
  - `public static PageSetup.Builder builder()`
  - `public static PageSetup a4()` — convenience for `builder().a4().build()`
  - Accessors `int pageWidthTwips()`, `int pageHeightTwips()`, `int topTwips()`, `int rightTwips()`, `int bottomTwips()`, `int leftTwips()`
  - `public org.docx4j.wml.SectPr toSectPr()`
  - Builder methods, each returning `Builder`: `a4()`, `pageSizeTwips(int width, int height)`, `marginsTwips(int all)`, `marginsCm(double all)`, `marginsInches(double all)`, `top(int)`, `right(int)`, `bottom(int)`, `left(int)`, and `build()`.

Builder defaults: A4 (`11906 × 16838`) and `851` twips on all four sides.

- [ ] **Step 1: Write the failing test**

Create `docx-service/src/test/java/com/example/docx/page/PageSetupTest.java`:

```java
package com.example.docx.page;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import com.example.docx.DocumentGenerationException;
import java.math.BigInteger;
import org.docx4j.wml.SectPr;
import org.junit.jupiter.api.Test;

class PageSetupTest {

    @Test
    void a4UsesLiteralWordDimensions() {
        PageSetup setup = PageSetup.a4();
        assertEquals(11906, setup.pageWidthTwips());
        assertEquals(16838, setup.pageHeightTwips());
    }

    @Test
    void defaultMarginsAre851Twips() {
        PageSetup setup = PageSetup.a4();
        assertEquals(851, setup.topTwips());
        assertEquals(851, setup.rightTwips());
        assertEquals(851, setup.bottomTwips());
        assertEquals(851, setup.leftTwips());
    }

    @Test
    void marginsCmIsNotTheSameAs851() {
        PageSetup setup = PageSetup.builder().a4().marginsCm(1.5).build();
        assertEquals(850, setup.topTwips());
    }

    @Test
    void perSideMarginsOverrideTheBulkSetter() {
        PageSetup setup = PageSetup.builder().a4().marginsTwips(851).top(1440).build();
        assertEquals(1440, setup.topTwips());
        assertEquals(851, setup.bottomTwips());
    }

    @Test
    void marginsInches() {
        PageSetup setup = PageSetup.builder().a4().marginsInches(1.0).build();
        assertEquals(1440, setup.leftTwips());
    }

    @Test
    void emitsSectPrWithSizeAndMargins() {
        SectPr sectPr = PageSetup.a4().toSectPr();
        assertEquals(BigInteger.valueOf(11906), sectPr.getPgSz().getW());
        assertEquals(BigInteger.valueOf(16838), sectPr.getPgSz().getH());
        assertEquals(BigInteger.valueOf(851), sectPr.getPgMar().getTop());
        assertEquals(BigInteger.valueOf(851), sectPr.getPgMar().getRight());
        assertEquals(BigInteger.valueOf(851), sectPr.getPgMar().getBottom());
        assertEquals(BigInteger.valueOf(851), sectPr.getPgMar().getLeft());
    }

    @Test
    void rejectsNegativeMargin() {
        assertThrows(DocumentGenerationException.class,
                () -> PageSetup.builder().a4().top(-1).build());
    }

    @Test
    void rejectsMarginsWiderThanThePage() {
        assertThrows(DocumentGenerationException.class,
                () -> PageSetup.builder().a4().left(6000).right(6000).build());
    }

    @Test
    void rejectsMarginsTallerThanThePage() {
        assertThrows(DocumentGenerationException.class,
                () -> PageSetup.builder().a4().top(9000).bottom(9000).build());
    }

    @Test
    void rejectsNonPositivePageSize() {
        assertThrows(DocumentGenerationException.class,
                () -> PageSetup.builder().pageSizeTwips(0, 16838).build());
    }
}
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd docx-service && JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test`
Expected: FAIL — compilation error, `cannot find symbol: class PageSetup`.

- [ ] **Step 3: Write minimal implementation**

Create `docx-service/src/main/java/com/example/docx/page/PageSetup.java`:

```java
package com.example.docx.page;

import com.example.docx.DocumentGenerationException;
import com.example.docx.Units;
import java.math.BigInteger;
import org.docx4j.jaxb.Context;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.SectPr;

/**
 * Immutable page geometry: size and the four margins, all in twips.
 *
 * <p>A4's dimensions are literals, not conversions. Deriving 297 mm gives 16838.7,
 * which rounds to 16839 and does not match the value Word writes.
 */
public final class PageSetup {

    /** A4 portrait width in twips, as written by Word. */
    public static final int A4_WIDTH_TWIPS = 11906;

    /** A4 portrait height in twips, as written by Word. */
    public static final int A4_HEIGHT_TWIPS = 16838;

    /** Default margin, ~1.5 cm. Note {@code marginsCm(1.5)} yields 850, not this. */
    public static final int DEFAULT_MARGIN_TWIPS = 851;

    private final int pageWidthTwips;
    private final int pageHeightTwips;
    private final int topTwips;
    private final int rightTwips;
    private final int bottomTwips;
    private final int leftTwips;

    private PageSetup(Builder b) {
        this.pageWidthTwips = b.pageWidthTwips;
        this.pageHeightTwips = b.pageHeightTwips;
        this.topTwips = b.topTwips;
        this.rightTwips = b.rightTwips;
        this.bottomTwips = b.bottomTwips;
        this.leftTwips = b.leftTwips;
    }

    public static Builder builder() {
        return new Builder();
    }

    /** A4 portrait with the default 851-twip margins. */
    public static PageSetup a4() {
        return builder().a4().build();
    }

    public int pageWidthTwips() {
        return pageWidthTwips;
    }

    public int pageHeightTwips() {
        return pageHeightTwips;
    }

    public int topTwips() {
        return topTwips;
    }

    public int rightTwips() {
        return rightTwips;
    }

    public int bottomTwips() {
        return bottomTwips;
    }

    public int leftTwips() {
        return leftTwips;
    }

    /** Builds the {@code w:sectPr} describing this page. */
    public SectPr toSectPr() {
        ObjectFactory factory = Context.getWmlObjectFactory();
        SectPr sectPr = factory.createSectPr();

        SectPr.PgSz pgSz = factory.createSectPrPgSz();
        pgSz.setW(BigInteger.valueOf(pageWidthTwips));
        pgSz.setH(BigInteger.valueOf(pageHeightTwips));
        sectPr.setPgSz(pgSz);

        SectPr.PgMar pgMar = factory.createSectPrPgMar();
        pgMar.setTop(BigInteger.valueOf(topTwips));
        pgMar.setRight(BigInteger.valueOf(rightTwips));
        pgMar.setBottom(BigInteger.valueOf(bottomTwips));
        pgMar.setLeft(BigInteger.valueOf(leftTwips));
        sectPr.setPgMar(pgMar);

        return sectPr;
    }

    /** Fluent builder. Every field is defaulted; {@code PageSetup.a4()} is valid alone. */
    public static final class Builder {

        private int pageWidthTwips = A4_WIDTH_TWIPS;
        private int pageHeightTwips = A4_HEIGHT_TWIPS;
        private int topTwips = DEFAULT_MARGIN_TWIPS;
        private int rightTwips = DEFAULT_MARGIN_TWIPS;
        private int bottomTwips = DEFAULT_MARGIN_TWIPS;
        private int leftTwips = DEFAULT_MARGIN_TWIPS;

        private Builder() {
        }

        public Builder a4() {
            this.pageWidthTwips = A4_WIDTH_TWIPS;
            this.pageHeightTwips = A4_HEIGHT_TWIPS;
            return this;
        }

        public Builder pageSizeTwips(int width, int height) {
            this.pageWidthTwips = width;
            this.pageHeightTwips = height;
            return this;
        }

        public Builder marginsTwips(int all) {
            this.topTwips = all;
            this.rightTwips = all;
            this.bottomTwips = all;
            this.leftTwips = all;
            return this;
        }

        public Builder marginsCm(double all) {
            return marginsTwips(Units.cmToTwips(all));
        }

        public Builder marginsInches(double all) {
            return marginsTwips(Units.inchesToTwips(all));
        }

        public Builder top(int twips) {
            this.topTwips = twips;
            return this;
        }

        public Builder right(int twips) {
            this.rightTwips = twips;
            return this;
        }

        public Builder bottom(int twips) {
            this.bottomTwips = twips;
            return this;
        }

        public Builder left(int twips) {
            this.leftTwips = twips;
            return this;
        }

        public PageSetup build() {
            requirePositive(pageWidthTwips, "page width");
            requirePositive(pageHeightTwips, "page height");
            requireNonNegative(topTwips, "top margin");
            requireNonNegative(rightTwips, "right margin");
            requireNonNegative(bottomTwips, "bottom margin");
            requireNonNegative(leftTwips, "left margin");

            if (leftTwips + rightTwips >= pageWidthTwips) {
                throw new DocumentGenerationException(
                        "left + right margins (" + (leftTwips + rightTwips)
                                + " twips) leave no width on a " + pageWidthTwips + "-twip page");
            }
            if (topTwips + bottomTwips >= pageHeightTwips) {
                throw new DocumentGenerationException(
                        "top + bottom margins (" + (topTwips + bottomTwips)
                                + " twips) leave no height on a " + pageHeightTwips + "-twip page");
            }
            return new PageSetup(this);
        }

        private static void requirePositive(int value, String field) {
            if (value <= 0) {
                throw new DocumentGenerationException(
                        field + " must be greater than zero twips, got " + value);
            }
        }

        private static void requireNonNegative(int value, String field) {
            if (value < 0) {
                throw new DocumentGenerationException(
                        field + " must not be negative, got " + value);
            }
        }
    }
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd docx-service && JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test`
Expected: PASS — `Tests run: 18, Failures: 0, Errors: 0, Skipped: 0`.

- [ ] **Step 5: Commit**

```bash
git add docx-service/src/main/java/com/example/docx/page/PageSetup.java docx-service/src/test/java/com/example/docx/page/PageSetupTest.java
git commit -m "feat: add PageSetup with A4 defaults and margin validation"
```

---

### Task 4: `HeadingStyle`

**Files:**
- Create: `docx-service/src/main/java/com/example/docx/style/HeadingStyle.java`
- Test: `docx-service/src/test/java/com/example/docx/style/HeadingStyleTest.java`

**Interfaces:**
- Consumes: `Units`, `DocumentGenerationException` from Tasks 1–2.
- Produces:
  - `public static HeadingStyle.Builder builder()`
  - `public static HeadingStyle defaults()`
  - Accessors `String fontFamily()`, `double sizePt()`, `boolean bold()`, `boolean italic()`, `String colorHex()` — `colorHex()` returns bare uppercase `RRGGBB`
  - `public org.docx4j.wml.RPr toRPr()`
  - Builder methods returning `Builder`: `font(String)`, `sizePt(double)`, `bold(boolean)`, `italic(boolean)`, `color(String)`, and `build()`.

Defaults: font `Calibri Light`, 20 pt, bold `true`, italic `false`, colour `1F4E79`.

- [ ] **Step 1: Write the failing test**

Create `docx-service/src/test/java/com/example/docx/style/HeadingStyleTest.java`:

```java
package com.example.docx.style;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.example.docx.DocumentGenerationException;
import java.math.BigInteger;
import org.docx4j.wml.RPr;
import org.junit.jupiter.api.Test;

class HeadingStyleTest {

    @Test
    void defaultsMatchTheSpec() {
        HeadingStyle style = HeadingStyle.defaults();
        assertEquals("Calibri Light", style.fontFamily());
        assertEquals(20.0, style.sizePt());
        assertTrue(style.bold());
        assertEquals(false, style.italic());
        assertEquals("1F4E79", style.colorHex());
    }

    @Test
    void stripsLeadingHashAndUppercases() {
        assertEquals("1F4E79", HeadingStyle.builder().color("#1f4e79").build().colorHex());
        assertEquals("1F4E79", HeadingStyle.builder().color("1f4e79").build().colorHex());
        assertEquals("ABCDEF", HeadingStyle.builder().color("#AbCdEf").build().colorHex());
    }

    @Test
    void emitsHalfPointSize() {
        RPr rPr = HeadingStyle.builder().sizePt(20).build().toRPr();
        assertEquals(BigInteger.valueOf(40), rPr.getSz().getVal());
        assertEquals(BigInteger.valueOf(40), rPr.getSzCs().getVal());
    }

    @Test
    void emitsFontOnAsciiAndHAnsi() {
        RPr rPr = HeadingStyle.builder().font("Arial").build().toRPr();
        assertEquals("Arial", rPr.getRFonts().getAscii());
        assertEquals("Arial", rPr.getRFonts().getHAnsi());
    }

    @Test
    void emitsColourWithoutHash() {
        RPr rPr = HeadingStyle.builder().color("#1F4E79").build().toRPr();
        assertEquals("1F4E79", rPr.getColor().getVal());
    }

    @Test
    void boldAndItalicAreOmittedWhenFalse() {
        RPr rPr = HeadingStyle.builder().bold(false).italic(false).build().toRPr();
        assertNull(rPr.getB());
        assertNull(rPr.getI());
    }

    @Test
    void boldAndItalicArePresentWhenTrue() {
        RPr rPr = HeadingStyle.builder().bold(true).italic(true).build().toRPr();
        assertTrue(rPr.getB().isVal());
        assertTrue(rPr.getI().isVal());
    }

    @Test
    void rejectsMalformedColour() {
        assertThrows(DocumentGenerationException.class,
                () -> HeadingStyle.builder().color("blue").build());
        assertThrows(DocumentGenerationException.class,
                () -> HeadingStyle.builder().color("#12345").build());
        assertThrows(DocumentGenerationException.class,
                () -> HeadingStyle.builder().color("#1234567").build());
        assertThrows(DocumentGenerationException.class,
                () -> HeadingStyle.builder().color("#12345G").build());
        assertThrows(DocumentGenerationException.class,
                () -> HeadingStyle.builder().color(null).build());
    }

    @Test
    void rejectsNonPositiveSize() {
        assertThrows(DocumentGenerationException.class,
                () -> HeadingStyle.builder().sizePt(0).build());
        assertThrows(DocumentGenerationException.class,
                () -> HeadingStyle.builder().sizePt(-4).build());
    }

    @Test
    void rejectsBlankFont() {
        assertThrows(DocumentGenerationException.class,
                () -> HeadingStyle.builder().font("   ").build());
        assertThrows(DocumentGenerationException.class,
                () -> HeadingStyle.builder().font(null).build());
    }
}
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd docx-service && JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test`
Expected: FAIL — compilation error, `cannot find symbol: class HeadingStyle`.

- [ ] **Step 3: Write minimal implementation**

Create `docx-service/src/main/java/com/example/docx/style/HeadingStyle.java`:

```java
package com.example.docx.style;

import com.example.docx.DocumentGenerationException;
import com.example.docx.Units;
import java.math.BigInteger;
import java.util.regex.Pattern;
import org.docx4j.jaxb.Context;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.Color;
import org.docx4j.wml.HpsMeasure;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPr;

/**
 * Immutable run formatting for the heading, emitted as direct formatting.
 *
 * <p>No built-in {@code Heading1} style id is referenced, so the heading will not
 * appear in Word's Navigation pane. That is the phase-one trade-off for rendering
 * identically regardless of the document stylesheet.
 */
public final class HeadingStyle {

    private static final Pattern HEX = Pattern.compile("[0-9A-Fa-f]{6}");

    public static final String DEFAULT_FONT = "Calibri Light";
    public static final double DEFAULT_SIZE_PT = 20.0;
    public static final String DEFAULT_COLOR = "1F4E79";

    private final String fontFamily;
    private final double sizePt;
    private final boolean bold;
    private final boolean italic;
    private final String colorHex;

    private HeadingStyle(Builder b) {
        this.fontFamily = b.fontFamily;
        this.sizePt = b.sizePt;
        this.bold = b.bold;
        this.italic = b.italic;
        this.colorHex = b.colorHex;
    }

    public static Builder builder() {
        return new Builder();
    }

    public static HeadingStyle defaults() {
        return builder().build();
    }

    public String fontFamily() {
        return fontFamily;
    }

    public double sizePt() {
        return sizePt;
    }

    public boolean bold() {
        return bold;
    }

    public boolean italic() {
        return italic;
    }

    /** Bare uppercase {@code RRGGBB}, never prefixed with {@code #}. */
    public String colorHex() {
        return colorHex;
    }

    /** Builds the {@code w:rPr} for this style. */
    public RPr toRPr() {
        ObjectFactory factory = Context.getWmlObjectFactory();
        RPr rPr = factory.createRPr();

        RFonts fonts = factory.createRFonts();
        fonts.setAscii(fontFamily);
        fonts.setHAnsi(fontFamily);
        rPr.setRFonts(fonts);

        HpsMeasure size = factory.createHpsMeasure();
        size.setVal(BigInteger.valueOf(Units.pointsToHalfPoints(sizePt)));
        rPr.setSz(size);
        rPr.setSzCs(size);

        if (bold) {
            BooleanDefaultTrue on = factory.createBooleanDefaultTrue();
            on.setVal(Boolean.TRUE);
            rPr.setB(on);
        }
        if (italic) {
            BooleanDefaultTrue on = factory.createBooleanDefaultTrue();
            on.setVal(Boolean.TRUE);
            rPr.setI(on);
        }

        Color color = factory.createColor();
        color.setVal(colorHex);
        rPr.setColor(color);

        return rPr;
    }

    /** Fluent builder. Every field is defaulted; {@code HeadingStyle.defaults()} is valid alone. */
    public static final class Builder {

        private String fontFamily = DEFAULT_FONT;
        private double sizePt = DEFAULT_SIZE_PT;
        private boolean bold = true;
        private boolean italic = false;
        private String colorHex = DEFAULT_COLOR;

        private Builder() {
        }

        public Builder font(String fontFamily) {
            this.fontFamily = fontFamily;
            return this;
        }

        public Builder sizePt(double sizePt) {
            this.sizePt = sizePt;
            return this;
        }

        public Builder bold(boolean bold) {
            this.bold = bold;
            return this;
        }

        public Builder italic(boolean italic) {
            this.italic = italic;
            return this;
        }

        /** Accepts {@code #RRGGBB} or {@code RRGGBB}, any case. */
        public Builder color(String hex) {
            this.colorHex = hex;
            return this;
        }

        public HeadingStyle build() {
            if (fontFamily == null || fontFamily.isBlank()) {
                throw new DocumentGenerationException("font family must not be blank");
            }
            if (!(sizePt > 0) || !Double.isFinite(sizePt)) {
                throw new DocumentGenerationException(
                        "font size must be greater than zero points, got " + sizePt);
            }
            this.fontFamily = fontFamily.trim();
            this.colorHex = normaliseColour(colorHex);
            return new HeadingStyle(this);
        }

        private static String normaliseColour(String hex) {
            if (hex == null) {
                throw new DocumentGenerationException("colour must not be null");
            }
            String bare = hex.startsWith("#") ? hex.substring(1) : hex;
            if (!HEX.matcher(bare).matches()) {
                throw new DocumentGenerationException(
                        "colour must be 6 hex digits, optionally prefixed with '#', got '" + hex + "'");
            }
            return bare.toUpperCase(java.util.Locale.ROOT);
        }
    }
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd docx-service && JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test`
Expected: PASS — `Tests run: 28, Failures: 0, Errors: 0, Skipped: 0`.

- [ ] **Step 5: Commit**

```bash
git add docx-service/src/main/java/com/example/docx/style/HeadingStyle.java docx-service/src/test/java/com/example/docx/style/HeadingStyleTest.java
git commit -m "feat: add HeadingStyle with hex normalisation and half-point sizing"
```

---

### Task 5: `Headings` content factory

**Files:**
- Create: `docx-service/src/main/java/com/example/docx/content/Headings.java`
- Test: `docx-service/src/test/java/com/example/docx/content/HeadingsTest.java`

**Interfaces:**
- Consumes: `HeadingStyle` (Task 4), `DocumentGenerationException` (Task 1).
- Produces: `public static org.docx4j.wml.P heading(String text, HeadingStyle style)` on a final, non-instantiable class.

**Object-model trap, verified by running it.** Whether a `Text` is wrapped in a `jakarta.xml.bind.JAXBElement` depends on how it got into the tree, not on the model:

| Origin | `run.getContent().get(0)` is |
| --- | --- |
| Freshly built in memory (this task) | a bare `org.docx4j.wml.Text` |
| Reloaded via `WordprocessingMLPackage.load()` (Task 6) | a `JAXBElement<Text>` |

So this task's test reads the value directly, while Task 6's must call `getValue()`. The helper below tolerates both, which is what makes it safe to copy between the two test classes.

- [ ] **Step 1: Write the failing test**

Create `docx-service/src/test/java/com/example/docx/content/HeadingsTest.java`:

```java
package com.example.docx.content;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import com.example.docx.DocumentGenerationException;
import com.example.docx.style.HeadingStyle;
import jakarta.xml.bind.JAXBElement;
import java.math.BigInteger;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Text;
import org.junit.jupiter.api.Test;

class HeadingsTest {

    /** Freshly built runs hold a bare Text; reloaded ones hold a JAXBElement. */
    private static Text textOf(R run) {
        Object first = run.getContent().get(0);
        return (Text) (first instanceof JAXBElement<?> je ? je.getValue() : first);
    }

    @Test
    void buildsParagraphWithOneStyledRun() {
        P p = Headings.heading("Quarterly Report", HeadingStyle.defaults());

        assertEquals(1, p.getContent().size());
        R run = (R) p.getContent().get(0);

        assertEquals("Quarterly Report", textOf(run).getValue());
        assertEquals(BigInteger.valueOf(40), run.getRPr().getSz().getVal());
        assertEquals("1F4E79", run.getRPr().getColor().getVal());
    }

    @Test
    void preservesSurroundingWhitespace() {
        P p = Headings.heading("  spaced  ", HeadingStyle.defaults());
        R run = (R) p.getContent().get(0);
        assertEquals("preserve", textOf(run).getSpace());
        assertEquals("  spaced  ", textOf(run).getValue());
    }

    @Test
    void rejectsBlankText() {
        assertThrows(DocumentGenerationException.class,
                () -> Headings.heading("   ", HeadingStyle.defaults()));
        assertThrows(DocumentGenerationException.class,
                () -> Headings.heading(null, HeadingStyle.defaults()));
    }

    @Test
    void rejectsNullStyle() {
        assertThrows(DocumentGenerationException.class,
                () -> Headings.heading("Title", null));
    }
}
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd docx-service && JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test`
Expected: FAIL — compilation error, `cannot find symbol: class Headings`.

- [ ] **Step 3: Write minimal implementation**

Create `docx-service/src/main/java/com/example/docx/content/Headings.java`:

```java
package com.example.docx.content;

import com.example.docx.DocumentGenerationException;
import com.example.docx.style.HeadingStyle;
import org.docx4j.jaxb.Context;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Text;

/**
 * Builds heading paragraphs.
 *
 * <p>A pure function of its arguments: it never touches the package, so it composes
 * freely and needs no fixtures.
 */
public final class Headings {

    private Headings() {
    }

    /** A {@code w:p} holding one styled {@code w:r} with the given text. */
    public static P heading(String text, HeadingStyle style) {
        if (text == null || text.isBlank()) {
            throw new DocumentGenerationException("heading text must not be blank");
        }
        if (style == null) {
            throw new DocumentGenerationException("heading style must not be null");
        }

        ObjectFactory factory = Context.getWmlObjectFactory();

        Text value = factory.createText();
        value.setValue(text);
        // Without xml:space=preserve, leading and trailing spaces vanish silently.
        value.setSpace("preserve");

        R run = factory.createR();
        run.getContent().add(value);
        run.setRPr(style.toRPr());

        P paragraph = factory.createP();
        paragraph.getContent().add(run);
        return paragraph;
    }
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd docx-service && JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test`
Expected: PASS — `Tests run: 32, Failures: 0, Errors: 0, Skipped: 0`.

- [ ] **Step 5: Commit**

```bash
git add docx-service/src/main/java/com/example/docx/content/Headings.java docx-service/src/test/java/com/example/docx/content/HeadingsTest.java
git commit -m "feat: add Headings paragraph factory"
```

---

### Task 6: `WordDocument` facade

**Files:**
- Create: `docx-service/src/main/java/com/example/docx/WordDocument.java`
- Test: `docx-service/src/test/java/com/example/docx/WordDocumentRoundTripTest.java`

**Interfaces:**
- Consumes: `PageSetup` (Task 3), `HeadingStyle` (Task 4), `Headings` (Task 5), `DocumentGenerationException` (Task 1).
- Produces:
  - `public static WordDocument.Builder builder()`
  - Builder methods returning `Builder`: `pageSetup(PageSetup)`, `heading(String text, HeadingStyle style)`, `heading(String text)`, and `build()` returning `WordDocument`
  - `public byte[] toByteArray()`
  - `public void writeTo(java.io.OutputStream out)`

`build()` performs all docx4j work and holds the finished package; `toByteArray()` and `writeTo()` serialise it. Not thread safe — one instance per document.

- [ ] **Step 1: Write the failing test**

Create `docx-service/src/test/java/com/example/docx/WordDocumentRoundTripTest.java`:

```java
package com.example.docx;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
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
    void writeToMatchesToByteArray() throws Exception {
        WordDocument doc = WordDocument.builder()
                .pageSetup(PageSetup.a4())
                .heading("Quarterly Report", HeadingStyle.defaults())
                .build();

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        doc.writeTo(out);

        assertArrayEquals(doc.toByteArray(), out.toByteArray());
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd docx-service && JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test`
Expected: FAIL — compilation error, `cannot find symbol: class WordDocument`.

- [ ] **Step 3: Write minimal implementation**

Create `docx-service/src/main/java/com/example/docx/WordDocument.java`:

```java
package com.example.docx;

import com.example.docx.content.Headings;
import com.example.docx.page.PageSetup;
import com.example.docx.style.HeadingStyle;
import java.io.ByteArrayOutputStream;
import java.io.OutputStream;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Body;

/**
 * Fluent facade over a single Word document.
 *
 * <p>Wraps mutable docx4j state and is <strong>not thread safe</strong>. Build one
 * per document; construction is cheap and that is the natural lifetime in a request
 * handler.
 *
 * <p>The first document generated in a JVM pays a one-off JAXB context
 * initialisation of roughly one second. Warm it at startup with
 * {@code Context.getWmlObjectFactory()} if latency matters.
 */
public final class WordDocument {

    private final WordprocessingMLPackage pkg;

    private WordDocument(WordprocessingMLPackage pkg) {
        this.pkg = pkg;
    }

    public static Builder builder() {
        return new Builder();
    }

    /** Serialises the document. */
    public byte[] toByteArray() {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        writeTo(out);
        return out.toByteArray();
    }

    /** Serialises the document to {@code out}, without buffering the whole file. */
    public void writeTo(OutputStream out) {
        if (out == null) {
            throw new DocumentGenerationException("output stream must not be null");
        }
        try {
            pkg.save(out);
        } catch (Docx4JException e) {
            throw new DocumentGenerationException("failed to serialise the document", e);
        }
    }

    /** Fluent builder. All docx4j work happens in {@link #build()}. */
    public static final class Builder {

        private PageSetup pageSetup = PageSetup.a4();
        private String headingText;
        private HeadingStyle headingStyle = HeadingStyle.defaults();

        private Builder() {
        }

        public Builder pageSetup(PageSetup pageSetup) {
            if (pageSetup == null) {
                throw new DocumentGenerationException("page setup must not be null");
            }
            this.pageSetup = pageSetup;
            return this;
        }

        /** Sets the heading text and its style. */
        public Builder heading(String text, HeadingStyle style) {
            if (style == null) {
                throw new DocumentGenerationException("heading style must not be null");
            }
            this.headingText = text;
            this.headingStyle = style;
            return this;
        }

        /** Sets the heading text, keeping {@link HeadingStyle#defaults()}. */
        public Builder heading(String text) {
            this.headingText = text;
            return this;
        }

        public WordDocument build() {
            if (headingText == null || headingText.isBlank()) {
                throw new DocumentGenerationException(
                        "a heading is required; call heading(String) before build()");
            }
            try {
                WordprocessingMLPackage pkg = WordprocessingMLPackage.createPackage();
                MainDocumentPart mainDocumentPart = pkg.getMainDocumentPart();

                Body body = mainDocumentPart.getJaxbElement().getBody();
                body.setSectPr(pageSetup.toSectPr());

                mainDocumentPart.getContent().add(Headings.heading(headingText, headingStyle));

                return new WordDocument(pkg);
            } catch (Docx4JException e) {
                throw new DocumentGenerationException("failed to create the document package", e);
            }
        }
    }
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd docx-service && JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B test`
Expected: PASS — `Tests run: 40, Failures: 0, Errors: 0, Skipped: 0`.

- [ ] **Step 5: Commit**

```bash
git add docx-service/src/main/java/com/example/docx/WordDocument.java docx-service/src/test/java/com/example/docx/WordDocumentRoundTripTest.java
git commit -m "feat: add WordDocument facade with round-trip coverage"
```

---

### Task 7: Runnable sample and README

**Files:**
- Create: `docx-service/src/main/java/com/example/docx/sample/SampleMain.java`
- Create: `docx-service/src/main/resources/simplelogger.properties`
- Create: `docx-service/README.md`

**Interfaces:**
- Consumes: `WordDocument`, `PageSetup`, `HeadingStyle`.
- Produces: nothing other tasks depend on. This is the terminal task.

- [ ] **Step 1: Write the sample**

Create `docx-service/src/main/java/com/example/docx/sample/SampleMain.java`:

```java
package com.example.docx.sample;

import com.example.docx.WordDocument;
import com.example.docx.page.PageSetup;
import com.example.docx.style.HeadingStyle;
import java.io.IOException;
import java.io.OutputStream;
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
                .build();

        Path target = Path.of("target", "sample.docx");
        Files.createDirectories(target.getParent());
        try (OutputStream out = Files.newOutputStream(target)) {
            document.writeTo(out);
        }
        System.out.println("Wrote " + target.toAbsolutePath());
    }
}
```

- [ ] **Step 2: Quieten the sample's logging**

Create `docx-service/src/main/resources/simplelogger.properties`:

```properties
org.slf4j.simpleLogger.defaultLogLevel=warn
```

- [ ] **Step 3: Run the sample**

Run: `cd docx-service && JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B compile exec:java`
Expected: `BUILD SUCCESS` and a line reading `Wrote /…/docx-service/target/sample.docx`.

- [ ] **Step 4: Verify the output is a real docx**

Run: `cd docx-service && unzip -l target/sample.docx`
Expected: the listing contains `word/document.xml`, `[Content_Types].xml` and `_rels/.rels`.

Run: `cd docx-service && unzip -p target/sample.docx word/document.xml | grep -o 'w:top="[0-9]*"'`
Expected: `w:top="851"`.

- [ ] **Step 5: Write the README**

Create `docx-service/README.md`:

````markdown
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
JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn compile exec:java
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

For large documents prefer `writeTo(out)` over `toByteArray()`, so the whole
file never sits in heap.

## Layout

| Package | Contents |
| --- | --- |
| `com.example.docx` | `WordDocument` (facade), `Units`, `DocumentGenerationException` |
| `…​.page` | `PageSetup` — page size and margins |
| `…​.style` | `HeadingStyle` — run formatting |
| `…​.content` | `Headings` — stateless paragraph factory |
| `…​.sample` | Runnable `main` |

`content` and `style` never touch `WordprocessingMLPackage`, so they are pure
functions of their arguments and need no fixtures. `page` and the facade own
the package.

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
````

- [ ] **Step 6: Run the full suite one last time**

Run: `cd docx-service && JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -B clean test`
Expected: `Tests run: 40, Failures: 0, Errors: 0, Skipped: 0` and `BUILD SUCCESS`.

- [ ] **Step 7: Commit**

```bash
git add docx-service/src/main/java/com/example/docx/sample/SampleMain.java docx-service/src/main/resources/simplelogger.properties docx-service/README.md
git commit -m "feat: add runnable sample and README"
```

---

## Definition of Done

- `mvn clean test` passes with 40 tests.
- `mvn compile exec:java` writes `target/sample.docx`, and `word/document.xml` inside it contains `w:top="851"` and `w:w="11906"`.
- The document opens in Word without a repair prompt.
- No class outside `page/` and `WordDocument` references `WordprocessingMLPackage`.
