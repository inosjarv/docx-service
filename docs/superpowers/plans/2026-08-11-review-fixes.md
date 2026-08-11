# docx-service Review Fixes Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Close the two review findings from the `feat/phase-one-docx` whole-branch review (R1: `WordDocument.Builder.table()` reads `pageSetup` eagerly instead of deferring to `build()`; R2: `ImageParts.attachSvgExtension`'s XML parser has no XXE hardening), then merge `feat/phase-one-docx` into `main`.

**Architecture:** No new files, no public API changes, no new dependencies. R1 changes the no-width `table(...)` overload in `WordDocument.Builder` to close over `pageSetup` inside a build-time lambda, mirroring the existing `svgImage(...)` pattern in the same class. R2 adds two hardening calls to the `DocumentBuilderFactory` already used in `ImageParts`.

**Tech Stack:** Java 25, Maven, docx4j 17.0.2, JUnit 6.1.3 (Jupiter).

## Global Constraints

- Java release level is exactly `25`. Run Maven with `JAVA_HOME=$(/usr/libexec/java_home -v 25)`.
- docx4j is `org.docx4j:docx4j-JAXB-ReferenceImpl:17.0.2`.
- `DocumentGenerationException extends RuntimeException` is the only exception type this module throws deliberately. Neither fix in this plan introduces a new throw site.
- Package root is `com.example.docx`. Module directory is `docx-service/`.
- Both fixes preserve existing public method signatures — this is a behavior-only change (R1) and a defensive-config-only change (R2), not an API change.

## File Structure

| File | Change |
| --- | --- |
| `src/main/java/com/example/docx/WordDocument.java` | R1: defer `table(headers, rows, style)`'s width resolution to `build()` |
| `src/test/java/com/example/docx/TableDocumentTest.java` | R1: add a test proving `table()` before `pageSetup()` still picks up the final page setup |
| `DOCUMENTATION.md` | R1: correct the now-stale "reads the page setup at call time" caveat |
| `src/main/java/com/example/docx/part/ImageParts.java` | R2: harden the `DocumentBuilderFactory` used in `attachSvgExtension` |

---

### Task 1: Defer `table()`'s page-setup read to `build()`

**Files:**
- Modify: `docx-service/src/main/java/com/example/docx/WordDocument.java:136-158` (the `svgImage` Javadoc/method for reference, and the no-width `table` overload)
- Modify: `docx-service/src/test/java/com/example/docx/TableDocumentTest.java`
- Modify: `docx-service/DOCUMENTATION.md:221-223`

**Interfaces:**
- Consumes: `PageSetup.usableWidthTwips()` (existing, unchanged), `Tables.of(List<String>, List<List<String>>, TableStyle, int)` (existing, unchanged), `Context.getWmlObjectFactory()` (existing, unchanged).
- Produces: no new public signatures. `Builder.table(List<String> headers, List<List<String>> rows, TableStyle style)` keeps its exact signature and return type (`Builder`); only its internal timing changes.

- [ ] **Step 1: Write the failing test**

Open `docx-service/src/test/java/com/example/docx/TableDocumentTest.java` and add this test at the end of the class, just before the final closing `}`:

```java
    @Test
    void tableWidthReflectsPageSetupSetAfterTable() throws Exception {
        PageSetup wide = PageSetup.builder()
                .pageSizeTwips(20000, 15840)
                .marginsTwips(851)
                .build();

        byte[] bytes = WordDocument.builder()
                .table(HEADERS, ROWS, TableStyle.defaults())
                .pageSetup(wide)
                .build()
                .toByteArray();

        Tbl table = (Tbl) unwrapped(reload(bytes)).stream()
                .filter(o -> o instanceof Tbl).findFirst().orElseThrow();

        int sum = table.getTblGrid().getGridCol().stream()
                .mapToInt(c -> c.getW().intValue()).sum();
        assertEquals(wide.usableWidthTwips(), sum,
                "table() called before pageSetup() should still size to the final page setup");
    }
```

No new imports are needed — `PageSetup`, `TableStyle`, `Tbl`, `WordDocument`, `assertEquals` are all already imported in this file.

- [ ] **Step 2: Run the test to verify it fails**

Run:
```bash
cd docx-service && JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -q test -Dtest=TableDocumentTest#tableWidthReflectsPageSetupSetAfterTable
```
Expected: FAIL. The assertion reports the sum equal to `PageSetup.a4().usableWidthTwips()` (10204), not `wide.usableWidthTwips()` (18298) — proving `table()` locked in the default page width before `pageSetup(wide)` ran.

- [ ] **Step 3: Implement the fix**

In `docx-service/src/main/java/com/example/docx/WordDocument.java`, replace the no-width `table` overload and its Javadoc:

Replace this:
```java
        /**
         * Appends a table spanning the full usable page width.
         *
         * <p>The width is read from the current page setup <em>now</em>, not at
         * {@code build()}. Call {@link #pageSetup(PageSetup)} before this, or use the
         * four-argument overload with an explicit width — otherwise a later
         * {@code pageSetup} call leaves the table sized for the old page.
         */
        public Builder table(List<String> headers, List<List<String>> rows, TableStyle style) {
            return table(headers, rows, style, pageSetup.usableWidthTwips());
        }
```

With this:
```java
        /**
         * Appends a table spanning the full usable page width, followed by a spacer
         * paragraph.
         *
         * <p>The width is resolved from {@link #pageSetup(PageSetup)} at {@code build()}
         * time, so this method may be called before or after {@code pageSetup(...)} —
         * whichever {@code pageSetup} is in effect when {@code build()} runs wins. Use the
         * four-argument overload to pin an explicit width regardless of page setup.
         */
        public Builder table(List<String> headers, List<List<String>> rows, TableStyle style) {
            content.add(pkg -> Tables.of(headers, rows, style, pageSetup.usableWidthTwips()));
            // Two adjacent tables merge into one in Word, and a body ending in a table
            // rather than a paragraph is irregular. A spacer prevents both.
            P spacer = Context.getWmlObjectFactory().createP();
            content.add(pkg -> spacer);
            return this;
        }
```

The four-argument overload (`table(headers, rows, style, widthTwips)`) is unchanged — its width is already explicit at the call site, so there is nothing to defer.

- [ ] **Step 4: Run the test to verify it passes**

Run:
```bash
cd docx-service && JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -q test -Dtest=TableDocumentTest
```
Expected: PASS, all tests in `TableDocumentTest` (including the two pre-existing ones).

- [ ] **Step 5: Correct the stale caveat in DOCUMENTATION.md**

In `docx-service/DOCUMENTATION.md`, replace:

```markdown
**`table(headers, rows, style)` reads the page setup at call time**, unlike `svgImage`
which defers to `build()`. Call `pageSetup(...)` first, or pass an explicit width —
otherwise a later `pageSetup` call leaves the table sized for the old page.
```

With:

```markdown
**`table(headers, rows, style)` defers the page setup to `build()`**, the same as
`svgImage`. It may be called before or after `pageSetup(...)` with the same result. Pass
an explicit width with the four-argument overload to pin the width regardless of page
setup.
```

- [ ] **Step 6: Run the full test suite**

Run:
```bash
cd docx-service && JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn clean test
```
Expected: `BUILD SUCCESS`, all tests pass (90 existing + 1 new = 91).

- [ ] **Step 7: Commit**

```bash
cd docx-service
git add src/main/java/com/example/docx/WordDocument.java \
        src/test/java/com/example/docx/TableDocumentTest.java \
        DOCUMENTATION.md
git commit -m "fix: defer table() page-setup read to build(), matching svgImage"
```

---

### Task 2: Harden `ImageParts`'s XML parser against XXE

**Files:**
- Modify: `docx-service/src/main/java/com/example/docx/part/ImageParts.java:121-134`

**Interfaces:**
- Consumes: nothing new.
- Produces: no signature change. `attachSvgExtension(Inline, String)` stays `private static`, same parameters, same return type (`void`), same declared `throws Exception`.

- [ ] **Step 1: Add XXE hardening to the `DocumentBuilderFactory`**

In `docx-service/src/main/java/com/example/docx/part/ImageParts.java`, in `attachSvgExtension`, replace:

```java
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        // Mandatory, and false by default. Left off, the prefixed name parses as a
        // literal element in NO namespace and Word silently ignores the SVG.
        factory.setNamespaceAware(true);
```

With:

```java
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        // Mandatory, and false by default. Left off, the prefixed name parses as a
        // literal element in NO namespace and Word silently ignores the SVG.
        factory.setNamespaceAware(true);
        // The string parsed here is fully library-constructed today (fixed namespaces
        // plus a docx4j-generated relationship id), so this isn't reachable yet — but
        // hardening now means a future edit that threads less-trusted input through this
        // factory doesn't reintroduce XXE by accident.
        factory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
        factory.setFeature("http://xml.org/sax/features/external-general-entities", false);
        factory.setFeature("http://xml.org/sax/features/external-parameter-entities", false);
        factory.setXIncludeAware(false);
        factory.setExpandEntityReferences(false);
```

- [ ] **Step 2: Run the existing image tests to confirm nothing broke**

Run:
```bash
cd docx-service && JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn -q test -Dtest=ImagePartsTest
```
Expected: PASS, all existing tests in `ImagePartsTest` — the hardening flags reject external/DTD content, not the well-formed, DTD-free, non-DOCTYPE snippet this method builds internally, so behavior is unchanged.

- [ ] **Step 3: Run the full test suite**

Run:
```bash
cd docx-service && JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn clean test
```
Expected: `BUILD SUCCESS`, all 91 tests pass.

- [ ] **Step 4: Commit**

```bash
cd docx-service
git add src/main/java/com/example/docx/part/ImageParts.java
git commit -m "fix: harden ImageParts' DocumentBuilderFactory against XXE"
```

---

### Task 3: Merge `feat/phase-one-docx` into `main`

**Files:** none (repository operation only).

**Interfaces:** none.

- [ ] **Step 1: Confirm the working tree is clean and the branch is up to date**

Run:
```bash
cd docx-service && git status
```
Expected: `On branch feat/phase-one-docx` and `nothing to commit, working tree clean`.

- [ ] **Step 2: Confirm `main` has not diverged**

Run:
```bash
cd docx-service && git merge-base --is-ancestor main feat/phase-one-docx && echo "main is an ancestor: fast-forward is possible"
```
Expected: prints `main is an ancestor: fast-forward is possible`. (If this fails, stop and re-plan — it means `main` gained commits this plan didn't account for.)

- [ ] **Step 3: Switch to `main` and fast-forward merge**

Run:
```bash
cd docx-service
git checkout main
git merge --ff-only feat/phase-one-docx
```
Expected: `Fast-forward` output, no merge commit created, `main` now points at the same commit as `feat/phase-one-docx`.

- [ ] **Step 4: Verify the full suite still passes on `main`**

Run:
```bash
cd docx-service && JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn clean test
```
Expected: `BUILD SUCCESS`, all 91 tests pass.

- [ ] **Step 5: Confirm branch state**

Run:
```bash
cd docx-service && git log --oneline -1 main && git log --oneline -1 feat/phase-one-docx && git branch -vv
```
Expected: both branches at the same commit (the `fix: harden ImageParts...` commit from Task 2).

No further commit is needed for this task — `git merge --ff-only` only moves the `main` ref, it does not create new content to commit.
