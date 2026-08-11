# docx-service — Phase-One Review Fixes

**Date:** 2026-08-11
**Status:** Approved for planning
**Builds on:** [2026-08-10-docx-service-design.md](2026-08-10-docx-service-design.md), [2026-08-11-tables-design.md](2026-08-11-tables-design.md), [2026-08-11-svg-images-design.md](2026-08-11-svg-images-design.md)

## Goal

Close the two review findings from the whole-branch review of `feat/phase-one-docx`
(commits `2749e9b..9ebd5e9`), then merge the branch into `main`. Scope is deliberately
narrow — the reviewer's other three Minor findings (renaming `TableBorderStyle.box()`,
typing `DocumentContent.toBodyElement`'s return value, and a multi-image round-trip
test) are explicitly out of scope for this pass.

## Findings to fix

| # | Finding | Fix |
| --- | --- | --- |
| R1 | `WordDocument.Builder.table()` calls `pageSetup.usableWidthTwips()` **eagerly**, at the `table(...)` call site, while `svgImage()` defers its use of `pageSetup` into the `pkg -> ...` lambda that only runs at `build()`. If a caller calls `.table(...)` before `.pageSetup(...)` — a natural ordering — the table silently locks in the default page width instead of the caller's configured one, with no error. | Make `table()` match `svgImage()`: store the raw `headers`/`rows`/`style` arguments in the builder, and move the `pageSetup.usableWidthTwips()` call inside the build-time lambda so it always sees the final `pageSetup`. No public API change — behavior only changes for the previously-broken call ordering. |
| R2 | `ImageParts.attachSvgExtension` builds a `DocumentBuilderFactory` via `newInstance()` with no XXE hardening. Not exploitable today (the parsed string is fully library-constructed from fixed namespaces plus a docx4j-generated relationship id), but the factory has nothing to catch a future edit that threads a less-trusted string through the same code path. | Disable DTD processing on the factory before use: `setFeature("http://apache.org/xml/features/disallow-doctype-decl", true)`, plus the standard external-general/parameter-entity features set to `false`, matching the usual "hardened `DocumentBuilderFactory`" recipe. |

## Testing

- **R1:** Add a test that calls `.table(...)` *before* `.pageSetup(...)` on the builder,
  then builds and reloads the document, asserting the table's column widths sum to the
  usable width implied by the *final* page setup, not the default. Place alongside the
  existing `TableDocumentTest` / `WordDocument` builder tests.
- **R2:** No new test — this is a defensive factory-configuration change with no
  behavioral surface to assert against (the existing SVG round-trip tests already cover
  correct-input behavior and continue to pass unchanged).

## Out of scope

- `TableBorderStyle.box()` rename, `DocumentContent` return-type tightening, and a
  multi-image round-trip test — left for a future pass.
- No new dependencies, no public API removals.

## Merge

Fix on `feat/phase-one-docx`, run the full suite (`JAVA_HOME=$(/usr/libexec/java_home -v 25) mvn clean test`),
then merge into `main` locally. No git remote is configured for this repo, so this is a
local-only merge — no PR.
