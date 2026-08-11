# docx-service — Tables

**Date:** 2026-08-11
**Status:** Approved for planning
**Builds on:** [2026-08-11-multi-section-design.md](2026-08-11-multi-section-design.md)

## Goal

Add a table: one header row plus data rows, with borders configurable independently for
the header and for the body — each choosing which cell edges it draws, in a chosen
colour, width and line style.

## The interface change this forces

`DocumentContent` currently returns a paragraph:

```java
P toParagraph(WordprocessingMLPackage pkg);
```

A table is a `Tbl`, not a `P`. In OOXML a body holds a heterogeneous sequence of block
elements — `w:p` and `w:tbl` are both valid children — and docx4j models that as
`List<Object>`. So the method widens:

```java
Object toBodyElement(WordprocessingMLPackage pkg);   // a P or a Tbl
```

`Object` is loose, but it is exactly the type `MainDocumentPart.getContent()` already
uses. A sealed interface would impose typing that OOXML itself does not have, on a union
of two cases.

## Architecture

```
style/BorderLine.java         curated line styles -> STBorder
style/Edge.java               TOP, BOTTOM, LEFT, RIGHT
style/TableBorderStyle.java   colour, widthPt, line, and which edges are drawn
style/TableStyle.java         headerBorder, bodyBorder, header and body TextStyle
content/Tables.java           (headers, rows, style, widthTwips) -> Tbl.  Pure.
DocumentContent.java          toParagraph -> toBodyElement, returning Object
WordDocument.java             table(...) builder methods
```

`Tables` belongs in `content/` because a table needs no package: column widths arrive as
an `int`, not a `PageSetup`. The layering rule holds.

## Border model

Borders are per **cell edge**. Each group — header, body — picks which edges every cell
in it draws.

```java
TableStyle.builder()
        .headerBorder(TableBorderStyle.builder()
                .color("#1F4E79")
                .widthPt(1.0)
                .line(BorderLine.SINGLE)
                .edges(Edge.BOTTOM)
                .build())
        .bodyBorder(TableBorderStyle.builder()
                .color("#BFBFBF")
                .widthPt(0.5)
                .line(BorderLine.DOTTED)
                .edges(Edge.BOTTOM)
                .build())
        .build();
```

`edges(...)` is varargs. No edges, or `TableBorderStyle.none()`, draws nothing.
Convenience presets: `bottomOnly(colour, widthPt)`, `box(...)`, `grid(...)`.

| Wanted | Set |
| --- | --- |
| Rule under the header | header → `BOTTOM` |
| Rules between body rows | body → `BOTTOM` |
| Full grid | body → all four |
| Outer box only | header → `TOP, LEFT, RIGHT`; body → `LEFT, RIGHT, BOTTOM` |

### Why cell edges rather than OOXML's insideH/insideV

`w:tblBorders` is table-wide and cannot distinguish the header row from body rows, so
giving the header its own border *requires* `w:tcBorders` on each cell. At cell level
`insideH` and `insideV` no longer mean what their names suggest — they apply to a cell's
own internal splits, not to the lines between rows. Four cell edges is therefore the
honest model.

Consequence: adjacent cells each own their edges, so body `LEFT` + `RIGHT` puts two
borders between neighbouring columns. Word's conflict resolution collapses them
visually and the thicker wins, so mixing widths across edges can look uneven. That is
inherent to cell-level borders, not a defect.

A true outer-box-only border is expressible (see the table above) but fiddly. A
table-level `outerBorder` using `w:tblBorders` would do it cleanly and is deliberately
out of scope until asked for.

### `BorderLine`

`STBorder` has roughly 180 values, the large majority of them decorative page-border art
(`GINGERBREAD_MAN`, `SCARED_CAT`, `PUMPKIN_1`). Exposing that set for a table border
would be absurd, so `BorderLine` is a curated enum — `SINGLE`, `THICK`, `DOUBLE`,
`DOTTED`, `DASHED`, `DOT_DASH` — mapped to `STBorder`, the same approach `Alignment`
takes to `JcEnumeration`. Default is `SINGLE`.

There is deliberately no `NONE` member: absence of a border is already expressed two
ways (`none()`, or no edges), and a third would be one too many.

## Units

Border width is in **eighths of a point**: `w:sz="8"` is 1 pt. This is the project's
third unit, after twips for geometry and half-points for font size. The conversion
belongs in `Units` beside the others — `pointsToEighths(double)` — so it is never
invented inline. A 0.5 pt hairline is `sz="4"`.

## Four traps

1. **`w:tblGrid` is mandatory, with fixed layout.** Without both, Word auto-fits to
   content and the same table renders differently in different clients.
   `CTTblLayoutType` set to `fixed` makes Word honour the grid.
2. **An empty `w:tc` makes the document unopenable.** Not misrendered — unopenable.
   Every cell gets a paragraph even when its text is blank. This is the single most
   important rule here, and the blank-cell test exists to hold it.
3. **Border width is in eighths of a point**, per Units above.
4. **A table is followed by an empty paragraph.** Two adjacent tables with nothing
   between them are merged by Word into one, and a body ending in a table rather than a
   paragraph is irregular. Emitting a paragraph after every table avoids both. This
   follows established OOXML convention rather than an observation of Word's behaviour,
   which cannot be checked in this environment; it costs one empty paragraph and removes
   a class of bug.

## Page spanning

Header rows carry `w:tblHeader`, so the header repeats at the top of every page a table
spans. This matters now that documents run to several pages.

## Column widths

Columns share the usable page width equally: `widthTwips / columnCount`, with any
remainder from integer division added to the last column so the widths sum exactly to
the table width. Default width is the full usable page width; an overload takes an
explicit width.

## Cell content

Header and body cells each take a `TextStyle` from `TableStyle`. Cell paragraphs use a
fixed compact `ParagraphStyle` — `spaceBeforeTwips 0`, `spaceAfterTwips 0`, no
`keepWithNext` — because the spacing designed for body prose leaves visible gaps inside
a cell.

Cell padding is `w:tblCellMar` on the table, fixed at **80 twips** (about 0.14 cm) left
and right, 0 top and bottom. Not configurable: no requirement calls for varying it, and
it can be promoted to `TableStyle` later without breaking callers.

## Error handling

Unchanged contract: `DocumentGenerationException` only, with a message naming the field
and no cause for validation failures.

Rejected: a null or empty header list together with no rows; any row whose length
differs from the header count, naming the offending row index and both lengths; a
non-positive width; a non-positive border width; a malformed hex colour; a null
`TableStyle`.

Null cell text is treated as empty, not rejected — a null in one cell of a data set
should not fail the whole document. Empty cells still get their paragraph.

## Testing

The round trip is the test that matters: build a table, reload the bytes, and assert on
the reloaded `Tbl`.

- Grid: `w:tblGrid` has one `gridCol` per column, and the widths sum to the requested
  table width exactly.
- **Every cell contains at least one paragraph**, including cells given `""` and `null`.
  This is the unopenable case.
- Header cells carry the header border colour and only its chosen edges; body cells
  carry the body colour and its edges. Asserting both distinguishes the two groups,
  which a single-border implementation would fail.
- `w:tblHeader` is present on the header row and absent on body rows.
- Border width conversion: 1.0 pt emits `sz="8"`, 0.5 pt emits `sz="4"`.
- `BorderLine.DOTTED` emits `w:val="dotted"`.
- `TableBorderStyle.none()` and an empty edge set emit no `w:tcBorders` at all.
- A row of the wrong length is rejected, and the message names the row index.
- Mixed content — paragraph, table, paragraph — reloads in insertion order, proving the
  widened `DocumentContent` did not break ordering.

## Out of scope

Cell shading and fills; merged cells, whether horizontal or vertical; per-cell style
overrides; explicit per-column widths; a table-level `outerBorder`; nested tables;
column alignment beyond what `TextStyle` and the cell's `ParagraphStyle` already give;
table captions; and `w:cantSplit` to keep a row off a page boundary.
