# Bulk DOCX Updater

Command-line tool to bulk update DOCX files.

## Usage

Basic form

- `python main.py PATH [options]`

Options

- `-c, --config PATH`  JSON config file (dict format)
- `-s, --search TEXT`  Search text
- `-r, --replace TEXT` Replacement text
- `--recursive`  Recurse into subdirectories
- `--pattern GLOB`  File pattern (default: `*.docx`)
- `-j, --workers N`  Number of parallel workers (default: 1)
- `--dry-run`  Show diffs without modifying files
- `--xml-diff`  Include XML-level diffs in dry-run output
- `--diff-context N`  Unified diff context lines (default: 3)
- `--verbose`  Enable verbose logging

Examples

- Config file: `python src/main.py ./docs --config replace.json`
- Single replace: `python src/main.py ./docs --search "old" --replace "new"`
- Dry run: `python src/main.py ./docs --config replace.json --dry-run`
- Recursive: `python src/main.py ./docs --config replace.json --recursive`
- Parallel: `python src/main.py ./docs --config replace.json -j 4`

## Config Format

JSON object with operation names as keys:

```json
{
  "replace": [
    ["old text", "new text"],
    ["another", "replacement"]
  ],
  "set_comments": "Template: {{FILENAME}}",
  "clear_properties": ["author", "company", "title"]
}
```

### Settings

Optional settings keys in the config object:

```json
{
  "preserve_formatting": false,
  "margins": "1,1,1.5,1.5"
}
```

- `preserve_formatting` (bool): Preserve existing run formatting during replace (default: `true`)
- `margins`: Set page margins in inches. String `"top,bottom,left,right"` or dict `{"top": 1.0, "bottom": 1.0, "left": 1.5, "right": 1.5}`

### Operations

**Replace text**

```json
{ "replace": [["Old Text", "New Text"], ["foo", "bar"]] }
```

Each entry is `[search, replace]` or `[search, replace, {options}]`. Options dict supports `regex: true`.

Dict form also accepted per-entry:

```json
{ "replace": [{"search": "old", "replace": "new", "regex": true}] }
```

**Limiting replacements:**

- `count`: Maximum number of replacements (default: 0 = unlimited)
- `occurrence`: Target a specific match (1-based). `1` = first match only, `2` = second only, etc.

```json
{ "replace": [{"search": "Address", "replace": "123 Main St", "occurrence": 1}] }
{ "replace": [{"search": "Address", "replace": "City, ST 12345", "occurrence": 2}] }
```

**Replace XML**
```json
{ "xml_replace": [{"search": "<w:t>old</w:t>", "replace": "<w:t>new</w:t>"}] }
{ "xml_replace": [{"search_file": "search.xml", "replace_file": "replace.xml"}] }
```

**Table operations**
```json
{ "replace_table_cell": {"table_header": "Phase, Time", "row": 0, "column": 1, "replace": "Time"} }
{ "replace_table_cell": {"table_header": "Col1, Col2", "header_row": 1, "row": 0, "column": 0, "replace": "New Title"} }
{ "replace_table_cell": {"table_index": 2, "row": 0, "column": 0, "replace": "New Title"} }
{ "set_table_column_widths": {"table_header": "Phase, Time", "column_widths": [1.5, 2.0]} }
{ "table_header_repeat": true }
{ "table_header_repeat": "Phase, Time" }
{ "align_table_cells": {"patterns": ["text1", "text2"], "alignment": "left"} }
```

Use a list value to apply multiple instances of `replace_table_cell`, `set_table_column_widths`, `align_table_cells`, `replace_in_table`, or `replace_table`.

**Replace whole table**

Swaps an entire `<w:tbl>` element for new table XML. Unlike `replace_table_cell` (which only edits cell text), the replacement may have a completely different shape, orientation, or docxtpl loop tags — useful for turning a fixed-column table into a dynamic `{%tr for ... %}` one.

```json
{ "replace_table": {"match": "reading1_co", "replace_file": "mdc_table.xml"} }
{ "replace_table": {"table_index": 22, "replace_file": "mdc_table.xml"} }
{ "replace_table": {"table_header": "Phase, Time", "replace": "<w:tbl>...</w:tbl>"} }
```

Locate the target table with one of:

- `match`: substring found anywhere in the table's text (most robust across templates whose XML differs by rsid/paraId)
- `table_index`: 0-based table index
- `table_header`: header-row text match (`header_row` selects the header row, default 0)

Provide the replacement table as inline `replace` XML or via `replace_file` (resolved relative to the config file). The XML root must be a `<w:tbl>`. A table copied straight out of Word already carries its own `xmlns` declarations; for hand-written XML that uses only `w:`/`w14:`/`mc:`/etc. prefixes, the standard declarations are injected automatically.

`align_table_cells` aligns table cells containing specific text patterns. Supported alignments: `left`, `center`, `right`, `justify` (defaults to `left`).

**Merge repeated tables into one**

Documents rendered from split templates often repeat the same table (identical title + header block) once per page. `merge_tables` folds every matching table into the first, appending each continuation's data rows and dropping the duplicated leading header rows, so the result is a single continuous table with no repeated rows. The emptied continuation tables and the blank (page-break) paragraphs that separated them are removed.

```json
{ "merge_tables": {"match": "NMNEHC Test Results"} }
{ "merge_tables": {"table_header": "NMNEHC Test Results", "skip_rows": 11} }
```

Locate the tables with `match` (substring) or `table_header` (header-row text match, with `header_row` for the header row index) — the same matching as `replace_table`, but `table_index` isn't accepted since merging needs two or more tables. By default the duplicated leading rows are auto-detected (the identical prefix each continuation shares with the first table, compared with whitespace normalized so stray spacer/non-breaking-space differences don't defeat detection); set `skip_rows` to drop a fixed number instead. Re-running is idempotent — once merged, only one table matches, so nothing changes.

**Insert a new block (paragraphs + tables)**

Inserts brand-new body-level content at an anchor paragraph located by text.
Unlike `replace_table` (which swaps an *existing* `<w:tbl>`), this *adds* content,
so it can introduce a section that didn't exist before — e.g. a new raw-data
appendix.

```json
{ "insert_block": {"before": "SITE PHOTOS", "replace_file": "ecom_rawdata_table.xml", "skip_if_present": "ANALYZER RAW DATA"} }
{ "insert_block": {"after": "CALIBRATION CERTIFICATES", "replace": "<block><w:p>...</w:p><w:tbl>...</w:tbl></block>"} }
{ "insert_block": [ {"before": "SITE PHOTOS", "replace_file": "a.xml"}, {"after": "NOTES", "replace_file": "b.xml"} ] }
```

- `before` / `after` (exactly one): text of the anchor paragraph. Exact (stripped)
  match preferred, falls back to the first paragraph containing the text. The block
  is inserted immediately before / after that paragraph.
- `replace` / `replace_file`: the XML to insert. Several top-level elements
  (paragraphs, tables) must be wrapped in a single root element (e.g.
  `<block> ... </block>`); the root's children are inserted in order and the
  wrapper is discarded. Standard Word namespace prefixes are injected if the root
  doesn't declare them.
- `skip_if_present` (optional): if this text already appears anywhere in the
  document body, the insert is skipped — making re-runs idempotent.

Runs before `landscape_table`, so a freshly-inserted table can be located and
rotated in the same config.

**Remove a page break from a paragraph**

Strips every `<w:br w:type="page"/>` from the paragraph located by text (and drops
the run if that leaves it empty). Operates on the element tree, so it's robust to
XML whitespace/serialization — unlike a literal `xml_replace`. A
`<w:lastRenderedPageBreak/>` render hint is left untouched.

```json
{ "remove_page_break": "{% for img in cylinder_certs %}" }
{ "remove_page_break": {"in_paragraph": "CALIBRATION CERTIFICATES"} }
{ "remove_page_break": [ {"in_paragraph": "foo"}, {"in_paragraph": "bar"} ] }
```

- `in_paragraph`: text identifying the paragraph (exact stripped match preferred,
  falls back to the first paragraph containing the text). The string form is
  shorthand for `{"in_paragraph": ...}`.

Runs after `insert_block`, so a freshly-inserted section can also be targeted.

**Landscape table**

Wraps a located table in its own landscape section, leaving the surrounding content in its original orientation. Useful for wide tables (many columns) that overflow a portrait page and wrap unreadably. Inserts a section break before the table and a landscape section break after it.

```json
{ "landscape_table": {"match": "for ftir in run1"} }
{ "landscape_table": {"table_header": "Spectrum, Time, Phase", "margins": "0.5,0.5,0.5,0.5"} }
{ "landscape_table": [ {"match": "run1"}, {"match": "run2"} ] }
```

Locate the table the same way as `replace_table` (`match` / `table_index` / `table_header`, with `header_row` for header matching). Optional `margins` sets the landscape section's margins as `"top,bottom,left,right"` inches or a dict; defaults to `0.5` all round. Runs after `replace_table`, so a freshly-swapped table can be located and rotated in the same config. Re-running is idempotent — a table already in a landscape section is left untouched. Only tables that are direct children of the document body are supported (not nested tables).

**Section break before a heading**

Makes a matched paragraph start its own page/section by moving the section break that currently *follows* it to immediately *before* it. Fixes templates where a heading is stranded at the tail of the previous section — e.g. an `O2 RAW DATA` heading left inside the landscape FTIR rawdata section, so it renders at the end of those pages instead of heading its own page. The relocated break keeps its orientation, so the preceding content stays as-is and the heading begins a new page in the next section's orientation.

```json
{ "section_break_before": {"match": "O2 RAW DATA"} }
{ "section_break_before": [ {"match": "O2 RAW DATA"}, {"match": "SITE PHOTOS"} ] }
```

`match` is the paragraph text (exact match preferred, falls back to substring — so an appendix-list entry containing the same words isn't mistaken for the heading). Idempotent: a paragraph already preceded by a section break is left untouched. If no section break follows the paragraph, it's a no-op.

**Image replacement**
```json
{ "replace_image": {"image_path": "path/to/logo.png"} }
{ "replace_image": {"image_path": "logo.png", "scale": 0.5, "center": true} }
```

Replaces first image, maintaining aspect ratio. Optional: `scale` (0.5 = 50%), `center` (true/false). Advanced: `name`/`alt_text`/`index` to target specific images.

**Set comments**
```json
{ "set_comments": "{{FILENAME}}" }
```

Sets Comments field (File → Info → Properties). Placeholders: `{{FILENAME}}`, `{{BASENAME}}`, `{{EXTENSION}}`, `{{PARENT_DIR}}`.

**Clear properties**
```json
{ "clear_properties": ["author", "company"] }
{ "clear_properties": true }
```

Clears document properties. Use `true` to clear all common properties.

**Supported:** `title`, `subject`, `author`, `keywords`, `comments`, `last_modified_by`, `category`, `content_status`, `company`

**Other**
```json
{ "cleanup_empty_after": "HEADER" }
{ "cleanup_empty_after": ["HEADER1", "HEADER2"] }
{ "font_size": {"from": 8, "to": 10} }
```

## Formatting

Use `{format:options}text{/format}` for inline formatting:

```json
{ "replace": [["{format:bold,center,size16}New Title{/format}", "Title"]] }
```

**Options:** `bold`, `italic`, `underline`, `left`, `center`, `right`, `justify`, `size8`-`size24`, `spacing0`-`spacing24`

**Tokens:** `pagebreak`, `linebreak`, `paragraphbreak`

```json
{ "replace": [["PHOTOS", "Photo1paragraphbreakPhoto2paragraphbreakPhoto3"]] }
```

## Example recipes

Real configs used to retrofit report templates, kept here as reference. Each is a
standalone dict config you pass with `-c`. Recipes that swap or insert tables
reference an external XML fragment (a `<w:tbl>`, or for `insert_block` a `<block>`
wrapper of paragraphs + tables, exported from Word) — see `replace_table` /
`insert_block` above for the fragment format. Locate-by-`match` uses a unique
substring of the table's text (often its docxtpl loop tag), which survives
template edits that change rsid/paraId.

**Swap an FTIR raw-data table for a dynamic landscape table**

Replaces a fixed table with a docxtpl `{%tr for ... %}` table, rotates just that
table's section to landscape, then moves the trailing `O2 RAW DATA` heading onto
its own vertically-centered page.

```json
{
  "replace_table": { "match": "for ftir in ftir_rawdata", "replace_file": "ftir_rawdata_table.xml" },
  "landscape_table": { "match": "for ftir in ftir_rawdata", "margins": "0.5,0.5,0.5,0.5" },
  "section_break_before": { "match": "O2 RAW DATA" },
  "divider": { "match": "O2 RAW DATA" }
}
```

Per-run variant — the same idea applied to a calibration plus three runs via list
values:

```json
{
  "replace_table": [
    { "match": "for ftir in calibration", "replace_file": "ftir_rawdata_table_calibration.xml" },
    { "match": "for ftir in run1", "replace_file": "ftir_rawdata_table_run1.xml" },
    { "match": "for ftir in run2", "replace_file": "ftir_rawdata_table_run2.xml" },
    { "match": "for ftir in run3", "replace_file": "ftir_rawdata_table_run3.xml" }
  ],
  "landscape_table": [
    { "match": "for ftir in calibration", "margins": "0.5,0.5,0.5,0.5" },
    { "match": "for ftir in run1", "margins": "0.5,0.5,0.5,0.5" },
    { "match": "for ftir in run2", "margins": "0.5,0.5,0.5,0.5" },
    { "match": "for ftir in run3", "margins": "0.5,0.5,0.5,0.5" }
  ],
  "section_break_before": { "match": "O2 RAW DATA" },
  "divider": { "match": "O2 RAW DATA" }
}
```

**Insert an analyzer raw-data appendix**

Adds a new section ahead of `SITE PHOTOS` (idempotent via `skip_if_present`) and
strips a stray page break from a docxtpl loop paragraph.

```json
{
  "insert_block": { "before": "SITE PHOTOS", "replace_file": "ecom_rawdata_table.xml", "skip_if_present": "ANALYZER RAW DATA" },
  "remove_page_break": { "in_paragraph": "{% for img in cylinder_certs %}" }
}
```

**Swap several result tables in one pass**

Locates each table by a unique substring of its text and replaces the whole
`<w:tbl>`:

```json
{
  "replace_table": [
    { "match": "Oil as Octane (200) 191C", "replace_file": "voc_raw_table.xml" },
    { "match": "Acetaldehyde (C2H4O) Emission Results", "replace_file": "voc_nmnehc_acetaldehyde.xml" },
    { "match": "Ethylene (C2H4) Emission Results", "replace_file": "voc_nmnehc_ethylene.xml" },
    { "match": "Total VOC (as C3H8)", "replace_file": "voc_total_table.xml" }
  ]
}
```

**Tighten small in-template tables**

Left-justifies every cell and sets tight (28-twip) cell margins on the O2 / THC /
fuel raw-data tables, located by their docxtpl loop tags:

```json
{
  "format_table": [
    { "match": "for o2 in o2_rawdata", "align": "left", "cell_margins": "28" },
    { "match": "for thc in thc_rawdata", "align": "left", "cell_margins": "28" },
    { "match": "for fuel in run1", "align": "left", "cell_margins": "28" }
  ]
}
```

**Rewrite a docxtpl expression**

Plain text replacements can edit docxtpl expressions in place — e.g. dropping a
`.strftime(...)` call so the template emits the raw timestamp:

```json
{
  "replace": [
    ["{{ o2.ReadingTimestamp.strftime('%H:%M:%S') }}", "{{ o2.ReadingTimestamp }}"],
    ["{{ ftir.time_fmtd().split(' ')[-1] }}", "{{ ftir.time_fmtd() }}"]
  ]
}
```
