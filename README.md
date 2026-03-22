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

Use a list value to apply multiple instances of `replace_table_cell`, `set_table_column_widths`, `align_table_cells`, or `replace_in_table`.

`align_table_cells` aligns table cells containing specific text patterns. Supported alignments: `left`, `center`, `right`, `justify` (defaults to `left`).

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
